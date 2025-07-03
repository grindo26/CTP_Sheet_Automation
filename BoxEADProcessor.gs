// ===============================
// OAuth2 & Box API Setup
// ===============================

/**
 * Keys for retrieving secrets from Script Properties.
 */
var SCRIPT_PROP_KEY_CLIENT_ID = "BOX_CLIENT_ID";
var SCRIPT_PROP_KEY_CLIENT_SECRET = "BOX_CLIENT_SECRET";

/**
 * Returns an OAuth2 service for the Box API.
 */
function getOAuthService() {
    return (
        OAuth2.createService("box")
            .setAuthorizationBaseUrl("https://account.box.com/api/oauth2/authorize")
            .setTokenUrl("https://api.box.com/oauth2/token")
            // Load credentials from Script Properties:
            .setClientId(PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEY_CLIENT_ID))
            .setClientSecret(PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_KEY_CLIENT_SECRET))
            .setCallbackFunction("authCallback")
            .setPropertyStore(PropertiesService.getUserProperties())
            .setScope("root_readwrite")
    );
}

/**
 * Callback for OAuth2 authorization.
 */
function authCallback(request) {
    var boxService = getOAuthService();
    if (boxService.handleCallback(request)) {
        return HtmlService.createHtmlOutput("Success! You can close this tab.");
    } else {
        return HtmlService.createHtmlOutput("Authorization denied. You can close this tab.");
    }
}

/**
 * Retrieves a valid Box access token.
 * If authorization has not been granted, logs the authorization URL.
 */
function getBoxAccessToken() {
    var service = getOAuthService();
    if (service.hasAccess()) {
        return service.getAccessToken();
    } else {
        var authorizationUrl = service.getAuthorizationUrl();
        Logger.log("Authorization is required. Open the following URL and re-run the script: %s", authorizationUrl);
        throw new Error("Authorization required.");
    }
}

// ===============================
// Sheet & Batch Settings
// ===============================

// === Sheet columns (1-indexed) ===
var COL_FIRST_NAME = 2; // B
var COL_MIDDLE_NAME = 3; // C
var COL_LAST_NAME = 4; // D
var COL_NICK_NAME = 6; // F
var COL_EMAIL = 5;
var COL_EAD_URL1 = 10; // J
var COL_EAD_URL2 = 11; // K
var COL_FOLDER_NAME = 12; // L
var COL_SUBMISSION_DT = 13; // M
var COL_FOLDER_LINK = 14; // N — where we'll write the Box folder URL

// === Batch processing settings ===
var BATCH_SIZE = 20; // rows per execution
var PROPERTY_LAST_ROW = "BOX_LAST_PROCESSED_ROW"; // key in script properties

// ===============================
// Box Folder & File Operations
// ===============================

/**
 * Checks whether a given Box folder already contains any files.
 */
function folderHasFiles(folderId) {
    var accessToken = getBoxAccessToken();
    var folderUrl = "https://api.box.com/2.0/folders/" + folderId + "/items";
    var options = {
        method: "GET",
        headers: {
            Authorization: "Bearer " + accessToken,
            "Content-Type": "application/json",
        },
        muteHttpExceptions: true,
    };

    try {
        var response = UrlFetchApp.fetch(folderUrl, options);
        var data = JSON.parse(response.getContentText());
        if (data.entries && data.entries.length) {
            return data.entries.some(function (ent) {
                return ent.type === "file";
            });
        }
    } catch (e) {
        Logger.log("Error checking folder contents: " + e);
    }
    return false;
}

/**
 * Creates or retrieves an existing Box folder under the specified parent folder.
 */
function createOrGetBoxFolder(parentFolderId, folderName) {
    var accessToken = getBoxAccessToken();
    var folderUrl = "https://api.box.com/2.0/folders/" + parentFolderId + "/items";
    var options = {
        method: "GET",
        headers: {
            Authorization: "Bearer " + accessToken,
            "Content-Type": "application/json",
        },
        muteHttpExceptions: true,
    };

    try {
        var response = UrlFetchApp.fetch(folderUrl, options);
        var resp = JSON.parse(response.getContentText());
        var items = Array.isArray(resp.entries) ? resp.entries : [];

        // If folder exists, return it
        for (var i = 0; i < items.length; i++) {
            if (items[i].type === "folder" && items[i].name === folderName) {
                return items[i];
            }
        }

        // Otherwise create
        var payload = {
            name: folderName,
            parent: { id: parentFolderId },
        };
        var createOptions = {
            method: "POST",
            headers: {
                Authorization: "Bearer " + accessToken,
                "Content-Type": "application/json",
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true,
        };
        var createResp = UrlFetchApp.fetch("https://api.box.com/2.0/folders", createOptions);
        var folder = JSON.parse(createResp.getContentText());
        return folder.id ? folder : null;
    } catch (e) {
        Logger.log("Error in createOrGetBoxFolder: " + e);
        return null;
    }
}

/**
 * Downloads a file from URL and uploads it to Box under the given folder.
 */
function downloadAndUploadFileToBox(fileUrl, folderId, fileName) {
    var accessToken = getBoxAccessToken();
    try {
        var resp = UrlFetchApp.fetch(fileUrl, { muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) {
            Logger.log("Download failed (" + resp.getResponseCode() + "): " + fileUrl);
            return;
        }
        var blob = resp.getBlob().setName(fileName);

        var attributes = JSON.stringify({
            name: fileName,
            parent: { id: folderId },
        });
        var formData = {
            attributes: Utilities.newBlob(attributes, "application/json"),
            file: blob,
        };
        var uploadOptions = {
            method: "POST",
            headers: { Authorization: "Bearer " + accessToken },
            payload: formData,
            muteHttpExceptions: true,
        };
        var up = UrlFetchApp.fetch("https://upload.box.com/api/2.0/files/content", uploadOptions);
        Logger.log("Upload response: " + up.getResponseCode());
    } catch (e) {
        Logger.log("Error in downloadAndUploadFileToBox: " + e);
    }
}

// ===============================
// EAD Card Processing
// ===============================

/**
 * Trigger on edits/inserts in the "EAD card" sheet.
 */
function onChangeEAD(e) {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== "EAD card") return;
    if (e.changeType !== "EDIT" && e.changeType !== "INSERT_ROW") return;
    var r = sheet.getActiveRange().getRow();
    if (r === 1) return; // skip header
    processRowEAD(sheet, r);
}

/**
 * Process a single row: create/get Box folder, write its link, and upload EAD files.
 */
function processRowEAD(sheet, row) {
    try {
        var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        var firstName = data[COL_FIRST_NAME - 1];
        var middleName = data[COL_MIDDLE_NAME - 1];
        var lastName = data[COL_LAST_NAME - 1];
        var email = data[COL_EMAIL - 1];
        var nickName = data[COL_NICK_NAME - 1];
        var folderName = data[COL_FOLDER_NAME - 1];
        var eadUrl1 = getPlainUrl(sheet, row, COL_EAD_URL1);
        var eadUrl2 = getPlainUrl(sheet, row, COL_EAD_URL2);

        if (!firstName || !lastName || !email || !eadUrl1) {
            Logger.log("Row %s missing key data, skipping.", row);
            return;
        }

        // 1) Under your top‐level “root” folder, find / create the A-Z bucket
        var ROOT_PARENT_ID = "315593622123";
        var firstLetter = firstName.charAt(0).toUpperCase();
        var letterFolder = createOrGetBoxFolder(ROOT_PARENT_ID, firstLetter);
        if (!letterFolder || !letterFolder.id) {
            Logger.log("Couldn’t get or create letter bucket for %s", firstLetter);
            return;
        }

        // 2) Inside that, find / create the person’s folder
        var personFolderName = firstName + " " + lastName + " " + email;
        var personFolder = createOrGetBoxFolder(letterFolder.id, personFolderName);
        if (!personFolder || !personFolder.id) {
            Logger.log("Couldn’t get/create person folder %s", personFolderName);
            return;
        }

        // 3) Write the folder link back to the sheet
        var folderLink = "https://app.box.com/folder/" + personFolder.id;
        sheet.getRange(row, COL_FOLDER_LINK).setValue(folderLink);

        // 4) If it’s empty, upload the EAD files into that person‐folder
        if (!folderHasFiles(personFolder.id)) {
            // build filenames
            var fullName = firstName + (nickName ? " (" + nickName + ")" : "") + (middleName ? " " + middleName : "") + " " + lastName;
            var ext1 = getFileExtension(eadUrl1);
            var fileName1 = fullName + " - EAD Card 1" + ext1;
            downloadAndUploadFileToBox(eadUrl1, personFolder.id, fileName1);

            if (eadUrl2) {
                var ext2 = getFileExtension(eadUrl2);
                var fileName2 = fullName + " - EAD Card 2" + ext2;
                downloadAndUploadFileToBox(eadUrl2, personFolder.id, fileName2);
            }
        } else {
            Logger.log("Person folder %s already has files – skipping upload.", personFolderName);
        }
    } catch (e) {
        Logger.log("Error processing row %s: %s", row, e);
    }
}

/**
 * Extracts a plain URL from a cell, handling HYPERLINK formulas.
 */
function getPlainUrl(sheet, row, col) {
    var cell = sheet.getRange(row, col);
    var formula = cell.getFormula();
    if (formula && formula.indexOf("HYPERLINK") !== -1) {
        var m = formula.match(/"(https?:\/\/[^"]+)"/);
        return m ? m[1] : null;
    } else {
        return cell.getValue();
    }
}

/**
 * Returns the file extension from a URL.
 */
function getFileExtension(url) {
    var m = url.match(/\.([^.\/?]+)(?:[?#]|$)/);
    return m ? "." + m[1] : "";
}

// ===============================
// Batch Processing
// ===============================

/**
 * Processes rows in batches of BATCH_SIZE, then re-schedules itself until done.
 */
function processAllRowsBatch() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("EAD card");
    if (!sheet) throw new Error("Sheet 'EAD card' not found!");

    var props = PropertiesService.getScriptProperties();
    var startRow = parseInt(props.getProperty(PROPERTY_LAST_ROW), 10) || 2;
    var lastRow = sheet.getLastRow();
    if (startRow > lastRow) {
        props.deleteProperty(PROPERTY_LAST_ROW);
        return;
    }

    var endRow = Math.min(startRow + BATCH_SIZE - 1, lastRow);
    for (var r = startRow; r <= endRow; r++) {
        processRowEAD(sheet, r);
    }

    props.setProperty(PROPERTY_LAST_ROW, (endRow + 1).toString());

    if (endRow < lastRow) {
        // clear old trigger
        ScriptApp.getProjectTriggers()
            .filter((t) => t.getHandlerFunction() === "processAllRowsBatch")
            .forEach(ScriptApp.deleteTrigger);
        // schedule next batch in 1 minute
        ScriptApp.newTrigger("processAllRowsBatch")
            .timeBased()
            .after(1 * 60 * 1000)
            .create();
    }
}

/**
 * Clears any previous batch progress and kicks off batch processing immediately.
 */
function startBatchProcessing() {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty(PROPERTY_LAST_ROW);

    // delete old triggers
    ScriptApp.getProjectTriggers()
        .filter((t) => t.getHandlerFunction() === "processAllRowsBatch")
        .forEach(ScriptApp.deleteTrigger);

    // start first batch
    processAllRowsBatch();
}

// ===============================
// Testing Helpers
// ===============================

/**
 * Manually test a single row (row 2 by default).
 */
function testProcessRowEAD() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("EAD card");
    processRowEAD(sheet, 2);
}

/**
 * Process all rows at once (not recommended for large sheets).
 */
function processAllRows() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("EAD card");
    var lastRow = sheet.getLastRow();
    var emptyCount = 0;

    for (var r = 2; r <= lastRow; r++) {
        var folderName = sheet.getRange(r, COL_FOLDER_NAME).getValue();
        var eadUrl1 = getPlainUrl(sheet, r, COL_EAD_URL1);

        // If both folderName and EAD URL1 are blank, count as empty
        if (!folderName && !eadUrl1) {
            emptyCount++;
            if (emptyCount >= 2) break; // exit after 2 in a row
            continue;
        }

        emptyCount = 0; // reset when we hit data
        processRowEAD(sheet, r);
    }
}

/**
 * Recursively collects every Box folder under the given parent.
 *
 * @param {string} parentId  Box folder ID to start from (use "0" for root).
 * @return {Array<Object>}   Array of folder objects with { id, name }.
 */
function getAllBoxFolders(parentId) {
    var accessToken = getBoxAccessToken();
    var folders = [];
    var offset = 0;
    var limit = 1000; // max per request

    do {
        var url = ["https://api.box.com/2.0/folders/", parentId, "/items?limit=", limit, "&offset=", offset, "&fields=type,id,name,total_count"].join(
            ""
        );
        var resp = UrlFetchApp.fetch(url, {
            method: "GET",
            headers: { Authorization: "Bearer " + accessToken },
            muteHttpExceptions: true,
        });
        var js = JSON.parse(resp.getContentText());

        js.entries.forEach(function (ent) {
            if (ent.type === "folder") {
                folders.push({ id: ent.id, name: ent.name });
                folders = folders.concat(getAllBoxFolders(ent.id));
            }
        });

        offset += js.entries.length;
    } while (offset < js.total_count);

    return folders;
}

/**
 * Dumps every folder’s name and link into "Sheet2".
 */
function exportFoldersToSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet2");
    if (!sheet) {
        sheet = ss.insertSheet("Sheet2");
    } else {
        sheet.clearContents();
    }

    // header row
    sheet.appendRow(["Folder Name", "Folder Link"]);

    // collect all folders under root (id="0")
    var allFolders = getAllBoxFolders("0");

    allFolders.forEach(function (folder) {
        var link = "https://app.box.com/folder/" + folder.id;
        sheet.appendRow([folder.name, link]);
    });
}
