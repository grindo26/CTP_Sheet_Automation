# Box ↔ Google Sheets Automations

This repository contains two Google Apps Script automations that integrate Box.com and Google Sheets:

1. **Folder List Extractor**  
   Fetches all sub-folders under a given Box folder, writes their names & URLs into a Sheet, and lets you refresh with one click.

2. **EAD Card Uploader**  
   Watches your “EAD card” sheet for new entries, creates candidate folders in Box (bucketed by first-name initial → “First Last email”), and uploads EAD-card files from sheet URLs.

---

## Table of Contents

-   [Features](#features)
-   [Prerequisites](#prerequisites)
-   [Setup](#setup)
-   [Configuration](#configuration)
-   [Usage](#usage)
    -   [Folder List Extractor](#folder-list-extractor)
    -   [EAD Card Uploader](#ead-card-uploader)
-   [Triggers & Buttons](#triggers--buttons)
-   [Error Handling & Limits](#error-handling--limits)
-   [Next Steps](#next-steps)
-   [Contact](#contact)

---

## Features

### Folder List Extractor

-   Pagination-safe retrieval of up to 1,000 Box sub-folders per request
-   Writes **folder name** & **clickable URL** to a Google Sheet
-   One-click **“Refresh”** via custom menu or on-sheet button

### EAD Card Uploader

-   On-edit / on-insert trigger for your **EAD card** sheet
-   Two-level Box folder hierarchy:
    1. **Letter bucket** (A, B, C…) under your root folder
    2. **Candidate folder** named `First Last email`
-   Skips upload if the candidate folder already contains files
-   Uploads **EAD Card 1 & 2** from sheet URLs, renaming with full name + card number

---

## Prerequisites

-   A **Box.com** developer app with OAuth2 (Client ID & Secret)
-   A **Google Workspace** account with access to Google Sheets & Apps Script
-   **OAuth2 for Apps Script** library added to your project:
    1. In Apps Script: **Libraries** → add ID `1BPDX…Oxjo`
    2. Set identifier to `OAuth2` and **Add**

---

## Setup

1. **Copy** the `.gs` files into your Apps Script project
2. In **Project Settings → Script properties**, add:
    - `BOX_CLIENT_ID` → your Box OAuth2 Client ID
    - `BOX_CLIENT_SECRET` → your Box OAuth2 Client Secret
3. **Enable** the Apps Script **Sheets**, **Drive**, and **URL Fetch** APIs if prompted

---

## Configuration

Edit the constants at the top of each script:

```js
// Root Box folder ID under which everything lives:
var ROOT_PARENT_ID = "315593622123";

// Folder List Extractor:
var FOLDER_LIST_SHEET_ID = "<YOUR_SHEET_ID>";
var FOLDER_LIST_SHEET_NAME = "Sheet1";

// EAD Card Uploader:
var EAD_SHEET_NAME = "EAD card";
var COL_FIRST_NAME = 2; // B
var COL_EMAIL = 5; // E
var COL_EAD_URL1 = 10; // J
// …etc…
```

---

## Usage

### Folder List Extractor

1. Open Apps Script editor
2. Select function `extractAFolder` → ▶️ **Run**
3. Check your sheet for updated folder list

### EAD Card Uploader

-   Automatically triggers on row edits/inserts in **EAD card**
-   To test manually:

    1. Open Apps Script editor
    2. Run `testProcessRowEAD` (processes row 2)

---

## Triggers & Buttons

-   **Sheet Button**: Insert a drawing/image → right-click → **Assign script** → enter `extractAFolder` or `startBatchProcessing`
-   **Batch Continuation**: EAD uploader schedules next chunk with `ScriptApp.newTrigger('processAllRowsBatch')…`

---

## Error Handling & Limits

-   **6-minute execution limit** in Apps Script

    -   Workaround: batch processing or offload to Cloud Functions

-   Extensive `Logger.log()` for debugging
-   OAuth2 token refresh handled by the `OAuth2` library

---

## Next Steps

-   Offload large-scale folder retrieval to Google Cloud Functions + Sheets API
-   Add retry/back-off logic for Box API rate limits
-   Modularize scripts into separate files (`oauth.gs`, `extract.gs`, `upload.gs`) as needed

---

## Contact

Prepared by **Pratik Sangle**
