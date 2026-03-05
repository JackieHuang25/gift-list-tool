# Gift List Tool

## Overview
This project provides a simple web app for customers to review and update their shipping address, backed by an Excel workbook stored in OneDrive/SharePoint. A token generator script writes tokenized links into the workbook and outputs a JSON mapping for internal use.

## Folder Structure
```
.
├── data/
│   ├── input/
│   │   ├── gift-list.csv
│   │   └── gift-list.xlsx
│   └── output/
│       └── gift-list-token.json
├── pages/
│   ├── api/
│   │   ├── get-address.js
│   │   └── update-address.js
│   └── index.js
├── public/
├── scripts/
│   └── generate-tokens.js
├── package.json
└── .env.local
```

## Data Flow
1. Run the token generator in `scripts/generate-tokens.js`.
2. It downloads the workbook using Microsoft Graph, appends a unique token link per row, and uploads the workbook back.
3. The script writes `data/output/gift-list-token.json` for internal reference.
4. Customers open a token link that loads the Next.js page in `pages/index.js`.
5. The frontend calls the API routes to read or update a row in the workbook based on the token.

## Logic Details

### Token Generation
- File: `scripts/generate-tokens.js`
- Reads Graph credentials and `SHARE_LINK` from `.env.local`.
- Downloads the workbook, generates a token and link per row, and writes the link into the `Token Link` column.
- Uploads the updated workbook to the same SharePoint/OneDrive file.
- Outputs a JSON file with token metadata at `data/output/gift-list-token.json`.

### Read Address API
- File: `pages/api/get-address.js`
- Downloads the workbook, reads rows, and finds a matching record by token.
- Returns a normalized address record to the frontend.

### Update Address API
- File: `pages/api/update-address.js`
- Downloads the workbook, finds the matching row by token, and compares changes.
- Inserts an updated row with change metadata (changed fields/values, match status, submit time).
- Uploads the workbook back to SharePoint/OneDrive.

### Frontend
- File: `pages/index.js`
- Reads the token from the URL query string.
- Calls `GET /api/get-address` to load the address.
- Lets the user edit allowed fields, then calls `POST /api/update-address`.
- Displays status, change summary, and last updated time.

## Environment Variables
Create `.env.local` with the following values:
```
TENANT_ID=...
CLIENT_ID=...
CLIENT_SECRET=...
SHARE_LINK=...
```

## Commands
- Install: `npm install`
- Dev server: `npm run dev`
- Token generation: `node scripts/generate-tokens.js`
