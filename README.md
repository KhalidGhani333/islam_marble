# Islam Marble Gate Pass System

The app runs entirely in the browser. Gate passes, material rows, and company settings are stored in the `.xlsx` file you select. The app does not send records to a database, Google Sheets, or an API.

## Use the Excel file

1. Open the app in Chrome or Edge on `localhost` or an HTTPS deployment.
2. Choose **New Excel** to create a workbook, or **Open Excel** to select an existing workbook created by this app.
3. Keep the workbook on your computer and use the same file when returning to the app. After reloading the browser, select the file again.
4. The app writes every saved, edited, or deleted gate pass directly to that workbook. **Download Excel Copy** creates an extra copy.

The workbook contains `Gate Passes` and `Settings` sheets. Keep the header rows unchanged. Other worksheets in the selected file are preserved.

Direct file access requires a browser that supports the File System Access API (Chrome or Edge) and a secure context (`localhost` or HTTPS). Each computer/browser has its own selected file; changes are not shared across devices unless you separately share the Excel file.

## Features

- Gate-pass records and company settings saved in the selected Excel workbook.
- Excel report export, plus gate-pass print, PDF, and PNG download.
- Automatic amount calculation (feet × rate).

## Run locally

```sh
npm install
npm run dev
```

Open the local URL printed by Vite, then create or open an Excel workbook.

## Deploy from GitHub to Vercel

Import the repository into Vercel and deploy it as a Vite app. No API routes, database, Google credentials, or environment variables are required. The deployed site only edits a workbook after the user selects it in their browser.
