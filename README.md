# Islam Marble Gate Pass System

The app stores gate passes and company settings directly in the Google spreadsheet configured by `GOOGLE_SHEET_ID`. It does not use a separate database. The spreadsheet must be a native Google Sheet, such as the file at `https://docs.google.com/spreadsheets/d/1jdVg33v88AQlnObha0kxecQW3hah3kGmnUdhfA9ipGU`.

## Google Sheets access

1. In Google Cloud Console, create a project, enable the Google Sheets API, and create a service account with a JSON key.
2. Share the spreadsheet with the service account's `client_email` and give it **Editor** permission. Do not make the sheet public.
3. Configure these environment variables in Vercel for Production, Preview, and Development:
   - `GOOGLE_SHEET_ID`: `1jdVg33v88AQlnObha0kxecQW3hah3kGmnUdhfA9ipGU`
   - `GOOGLE_SERVICE_ACCOUNT_EMAIL`: the service account's `client_email`
   - `GOOGLE_PRIVATE_KEY`: the service account's `private_key` (keep newline characters or enter them as `\n`)
   - `APP_PASSWORD`: a private app sign-in password of at least 12 characters
4. Redeploy after setting the variables. Open the Vercel URL and sign in with `APP_PASSWORD`.

On its first authenticated request, the app creates `Gate Passes` and `Settings` tabs and adds any missing app columns to their header rows. Existing columns and other tabs are left in place. Keep the app's header names unchanged.

The app-password sign-in protects spreadsheet reads and writes. The password is only stored in Vercel environment variables; never commit service-account credentials or `.env` to GitHub.

## Run locally

1. Copy `.env.example` to `.env` and set the service-account email, private key, and a private app password.
2. Share the spreadsheet with the service account as an Editor.
3. Run `npm install`, then `npm run dev`.

## Features

- Gate passes and company settings saved in the Google spreadsheet.
- Excel report export and gate-pass print, PDF, and PNG download.
- Automatic amount calculation (feet × rate).
- No separate database.
