# Yubaraj Finance & Stock

## Contents
- `www/index.html`: offline-first starter app
- `www/manifest.json` and `www/service-worker.js`: PWA support
- `www/icon.png`: supplied app icon image
- `Code.gs`: Google Apps Script endpoint; creates a spreadsheet on first sync and upserts rows by record ID
- `.github/workflows/build-apk.yml`: GitHub Actions debug APK build

## Build APK
1. Upload all files and folders to a GitHub repository (keep hidden `.github` folder).
2. Open **Actions** and run **Build Yubaraj Finance & Stock APK** (or push to `main`).
3. Open the completed workflow run and download artifact `yubaraj-finance-stock-debug-apk`.
4. Extract it and install `app-debug.apk` on Android (allow install from this source).

## Google Sheets setup
1. Open `Code.gs` in `script.google.com` as a new Apps Script project.
2. Deploy > New deployment > Web app.
3. Execute as: **Me**. Access: choose the access option appropriate for your account/use case.
4. Authorize the script, deploy, and copy the Web App URL ending in `/exec`.
5. In the app, open **Online Sync / Backup**, paste URL, Save URL, then Save Online / Sync Now.
The first successful request creates a Google spreadsheet in the account that deployed the script.

## Important starter limitations
This is a functional starter, not yet audited accounting software. Test with sample data and keep backups. Loan schedule currently provides an estimate and records principal/interest payments separately; verify formulas and local lending requirements before live use. Savings and stock are stored separately. This first version is local-device storage; Google sync is an explicit action. For production use, authentication, server-side validation, multi-device conflict handling, encrypted backup and full ledger reconciliation should be added.
