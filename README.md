# FundRaja + OpenUPI Watch: payment-detection MVP

This bundle updates the existing vanilla HTML/JavaScript frontend and Google Apps Script backend and includes a starter native Android notification listener.

## Files
- `fundraja_index.html` — updated FundRaja frontend with detector pairing, creator donation review, and acknowledgement UI.
- `backend.gs` — updated Apps Script backend.
- `OpenUPIWatch-Android-MVP/` — native Kotlin Android app and GitHub Actions workflow.

## Backend deployment
1. Back up the current Apps Script project and Google Sheet.
2. Open the existing Apps Script project attached to the FundRaja spreadsheet.
3. Replace the current backend code with `backend.gs`. If this is not a container-bound script, set `SHEET_ID` to the correct spreadsheet ID before deploying. Do not use another project's Sheet ID.
4. Save, then use **Deploy → Manage deployments → Edit → New version → Deploy** so the existing Web App URL can remain in use.
5. Log out of FundRaja and log back in once. New server-issued session tokens are required for creator pairing and acknowledgement. The token hash is stored in Users column L (`session_token_hash`).
6. Test on a copy of the spreadsheet first. The backend automatically adds new columns to `Donations` and creates `PaymentEvents` and `DetectorDevices` sheets as needed.

## Frontend deployment
Replace the currently hosted `fundraja_index.html` with this updated file. Keep the existing `APPS_SCRIPT_URL` and Google client ID unless your deployment uses different values. Test in a staging copy before production.

## Build the Android APK without local Android tooling
1. Create a new public GitHub repository named `openupi-watch`.
2. Upload the contents of `OpenUPIWatch-Android-MVP/` to the repository root, preserving paths.
3. Open **Actions**, enable workflows if asked, and run **Build OpenUPI Watch APK**.
4. Download the `openupi-watch-debug-apk` artifact from the successful workflow run and extract `app-debug.apk`.
5. Install it on an Android phone. Open OpenUPI Watch, paste the FundRaja Apps Script Web App URL, pairing token, and recipient UPI ID; save; then grant notification access in Android Settings.
6. In FundRaja, open **My Profile → Automatic UPI Payment Detection → Pair Android Detector**. Copy the one-time token into the Android app. Do not share it.

## Current behavior and limits
- Donor UTR submission creates a pending donation claim; it no longer increments the campaign total immediately.
- The backend rejects duplicate submitted UTRs and assigns a unique donation ID.
- Android notification events are deduplicated by event ID and recorded in `PaymentEvents`.
- Matching uses the configured recipient UPI, parsed amount, and UTR when available. Ambiguous/unparseable notifications remain review items.
- A raw notification never automatically increases campaign totals. A creator must check their UPI/bank record and acknowledge receipt. Acknowledgement is idempotent and increments the campaign total once.
- Notification parsing is generic and will need app-specific tuning. Some payment apps do not expose enough details in notifications; some payments will not be detected. This is not independent bank-level verification.
- The app requests notification-listener access only. It does not request SMS permission or read unrelated app data by design.
- This is an MVP. Apps Script + Sheets is not a transactional database; test concurrency and recovery carefully before using for real fundraising.

## Security notes before production
- The new pairing, detector status, donation-list, and acknowledgement actions require a server-issued session token; only its hash is stored in the Users sheet. Existing browser sessions must log in again.
- Google sign-in ID tokens are checked server-side using Google's tokeninfo endpoint and the configured OAuth client ID.
- Never put detector tokens in public frontend code, screenshots, GitHub issues, or source control. The Android app stores its token locally on the paired device.
- Review the full existing app authentication and campaign moderation model before production. This MVP does not constitute a full KYC system, fraud-control system, or bank verification integration.
