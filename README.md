# FundRaja

FundRaja is a lightweight serverless crowdfunding prototype for campaign discovery, creator profiles, donation claims, and creator-side payment review. The frontend is static HTML/CSS/vanilla JavaScript, the API runs on Google Apps Script, and Google Sheets stores application data. OpenUPI Watch is an optional Android companion that reports payment-app notification evidence; it is not a bank integration.

> **Payment safety:** A donor-submitted claim or Android notification is not proof that funds reached the recipient. A creator must independently check their UPI app or bank record and manually acknowledge receipt before FundRaja credits the campaign total.

## Technology stack

- **Frontend:** HTML, CSS, vanilla JavaScript; responsive landing page, modals, campaign cards, and creator dashboard.
- **Hosting:** GitHub Pages over HTTPS.
- **Backend/API:** Google Apps Script Web App using doGet and doPost.
- **Data:** Google Sheets, with Apps Script LockService protecting selected payment writes from concurrent processing.
- **Identity:** phone + PIN login and Google Identity Services sign-in. Email verification is a separate profile trust signal.
- **Media:** Google Drive image uploads when Drive permissions and storage configuration are set.
- **Optional payment evidence:** Kotlin Android app using Android NotificationListenerService, paired with a creator account using a random device token.
- **Automation:** GitHub Actions for Pages deployment and Android APK builds in the companion repository.

No Node.js, React build, Docker, or separate server hosting is required for the current frontend.

## Repository layout

- **index.html** — complete static site, styles, modals, dashboard, and client-side API calls.
- **backend.gs** — Apps Script backend; copy this file into the Apps Script project attached to the intended spreadsheet.
- **README.md** — setup, architecture, safety notes, and upgrade plan.
- Android companion source and build workflows live in [OpenUPIWatch-Android](https://github.com/deebong/OpenUPIWatch-Android).

## Features

- Responsive landing page, campaign discovery, search, category filtering, and campaign cards.
- Account creation, phone/PIN login, Google sign-in, server-issued session tokens, and profile editing.
- Email verification codes sent by Apps Script Mail service.
- Public creator profiles, profile photos, optional social links, and campaign images.
- Campaign creation and creator-owned campaign editing.
- Donation claims with amount and 12-digit UTR validation and duplicate-UTR rejection.
- Creator review of donation claims and optional OpenUPI Watch device status.
- Detector pairing, revocation, event deduplication, basic notification amount/UTR parsing, and candidate matching.
- Manual creator acknowledgement; the same donation is credited at most once by the acknowledgement handler.
- In-site notice and confirmation modals.

Features depend on the deployed Apps Script version, Google permissions, and correct sheet headers. The public repository alone does not prove that the live Apps Script backend has been updated.

## Architecture and payment lifecycle

1. A donor transfers money directly through their own UPI app, outside FundRaja.
2. The donor submits a donation claim and UTR. FundRaja records it as pending evidence; it does not increase the campaign total at this stage.
3. If enabled, OpenUPI Watch may report a payment notification as optional evidence. Notifications can be missing, delayed, misparsed, spoofed, or unrelated.
4. FundRaja may mark a candidate for creator review when notification details match a pending claim. Matching is advisory only.
5. The creator checks the actual UPI/bank transaction history and manually acknowledges receipt.
6. Only that creator acknowledgement updates the displayed campaign total. This is not an independent verification service and cannot guarantee the creator has honestly checked the bank record.

## Install and deploy

### 1. Prepare Google Sheets

1. Create a dedicated spreadsheet for this deployment. Do not use a spreadsheet containing unrelated personal or business data.
2. Keep the spreadsheet private; do not publish it to the web or share it with donors.
3. Create the base sheets and headers described in [Google Sheets schema](#google-sheets-schema). The backend can create detector/event sheets and add payment-related columns when needed.
4. Back up the spreadsheet before deploying updates.

### 2. Deploy the Apps Script backend

1. Open **Extensions → Apps Script** from the spreadsheet, or create a standalone Apps Script project.
2. Paste the complete current backend.gs into the script editor.
3. For a standalone script, set SHEET_ID in backend.gs to the intended spreadsheet ID. For a spreadsheet-bound script, it can remain blank.
4. Review GOOGLE_CLIENT_ID and use the OAuth client registered for your deployment if different.
5. Save the project.
6. For email and Drive features, select authorizeFundRajaServices in the function selector and click **Run**. Review the requested permissions and authorize with the account that owns the deployment. This function checks Mail quota and prepares the image folder.
7. Select **Deploy → New deployment → Web app**. Choose an execution identity that can access the spreadsheet and set access according to your public-app model. Copy the Web App URL.
8. For backend updates, use **Deploy → Manage deployments → Edit → New version → Deploy**. Keep the Web App URL stable if the frontend points to it.

**Important:** A script executing as its owner can access that owner's spreadsheet and Drive. Treat the Web App URL as public, validate every action server-side, and never expose the spreadsheet itself.

### 3. Configure and deploy the frontend

1. Update APPS_SCRIPT_URL in index.html to the deployed Web App URL.
2. Ensure the Google Identity Services client ID in the HTML matches backend.gs.
3. Push index.html and this README to the repository's main branch. GitHub Pages should deploy the static site via the repository's Pages workflow/settings.
4. Wait for the Pages Actions run to succeed, then open the published site and perform a hard refresh.
5. Test using a separate spreadsheet and non-production accounts before using real campaigns.

### 4. First-run verification checklist

- Create an account and log out/in; confirm the session works.
- Verify an email and confirm the verification controls disappear immediately.
- Create a test campaign and confirm its status.
- Confirm public creator profiles do not reveal private phone or email values.
- Submit a test donation claim with a non-sensitive test UTR and confirm it remains pending.
- Test pairing/revoking OpenUPI Watch and confirm events do not automatically credit a campaign.
- Acknowledge a test claim and confirm repeated acknowledgement does not add the amount twice.
- Test narrow phone, tablet, and desktop widths, including opening and closing the dashboard sidebar.
- Review browser console and Apps Script execution logs. Do not use real payment data for test cases.

## Google Sheets schema

The existing workbook may contain legacy columns; header-based helpers preserve many existing headers and add missing ones.

### Users

Core fields include user ID, phone, PIN hash, PIN salt, display name, created timestamp, account status, email, about/bio, socials, photo URL, session token hash, and email verification state. The session token itself should not be stored in the sheet; only its hash is stored.

### Campaigns

The current frontend/backend use campaign ID, title, description/story, target amount, raised amount, recipient UPI ID, creator display name, category, end date, rewards/thank-you note, moderation status, creator user ID, and optional image URL. Creator user ID is used for ownership checks.

### Donations

The backend starts with legacy donation fields and adds unique donation ID, payment status, detected event ID, creator acknowledgement timestamp/user, credited timestamp, evidence note, donor marked-paid timestamp, detector report timestamp, and donor user ID. Header names matter; retain existing data and back up before changing columns.

### DetectorDevices

Stores device ID, creator user ID, SHA-256 token hash, creation/revocation timestamps, last-seen and last-event timestamps, recipient UPI ID, enabled state, device name, and pairing metadata. The raw pairing token is returned once during pairing and must be treated as a secret.

### PaymentEvents

Stores deduplicated event ID, receipt timestamp, paired creator/device, source package, parsed amount and UTR, recipient UPI ID, review status, optional matched donation ID, notification title/text, posted timestamp, and review note. Notification text can contain sensitive transaction details; restrict access to this sheet.

## OpenUPI Watch setup

1. Open [OpenUPIWatch-Android](https://github.com/deebong/OpenUPIWatch-Android) and run its Android build workflow, or use a verified APK artifact/release produced by that repository.
2. Install the APK on a compatible Android phone.
3. In FundRaja, open the creator dashboard and choose the detector pairing option to generate pairing configuration.
4. Import the pairing information in OpenUPI Watch, save it, and grant notification access in Android Settings.
5. Confirm the app reports that the backend connected and registered the device.
6. Use controlled test notifications where appropriate. Never treat a notification as a bank statement or payment confirmation.

The Android app should request only permissions necessary for notification listening and backend communication. Review the manifest and Android source before building a release. Do not share pairing QR codes, tokens, screenshots showing tokens, or production notification data.

## Security model and limitations

This repository is a prototype built on static hosting, Google Apps Script, and Google Sheets. These controls reduce common risks but do not make the service equivalent to a professionally operated financial platform.

Implemented in source:
- Sensitive actions check server-issued session tokens; only session-token hashes are stored.
- Public profile responses omit private phone and email fields; authenticated profile retrieval requires the owner's session.
- PINs use a per-user salt and a secret pepper stored in Apps Script Script Properties. Existing legacy salted-SHA-256 PIN hashes are upgraded after a successful login.
- New PINs must contain 8–12 digits. Existing users may continue logging in with their existing PIN and should change it to a stronger PIN.
- Login attempts are limited per phone identifier with persistent Script Properties counters; signup, Google sign-in, and email verification requests also have server-side throttles.
- Detector tokens are stored as hashes, and notification events remain advisory evidence rather than proof of payment.
- Creator ownership checks and locks protect selected campaign and payment writes.

### Remaining production blockers

- **PIN hashing:** HMAC-SHA-256 with a secret pepper is stronger than an unpeppered fast hash, but it is not a memory-hard password hash. Numeric PINs remain low entropy. For accounts holding sensitive information or supporting real-money workflows, migrate to a managed identity provider or a server capable of Argon2id, scrypt, or bcrypt, and require stronger credentials.
- **Rate-limit scale and abuse resistance:** Script Properties quotas and Apps Script execution limits make these controls a baseline, not an enterprise-grade abuse-control system. Use an external API gateway/WAF and durable datastore-backed throttling for a production financial service. Apps Script does not reliably expose client IP addresses to this application.
- **Verification codes:** Email codes still need durable attempt counting, strict single-use semantics, and monitoring. Cache-backed verification state can be evicted.
- **Operational security:** Limit Apps Script editor access, review deployment permissions and OAuth scopes, protect the spreadsheet, rotate secrets carefully, and configure monitoring and incident response.
- **Data and payments:** Google Sheets is not a transactional financial ledger. Add independent reconciliation, immutable audit logs, backups, privacy/retention controls, and formal payment-provider integrations before handling material funds.
- **Security assurance:** No independent penetration test, full threat model, or live Apps Script deployment test has been performed. GitHub Pages deployment does not deploy `backend.gs`; the Apps Script owner must copy the updated file into Apps Script and deploy a new version before the backend changes take effect.

Do not describe this repository as production-ready for regulated or high-value financial activity until the remaining blockers are addressed and independently tested.

## Recommended production upgrades

1. Replace PIN-only authentication with a managed identity system; add account recovery and revocation of all active sessions.
2. Move payments, accounts, and audit records from Sheets to a database with transactions, access control, backups, and retention policies.
3. Add durable rate limiting, anti-abuse controls, email verification attempt limits, and server-side request validation for every action.
4. Add moderation/admin tooling, campaign reporting, immutable audit events, and a creator dispute workflow.
5. Add explicit consent, configurable notification-data minimization, retention/deletion controls, and a privacy/security contact.
6. Add automated tests for authorization boundaries, duplicate submissions, race conditions, malformed inputs, and XSS payloads; run dependency and secret scanning in CI.
7. Add monitoring, redacted error reporting, scheduled backups, restore drills, and incident response documentation.
8. Add accessibility checks and real-device responsive/browser testing.
9. Add provider-specific notification parsers only where permitted and supported; continue to require manual creator confirmation.
10. Before public launch, commission an independent security review and test the deployed Apps Script permissions and spreadsheet sharing settings.

## Troubleshooting

- **Campaigns do not load:** Check the Apps Script Web App URL, deployment version, execution logs, spreadsheet ID, and sheet headers. Inspect the browser console/network response.
- **Email verification does not send:** Run authorizeFundRajaServices, check Mail quota, confirm the deployed backend is current, and inspect Apps Script execution logs.
- **Image upload fails:** Confirm Drive authorization and image folder configuration. Image URLs are public by design; upload only content intended for public viewing.
- **Session expired:** Sign in again. The backend rotates the stored session-token hash on login.
- **OpenUPI Watch does not connect:** Confirm Web App URL, pairing token, recipient UPI ID, internet access, and Android notification-access setting. Re-pair if the token was revoked.
- **A detector event appears:** Treat it as review evidence only. Confirm receipt directly in the UPI app or bank record before acknowledging a donation.
- **Pages changes are not visible:** Check the GitHub Pages Actions run and hard-refresh after deployment completes.

## License

FundRaja is released under the [MIT License](LICENSE). The license text is included in the repository root. Contributions and reuse are subject to the terms of that license.
