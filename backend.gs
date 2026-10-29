// FundRaja Google Apps Script backend
// Payment detection MVP: notification evidence is never treated as bank-level proof.

const SHEET_ID = ''; // Keep the active spreadsheet ID here if your deployment uses a bound script.
const DRIVE_FOLDER_ID = ''; // Optional public profile-image folder ID.
const GOOGLE_CLIENT_ID = '235867347996-oqv3jdtfopsbh0gdha5v726nmogoo9fg.apps.googleusercontent.com';

const DONATION_EXTRA_HEADERS = [
  'donation_id', 'payment_status', 'detected_event_id',
  'creator_ack_at', 'creator_ack_by', 'credited_at', 'evidence_note',
  'donor_marked_paid_at', 'detector_reported_at', 'donor_user_id'
];
const PAYMENT_EVENT_HEADERS = [
  'event_id', 'received_at', 'device_user_id', 'package_name', 'event_amount',
  'event_utr', 'recipient_upi', 'event_status', 'matched_donation_id',
  'event_title', 'event_text', 'posted_at', 'review_note'
];
const DETECTOR_HEADERS = [
  'device_id', 'user_id', 'token_hash', 'created_at', 'revoked_at',
  'last_seen_at', 'recipient_upi', 'enabled', 'device_name', 'last_event_at'
];

function doGet(e) {
  try {
    const op = e && e.parameter ? e.parameter.op : '';
    const ss = getSpreadsheet_();

    if (op === 'getCampaigns') {
      const sheet = ss.getSheetByName('Campaigns');
      if (!sheet || sheet.getLastRow() < 1) return response([]);
      const data = sheet.getDataRange().getValues();
      data.shift();
      const campaigns = data.filter(row => row[0]).map(row => ({
        id: row[0], title: row[1], description: row[2], target_amount: row[3],
        raised_amount: row[4], upi_id: row[5], creator_name: row[6],
        category: row[7], end_date: row[8], rewards: row[9],
        status: row[10], creator_id: row[11]
      }));
      return response(campaigns);
    }

    if (op === 'getProfile') {
      const userId = String(e.parameter.uid || '');
      const userSheet = ss.getSheetByName('Users');
      if (!userSheet) return response({ error: 'Users sheet not found' });
      const users = userSheet.getDataRange().getValues();
      const user = users.find((row, i) => i > 0 && String(row[0]) === userId);
      if (user) return response({ status: 'success', data: {
        user_id: user[0], phone: user[1], name: user[4], email: user[7],
        about: user[8], socials: user[9], photo_url: user[10]
      }});
      return response({ error: 'User not found' });
    }

    return response({ error: 'Unknown operation' });
  } catch (err) {
    return response({ error: String(err) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(data.action || '');
    // The detector endpoint authenticates before processing any event data.
    if (action === 'detector_event') return receiveDetectorEvent_(data);
    if (action === 'detector_ping') return detectorPing_(getSpreadsheet_(), data);

    const ss = getSpreadsheet_();

    // --- IMAGE UPLOAD ---
    if (action === 'upload_image') {
      if (!data.image || !data.userId) return response({ error: 'Missing image data' });
      if (!DRIVE_FOLDER_ID) return response({ error: 'Profile image storage is not configured' });
      const parts = String(data.image).split(',');
      if (parts.length !== 2 || !/^data:image\/(png|jpeg|webp);base64$/i.test(parts[0])) {
        return response({ error: 'Unsupported image format' });
      }
      const contentType = parts[0].split(':')[1].split(';')[0];
      const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), contentType, data.userId + '_profile');
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return response({ status: 'success', url: 'https://lh3.googleusercontent.com/d/' + file.getId() });
    }

    // --- AUTH ACTIONS (preserves the existing FundRaja auth contract) ---
    if (action === 'signup') {
      const userSheet = ss.getSheetByName('Users');
      if (!userSheet) return response({ error: 'Users sheet not found' });
      const users = userSheet.getDataRange().getValues();
      const phone = String(data.phone || '');
      if (!phone || !data.name || !/^\d{4,8}$/.test(String(data.pin || ''))) return response({ error: 'Enter a name, phone, and 4–8 digit PIN' });
      if (users.some((row, i) => i > 0 && String(row[1]) === phone)) return response({ error: 'Phone number already registered' });
      const salt = Utilities.getUuid();
      const userId = 'USR-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
      userSheet.appendRow([userId, phone, hashPin(data.pin, salt), salt, data.name, new Date(), 'active', '', '', '', '', '']);
      const sessionToken = issueSession_(userSheet, userSheet.getLastRow());
      return response({ status: 'success', userId: userId, name: data.name, sessionToken: sessionToken });
    }

    if (action === 'login') {
      const phone = String(data.phone || '');
      const cache = CacheService.getScriptCache();
      const attemptsKey = 'attempts_' + phone;
      const failedAttempts = parseInt(cache.get(attemptsKey) || '0', 10);
      if (failedAttempts >= 3) return response({ error: 'Account locked. Try again in 15 mins.' });
      const userSheet = ss.getSheetByName('Users');
      if (!userSheet) return response({ error: 'Users sheet not found' });
      const users = userSheet.getDataRange().getValues();
      const userRow = users.find((row, i) => i > 0 && String(row[1]) === phone);
      if (!userRow) return response({ error: 'User not found' });
      if (hashPin(data.pin || '', userRow[3]) === userRow[2]) {
        cache.remove(attemptsKey);
        const rowIndex = users.findIndex((row, i) => i > 0 && String(row[0]) === String(userRow[0])) + 1;
        const sessionToken = issueSession_(userSheet, rowIndex);
        return response({ status: 'success', userId: userRow[0], name: userRow[4], sessionToken: sessionToken });
      }
      cache.put(attemptsKey, String(failedAttempts + 1), 900);
      return response({ error: 'Invalid PIN.' });
    }

    if (action === 'google_auth') {
      // Verify the ID token with Google's tokeninfo endpoint; never trust client-decoded claims.
      const credential = String(data.credential || '');
      if (!credential) return response({ error: 'Google credential required' });
      const verifyResponse = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential), { muteHttpExceptions: true });
      if (verifyResponse.getResponseCode() !== 200) return response({ error: 'Google sign-in token is invalid or expired' });
      const claims = JSON.parse(verifyResponse.getContentText());
      if (String(claims.aud || '') !== GOOGLE_CLIENT_ID || String(claims.email_verified) !== 'true' || !claims.email) {
        return response({ error: 'Google account verification failed' });
      }
      const email = String(claims.email).trim().toLowerCase();
      const displayName = String(claims.name || email).slice(0, 120);
      const userSheet = ss.getSheetByName('Users');
      const users = userSheet.getDataRange().getValues();
      const userRow = users.find((row, i) => i > 0 && String(row[7] || '').toLowerCase() === email);
      if (userRow) {
        const rowIndex = users.findIndex((row, i) => i > 0 && String(row[0]) === String(userRow[0])) + 1;
        const sessionToken = issueSession_(userSheet, rowIndex);
        return response({ status: 'success', userId: userRow[0], name: userRow[4], sessionToken: sessionToken });
      }
      const userId = 'GUSR-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
      userSheet.appendRow([userId, '', '', '', displayName, new Date(), 'active', email, '', '', '', '']);
      const sessionToken = issueSession_(userSheet, userSheet.getLastRow());
      return response({ status: 'success', userId: userId, name: displayName, sessionToken: sessionToken });
    }

    if (action === 'update_profile') {
      const userSheet = ss.getSheetByName('Users');
      const users = userSheet.getDataRange().getValues();
      const rowIndex = users.findIndex((row, i) => i > 0 && String(row[0]) === String(data.userId)) + 1;
      if (rowIndex <= 1) return response({ error: 'User not found' });
      if (data.about && String(data.about).length > 500) return response({ error: 'Bio too long (max 500 chars)' });
      if (data.about !== undefined) userSheet.getRange(rowIndex, 9).setValue(data.about);
      if (data.email !== undefined) userSheet.getRange(rowIndex, 8).setValue(data.email);
      if (data.socials !== undefined) userSheet.getRange(rowIndex, 10).setValue(data.socials);
      if (data.photoUrl !== undefined) userSheet.getRange(rowIndex, 11).setValue(data.photoUrl);
      if (data.phone && data.pin) {
        const phone = String(data.phone);
        if (users.some((row, idx) => idx > 0 && String(row[1]) === phone && idx + 1 !== rowIndex)) return response({ error: 'Phone already in use' });
        const salt = Utilities.getUuid();
        userSheet.getRange(rowIndex, 2).setValue(phone);
        userSheet.getRange(rowIndex, 3).setValue(hashPin(data.pin, salt));
        userSheet.getRange(rowIndex, 4).setValue(salt);
      }
      return response({ status: 'success' });
    }

    // --- CAMPAIGN ACTIONS ---
    if (action === 'createCampaign') {
      const userId = String(data.userId || '');
      if (!userId) return response({ error: 'User ID required' });
      if (!data.title || String(data.title).length > 150) return response({ error: 'Title invalid or too long' });
      if (!data.desc || String(data.desc).length > 2000) return response({ error: 'Description invalid or too long' });
      const cache = CacheService.getScriptCache();
      const key = 'last_post_' + userId;
      if (cache.get(key)) return response({ error: 'Rate limit exceeded. Please wait 10 minutes before posting again.' });
      const sheet = ss.getSheetByName('Campaigns');
      const id = 'CMP-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
      sheet.appendRow([id, data.title, data.desc, Number(data.target) || 0, 0, data.upi, data.creator, data.category, data.endDate, data.rewards, 'Pending', userId]);
      cache.put(key, 'true', 600);
      return response({ status: 'success', id: id });
    }

    // Donor submission creates a pending claim only. It does not increase raised_amount.
    if (action === 'donate') return submitDonation_(ss, data);

    if (action === 'pair_detector') return pairDetector_(ss, data);
    if (action === 'revoke_detector') return revokeDetector_(ss, data);
    if (action === 'get_detector_status') return getDetectorStatus_(ss, data);
    if (action === 'get_creator_donations') return getCreatorDonations_(ss, data);
    if (action === 'get_my_donations') return getMyDonations_(ss, data);
    if (action === 'acknowledge_donation') return acknowledgeDonation_(ss, data);

    return response({ error: 'Unknown action' });
  } catch (err) {
    return response({ error: String(err && err.message ? err.message : err) });
  }
}

function submitDonation_(ss, data) {
  const amount = Number(data.amount);
  const campaignId = String(data.campId || '');
  const utr = String(data.ref || '').trim();
  if (!campaignId || !Number.isFinite(amount) || amount <= 0 || amount > 10000000) return response({ error: 'Enter a valid donation amount' });
  if (!/^\d{12}$/.test(utr)) return response({ error: 'UTR must contain exactly 12 digits' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const campaigns = ss.getSheetByName('Campaigns');
    const campaignRows = campaigns.getDataRange().getValues();
    const campaignIndex = campaignRows.findIndex((r, i) => i > 0 && String(r[0]) === campaignId);
    if (campaignIndex < 1) return response({ error: 'Campaign not found' });
    const campaign = campaignRows[campaignIndex];
    if (String(campaign[10] || '').toLowerCase() !== 'approved') return response({ error: 'This campaign is not accepting donations yet' });

    const donations = ensureDonationSheet_(ss);
    const map = headerMap_(donations);
    const existing = donations.getDataRange().getValues();
    const refCol = map.transaction_ref;
    if (refCol !== undefined && existing.some((row, i) => i > 0 && String(row[refCol] || '').trim() === utr)) {
      return response({ error: 'This UTR has already been submitted' });
    }
    const donationId = 'DON-' + Utilities.getUuid().replace(/-/g, '').toUpperCase();
    const values = new Array(donations.getLastColumn()).fill('');
    setByHeader_(values, map, 'timestamp', new Date());
    setByHeader_(values, map, 'campaign_id', campaignId);
    setByHeader_(values, map, 'donor_name', String(data.donor || '').slice(0, 120));
    setByHeader_(values, map, 'amount', amount);
    setByHeader_(values, map, 'transaction_ref', utr);
    setByHeader_(values, map, 'donation_id', donationId);
    setByHeader_(values, map, 'payment_status', 'AWAITING_EVIDENCE');
    setByHeader_(values, map, 'donor_marked_paid_at', new Date());
    const donorUserId = String(data.userId || '');
    if (donorUserId && requireSession_(ss, donorUserId, data.sessionToken)) setByHeader_(values, map, 'donor_user_id', donorUserId);
    setByHeader_(values, map, 'evidence_note', 'Donator marked as paid; awaiting optional detector evidence and creator verification.');
    donations.appendRow(values);
    return response({ status: 'success', donationId: donationId, paymentStatus: 'AWAITING_EVIDENCE', message: 'Donation claim saved; awaiting creator confirmation.' });
  } finally {
    lock.releaseLock();
  }
}

function pairDetector_(ss, data) {
  const userId = String(data.userId || '');
  if (!requireSession_(ss, userId, data.sessionToken)) return response({ error: 'Session expired. Log in again to pair a detector.' });
  const recipientUpi = String(data.recipientUpi || '').trim().toLowerCase();
  if (!userId || !recipientUpi || !/^[^\s@]+@[^\s@]+$/.test(recipientUpi)) return response({ error: 'A valid creator ID and recipient UPI ID are required' });
  const campaignRows = ss.getSheetByName('Campaigns').getDataRange().getValues();
  const owned = campaignRows.filter((r, i) => i > 0 && String(r[11]) === userId);
  if (!owned.length) return response({ error: 'Create a campaign before pairing a detector' });
  if (!owned.some(r => String(r[5] || '').trim().toLowerCase() === recipientUpi)) return response({ error: 'This UPI ID does not match any of your campaigns' });

  const sheet = ensureDetectorSheet_(ss);
  const rows = sheet.getDataRange().getValues();
  const rawToken = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const deviceId = 'DEV-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
  const now = new Date();
  // Disable older devices for this creator/UPI to make pairing/re-pairing predictable.
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === userId && String(rows[i][6] || '').toLowerCase() === recipientUpi && String(rows[i][4] || '') === '' && String(rows[i][7]).toUpperCase() === 'TRUE') {
      sheet.getRange(i + 1, 5).setValue(now);
      sheet.getRange(i + 1, 8).setValue(false);
    }
  }
  sheet.appendRow([deviceId, userId, sha256_(rawToken), now, '', '', recipientUpi, true]);
  return response({ status: 'success', deviceId: deviceId, token: rawToken, recipientUpi: recipientUpi, message: 'Copy this token into OpenUPI Watch. It is shown only once.' });
}

function revokeDetector_(ss, data) {
  const userId = String(data.userId || '');
  if (!requireSession_(ss, userId, data.sessionToken)) return response({ error: 'Session expired. Log in again.' });
  const deviceId = String(data.deviceId || '');
  const sheet = ensureDetectorSheet_(ss);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === deviceId && String(rows[i][1]) === userId && String(rows[i][4] || '') === '') {
      sheet.getRange(i + 1, 5).setValue(new Date());
      sheet.getRange(i + 1, 8).setValue(false);
      return response({ status: 'success' });
    }
  }
  return response({ error: 'Active detector not found' });
}

function getDetectorStatus_(ss, data) {
  const userId = String(data.userId || '');
  if (!requireSession_(ss, userId, data.sessionToken)) return response({ error: 'Session expired. Log in again.' });
  const sheet = ensureDetectorSheet_(ss);
  const rows = sheet.getDataRange().getValues();
  const map = headerMap_(sheet);
  const devices = rows.slice(1).filter(r =>
    String(getByHeader_(r, map, 'user_id')) === userId &&
    String(getByHeader_(r, map, 'enabled')).toUpperCase() === 'TRUE' &&
    !getByHeader_(r, map, 'revoked_at')
  ).map(r => ({
    deviceId: getByHeader_(r, map, 'device_id'),
    deviceName: String(getByHeader_(r, map, 'device_name') || ''),
    recipientUpi: getByHeader_(r, map, 'recipient_upi'),
    createdAt: getByHeader_(r, map, 'created_at'),
    lastSeenAt: getByHeader_(r, map, 'last_seen_at') || null,
    lastEventAt: getByHeader_(r, map, 'last_event_at') || null
  }));
  return response({ status: 'success', devices: devices });
}

function detectorPing_(ss, data) {
  const token = String(data.token || '');
  const deviceName = String(data.deviceName || '').trim().slice(0, 100);
  const recipientUpi = String(data.recipientUpi || '').trim().toLowerCase();
  if (!token || !deviceName || !recipientUpi) return response({ error: 'Missing detector token, device name, or recipient UPI ID' });
  const sheet = ensureDetectorSheet_(ss);
  const rows = sheet.getDataRange().getValues();
  const map = headerMap_(sheet);
  const tokenHash = sha256_(token);
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (String(getByHeader_(row, map, 'token_hash')) !== tokenHash ||
        String(getByHeader_(row, map, 'enabled')).toUpperCase() !== 'TRUE' ||
        getByHeader_(row, map, 'revoked_at')) continue;
    if (String(getByHeader_(row, map, 'recipient_upi') || '').trim().toLowerCase() !== recipientUpi) {
      return response({ error: 'Recipient UPI does not match this paired device' });
    }
    const rowNumber = i + 1;
    sheet.getRange(rowNumber, map.last_seen_at + 1).setValue(new Date());
    sheet.getRange(rowNumber, map.device_name + 1).setValue(deviceName);
    return response({ status: 'success', deviceId: getByHeader_(row, map, 'device_id'), deviceName: deviceName, checkedAt: new Date() });
  }
  return response({ error: 'Invalid or revoked detector token' });
}

function getCreatorDonations_(ss, data) {
  const userId = String(data.userId || '');
  if (!userId) return response({ error: 'User ID required' });
  if (!requireSession_(ss, userId, data.sessionToken)) return response({ error: 'Session expired. Log in again.' });
  const campaignRows = ss.getSheetByName('Campaigns').getDataRange().getValues();
  const ownedIds = {};
  campaignRows.forEach((r, i) => { if (i > 0 && String(r[11]) === userId) ownedIds[String(r[0])] = true; });
  const sheet = ensureDonationSheet_(ss);
  const map = headerMap_(sheet);
  const rows = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const campId = String(getByHeader_(r, map, 'campaign_id') || '');
    if (!ownedIds[campId]) continue;
    result.push({
      rowNumber: i + 1,
      donationId: String(getByHeader_(r, map, 'donation_id') || ''),
      campaignId: campId,
      donor: String(getByHeader_(r, map, 'donor_name') || ''),
      amount: Number(getByHeader_(r, map, 'amount')) || 0,
      utr: String(getByHeader_(r, map, 'transaction_ref') || ''),
      status: String(getByHeader_(r, map, 'payment_status') || 'LEGACY'),
      note: String(getByHeader_(r, map, 'evidence_note') || ''),
      detectedEventId: String(getByHeader_(r, map, 'detected_event_id') || ''),
      credited: Boolean(getByHeader_(r, map, 'credited_at')),
      donorMarkedPaid: Boolean(getByHeader_(r, map, 'donor_marked_paid_at')),
      donorMarkedPaidAt: getByHeader_(r, map, 'donor_marked_paid_at') || null,
      detectorReported: Boolean(getByHeader_(r, map, 'detector_reported_at')),
      detectorReportedAt: getByHeader_(r, map, 'detector_reported_at') || null,
      creatorAcknowledged: Boolean(getByHeader_(r, map, 'creator_ack_at')),
      creatorAcknowledgedAt: getByHeader_(r, map, 'creator_ack_at') || null
    });
  }
  result.sort((a, b) => b.rowNumber - a.rowNumber);
  return response({ status: 'success', donations: result.slice(0, 100) });
}

function getMyDonations_(ss, data) {
  const userId = String(data.userId || '');
  if (!userId || !requireSession_(ss, userId, data.sessionToken)) return response({ error: 'Session expired. Log in again.' });
  const sheet = ensureDonationSheet_(ss);
  const map = headerMap_(sheet);
  const rows = sheet.getDataRange().getValues();
  const donations = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(getByHeader_(r, map, 'donor_user_id') || '') !== userId) continue;
    donations.push({
      donationId: String(getByHeader_(r, map, 'donation_id') || ''),
      campaignId: String(getByHeader_(r, map, 'campaign_id') || ''),
      donor: String(getByHeader_(r, map, 'donor_name') || ''),
      amount: Number(getByHeader_(r, map, 'amount')) || 0,
      status: String(getByHeader_(r, map, 'payment_status') || 'LEGACY'),
      donorMarkedPaid: Boolean(getByHeader_(r, map, 'donor_marked_paid_at')),
      detectorReported: Boolean(getByHeader_(r, map, 'detector_reported_at')),
      creatorAcknowledged: Boolean(getByHeader_(r, map, 'creator_ack_at')),
      credited: Boolean(getByHeader_(r, map, 'credited_at'))
    });
  }
  donations.sort((a, b) => String(b.donationId).localeCompare(String(a.donationId)));
  return response({ status: 'success', donations: donations.slice(0, 50) });
}

function acknowledgeDonation_(ss, data) {
  const userId = String(data.userId || '');
  const donationId = String(data.donationId || '');
  if (!userId || !donationId) return response({ error: 'User and donation IDs are required' });
  if (!requireSession_(ss, userId, data.sessionToken)) return response({ error: 'Session expired. Log in again before acknowledging a payment.' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const campaignRows = ss.getSheetByName('Campaigns').getDataRange().getValues();
    const donations = ensureDonationSheet_(ss);
    const map = headerMap_(donations);
    const rows = donations.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(getByHeader_(rows[i], map, 'donation_id') || '') === donationId) { rowIndex = i + 1; break; }
    }
    if (rowIndex < 2) return response({ error: 'Donation not found' });
    const row = rows[rowIndex - 1];
    const campId = String(getByHeader_(row, map, 'campaign_id') || '');
    const campaignIndex = campaignRows.findIndex((r, i) => i > 0 && String(r[0]) === campId);
    if (campaignIndex < 1 || String(campaignRows[campaignIndex][11]) !== userId) return response({ error: 'You cannot acknowledge this donation' });
    if (getByHeader_(row, map, 'credited_at')) return response({ status: 'success', alreadyCredited: true, message: 'This donation was already credited.' });
    const amount = Number(getByHeader_(row, map, 'amount'));
    if (!Number.isFinite(amount) || amount <= 0) return response({ error: 'Invalid donation amount' });
    const now = new Date();
    donations.getRange(rowIndex, map.payment_status + 1).setValue('PAID');
    donations.getRange(rowIndex, map.creator_ack_at + 1).setValue(now);
    donations.getRange(rowIndex, map.creator_ack_by + 1).setValue(userId);
    donations.getRange(rowIndex, map.credited_at + 1).setValue(now);
    donations.getRange(rowIndex, map.evidence_note + 1).setValue('Creator manually acknowledged receipt. This is not independent bank verification.');
    const campaignRowIndex = campaignIndex + 1;
    const currentRaised = Number(campaignRows[campaignIndex][4]) || 0;
    ss.getSheetByName('Campaigns').getRange(campaignRowIndex, 5).setValue(currentRaised + amount);
    return response({ status: 'success', paymentStatus: 'PAID', message: 'Creator acknowledgement recorded and campaign total updated once.' });
  } finally {
    lock.releaseLock();
  }
}

function receiveDetectorEvent_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return receiveDetectorEventLocked_(data);
  } finally {
    lock.releaseLock();
  }
}

function receiveDetectorEventLocked_(data) {
  const token = String(data.token || '');
  const eventId = String(data.eventId || '').slice(0, 160);
  if (!token || !eventId) return response({ error: 'Missing detector token or event ID' });
  const ss = getSpreadsheet_();
  const devices = ensureDetectorSheet_(ss);
  const deviceRows = devices.getDataRange().getValues();
  const tokenHash = sha256_(token);
  let device = null, deviceRowIndex = -1;
  for (let i = 1; i < deviceRows.length; i++) {
    if (String(deviceRows[i][2]) === tokenHash && String(deviceRows[i][4] || '') === '' && String(deviceRows[i][7]).toUpperCase() === 'TRUE') {
      device = deviceRows[i]; deviceRowIndex = i + 1; break;
    }
  }
  if (!device) return response({ error: 'Invalid or revoked detector token' });

  const eventSheet = ensurePaymentEventSheet_(ss);
  const eventMap = headerMap_(eventSheet);
  const eventRows = eventSheet.getDataRange().getValues();
  if (eventRows.some((r, i) => i > 0 && String(getByHeader_(r, eventMap, 'event_id')) === eventId)) {
    return response({ status: 'success', duplicate: true, message: 'Event already processed' });
  }

  const title = String(data.title || '').slice(0, 180);
  const text = String(data.text || '').slice(0, 1200);
  const combined = (title + ' ' + text).slice(0, 1400);
  const amount = parseNotificationAmount_(combined);
  const utr = parseNotificationUtr_(combined);
  const recipient = String(data.recipientUpi || device[6] || '').trim().toLowerCase();
  if (!recipient || recipient !== String(device[6] || '').trim().toLowerCase()) return response({ error: 'Recipient UPI does not match this paired device' });
  const detectorMap = headerMap_(devices);
  devices.getRange(deviceRowIndex, detectorMap.last_seen_at + 1).setValue(new Date());
  devices.getRange(deviceRowIndex, detectorMap.last_event_at + 1).setValue(new Date());
  const postedAt = Number(data.postedAt) || Date.now();
  const campaignRows = ss.getSheetByName('Campaigns').getDataRange().getValues();
  const ownerCampaigns = campaignRows.filter((r, i) => i > 0 && String(r[11]) === String(device[1]) && String(r[5] || '').trim().toLowerCase() === recipient);
  let status = 'UNMATCHED', matchedDonationId = '', note = 'Could not safely match this notification to a pending donation.';

  const donationSheet = ensureDonationSheet_(ss);
  const dm = headerMap_(donationSheet);
  const donationRows = donationSheet.getDataRange().getValues();
  const ownedCampaignIds = {};
  ownerCampaigns.forEach(r => { ownedCampaignIds[String(r[0])] = true; });
  let candidates = [];
  if (amount !== null && recipient && ownerCampaigns.length) {
    candidates = donationRows.map((r, i) => ({ r: r, i: i + 1 })).filter(item => {
      if (item.i === 1) return false;
      const campaignId = String(getByHeader_(item.r, dm, 'campaign_id') || '');
      const state = String(getByHeader_(item.r, dm, 'payment_status') || '');
      const donationAmount = Number(getByHeader_(item.r, dm, 'amount'));
      const credited = getByHeader_(item.r, dm, 'credited_at');
      if (!ownedCampaignIds[campaignId] || credited || getByHeader_(item.r, dm, 'detected_event_id') || !['AWAITING_EVIDENCE', 'NEEDS_REVIEW', 'DETECTED_REVIEW'].includes(state)) return false;
      if (Math.round(donationAmount * 100) !== Math.round(amount * 100)) return false;
      const donationUtr = String(getByHeader_(item.r, dm, 'transaction_ref') || '').trim();
      if (utr && donationUtr && utr !== donationUtr) return false;
      return true;
    });
  }
  if (candidates.length === 1) {
    const candidate = candidates[0];
    matchedDonationId = String(getByHeader_(candidate.r, dm, 'donation_id') || '');
    const donationUtr = String(getByHeader_(candidate.r, dm, 'transaction_ref') || '').trim();
    status = 'MATCHED_REVIEW';
    note = (utr && donationUtr && utr === donationUtr)
      ? 'Amount and UTR match; creator must still confirm receipt.'
      : 'Single amount/recipient candidate; creator review required because notification matching is not bank verification.';
    if (matchedDonationId) {
      donationSheet.getRange(candidate.i, dm.payment_status + 1).setValue('DETECTED_REVIEW');
      donationSheet.getRange(candidate.i, dm.detected_event_id + 1).setValue(eventId);
      if (dm.detector_reported_at !== undefined) donationSheet.getRange(candidate.i, dm.detector_reported_at + 1).setValue(new Date());
      donationSheet.getRange(candidate.i, dm.evidence_note + 1).setValue(note);
    }
  } else if (candidates.length > 1) {
    status = 'AMBIGUOUS';
    note = 'Multiple pending donations have the same amount for this recipient; no automatic association made.';
  } else if (amount === null) {
    status = 'UNPARSEABLE';
    note = 'Could not parse a payment amount from this notification. App-specific parsing may be required.';
  }

  const eventValues = new Array(eventSheet.getLastColumn()).fill('');
  setByHeader_(eventValues, eventMap, 'event_id', eventId);
  setByHeader_(eventValues, eventMap, 'received_at', new Date());
  setByHeader_(eventValues, eventMap, 'device_user_id', device[1]);
  setByHeader_(eventValues, eventMap, 'package_name', String(data.packageName || '').slice(0, 180));
  setByHeader_(eventValues, eventMap, 'event_amount', amount === null ? '' : amount);
  setByHeader_(eventValues, eventMap, 'event_utr', utr || '');
  setByHeader_(eventValues, eventMap, 'recipient_upi', recipient);
  setByHeader_(eventValues, eventMap, 'event_status', status);
  setByHeader_(eventValues, eventMap, 'matched_donation_id', matchedDonationId);
  setByHeader_(eventValues, eventMap, 'event_title', title);
  setByHeader_(eventValues, eventMap, 'event_text', text);
  setByHeader_(eventValues, eventMap, 'posted_at', new Date(postedAt));
  setByHeader_(eventValues, eventMap, 'review_note', note);
  eventSheet.appendRow(eventValues);
  devices.getRange(deviceRowIndex, 6).setValue(new Date());
  return response({ status: 'success', eventStatus: status, matchedDonationId: matchedDonationId, amount: amount, note: note });
}

function ensureDonationSheet_(ss) {
  let sheet = ss.getSheetByName('Donations');
  if (!sheet) {
    sheet = ss.insertSheet('Donations');
    sheet.appendRow(['timestamp', 'campaign_id', 'donor_name', 'amount', 'transaction_ref']);
  }
  ensureHeaders_(sheet, DONATION_EXTRA_HEADERS);
  return sheet;
}
function ensurePaymentEventSheet_(ss) {
  let sheet = ss.getSheetByName('PaymentEvents');
  if (!sheet) { sheet = ss.insertSheet('PaymentEvents'); sheet.appendRow(PAYMENT_EVENT_HEADERS); }
  else ensureHeaders_(sheet, PAYMENT_EVENT_HEADERS);
  return sheet;
}
function ensureDetectorSheet_(ss) {
  let sheet = ss.getSheetByName('DetectorDevices');
  if (!sheet) { sheet = ss.insertSheet('DetectorDevices'); sheet.appendRow(DETECTOR_HEADERS); }
  else ensureHeaders_(sheet, DETECTOR_HEADERS);
  return sheet;
}
function ensureHeaders_(sheet, required) {
  if (sheet.getLastRow() === 0) { sheet.appendRow(required); return; }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  required.forEach(h => {
    if (current.indexOf(normalizeHeader_(h)) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      current.push(normalizeHeader_(h));
    }
  });
}
function headerMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { const key = normalizeHeader_(h); if (key) map[key] = i; });
  // Map legacy aliases without renaming existing columns.
  if (map.timestamp === undefined && map.created_at !== undefined) map.timestamp = map.created_at;
  if (map.campaign_id === undefined && map.campid !== undefined) map.campaign_id = map.campid;
  if (map.donor_name === undefined && map.donor !== undefined) map.donor_name = map.donor;
  if (map.transaction_ref === undefined && map.ref !== undefined) map.transaction_ref = map.ref;
  return map;
}
function normalizeHeader_(value) { return String(value || '').trim().toLowerCase().replace(/[\s-]+/g, '_'); }
function setByHeader_(row, map, key, value) { if (map[key] !== undefined) row[map[key]] = value; }
function getByHeader_(row, map, key) { return map[key] === undefined ? '' : row[map[key]]; }
function parseNotificationAmount_(text) {
  const patterns = [/(?:₹|INR\s*|Rs\.?\s*)([0-9][0-9,]*(?:\.\d{1,2})?)/i, /(?:received|credited|paid|payment of)\s*(?:₹|INR\s*|Rs\.?\s*)?([0-9][0-9,]*(?:\.\d{1,2})?)\s*(?:rupees|inr)?/i];
  for (let i = 0; i < patterns.length; i++) {
    const m = text.match(patterns[i]);
    if (m && m[1]) { const n = Number(m[1].replace(/,/g, '')); if (Number.isFinite(n) && n > 0 && n <= 10000000) return n; }
  }
  return null;
}
function parseNotificationUtr_(text) {
  const labelled = text.match(/(?:UTR|UPI\s*ref(?:erence)?|transaction\s*id|reference\s*no\.?)[\s:#-]*([A-Z0-9]{10,24})/i);
  if (labelled && labelled[1]) return labelled[1].toUpperCase();
  const digits = text.match(/\b\d{12}\b/);
  return digits ? digits[0] : '';
}
function sha256_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value));
  return bytes.map(b => { const n = b < 0 ? b + 256 : b; return ('0' + n.toString(16)).slice(-2); }).join('');
}
function getSpreadsheet_() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Set SHEET_ID in backend.gs or use a script bound to the FundRaja spreadsheet.');
  return active;
}
function issueSession_(userSheet, rowIndex) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  // Column L stores only the hash. A new login rotates the previous token.
  if (normalizeHeader_(userSheet.getRange(1, 12).getValue()) !== 'session_token_hash') {
    userSheet.getRange(1, 12).setValue('session_token_hash');
  }
  userSheet.getRange(rowIndex, 12).setValue(sha256_(token));
  return token;
}
function requireSession_(ss, userId, token) {
  if (!userId || !token) return false;
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet || userSheet.getLastRow() < 2) return false;
  const rows = userSheet.getDataRange().getValues();
  const user = rows.find((row, i) => i > 0 && String(row[0]) === String(userId));
  return !!(user && String(user[11] || '') && String(user[11]) === sha256_(String(token)));
}

function hashPin(pin, salt) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin) + String(salt));
  return digest.map(b => { const n = b < 0 ? b + 256 : b; return ('0' + n.toString(16)).slice(-2); }).join('');
}
function response(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
