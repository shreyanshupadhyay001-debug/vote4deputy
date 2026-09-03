/**
 * HOUSE OF THUNDERSOUL — DEPUTY HOUSE REPRESENTATIVE ELECTION 2026
 * Backend (Google Apps Script)
 *
 * SHEETS EXPECTED IN THIS SPREADSHEET:
 *   Voters   -> Name | Email | House | Eligible
 *   Votes    -> Timestamp | Voter Name | Email | Candidate | Vote ID | Status
 *   OTP      -> Email | OTP Code | Created At | Expires At | Attempts | Verified | Session Token | Session Expires At | Last Sent At | Consumed
 *   Settings -> Key | Value
 *
 * Nothing in this file should ever hand the full voter list, OTP codes, or
 * vote records back to the browser. Every function callable from the page
 * (google.script.run) re-checks eligibility, OTP validity and vote status
 * on the server, independent of anything the client claims.
 */

// ---------- CONFIG ----------
const SHEET_VOTERS = 'Voters';
const SHEET_VOTES = 'Votes';
const SHEET_OTP = 'OTP';
const SHEET_SETTINGS = 'Settings';

const OTP_LENGTH = 6;
const OTP_MAX_ATTEMPTS = 5;
const OTP_RESEND_COOLDOWN_SECONDS = 45;
const SESSION_VALID_MINUTES = 20; // how long a verified OTP session may be used to cast a vote

// ---------- WEB APP ENTRY POINT ----------
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('House of ThunderSoul — Election 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- SHEET HELPERS ----------
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function getSettings_() {
  const sh = sheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (key) map[String(key).trim()] = data[i][1];
  }
  return map;
}

function isElectionOpen_() {
  const s = getSettings_();
  return String(s['ElectionStatus'] || 'CLOSED').trim().toUpperCase() === 'OPEN';
}

// ---------- PUBLIC: ELECTION INFO (safe to expose) ----------
function getElectionInfo() {
  const s = getSettings_();
  return {
    status: String(s['ElectionStatus'] || 'CLOSED').trim().toUpperCase(),
    title: s['ElectionTitle'] || 'DEPUTY HOUSE REPRESENTATIVE ELECTION 2026',
    logoUrl: s['HouseLogoUrl'] || '',
    candidates: [
      { id: 'C1', name: s['Candidate1Name'] || 'Candidate One', imageUrl: s['Candidate1ImageUrl'] || '' },
      { id: 'C2', name: s['Candidate2Name'] || 'Candidate Two', imageUrl: s['Candidate2ImageUrl'] || '' }
    ]
  };
}

// ---------- VOTER LOOKUP ----------
function findVoterRow_(email) {
  const sh = sheet_(SHEET_VOTERS);
  const data = sh.getDataRange().getValues(); // Name | Email | House | Eligible
  const target = normalizeEmail_(email);
  for (let i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][1]) === target) {
      return { row: i + 1, name: data[i][0], email: data[i][1], house: data[i][2], eligible: data[i][3] };
    }
  }
  return null;
}

function isEligibleVoter_(email) {
  const v = findVoterRow_(email);
  if (!v) return { ok: false };
  const eligible = String(v.eligible).trim().toUpperCase();
  if (eligible !== 'TRUE' && eligible !== 'YES' && v.eligible !== true) return { ok: false };
  return { ok: true, voter: v };
}

// ---------- OTP SHEET HELPERS ----------
// OTP columns: Email | OTP Code | Created At | Expires At | Attempts | Verified | Session Token | Session Expires At | Last Sent At | Consumed
function findOtpRowIndex_(email) {
  const sh = sheet_(SHEET_OTP);
  const data = sh.getDataRange().getValues();
  const target = normalizeEmail_(email);
  for (let i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][0]) === target) return i + 1; // 1-based sheet row
  }
  return -1;
}

function writeOtpRow_(email, otpCode, createdAt, expiresAt) {
  const sh = sheet_(SHEET_OTP);
  const rowIdx = findOtpRowIndex_(email);
  const row = [normalizeEmail_(email), otpCode, createdAt, expiresAt, 0, false, '', '', createdAt, false];
  if (rowIdx === -1) {
    sh.appendRow(row);
  } else {
    sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  }
}

function readOtpRow_(email) {
  const idx = findOtpRowIndex_(email);
  if (idx === -1) return null;
  const sh = sheet_(SHEET_OTP);
  const vals = sh.getRange(idx, 1, 1, 10).getValues()[0];
  return {
    rowIndex: idx,
    email: vals[0],
    code: String(vals[1]),
    createdAt: vals[2],
    expiresAt: vals[3],
    attempts: Number(vals[4] || 0),
    verified: vals[5] === true,
    sessionToken: vals[6],
    sessionExpiresAt: vals[7],
    lastSentAt: vals[8],
    consumed: vals[9] === true
  };
}

function updateOtpRow_(rowIndex, updates) {
  // updates: {attempts, verified, sessionToken, sessionExpiresAt, consumed}
  const sh = sheet_(SHEET_OTP);
  const current = sh.getRange(rowIndex, 1, 1, 10).getValues()[0];
  if (updates.attempts !== undefined) current[4] = updates.attempts;
  if (updates.verified !== undefined) current[5] = updates.verified;
  if (updates.sessionToken !== undefined) current[6] = updates.sessionToken;
  if (updates.sessionExpiresAt !== undefined) current[7] = updates.sessionExpiresAt;
  if (updates.consumed !== undefined) current[9] = updates.consumed;
  sh.getRange(rowIndex, 1, 1, 10).setValues([current]);
}

function generateOtpCode_() {
  let code = '';
  for (let i = 0; i < OTP_LENGTH; i++) code += Math.floor(Math.random() * 10);
  return code;
}

// ---------- STEP 1: REQUEST OTP ----------
// Returns a generic, non-revealing status object to the client.
function requestOtp(rawEmail) {
  try {
    if (!isElectionOpen_()) {
      return { ok: false, code: 'ELECTION_CLOSED', message: 'Voting for the Deputy House Representative has now concluded.' };
    }

    const email = normalizeEmail_(rawEmail);
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return { ok: false, code: 'INVALID_EMAIL', message: 'Please enter a valid university email address.' };
    }

    const eligibility = isEligibleVoter_(email);
    if (!eligibility.ok) {
      return { ok: false, code: 'NOT_ELIGIBLE', message: 'You are not registered as a member of House ThunderSoul and are not eligible to vote in this election.' };
    }

    // Already voted? Don't waste an OTP, but don't reveal candidate.
    if (findVoteRow_(email)) {
      return { ok: false, code: 'ALREADY_VOTED', message: 'Our records show that your vote has already been submitted. Each ThunderSoul member is permitted to vote only once.' };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const existing = readOtpRow_(email);
      const now = new Date();

      if (existing && existing.lastSentAt) {
        const secondsSinceLast = (now.getTime() - new Date(existing.lastSentAt).getTime()) / 1000;
        if (secondsSinceLast < OTP_RESEND_COOLDOWN_SECONDS) {
          const wait = Math.ceil(OTP_RESEND_COOLDOWN_SECONDS - secondsSinceLast);
          return { ok: false, code: 'COOLDOWN', message: 'Please wait ' + wait + ' seconds before requesting a new code.', retryAfterSeconds: wait };
        }
      }

      const settings = getSettings_();
      const expiryMinutes = Number(settings['OtpExpiryMinutes'] || 5);
      const code = generateOtpCode_();
      const expiresAt = new Date(now.getTime() + expiryMinutes * 60000);

      writeOtpRow_(email, code, now, expiresAt);
      sendOtpEmail_(email, eligibility.voter.name, code, expiryMinutes);

      return { ok: true, message: 'A verification code has been sent to your university email.', expiryMinutes: expiryMinutes };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return { ok: false, code: 'SERVER_ERROR', message: 'Something went wrong while processing your request. Please try again.' };
  }
}

function sendOtpEmail_(email, name, code, expiryMinutes) {
  const settings = getSettings_();
  const title = settings['ElectionTitle'] || 'House of ThunderSoul Election';
  const subject = 'Your ThunderSoul Verification Code';
  const body =
    'Dear ' + (name || 'Member') + ',\n\n' +
    'Your verification code for the ' + title + ' is:\n\n' +
    code + '\n\n' +
    'This code will expire in ' + expiryMinutes + ' minutes. Do not share this code with anyone.\n\n' +
    'If you did not request this code, you can safely ignore this email.\n\n' +
    'Strike With Purpose. Rise With Power.\nHouse of ThunderSoul';
  MailApp.sendEmail(email, subject, body);
}

// ---------- STEP 2: VERIFY OTP ----------
function verifyOtp(rawEmail, code) {
  try {
    if (!isElectionOpen_()) {
      return { ok: false, code: 'ELECTION_CLOSED', message: 'Voting for the Deputy House Representative has now concluded.' };
    }

    const email = normalizeEmail_(rawEmail);
    const eligibility = isEligibleVoter_(email);
    if (!eligibility.ok) {
      return { ok: false, code: 'NOT_ELIGIBLE', message: 'You are not registered as a member of House ThunderSoul and are not eligible to vote in this election.' };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const record = readOtpRow_(email);
      if (!record) {
        return { ok: false, code: 'NO_OTP', message: 'Please request a verification code first.' };
      }

      if (record.consumed) {
        return { ok: false, code: 'OTP_USED', message: 'This verification code has already been used. Please request a new one.' };
      }

      const now = new Date();
      if (new Date(record.expiresAt).getTime() < now.getTime()) {
        return { ok: false, code: 'OTP_EXPIRED', message: 'This verification code has expired. Please request a new code.' };
      }

      if (record.attempts >= OTP_MAX_ATTEMPTS) {
        return { ok: false, code: 'TOO_MANY_ATTEMPTS', message: 'Too many incorrect attempts. Please request a new verification code.' };
      }

      const submitted = String(code || '').trim();
      if (submitted !== record.code) {
        updateOtpRow_(record.rowIndex, { attempts: record.attempts + 1 });
        return { ok: false, code: 'WRONG_CODE', message: 'The verification code is incorrect. Please try again.' };
      }

      // Success — issue a short-lived session token bound to this email.
      const sessionToken = Utilities.getUuid();
      const sessionExpiresAt = new Date(now.getTime() + SESSION_VALID_MINUTES * 60000);
      updateOtpRow_(record.rowIndex, { verified: true, sessionToken: sessionToken, sessionExpiresAt: sessionExpiresAt });

      return {
        ok: true,
        message: 'Your ThunderSoul membership has been verified.',
        sessionToken: sessionToken,
        firstName: String(eligibility.voter.name || '').trim().split(' ')[0]
      };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return { ok: false, code: 'SERVER_ERROR', message: 'Something went wrong while processing your request. Please try again.' };
  }
}

// ---------- VOTE LOOKUP ----------
function findVoteRow_(email) {
  const sh = sheet_(SHEET_VOTES);
  const data = sh.getDataRange().getValues(); // Timestamp | Voter Name | Email | Candidate | Vote ID | Status
  const target = normalizeEmail_(email);
  for (let i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][2]) === target && String(data[i][5]).trim().toUpperCase() === 'SUBMITTED') {
      return { row: i + 1 };
    }
  }
  return null;
}

// ---------- STEP 3: SUBMIT VOTE ----------
function submitVote(rawEmail, sessionToken, candidateId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { ok: false, code: 'SERVER_ERROR', message: 'Something went wrong while processing your request. Please try again.' };
  }

  try {
    if (!isElectionOpen_()) {
      return { ok: false, code: 'ELECTION_CLOSED', message: 'Voting for the Deputy House Representative has now concluded.' };
    }

    const email = normalizeEmail_(rawEmail);
    const eligibility = isEligibleVoter_(email);
    if (!eligibility.ok) {
      return { ok: false, code: 'NOT_ELIGIBLE', message: 'You are not registered as a member of House ThunderSoul and are not eligible to vote in this election.' };
    }

    const record = readOtpRow_(email);
    if (!record || !record.verified || record.sessionToken !== sessionToken) {
      return { ok: false, code: 'NOT_VERIFIED', message: 'Your session could not be verified. Please verify your email again.' };
    }
    if (record.consumed) {
      return { ok: false, code: 'ALREADY_VOTED', message: 'Our records show that your vote has already been submitted. Each ThunderSoul member is permitted to vote only once.' };
    }
    if (new Date(record.sessionExpiresAt).getTime() < new Date().getTime()) {
      return { ok: false, code: 'SESSION_EXPIRED', message: 'Your verification session has expired. Please verify your email again.' };
    }

    if (findVoteRow_(email)) {
      return { ok: false, code: 'ALREADY_VOTED', message: 'Our records show that your vote has already been submitted. Each ThunderSoul member is permitted to vote only once.' };
    }

    const info = getElectionInfo();
    const candidate = info.candidates.find(c => c.id === candidateId);
    if (!candidate) {
      return { ok: false, code: 'INVALID_CANDIDATE', message: 'Something went wrong while processing your request. Please try again.' };
    }

    const voteId = Utilities.getUuid();
    const sh = sheet_(SHEET_VOTES);
    sh.appendRow([new Date(), eligibility.voter.name, eligibility.voter.email, candidate.name, voteId, 'SUBMITTED']);

    // Consume the OTP session so it can never be used again.
    updateOtpRow_(record.rowIndex, { consumed: true });

    return { ok: true, message: 'Your vote has been successfully recorded.', voteId: voteId };
  } catch (err) {
    return { ok: false, code: 'SERVER_ERROR', message: 'Something went wrong while processing your request. Please try again.' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- CHECK VOTE STATUS (used so a returning/refreshing voter can't vote twice) ----------
function checkVoteStatus(rawEmail) {
  const email = normalizeEmail_(rawEmail);
  const eligibility = isEligibleVoter_(email);
  if (!eligibility.ok) {
    return { ok: false, code: 'NOT_ELIGIBLE', message: 'You are not registered as a member of House ThunderSoul and are not eligible to vote in this election.' };
  }
  const voted = !!findVoteRow_(email);
  return { ok: true, hasVoted: voted };
}

// =======================================================================
//  ADMIN TOOLS — run from the Google Sheet's custom menu, never exposed
//  to the voter-facing web app.
// =======================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ThunderSoul Admin')
    .addItem('Open Election', 'adminOpenElection')
    .addItem('Close Election', 'adminCloseElection')
    .addSeparator()
    .addItem('View Results', 'adminViewResults')
    .addSeparator()
    .addItem('Set Up Sheets (first-time only)', 'adminSetupSheets')
    .addItem('Import Voter List From This Project', 'adminSeedVoters')
    .addItem('Clean Up Expired OTPs', 'adminCleanupOtps')
    .addToUi();
}

function adminOpenElection() {
  setSetting_('ElectionStatus', 'OPEN');
  SpreadsheetApp.getUi().alert('Election is now OPEN.');
}

function adminCloseElection() {
  setSetting_('ElectionStatus', 'CLOSED');
  SpreadsheetApp.getUi().alert('Election is now CLOSED.');
}

function setSetting_(key, value) {
  const sh = sheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

// Admin-only: tallies results. Never callable from the web app.
function adminViewResults() {
  const votesSh = sheet_(SHEET_VOTES);
  const data = votesSh.getDataRange().getValues();
  const tally = {};
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][5]).trim().toUpperCase() === 'SUBMITTED') {
      const candidate = data[i][3];
      tally[candidate] = (tally[candidate] || 0) + 1;
      total++;
    }
  }

  const votersSh = sheet_(SHEET_VOTERS);
  const voterData = votersSh.getDataRange().getValues();
  let eligibleCount = 0;
  for (let i = 1; i < voterData.length; i++) {
    const eligible = String(voterData[i][3]).trim().toUpperCase();
    if (eligible === 'TRUE' || eligible === 'YES') eligibleCount++;
  }

  let msg = 'RESULTS (admin only)\n\n';
  Object.keys(tally).forEach(name => { msg += name + ': ' + tally[name] + '\n'; });
  msg += '\nTotal Votes Cast: ' + total;
  msg += '\nEligible Voters: ' + eligibleCount;
  msg += '\nTurnout: ' + (eligibleCount ? ((total / eligibleCount) * 100).toFixed(1) : '0') + '%';

  // Also write to a private Results sheet for a persistent record.
  let resultsSh = ss_().getSheetByName('Results');
  if (!resultsSh) resultsSh = ss_().insertSheet('Results');
  resultsSh.clear();
  resultsSh.appendRow(['Candidate', 'Votes']);
  Object.keys(tally).forEach(name => resultsSh.appendRow([name, tally[name]]));
  resultsSh.appendRow([]);
  resultsSh.appendRow(['Total Votes Cast', total]);
  resultsSh.appendRow(['Eligible Voters', eligibleCount]);
  resultsSh.appendRow(['Turnout %', eligibleCount ? ((total / eligibleCount) * 100).toFixed(1) : 0]);

  SpreadsheetApp.getUi().alert(msg);
}

function adminCleanupOtps() {
  const sh = sheet_(SHEET_OTP);
  const data = sh.getDataRange().getValues();
  const now = new Date().getTime();
  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    const expiresAt = data[i][3];
    const consumed = data[i][9] === true;
    // Keep consumed rows briefly for audit; delete anything expired > 1 day, or long-consumed rows.
    if (expiresAt && (now - new Date(expiresAt).getTime()) > 24 * 60 * 60 * 1000) {
      rowsToDelete.push(i + 1);
    } else if (consumed && (now - new Date(data[i][2]).getTime()) > 24 * 60 * 60 * 1000) {
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.reverse().forEach(r => sh.deleteRow(r));
  SpreadsheetApp.getUi().alert('Cleaned up ' + rowsToDelete.length + ' expired OTP record(s).');
}

// One-time setup: creates all sheets with correct headers if they don't exist.
function adminSetupSheets() {
  const spreadsheet = ss_();

  function ensure(name, headers) {
    let sh = spreadsheet.getSheetByName(name);
    if (!sh) sh = spreadsheet.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  ensure(SHEET_VOTERS, ['Name', 'Email', 'House', 'Eligible']);
  ensure(SHEET_VOTES, ['Timestamp', 'Voter Name', 'Email', 'Candidate', 'Vote ID', 'Status']);
  ensure(SHEET_OTP, ['Email', 'OTP Code', 'Created At', 'Expires At', 'Attempts', 'Verified', 'Session Token', 'Session Expires At', 'Last Sent At', 'Consumed']);

  const settingsSh = ensure(SHEET_SETTINGS, ['Key', 'Value']);
  if (settingsSh.getLastRow() === 1) {
    const defaults = [
      ['ElectionStatus', 'CLOSED'],
      ['ElectionTitle', 'DEPUTY HOUSE REPRESENTATIVE ELECTION 2026'],
      ['OtpExpiryMinutes', 5],
      ['HouseLogoUrl', ''],
      ['Candidate1Name', 'Prajna Kapilya'],
      ['Candidate1ImageUrl', ''],
      ['Candidate2Name', 'Suhani Yadav'],
      ['Candidate2ImageUrl', '']
    ];
    defaults.forEach(row => settingsSh.appendRow(row));
  }

  SpreadsheetApp.getUi().alert('Sheets are set up. Next: run "Import Voter List From This Project" or paste your own voter list into the Voters tab, then fill in image URLs in Settings.');
}

// One-time seed of the Voters tab from the official ThunderSoul email list
// (extracted from the uploaded house circular). Safe to re-run — it skips
// anyone already present. Edit or delete rows afterwards directly in the sheet;
// the sheet, not this code, is the source of truth going forward.
function adminSeedVoters() {
  const seed = [
    ['Aarushi Pal', 'aarushi.26@nlumeg.ac.in'],
    ['Abhijeet Kumar Jha', 'abhijeet@nlumeg.ac.in'],
    ['Aditi Arya', 'aditi.a@nlumeg.ac.in'],
    ['Aditya Trivedi', 'adityatrivedi.25@nlumeg.ac.in'],
    ['Aman Pc Joshi', 'aman.26@nlumeg.ac.in'],
    ['Anamika Kumar', 'anamika@nlumeg.ac.in'],
    ['Ananya Bora', 'ananya.b@nlumeg.ac.in'],
    ['Ananya Nandi', 'ananyanandi.25@nlumeg.ac.in'],
    ['Apeksha Pandey', 'apeksha.25@nlumeg.ac.in'],
    ['Aphibakordor Kharkongor', 'aphibakordor.k@nlumeg.ac.in'],
    ['Asher M Marak', 'asher@nlumeg.ac.in'],
    ['Awantika Neeraj', 'awantika.26@nlumeg.ac.in'],
    ['Ayana Chakraborty', 'ayana@nlumeg.ac.in'],
    ['Bhavini Singh Tanwar', 'bhavini.25@nlumeg.ac.in'],
    ['Dawanmi Sumer', 'dawanmi.26@nlumeg.ac.in'],
    ['Drishti Gautam', 'drishti.25@nlumeg.ac.in'],
    ['Emdor Sungoh', 'emdor.s@nlumeg.ac.in'],
    ['Goldstar Lyngdoh', 'goldstar.26@nlumeg.ac.in'],
    ['Jagriti Pandey', 'jagriti@nlumeg.ac.in'],
    ['Janik Sharma', 'janik.26@nlumeg.ac.in'],
    ['Kaushik Kumar Singh', 'kaushik.25@nlumeg.ac.in'],
    ['Kynsai Shaun', 'kynsai.s@nlumeg.ac.in'],
    ['Lakshya Gupta', 'lakshya.26@nlumeg.ac.in'],
    ['Lalwani Palash Rajesh', 'lalwani.p@nlumeg.ac.in'],
    ['Mannan Joshi', 'mannan.25@nlumeg.ac.in'],
    ['Mohammad Tariq', 'mohammad.26@nlumeg.ac.in'],
    ['Mohd Ahsan Siddiqui', 'ahsan@nlumeg.ac.in'],
    ['Naphi Dasien W Laloo', 'naphi.25@nlumeg.ac.in'],
    ['Neha Kumari', 'nehakumari@nlumeg.ac.in'],
    ['Prajna Kapilya', 'prajna.25@nlumeg.ac.in'],
    ['Prasenjeet Pradhan', 'prasenjeet@nlumeg.ac.in'],
    ['Rajdeep Kalita', 'rajdeep@nlumeg.ac.in'],
    ['Ramya Naga Sai Sree Khadarabad', 'ramya.26@nlumeg.ac.in'],
    ['Rishita Tripathy', 'rishita.t.26@nlumeg.ac.in'],
    ['Ritik Raj', 'ritik.25@nlumeg.ac.in'],
    ['Saif Ahmad', 'saif.26@nlumeg.ac.in'],
    ['Shameek Tripathi', 'shameek.26@nlumeg.ac.in'],
    ['Shreyansh Upadhyay', 'shreyansh.u@nlumeg.ac.in'],
    ['Sterlingmore Tyngkan', 'sterlingmore@nlumeg.ac.in'],
    ['Suhani Yadav', 'suhani.25@nlumeg.ac.in'],
    ['Tamanna Sharma', 'tammanna.s@nlumeg.ac.in'],
    ['Tanisha Nath', 'tanisha.n@nlumeg.ac.in'],
    ['Tanu Sheoran', 'tanu@nlumeg.ac.in'],
    ['Vagisha Banerjee', 'vagisha.25@nlumeg.ac.in'],
    ['Wan Sa Lamare', 'wansa.26@nlumeg.ac.in'],
    ['Zach Nongbri', 'zach@nlumeg.ac.in'],
    ['Dangeitskhem Kharsyntiew', 'dangeitskhem.26@nlumeg.ac.in'],
    ['Mardakini Lyngdoh', 'mardakini.26@nlumeg.ac.in']
    // NOTE: "Aditya Kumar" and "Umesh Kumar" were on the requested roster
    // but had no email address in the uploaded source document, so they
    // are NOT included here. Add them manually to the Voters tab with
    // their correct university email once you have it.
  ];

  const sh = sheet_(SHEET_VOTERS);
  const existing = sh.getDataRange().getValues();
  const existingEmails = new Set(existing.slice(1).map(r => normalizeEmail_(r[1])));

  let added = 0;
  seed.forEach(([name, email]) => {
    if (!existingEmails.has(normalizeEmail_(email))) {
      sh.appendRow([name, email, 'ThunderSoul', true]);
      added++;
    }
  });

  SpreadsheetApp.getUi().alert('Imported ' + added + ' new voter(s). ' +
    (seed.length - added) + ' already present. Remember: "Aditya Kumar" and ' +
    '"Umesh Kumar" need to be added manually with their real email addresses.');
}
