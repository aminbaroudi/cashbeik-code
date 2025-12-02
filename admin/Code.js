// ===============================
// File: Code.gs  (Admin App)
// ===============================

// ── Script Properties ──
const ASP = PropertiesService.getScriptProperties();
const USER_DB_ID_KEY    = 'USER_DB_ID';     // MemberAppData spreadsheet ID
const MERCHANT_DB_ID_KEY= 'MERCHANT_DB_ID'; // (optional) Merchant spreadsheet ID
// Master admin email (for self-reset flow)
const MASTER_ADMIN_EMAIL_KEY = 'MASTER_ADMIN_EMAIL';
// Campaign images storage (Drive folder)
// AFTER
const Campaign_IMG_FOLDER_ID_KEY = 'CAMPAIGN_IMG_FOLDER_ID';  // human key, not a folder ID



// Set once from the Config panel (or console):
function setupSetMasterAdminEmail(email){
  const e = String(email||'').trim();
  if (!e || !/@/.test(e)) throw new Error('Provide a valid email address.');
  ASP.setProperty(MASTER_ADMIN_EMAIL_KEY, e);
  return { ok:true, email:e };
}
function getMasterAdminEmail_(){
  return String(ASP.getProperty(MASTER_ADMIN_EMAIL_KEY)||'').trim();
}


// ── Local (admin project) sheets ──
const ADMINS_SHEET = 'Admins';    // [Username, PinHash, CreatedAt]
const A_SESS_SHEET = 'ASessions'; // [SID, Username, LastSeenMs, CreatedAt]

const A_SESSION_TTL_MS = 60 * 1000; // 1 minute (match user/merchant)

// For "active" session counts in visibility
const U_SESSION_TTL_MS = 60 * 1000; // User app sessions TTL (mirror user app)
const M_SESSION_TTL_MS = 60 * 1000; // Merchant app sessions TTL (mirror merchant app)


function doGet(e) {
  // Check for the reset action from the email link
  if (e && e.parameter && e.parameter.action === 'master-reset' && e.parameter.token) {
    
    // Serve a new, separate HTML file for resetting the PIN
    const tpl = HtmlService.createTemplateFromFile('Reset');
    tpl.token = e.parameter.token;         // Pass the token into the template
    tpl.appVersion = getAppVersion_();     // Version for footer and client logs
    
    return tpl.evaluate()
      .setTitle('Admin App - Reset PIN')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // If no action, serve the normal sign in page
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.appVersion = getAppVersion_();       // Version for footer and client logs

  return tpl.evaluate()
    .setTitle('Admin App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ───────────────────────────── Setup (link DBs)
function setupSetUserDbId(id) {
  if (!id) throw new Error('Pass the MemberAppData spreadsheet ID.');

  // 1. Validate ID and open
  let ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Failed to open Spreadsheet. Check the ID and permissions.');
  }

  // 2. Perform the "handshake" - check for a critical sheet
  if (!ss.getSheetByName('Users') || !ss.getSheetByName('Transactions')) {
    throw new Error('This does not appear to be the correct User DB. Sheets "Users" and "Transactions" were not found.');
  }

  // 3. Save property
  ASP.setProperty(USER_DB_ID_KEY, id);
  return { ok:true, id, name: ss.getName() };
}
function setupSetMerchantDbId(id) {
  if (!id) throw new Error('Pass the Merchant spreadsheet ID (the one your Merchant script is bound to).');
  
  // 1. Validate ID and open
  let ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Failed to open Spreadsheet. Check the ID and permissions.');
  }
  
  // 2. Perform the "handshake" - check for the "Staff" sheet
  if (!ss.getSheetByName('Staff')) {
    throw new Error('This does not appear to be the correct Merchant DB. Sheet "Staff" was not found.');
  }
  
  // 3. Save property
  ASP.setProperty(MERCHANT_DB_ID_KEY, id);
  return { ok:true, id, name: ss.getName() };
}

function getConfig() {
  const uid = ASP.getProperty(USER_DB_ID_KEY) || '';
  const mid = ASP.getProperty(MERCHANT_DB_ID_KEY) || '';

  const masterEmail = getMasterAdminEmail_();
  const masterUsername = getMasterAdminUsername_();
  const appVersion = getAppVersion_();

  return {
    userDb: uid ? { id: uid, name: SpreadsheetApp.openById(uid).getName() } : null,
    merchantDb: mid ? { id: mid, name: SpreadsheetApp.openById(mid).getName() } : null,
    masterEmail: masterEmail,
    masterUsername: masterUsername,
    appVersion: appVersion
  };
}

function ensureCampaignImageFolder_() {
  let folderId = ASP.getProperty(Campaign_IMG_FOLDER_ID_KEY) || '';
  try {
    if (folderId) {
      const f = DriveApp.getFolderById(folderId);
      if (f) return f;
    }
  } catch (_) { /* fall through to re-create */ }

  // Create (or re-create) the folder if missing
  const f = DriveApp.createFolder('CampaignImages');
  ASP.setProperty(Campaign_IMG_FOLDER_ID_KEY, f.getId());
  return f;
}

function publicViewUrlForFileId_(fileId) {
  // Works well in <img>: https://drive.google.com/uc?export=view&id=FILE_ID
  return 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(String(fileId||''));
}

function setupSetMasterAdminUsername(username){
  const u = sanitizeUsername_(username);
  if (!u) throw new Error('Provide a valid username.');
  ASP.setProperty('MASTER_ADMIN_USERNAME', u);
  return { ok:true, username:u };
}
function getMasterAdminUsername_(){ return String(ASP.getProperty('MASTER_ADMIN_USERNAME')||'').trim(); }

function adminMasterSendSelfResetLink(sid){
  const caller = ensureAdminActive_(sid);
  if (!isMasterAdminCaller_(caller.username)) throw new Error('Only master admin can request this link.');
  const email = getMasterAdminEmail_();
  if (!email) throw new Error('Master admin email not configured.');

  const token = Utilities.getUuid().replace(/-/g,'');
  const expMs = Date.now() + 60*60*1000;
  const { resetTokens, logs } = getUserDb_();
  resetTokens.appendRow([token, email, expMs, false, new Date().toISOString(), '']);

  const url = ScriptApp.getService().getUrl();
  const link = url + '?action=master-reset&token=' + encodeURIComponent(token);

  MailApp.sendEmail({
    to: email,
    subject: 'Cashbeik Admin — Master PIN reset',
    htmlBody: '<p>Your master PIN reset link (valid 60 minutes):</p><p><a href="'+link+'">'+link+'</a></p>'
  });

  try { logs.appendRow([new Date().toISOString(), 'admin_purge_send_master_link', JSON.stringify({ email })]); } catch(_){}
  return { ok:true };
}

function adminMasterCompleteReset(token, newPin){
  if (!isValidPin_(newPin)) throw new Error('PIN must be exactly 6 digits.');
  const { resetTokens, logs } = getUserDb_();
  const r = findRowByColumn_(resetTokens, 1, token);
  if (!r) throw new Error('Invalid or expired token.');

  const rowVals = resetTokens.getRange(r,1,1,6).getValues()[0];
  const email = String(rowVals[1]||'');
  const expMs = Number(rowVals[2]||0);
  const used  = String(rowVals[3]).toLowerCase()==='true';
  if (used || !expMs || Date.now()>expMs) throw new Error('Invalid or expired token.');

  const masterUsername = getMasterAdminUsername_();
  if (!masterUsername) throw new Error('Master admin username not configured.');

  const ref = getAdminRow_(masterUsername);
  if (!ref) throw new Error('Master admin username not found.');
  const { sheet, row, idx } = ref;

  const rec = makeSaltedAdminRecord_(newPin);
  const updates = {
    'Salt': rec.saltB64,'Hash': rec.hashB64,'Iter': rec.iter,'Algo': PIN_ALGO_DEFAULT,'Peppered': rec.peppered,'UpdatedAt': rec.updatedAt
  };
  Object.keys(updates).forEach(k=>{
    const c=idx[k]; if(c) sheet.getRange(row,c).setValue(updates[k]);
  });

  resetTokens.getRange(r,4).setValue(true);
  resetTokens.getRange(r,6).setValue(Date.now());

  try { logs.appendRow([new Date().toISOString(), 'admin_master_reset_done', JSON.stringify({ email })]); } catch(_){}
  return { ok:true, username: masterUsername };
}


// ───────────────────────────── Admin accounts (local to Admin app)
function adminCreateAdmin(username, pin) {
  username = sanitizeUsername_(username);
  if (!username) throw new Error('Valid username required.');
  if (!isValidPin_(pin)) throw new Error('PIN must be exactly 6 digits.');

  const { admins } = getLocalDb_();
  if (findRowByColumn_(admins, 1, username)) throw new Error('Username already exists.');

  const rec = makeSaltedAdminRecord_(pin);
  const nowIso = new Date().toISOString();
  // Default: Active=true, Role='admin'. Master is defined by MASTER_ADMIN_EMAIL (not by role here).
  admins.appendRow([
    username, '', nowIso,
    rec.saltB64, rec.hashB64, rec.iter, PIN_ALGO_DEFAULT, rec.peppered, rec.updatedAt,
    true, 'admin',
    true // <-- ADD THIS for MustChangePin
  ]);
  return { ok:true, username };
}

function signinAdmin(username, pin) {
  username = sanitizeUsername_(username);
  if (!username) throw new Error('Invalid credentials.');
  if (!isValidPin_(pin)) throw new Error('Invalid credentials.');

  const { admins, asessions } = getLocalDb_();
  const r = findRowByColumn_(admins, 1, username);
  if (!r) throw new Error('Invalid credentials.');

  const ok = verifyAdminPinAndMigrateIfNeeded_(admins, r, pin).ok;
  if (!ok) throw new Error('Invalid credentials.');

  // --- START: Added PIN change check ---
  const idx = adminHeaderIdx_(admins);
  const mustChangeCol = idx['MustChangePin'];
  const mustChange = mustChangeCol ? admins.getRange(r, mustChangeCol).getValue() : false;

  if (String(mustChange).toLowerCase() === 'true') {
    // Do NOT create a full session.
    // Return a special response indicating a forced change is required.
    return { sid: null, username: username, mustChangePin: true };
  }
  // --- END: Added PIN change check ---

  const sid = genSid_();
  const now = Date.now();
  asessions.appendRow([sid, username, now, new Date(now).toISOString()]);
  // Also add mustChangePin:false to the successful response
  return { sid, username, mustChangePin: false };
}

/**
 * Allows a user (sub-admin) to change their own PIN, typically after
 * being forced to by the 'MustChangePin' flag.
 * This verifies their old (temp) PIN and sets the new one.
 */
function adminChangeOwnPin(username, oldPin, newPin) {
  username = sanitizeUsername_(username);
  if (!username) throw new Error('Invalid credentials.');
  if (!isValidPin_(oldPin)) throw new Error('Invalid old PIN.');
  if (!isValidPin_(newPin)) throw new Error('New PIN must be exactly 6 digits.');
  if (oldPin === newPin) throw new Error('New PIN must be different from the old one.');

  const { admins } = getLocalDb_();
  const r = findRowByColumn_(admins, 1, username);
  if (!r) throw new Error('User not found.');

  // Verify the old (temp) PIN first
  const ok = verifyAdminPinAndMigrateIfNeeded_(admins, r, oldPin).ok;
  if (!ok) throw new Error('Invalid old PIN. Credentials do not match.');

  // Now, update to the new PIN and clear the 'MustChangePin' flag
  const rec = makeSaltedAdminRecord_(newPin);
  const idx = adminHeaderIdx_(admins);
  const updates = {
    'Salt': rec.saltB64,'Hash': rec.hashB64,'Iter': rec.iter,'Algo': PIN_ALGO_DEFAULT,'Peppered': rec.peppered,'UpdatedAt': rec.updatedAt,
    'MustChangePin': false // <-- Clear the flag
  };
  Object.keys(updates).forEach(k=>{
    const c = idx[k]; if (c) admins.getRange(r, c).setValue(updates[k]);
  });

  return { ok: true, username: username };
}

function validateASession(sid) {
  if (!sid) return { ok:false };
  const { asessions } = getLocalDb_();
  const r = findRowByColumn_(asessions, 1, sid);
  if (!r) return { ok:false };
  const [, username, lastSeenMs] = asessions.getRange(r,1,1,4).getValues()[0];
  const now = Date.now();
  if (now - Number(lastSeenMs||0) > A_SESSION_TTL_MS) {
    asessions.deleteRow(r);
    return { ok:false, reason:'expired' };
  }
  asessions.getRange(r,3).setValue(now);
  return { ok:true, sid, username };
}

function logoutASession(sid) {
  if (!sid) return { ok:true };
  const { asessions } = getLocalDb_();
  const r = findRowByColumn_(asessions, 1, sid);
  if (r) asessions.deleteRow(r);
  return { ok:true };
}

function requireAdmin_(sid) {
  const v = validateASession(sid);
  if (!v || !v.ok) throw new Error('Admin session required.');
  return v;
}

function getAdminRow_(username){
  const { admins } = getLocalDb_();
  const r = findRowByColumn_(admins, 1, username);
  return r ? { sheet: admins, row: r, idx: adminHeaderIdx_(admins) } : null;
}

function getAdminRecord_(username){
  const ref = getAdminRow_(username);
  if (!ref) return null;
  const { sheet, row, idx } = ref;
  const get = (name)=> { const c=idx[name]; return c? sheet.getRange(row, c).getValue() : ''; };
  return {
    username: String(get('Username')||''),
    active: String(get('Active')||'').toString().toLowerCase() !== 'false',
    role: String(get('Role')||'admin').toLowerCase()
  };
}

// Master is determined by configured email; sub-admins are any Admin rows (role column is informational).
function isMasterAdminCaller_(callerUsername) {
  const masterUsername = getMasterAdminUsername_();
  if (!masterUsername) return false; // Master username must be configured

  const caller = sanitizeUsername_(callerUsername);
  return caller === masterUsername;
}

function ensureAdminActive_(sid){
  const v = requireAdmin_(sid);
  const rec = getAdminRecord_(v.username);
  if (!rec || !rec.active) throw new Error('Admin inactive.');
  return { ...v, role: (rec?rec.role:'admin') };
}

// ───────────────────────────── Admin management (for UI)
function adminListAdmins(sid) {
  requireAdmin_(sid);
  const { admins } = getLocalDb_();
  const last = admins.getLastRow();
  if (last < 2) return [];

  const idx = adminHeaderIdx_(admins);
  const cols = Math.max(admins.getLastColumn() || 12, 12);
  const vals = admins.getRange(2,1,last-1,cols).getValues();

  return vals.map(r => ({ 
    username: String(r[0]||''), 
    createdAt: r[2] || '',
    // Read from the 'Active' column, default to true if blank
    active: String(r[(idx['Active']||10)-1]).toString().toLowerCase() !== 'false' 
  }));
}



function adminSetAdminActive(sid, targetUsername, active) {
  const caller = ensureAdminActive_(sid);
  if (!isMasterAdminCaller_(caller.username)) throw new Error('Only master admin can change admin status.');

  const uname = sanitizeUsername_(targetUsername);
  if (uname === getMasterAdminUsername_()) throw new Error('Cannot change the status of the master admin account.');
  if (uname === caller.username) throw new Error('You cannot change your own status.');

  const { admins } = getLocalDb_();
  const r = findRowByColumn_(admins, 1, uname);
  if (!r) throw new Error('Admin not found.');

  const idx = adminHeaderIdx_(admins);
  const activeCol = idx['Active'] || 10;
  admins.getRange(r, activeCol).setValue(!!active);

  return { ok: true, username: uname, active: !!active };
}

function adminDeleteAdmin(sid, targetUsername) {
  const caller = ensureAdminActive_(sid);
  if (!isMasterAdminCaller_(caller.username)) throw new Error('Only master admin can delete other admins.');

  const uname = sanitizeUsername_(targetUsername);
  if (uname === getMasterAdminUsername_()) throw new Error('The master admin account cannot be deleted.');
  if (uname === caller.username) throw new Error('You cannot delete yourself.');

  const { admins, asessions } = getLocalDb_();
  const r = findRowByColumn_(admins, 1, uname);
  if (!r) throw new Error('Admin not found.');

  // Delete the admin row
  admins.deleteRow(r);

  // Clean up any active sessions for that user
  const sessLast = asessions.getLastRow();
  if (sessLast >= 2) {
    const vals = asessions.getRange(2, 2, sessLast - 1, 1).getValues(); // Usernames column
    const toDel = [];
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).toLowerCase() === uname) {
        toDel.push(2 + i); // row index
      }
    }
    for (let i = toDel.length - 1; i >= 0; i--) {
      asessions.deleteRow(toDel[i]);
    }
  }

  return { ok: true, username: uname };
}

function adminAddAdmin(sid, username, pin) {
  requireAdmin_(sid);
  return adminCreateAdmin(username, pin);
}

// ── Bootstrap: allow creating the very first admin without a session ──
function isAdminProvisioned() {
  const { admins } = getLocalDb_();
  return admins.getLastRow() >= 2; // header + at least one admin row
}

function adminBootstrapFirstAdmin(username, pin) {
  // Only allowed if there are NO admins yet
  if (isAdminProvisioned()) throw new Error('Admins already exist.');
  return adminCreateAdmin(username, pin);
}


function adminResetSubAdminPin(sid, targetUsername, tempPin){
  const caller = ensureAdminActive_(sid);
  if (!isMasterAdminCaller_(caller.username)) throw new Error('Only master admin can reset sub-admin PINs.');
  if (!isValidPin_(tempPin)) throw new Error('PIN must be exactly 6 digits.');
  const uname = sanitizeUsername_(targetUsername);
  if (!uname) throw new Error('Invalid username.');

  const { admins } = getLocalDb_();
  const r = findRowByColumn_(admins, 1, uname);
  if (!r) throw new Error('Admin not found.');

  const v = validateASession(sid);
  if (v && v.ok && v.username === uname) throw new Error('Use master self-reset link for your own PIN.');

  const rec = makeSaltedAdminRecord_(tempPin);
  const idx = adminHeaderIdx_(admins);
  const updates = {
    'Salt': rec.saltB64,'Hash': rec.hashB64,'Iter': rec.iter,'Algo': PIN_ALGO_DEFAULT,'Peppered': rec.peppered,'UpdatedAt': rec.updatedAt,
    'MustChangePin': true // <-- ADD THIS LINE
  };
  Object.keys(updates).forEach(k=>{
    const c = idx[k]; if (c) admins.getRange(r, c).setValue(updates[k]);
  });
  return { ok:true, username: uname };
}


// ───────────────────────────── Admin actions (purge/exports/protect)
function adminPurgeUserSessions(sid, maxAgeMinutes) {
  requireAdmin_(sid);
  const { sessions } = getUserDb_();
  const ageMs = (Number(maxAgeMinutes)>0 ? Number(maxAgeMinutes):120)*60*1000;
  const last = sessions.getLastRow();
  if (last < 2) return { removed:0 };
  const now = Date.now();
  const vals = sessions.getRange(2,1,last-1,4).getValues();
  const toDel = [];
  for (let i=0;i<vals.length;i++){
    const lastSeen = Number(vals[i][2])||0;
    if (!lastSeen || (now-lastSeen)>ageMs) toDel.push(2+i);
  }
  for (let i=toDel.length-1;i>=0;i--) sessions.deleteRow(toDel[i]);
  try { const { logs } = getUserDb_(); logs.appendRow([new Date().toISOString(), 'admin_purge_user_sessions', String({ removed: toDel.length })]); } catch(_){}
  return { removed: toDel.length };
}
function adminPurgeResetTokens(sid, maxAgeMinutes) {
  requireAdmin_(sid);
  const { resetTokens } = getUserDb_();
  const ageMs = (Number(maxAgeMinutes) > 0 ? Number(maxAgeMinutes) : 60) * 60 * 1000;
  const last = resetTokens.getLastRow();
  if (last < 2) return { removed: 0 };
  const now = Date.now();
  const vals = resetTokens.getRange(2,1,last-1,6).getValues(); // Token, Email, ExpMs, Used, CreatedAt, VerifiedAtMs
  const toDel = [];
  for (let i = 0; i < vals.length; i++) {
    const expMs = Number(vals[i][2]) || 0;
    const used  = String(vals[i][3]).toLowerCase() === 'true';
    if ((expMs && (now - expMs) > ageMs) || used) toDel.push(2 + i);
  }
  for (let i = toDel.length - 1; i >= 0; i--) resetTokens.deleteRow(toDel[i]);
  try { const { logs } = getUserDb_(); logs.appendRow([new Date().toISOString(), 'admin_purge_reset_tokens', String({ removed: toDel.length })]); } catch(_){}
  return { removed: toDel.length };
}

function adminPurgeMerchantSessions(sid, maxAgeMinutes) {
  requireAdmin_(sid);
  const { msessions } = getMerchantDb_();
  const ageMs = (Number(maxAgeMinutes)>0 ? Number(maxAgeMinutes):120)*60*1000;
  const last = msessions.getLastRow();
  if (last < 2) return { removed:0 };
  const now = Date.now();
  const vals = msessions.getRange(2,1,last-1,4).getValues();
  const toDel = [];
  for (let i=0;i<vals.length;i++){
    const lastSeen = Number(vals[i][2])||0;
    if (!lastSeen || (now-lastSeen)>ageMs) toDel.push(2+i);
  }
  for (let i=toDel.length-1;i>=0;i--) msessions.deleteRow(toDel[i]);
  try { const { logs } = getUserDb_(); logs.appendRow([new Date().toISOString(), 'admin_purge_merchant_sessions', String({ removed: toDel.length })]); } catch(_){}
  return { removed: toDel.length };
}

function adminPurgeLinkTokens(sid, maxAgeMinutes) {
  requireAdmin_(sid);
  const { linkTokens } = getUserDb_();
  const ageMs = (Number(maxAgeMinutes)>0 ? Number(maxAgeMinutes):60)*60*1000;
  const last = linkTokens.getLastRow();
  if (last < 2) return { removed:0 };
  const now = Date.now();
  const vals = linkTokens.getRange(2,1,last-1,6).getValues();
  const toDel = [];
  for (let i=0;i<vals.length;i++){
    const expMs = Number(vals[i][3])||0;
    if (!expMs || expMs < (now - ageMs)) toDel.push(2+i);
  }
  for (let i=toDel.length-1;i>=0;i--) linkTokens.deleteRow(toDel[i]);
  try { const { logs } = getUserDb_(); logs.appendRow([new Date().toISOString(), 'admin_purge_link_tokens', String({ removed: toDel.length })]); } catch(_){}
  return { removed: toDel.length };
}

function exportBalancesCsv(sid) {
  requireAdmin_(sid);
  const { balances } = getUserDb_();
  const last = balances.getLastRow();
  let rows = [['MemberId','Points']];
  if (last >= 2) rows = rows.concat(balances.getRange(2,1,last-1,2).getValues());
  const csv = rows.map(r => r.map(csvEscape_).join(',')).join('\n');
  return { filename:'balances.csv', csv };
}
function exportTransactionsCsv(sid) {
  requireAdmin_(sid);
  const { tx } = getUserDb_();
  const last = tx.getLastRow();
  const cols = Math.max(7, tx.getLastColumn()||7);
  let rows = [['TxId','MemberId','MerchantId','Type','Points','AtMs','Staff']];
  if (last >= 2) rows = rows.concat(tx.getRange(2,1,last-1,cols).getValues());
  const csv = rows.map(r => r.map(csvEscape_).join(',')).join('\n');
  return { filename:'transactions.csv', csv };
}

function listLogs(sid, limit) {
  requireAdmin_(sid);
  const { logs } = getUserDb_();
  const last = logs.getLastRow();
  const n = Math.max(1, Math.min(Number(limit)||100, 1000));
  if (last < 2) return [];
  const startRow = Math.max(2, last - n + 1);
  const vals = logs.getRange(startRow,1,(last-startRow+1),3).getValues();
  return vals.map(r => ({ at:r[0], type:r[1], msg:r[2] })).reverse(); // newest first
}

function adminProtectHeaders(sid, warningOnly) {
  requireAdmin_(sid);
  const wOnly = String(warningOnly).toLowerCase() === 'true' || warningOnly === true;

  // Protect headers in User DB
  const { ss: userSS } = getUserDb_();
  protectAllSheetHeaders_(userSS, wOnly);

  // If Merchant DB linked, protect headers there too
  try {
    const m = getMerchantDb_();
    protectAllSheetHeaders_(m.ss, wOnly);
  } catch(_) {}

  return { ok:true };
}

// List members with campaign/comms fields + simple filters
function adminListMembers(sid, filters) {
  requireAdmin_(sid);
  const { ss } = getUserDb_();
  const users = ss.getSheetByName('Users');
  if (!users) return [];

  const last = users.getLastRow();
  if (last < 2) return [];

  const cols = Math.max(users.getLastColumn()||0, 21);
  const hdr = users.getRange(1,1,1,cols).getValues()[0].map(String);
  const i = (name)=> hdr.findIndex(h => String(h).toLowerCase() === name.toLowerCase());

  const idx = {
    Email: i('Email'),
    MemberId: i('MemberId'),
    First: i('First'),
    Last: i('Last'),
    CreatedAt: i('CreatedAt'),
    PrefComms: i('PrefComms'),
    PhoneE164: i('PhoneE164'),
    WaE164: i('WaE164'),
    City: i('City'),
    Country: i('Country'),
    Lang: i('Lang'),
    Referral: i('Referral'),
    OptMarketing: i('OptMarketing'),
    UTM_Source: i('UTM_Source'),
    UTM_Medium: i('UTM_Medium'),
    UTM_Campaign: i('UTM_Campaign'),
    UTM_Term: i('UTM_Term'),
    UTM_Content: i('UTM_Content'),
  };

  const vals = users.getRange(2,1,last-1,cols).getValues();

  const f = filters || {};
  const like = (a,b)=> !b || (String(a||'').toLowerCase().includes(String(b||'').toLowerCase()));
  const eq   = (a,b)=> !b || (String(a||'').toLowerCase() === String(b||'').toLowerCase());

  const yesNoAny = (v, sel) => {
    if (!sel || sel === 'any') return true;
    const truthy = String(v).toLowerCase() === 'true';
    return (sel === 'yes') ? truthy : !truthy;
  };

  const out = [];
  for (let r of vals) {
    const row = {
      email: r[idx.Email],
      memberId: r[idx.MemberId],
      first: r[idx.First],
      last: r[idx.Last],
      createdAt: r[idx.CreatedAt],
      prefComms: r[idx.PrefComms],
      phone: r[idx.PhoneE164],
      whatsapp: r[idx.WaE164],
      city: r[idx.City],
      country: r[idx.Country],
      lang: r[idx.Lang],
      referral: r[idx.Referral],
      optMarketing: r[idx.OptMarketing],
      utm_source: r[idx.UTM_Source],
      utm_medium: r[idx.UTM_Medium],
      utm_campaign: r[idx.UTM_Campaign],
      utm_term: r[idx.UTM_Term],
      utm_content: r[idx.UTM_Content],
    };

    // Filters
    if (!like(row.email, f.q) && !like(row.first, f.q) && !like(row.last, f.q) && !like(row.memberId, f.q)) continue;
    if (!eq(row.prefComms, f.prefComms)) continue;
    if (!like(row.utm_source, f.utm_source)) continue;
    if (!like(row.utm_medium, f.utm_medium)) continue;
    if (!like(row.utm_campaign, f.utm_campaign)) continue;
    if (!yesNoAny(row.optMarketing, f.optMarketing)) continue;
    // NEW: City/Country filters (contains match)
    if (!like(row.city, f.city)) continue;
    if (!like(row.country, f.country)) continue;


    out.push(row);
    if (out.length >= 1000) break; // cap
  }
  return out;
}

// Small summary for comms/utm breakdowns
function adminMembersSummary(sid) {
  requireAdmin_(sid);
  const rows = adminListMembers(sid, {}); // up to 1000
  const sum = {
    total: rows.length,
    pref: {}, // email/sms/whatsapp
    optMarketing: { yes:0, no:0 },
    utm_campaign: {}
  };
  for (const r of rows) {
    const p = String(r.prefComms||'').toLowerCase() || '—';
    sum.pref[p] = (sum.pref[p]||0) + 1;
    if (String(r.optMarketing).toLowerCase() === 'true') sum.optMarketing.yes++; else sum.optMarketing.no++;
    const c = r.utm_campaign || '—';
    sum.utm_campaign[c] = (sum.utm_campaign[c]||0) + 1;
  }
  return sum;
}

// ───────────────────────────── NEW: Create Merchant / Create Staff
function adminCreateMerchant(sid, name) {
  requireAdmin_(sid);
  name = String(name || '').trim();
  if (!name) throw new Error('Merchant name required.');
  const { merchants } = getUserDb_();

  const merchantId = 'MRC-' + Math.random().toString(36).slice(2, 8).toUpperCase();
  const secret = Utilities.getUuid().replace(/-/g,'').slice(0,24);

  if (findRowByColumn_(merchants, 1, merchantId)) {
    throw new Error('Collision on MerchantId. Try again.');
  }
  merchants.appendRow([merchantId, name, true, secret, new Date().toISOString()]);
  return { ok:true, merchantId, secret, name };
}

function adminCreateStaff(sid, merchantId, username, pin, role) {
  requireAdmin_(sid);
  merchantId = String(merchantId || '').trim().toUpperCase();
  username = sanitizeUsername_(username);
  role = String(role || 'staff').toLowerCase();

  if (!merchantId) throw new Error('merchantId required.');
  if (!username) throw new Error('valid username required.');
  if (!isValidPin_(pin)) throw new Error('PIN must be exactly 6 digits.');
  if (role !== 'staff' && role !== 'manager') throw new Error('role must be staff or manager.');

  const { merchants } = getUserDb_();
  const mr = findRowByColumn_(merchants, 1, merchantId);
  if (!mr) throw new Error('Merchant not found in User DB.');
  const active = String(merchants.getRange(mr, 3).getValue() || 'TRUE').toUpperCase() !== 'FALSE';
  if (!active) throw new Error('Merchant is inactive.');

  const { staff } = getMerchantDb_();
  if (findRowByColumn_(staff, 1, username)) throw new Error('Username already exists in Staff.');

  const rec = makeSaltedAdminRecord_(pin); // same format; reuse helpers
  staff.appendRow([
    username, '', merchantId, role, true, new Date().toISOString(),
    rec.saltB64, rec.hashB64, rec.iter, PIN_ALGO_DEFAULT, rec.peppered, rec.updatedAt,
    true // <-- ADD THIS for MustChangePin
  ]);

  return { ok:true, username, merchantId, role };
}

function adminSetMerchantActive(sid, targetMerchantId, active) {
  ensureAdminActive_(sid); // Any admin can do this

  const { merchants } = getUserDb_();
  const r = findRowByColumn_(merchants, 1, targetMerchantId);
  if (!r) throw new Error('Merchant not found.');

  merchants.getRange(r, 3).setValue(!!active); // Column 3 is 'Active'
  return { ok: true, merchantId: targetMerchantId, active: !!active };
}

function adminDeleteMerchant(sid, targetMerchantId) {
  ensureAdminActive_(sid);

  const { merchants } = getUserDb_();
  const r = findRowByColumn_(merchants, 1, targetMerchantId);
  if (!r) throw new Error('Merchant not found.');

  merchants.deleteRow(r);
  // WARNING: This does NOT delete associated staff or transactions.
  // Deactivating is strongly recommended over deleting.
  return { ok: true, merchantId: targetMerchantId };
}

function adminDeleteStaff(sid, targetUsername) {
  ensureAdminActive_(sid);

  const uname = sanitizeUsername_(targetUsername);
  if (!uname) throw new Error('Invalid username.');

  const { staff, msessions } = getMerchantDb_();

  const r = findRowByColumn_(staff, 1, uname);
  if (!r) throw new Error('Staff not found.');

  staff.deleteRow(r);

  // Clean up merchant sessions
  const sessLast = msessions.getLastRow();
  if (sessLast >= 2) {
    const vals = msessions.getRange(2, 2, sessLast - 1, 1).getValues();
    const toDel = [];
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).toLowerCase() === uname) {
        toDel.push(2 + i);
      }
    }
    for (let i = toDel.length - 1; i >= 0; i--) {
      msessions.deleteRow(toDel[i]);
    }
  }
  return { ok: true, username: uname };
}

function adminListMerchants(sid) {
  ensureAdminActive_(sid);
  const { merchants } = getUserDb_();
  const last = merchants.getLastRow();
  if (last < 2) return [];
  const cols = Math.max(5, merchants.getLastColumn()||5);
  const vals = merchants.getRange(2,1,last-1,cols).getValues();
  return vals.map(r => ({
    merchantId: String(r[0]||''),
    name: String(r[1]||''),
    active: String(r[2]).toString().toLowerCase() !== 'false',
    createdAt: r[4] || ''
  }));
}

function adminListStaff(sid, opts) {
  const caller = ensureAdminActive_(sid);
  const { staff } = getMerchantDb_();
  const last = staff.getLastRow();
  if (last < 2) return [];
  const cols = Math.max(13, staff.getLastColumn()||13);
  const vals = staff.getRange(2,1,last-1,cols).getValues();
  return vals.map(r => ({
    username: String(r[0]||''),
    merchantId: String(r[2]||''),
    role: String(r[3]||''),
    active: String(r[4]).toString().toLowerCase() !== 'false',
    createdAt: r[5] || '',
    mustChange: String(r[12]).toString().toLowerCase()==='true'
  }));
}

function adminSetStaffActive(sid, username, active) {
  const caller = ensureAdminActive_(sid);
  const { staff } = getMerchantDb_();
  const r = findRowByColumn_(staff, 1, sanitizeUsername_(username));
  if (!r) throw new Error('Staff not found.');
  const idx = adminHeaderIdx_(staff);
  staff.getRange(r, idx['Active']||5).setValue(!!active);
  return { ok:true };
}

function adminSetStaffRole(sid, username, role) {
  const caller = ensureAdminActive_(sid);
  role = String(role||'').toLowerCase();
  if (role !== 'staff' && role !== 'manager') throw new Error('role must be staff or manager.');
  const { staff } = getMerchantDb_();
  const r = findRowByColumn_(staff, 1, sanitizeUsername_(username));
  if (!r) throw new Error('Staff not found.');
  const idx = adminHeaderIdx_(staff);
  staff.getRange(r, idx['Role']||4).setValue(role);
  return { ok:true };
}

function adminResetStaffPin(sid, username, tempPin){
  const caller = ensureAdminActive_(sid);
  if (!isValidPin_(tempPin)) throw new Error('PIN must be exactly 6 digits.');
  const uname = sanitizeUsername_(username);
  if (!uname) throw new Error('Invalid username.');

  const { staff } = getMerchantDb_();
  const r = findRowByColumn_(staff, 1, uname);
  if (!r) throw new Error('Staff not found.');

  const rec = makeSaltedAdminRecord_(tempPin);
  const idx = adminHeaderIdx_(staff);
  const updates = {
    'Salt': rec.saltB64,
    'Hash': rec.hashB64,
    'Iter': rec.iter,
    'Algo': PIN_ALGO_DEFAULT,
    'Peppered': rec.peppered,
    'UpdatedAt': rec.updatedAt,
    'MustChangePin': true
  };
  Object.keys(updates).forEach(k=>{
    const c = idx[k]; if (c) staff.getRange(r, c).setValue(updates[k]);
  });
  return { ok:true, username: uname };
}

// ───────────────────────────── Campaigns / Campaigns (merchant-wide multipliers) ─────────────────────────────

function _genCampaignId_(){ return 'CPN-' + Math.random().toString(36).slice(2, 8).toUpperCase(); }
function _genRequestId_(){ return 'REQ-' + Math.random().toString(36).slice(2, 8).toUpperCase(); }
function _sanitizeReqType_(t){
  const v = String(t||'new').toLowerCase();
  return (['new','edit','renew','deactivate'].includes(v)) ? v : 'new';
}
function _sanitizeStatus_(s){
  const v = String(s||'pending').toLowerCase();
  return (['pending','approved','rejected','cancelled'].includes(v)) ? v : 'pending';
}

function _nowIso_(){ return new Date().toISOString(); }

function _assertMerchantExistsActive_(merchantId){
  const { merchants } = getUserDb_();
  const r = findRowByColumn_(merchants, 1, merchantId);
  if (!r) throw new Error('Merchant not found: ' + merchantId);
  const active = String(merchants.getRange(r, 3).getValue() || 'TRUE').toUpperCase() !== 'FALSE';
  if (!active) throw new Error('Merchant is inactive: ' + merchantId);
}

function _sanitizeBillingModel_(m){
  const v = String(m||'per_redemption').toLowerCase();
  return (v==='flat_monthly'||v==='per_redemption') ? v : 'per_redemption';
}

function adminCreateCampaign(sid, payload){
  ensureAdminActive_(sid);
  const p = payload || {};
  const merchantId = String(p.merchantId||'').trim().toUpperCase();
  const title      = String(p.title||'').trim();
  const type       = 'multiplier'; // future-proof (we only support merchant-wide multipliers now)
  const multiplier = Math.max(1, Number(p.multiplier||1));
  const startIso   = String(p.startIso||'').trim();
  const endIso     = String(p.endIso||'').trim();
  const minSpend   = Math.max(0, Number(p.minSpend||0));
  const maxRed     = Math.max(0, Number(p.maxRedemptions||0)); // 0 = unlimited
  const budgetCap  = Math.max(0, Number(p.budgetCap||0));      // 0 = no cap
  const maxPerCustomer= Math.max(0, Number(p.maxPerCustomer||0));
  const billModel  = _sanitizeBillingModel_(p.billingModel);
  const cpr        = Math.max(0, Number(p.costPerRedemption||0));
  const active     = String(p.active).toLowerCase()==='false' ? false : true;

  if (!merchantId) throw new Error('merchantId required.');
  if (!title) throw new Error('title required.');
  if (!startIso || !endIso) throw new Error('startIso and endIso required.');
  if (isNaN(Date.parse(startIso)) || isNaN(Date.parse(endIso))) throw new Error('Invalid date(s).');

  _assertMerchantExistsActive_(merchantId);

  const { campaigns } = getUserDb_();
  const id = _genCampaignId_();
  campaigns.appendRow([
    id, merchantId, title, type, multiplier,
    startIso, endIso, minSpend, maxRed, maxPerCustomer, budgetCap,
    billModel, cpr, !!active, _nowIso_(), _nowIso_()
  ]);
  return { ok:true, campaignId:id };
}

function adminListCampaigns(sid, opts){
  ensureAdminActive_(sid);
  const { campaigns } = getUserDb_();
  const last = campaigns.getLastRow();
  if (last < 2) return [];

  const cols = campaigns.getLastColumn() || 1;
  // const hdr = campaigns.getRange(1,1,1,cols).getValues()[0].map(h=>String(h||'').trim());
  // const h = (name) => hdr.findIndex(x => x.toLowerCase()===String(name||'').toLowerCase());
  const headers = adminHeaderIdx_(campaigns); // <--- Use the standard index helper
  const getIndex = (name) => headers[name] ? headers[name] - 1 : -1;

  const i = {
    campaignId: getIndex('CampaignId'),
    merchantId: getIndex('MerchantId'),
    title: getIndex('Title'),
    type: getIndex('Type'), // <--- PATCHED
    multiplier: getIndex('Multiplier'), // <--- PATCHED
    startIso: getIndex('StartIso'), // <--- PATCHED
    endIso: getIndex('EndIso'), // <--- PATCHED
    minSpend: getIndex('MinSpend'), // <--- PATCHED
    maxRedemptions: getIndex('MaxRedemptions'), // <--- PATCHED
    maxPerCustomer: getIndex('MaxPerCustomer'),
    budgetCap: getIndex('BudgetCap'), // <--- PATCHED
    billingModel: getIndex('BillingModel'), // <--- PATCHED
    costPerRedemption: getIndex('CostPerRedemption'), // <--- PATCHED
    active: getIndex('Active'), // <--- PATCHED
    createdAt: getIndex('CreatedAt'), // <--- PATCHED
    updatedAt: getIndex('UpdatedAt'), // <--- PATCHED
    imageUrl: getIndex('ImageUrl') // <--- PATCHED
  };

  const vals = campaigns.getRange(2,1,last-1,cols).getValues();
  const f = opts||{};
  const mid = String(f.merchantId||'').trim().toUpperCase();
  const q   = String(f.q||'').toLowerCase();

  const out = [];
  for (let r of vals){
    const row = {
      campaignId: String(i.campaignId>=0 ? r[i.campaignId] : ''),
      merchantId: String(i.merchantId>=0 ? r[i.merchantId] : ''),
      title:      String(i.title>=0 ? r[i.title] : ''),
      type:       String(i.type>=0 ? r[i.type] : ''),
      multiplier: Number(i.multiplier>=0 ? r[i.multiplier] : 0),
      startIso:   String(i.startIso>=0 ? r[i.startIso] : ''),
      endIso:     String(i.endIso>=0 ? r[i.endIso] : ''),
      minSpend:   Number(i.minSpend>=0 ? r[i.minSpend] : 0),
      maxRedemptions: Number(i.maxRedemptions>=0 ? r[i.maxRedemptions] : 0),
      maxPerCustomer: Number(i.maxPerCustomer>=0 ? r[i.maxPerCustomer] : 0),  
      budgetCap:  Number(i.budgetCap>=0 ? r[i.budgetCap] : 0),
      billingModel: String(i.billingModel>=0 ? r[i.billingModel] : 'per_redemption'),
      costPerRedemption: Number(i.costPerRedemption>=0 ? r[i.costPerRedemption] : 0),
      active:     String(i.active>=0 ? r[i.active] : 'TRUE').toLowerCase()!=='false',
      createdAt:  i.createdAt>=0 ? r[i.createdAt] : '',
      updatedAt:  i.updatedAt>=0 ? r[i.updatedAt] : '',
      imageUrl:   String(i.imageUrl>=0 ? (r[i.imageUrl]||'') : '')
    };
    if (mid && row.merchantId !== mid) continue;
    if (q && !(row.title.toLowerCase().includes(q) || row.campaignId.toLowerCase().includes(q))) continue;
    out.push(row);
  }
  out.sort((a,b)=> (Date.parse(b.updatedAt||b.createdAt||0) || 0) - (Date.parse(a.updatedAt||a.createdAt||0) || 0));
  return out;
}

function adminListCampaignRequests(sid, opts){
  ensureAdminActive_(sid);
  const { campaignRequests } = getUserDb_();
  const last = campaignRequests.getLastRow();
  if (last < 2) return [];

  const cols = Math.max(23, campaignRequests.getLastColumn()||23);
  const hdr = campaignRequests.getRange(1,1,1,cols).getValues()[0].map(String);
  const h = name => hdr.findIndex(x => x.toLowerCase() === String(name||'').toLowerCase());

  const i = {
    requestId: h('RequestId'),
    merchantId: h('MerchantId'),
    requestType: h('RequestType'),
    title: h('Title'),
    multiplier: h('Multiplier'),
    startIso: h('StartIso'),
    endIso: h('EndIso'),
    minSpend: h('MinSpend'),
    maxRedemptions: h('MaxRedemptions'),
    maxPerCustomer: h('MaxPerCustomer'),
    budgetCap: h('BudgetCap'),
    billingModel: h('BillingModel'),
    costPerRedemption: h('CostPerRedemption'),
    notes: h('Notes'),
    imageUrl: h('ImageUrl'),
    status: h('Status'),
    createdBy: h('CreatedBy'),
    createdAt: h('CreatedAt'),
    updatedAt: h('UpdatedAt'),
    decisionBy: h('DecisionBy'),
    decisionAt: h('DecisionAt'),
    decisionNotes: h('DecisionNotes'),
    linkedCampaignId: h('LinkedCampaignId')
  };

  const vals = campaignRequests.getRange(2,1,last-1,cols).getValues();
  const f = opts||{};
  const mid = String(f.merchantId||'').trim().toUpperCase();
  const q   = String(f.q||'').toLowerCase();
  const status = String(f.status||'').toLowerCase();
  const rtype  = String(f.requestType||'').toLowerCase();

  const out = [];
  for (const r of vals){
    const row = {
      requestId: String(i.requestId>=0 ? r[i.requestId] : ''),
      merchantId: String(i.merchantId>=0 ? r[i.merchantId] : ''),
      requestType: String(i.requestType>=0 ? r[i.requestType] : ''),
      title: String(i.title>=0 ? r[i.title] : ''),
      multiplier: Number(i.multiplier>=0 ? r[i.multiplier] : 0),
      startIso: String(i.startIso>=0 ? r[i.startIso] : ''),
      endIso: String(i.endIso>=0 ? r[i.endIso] : ''),
      minSpend: Number(i.minSpend>=0 ? r[i.minSpend] : 0),
      maxRedemptions: Number(i.maxRedemptions>=0 ? r[i.maxRedemptions] : 0),
      maxPerCustomer: Number(i.maxPerCustomer>=0 ? r[i.maxPerCustomer] : 0),
      budgetCap: Number(i.budgetCap>=0 ? r[i.budgetCap] : 0),
      billingModel: String(i.billingModel>=0 ? r[i.billingModel] : ''),
      costPerRedemption: Number(i.costPerRedemption>=0 ? r[i.costPerRedemption] : 0),
      notes: String(i.notes>=0 ? r[i.notes] : ''),
      imageUrl: String(i.imageUrl>=0 ? r[i.imageUrl] : ''),
      status: String(i.status>=0 ? r[i.status] : ''),
      createdBy: String(i.createdBy>=0 ? r[i.createdBy] : ''),
      createdAt: i.createdAt>=0 ? r[i.createdAt] : '',
      updatedAt: i.updatedAt>=0 ? r[i.updatedAt] : '',
      decisionBy: String(i.decisionBy>=0 ? r[i.decisionBy] : ''),
      decisionAt: i.decisionAt>=0 ? r[i.decisionAt] : '',
      decisionNotes: String(i.decisionNotes>=0 ? r[i.decisionNotes] : ''),
      linkedCampaignId: String(i.linkedCampaignId>=0 ? r[i.linkedCampaignId] : '')
    };

    if (mid && row.merchantId !== mid) continue;
    if (status && row.status.toLowerCase() !== status) continue;
    if (rtype && row.requestType.toLowerCase() !== rtype) continue;
    if (q && !(row.title.toLowerCase().includes(q) || row.requestId.toLowerCase().includes(q))) continue;

    out.push(row);
  }

  // newest first by UpdatedAt/CreatedAt
  out.sort((a,b)=> (Date.parse(b.updatedAt||b.createdAt||0)||0) - (Date.parse(a.updatedAt||a.createdAt||0)||0));
  return out;
}

function adminApproveCampaignRequest(sid, requestId){
  const caller = ensureAdminActive_(sid);
  const { campaignRequests, campaigns, logs } = getUserDb_();

  const r = findRowByColumn_(campaignRequests, 1, requestId);
  if (!r) throw new Error('Request not found.');
  const idx = adminHeaderIdx_(campaignRequests);

  const status = String(campaignRequests.getRange(r, idx['Status']).getValue()||'').toLowerCase();
  if (status !== 'pending') throw new Error('Only pending requests can be approved.');

  const get = (k)=> campaignRequests.getRange(r, idx[k]).getValue();

  const payload = {
    merchantId: String(get('MerchantId')||'').trim().toUpperCase(),
    title: String(get('Title')||'').trim(),
    multiplier: Number(get('Multiplier')||1),
    startIso: String(get('StartIso')||'').trim(),
    endIso: String(get('EndIso')||'').trim(),
    minSpend: Number(get('MinSpend')||0),
    maxRedemptions: Number(get('MaxRedemptions')||0),
    maxPerCustomer: Number(get('MaxPerCustomer')||0),
    budgetCap: Number(get('BudgetCap')||0),
    billingModel: _sanitizeBillingModel_(get('BillingModel')),
    costPerRedemption: Number(get('CostPerRedemption')||0),
    notes: String(get('Notes')||'').trim(),
    imageUrl: String(get('ImageUrl')||'').trim()
  };
  const reqType = _sanitizeReqType_(get('RequestType'));
  const nowIso = _nowIso_();

  let linkedCampaignId = String(get('LinkedCampaignId')||'').trim();
  if (reqType === 'deactivate'){
    linkedCampaignId = linkedCampaignId || String(get('LinkedCampaignId')||'').trim();
    if (!linkedCampaignId) throw new Error('Deactivate request missing LinkedCampaignId.');
    adminDeactivateOrDeleteCampaign(sid, linkedCampaignId);
  } else if (reqType === 'edit'){
    // Update an existing campaign
    linkedCampaignId = linkedCampaignId || String(get('LinkedCampaignId')||'').trim();
    if (!linkedCampaignId) throw new Error('Edit request missing LinkedCampaignId.');
    adminUpdateCampaign(sid, linkedCampaignId, {
      title: payload.title,
      multiplier: payload.multiplier,
      startIso: payload.startIso,
      endIso: payload.endIso,
      minSpend: payload.minSpend,
      maxRedemptions: payload.maxRedemptions,
      maxPerCustomer: payload.maxPerCustomer,
      budgetCap: payload.budgetCap,
      billingModel: payload.billingModel,
      costPerRedemption: payload.costPerRedemption
    });
    // optional image copy if provided
    if (payload.imageUrl){
      const cIdx = adminHeaderIdx_(campaigns);
      const cr = findRowByColumn_(campaigns, 1, linkedCampaignId);
      if (cr && cIdx['ImageUrl']) campaigns.getRange(cr, cIdx['ImageUrl']).setValue(payload.imageUrl);
      if (cr && cIdx['UpdatedAt']) campaigns.getRange(cr, cIdx['UpdatedAt']).setValue(nowIso);
    }
  } else {
    // 'new' or 'renew' => create a fresh campaign
    const created = adminCreateCampaign(sid, {
      merchantId: payload.merchantId,
      title: payload.title,
      multiplier: payload.multiplier,
      startIso: payload.startIso,
      endIso: payload.endIso,
      minSpend: payload.minSpend,
      maxRedemptions: payload.maxRedemptions,
      maxPerCustomer: payload.maxPerCustomer,
      budgetCap: payload.budgetCap,
      billingModel: payload.billingModel,
      costPerRedemption: payload.costPerRedemption,
      active: true
    });
    linkedCampaignId = created.campaignId;

    // copy image if provided
    if (payload.imageUrl && linkedCampaignId){
      const cIdx = adminHeaderIdx_(campaigns);
      const cr = findRowByColumn_(campaigns, 1, linkedCampaignId);
      if (cr && cIdx['ImageUrl']) campaigns.getRange(cr, cIdx['ImageUrl']).setValue(payload.imageUrl);
      if (cr && cIdx['UpdatedAt']) campaigns.getRange(cr, cIdx['UpdatedAt']).setValue(nowIso);
    }
  }

  // mark request Approved
  campaignRequests.getRange(r, idx['Status']).setValue('approved');
  campaignRequests.getRange(r, idx['DecisionBy']).setValue(caller.username);
  campaignRequests.getRange(r, idx['DecisionAt']).setValue(nowIso);
  if (idx['LinkedCampaignId']) campaignRequests.getRange(r, idx['LinkedCampaignId']).setValue(linkedCampaignId||'');
  if (idx['UpdatedAt']) campaignRequests.getRange(r, idx['UpdatedAt']).setValue(nowIso);

  try { logs.appendRow([nowIso, 'admin_request_approve', JSON.stringify({ requestId, linkedCampaignId })]); } catch(_){}
  return { ok:true, requestId, linkedCampaignId };
}

function adminRejectCampaignRequest(sid, requestId, reason){
  const caller = ensureAdminActive_(sid);
  const { campaignRequests, logs } = getUserDb_();

  const r = findRowByColumn_(campaignRequests, 1, requestId);
  if (!r) throw new Error('Request not found.');
  const idx = adminHeaderIdx_(campaignRequests);

  const status = String(campaignRequests.getRange(r, idx['Status']).getValue()||'').toLowerCase();
  if (status !== 'pending') throw new Error('Only pending requests can be rejected.');

  const nowIso = _nowIso_();
  campaignRequests.getRange(r, idx['Status']).setValue('rejected');
  campaignRequests.getRange(r, idx['DecisionBy']).setValue(caller.username);
  campaignRequests.getRange(r, idx['DecisionAt']).setValue(nowIso);
  if (idx['DecisionNotes']) campaignRequests.getRange(r, idx['DecisionNotes']).setValue(String(reason||''));
  if (idx['UpdatedAt']) campaignRequests.getRange(r, idx['UpdatedAt']).setValue(nowIso);

  try { logs.appendRow([nowIso, 'admin_request_reject', JSON.stringify({ requestId, reason:String(reason||'') })]); } catch(_){}
  return { ok:true, requestId };
}

function _isCampaignCurrentlyActive_(campaignsSheet, row){
  const idx = adminHeaderIdx_(campaignsSheet);
  const active = String(campaignsSheet.getRange(row, _colOrThrow_(idx,'Active')).getValue()||'TRUE')
                  .toLowerCase() !== 'false';
  const startIso = String(campaignsSheet.getRange(row, _colOrThrow_(idx,'StartIso')).getValue()||'');
  const endIso   = String(campaignsSheet.getRange(row, _colOrThrow_(idx,'EndIso')).getValue()||'');
  const now = Date.now();
  const s = startIso ? Date.parse(startIso) : 0;
  const e = endIso   ? Date.parse(endIso)   : 0;
  const inWindow = (!s || now>=s) && (!e || now<=e);
  return active && inWindow;
}

function adminSetCampaignActive(sid, campaignId, active){
  ensureAdminActive_(sid);
  const { campaigns } = getUserDb_();
  const r = findRowByColumn_(campaigns, 1, campaignId);
  if (!r) throw new Error('Campaign not found.');
  const idx = adminHeaderIdx_(campaigns);
  const cActive = idx['Active']||13;
  const cUpdated= idx['UpdatedAt']||15;
  campaigns.getRange(r,cActive).setValue(!!active);
  campaigns.getRange(r,cUpdated).setValue(_nowIso_());
  return { ok:true, campaignId, active:!!active };
}

function adminDeleteCampaign(sid, campaignId){
  ensureAdminActive_(sid);
  const { campaigns } = getUserDb_();
  const r = findRowByColumn_(campaigns, 1, campaignId);
  if (!r) throw new Error('Campaign not found.');
  campaigns.deleteRow(r);
  return { ok:true, campaignId };
}

function adminDeactivateOrDeleteCampaign(sid, campaignId){
  ensureAdminActive_(sid);
  const { campaigns } = getUserDb_();
  const r = findRowByColumn_(campaigns, 1, campaignId);
  if (!r) throw new Error('Campaign not found.');
  const live = _isCampaignCurrentlyActive_(campaigns, r);

  if (live){
    return adminSetCampaignActive(sid, campaignId, false); // deactivate if currently live
  } else {
    return adminDeleteCampaign(sid, campaignId);           // otherwise delete the stale row
  }
}

function adminUpdateCampaign(sid, campaignId, fields){
  ensureAdminActive_(sid);
  const { campaigns } = getUserDb_();
  const r = findRowByColumn_(campaigns, 1, campaignId);
  if (!r) throw new Error('Campaign not found.');
  const idx = adminHeaderIdx_(campaigns);

  const map = {};
  if ('title' in fields)        map['Title'] = String(fields.title||'').trim();
  if ('multiplier' in fields)   map['Multiplier'] = Math.max(1, Number(fields.multiplier||1));
  if ('startIso' in fields)     map['StartIso'] = String(fields.startIso||'').trim();
  if ('endIso' in fields)       map['EndIso']   = String(fields.endIso||'').trim();
  if ('minSpend' in fields)     map['MinSpend'] = Math.max(0, Number(fields.minSpend||0));
  if ('maxRedemptions' in fields)   map['MaxRedemptions']   = Math.max(0, Number(fields.maxRedemptions||0));
  if ('maxPerCustomer' in fields)   map['MaxPerCustomer']   = Math.max(0, Number(fields.maxPerCustomer||0));
  if ('budgetCap' in fields)        map['BudgetCap']        = Math.max(0, Number(fields.budgetCap||0));
  if ('billingModel' in fields) map['BillingModel'] = _sanitizeBillingModel_(fields.billingModel);
  if ('costPerRedemption' in fields) map['CostPerRedemption'] = Math.max(0, Number(fields.costPerRedemption||0));

  Object.keys(map).forEach(k=>{
    const c = idx[k]; if (c) campaigns.getRange(r,c).setValue(map[k]);
  });
  const cUpdated= idx['UpdatedAt']||15;
  campaigns.getRange(r,cUpdated).setValue(_nowIso_());
  return { ok:true, campaignId };
}

function adminUploadCampaignImage(sid, campaignId, fileName, dataUrl) {
  ensureAdminActive_(sid);
  if (!campaignId) throw new Error('campaignId required.');
  if (!dataUrl || !/^data:image\/(png|jpe?g|gif|webp);base64,/.test(String(dataUrl))) {
    throw new Error('Invalid image data. Expect data URL of type png/jpg/gif/webp.');
  }

  const { campaigns } = getUserDb_();
  const r = findRowByColumn_(campaigns, 1, campaignId);
  if (!r) throw new Error('Campaign not found.');

  // Parse Data URL
  const m = String(dataUrl).match(/^data:(image\/[a-z0-9+.-]+);base64,(.*)$/i);
  if (!m) throw new Error('Invalid data URL.');
  const mime = m[1];
  const b64  = m[2];

  const bytes = Utilities.base64Decode(b64);
  const cleanName = (String(fileName||'Campaign').replace(/[^\w.\-]+/g,'_') || 'Campaign') +
                    '_' + campaignId + '_' + Date.now();

  const blob = Utilities.newBlob(bytes, mime, cleanName);
  const folder = ensureCampaignImageFolder_();
  const file = folder.createFile(blob);

  // Make public (anyone with link, view)
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (_) {}

  const url = publicViewUrlForFileId_(file.getId());

  // Write to Campaigns.ImageUrl & bump UpdatedAt
  const idx = adminHeaderIdx_(campaigns);
  const cImage = idx['ImageUrl'];
  const cUpdated = idx['UpdatedAt'] || 15;
  if (cImage) campaigns.getRange(r, cImage).setValue(url);
  if (cUpdated) campaigns.getRange(r, cUpdated).setValue(new Date().toISOString());

  return { ok:true, campaignId, imageUrl:url, fileId:file.getId() };
}

function adminListCouponRequests(sid, opts){
  ensureAdminActive_(sid);
  const { couponRequests } = getUserDb_();
  const idx = adminHeaderIdx_(couponRequests);
  const last = couponRequests.getLastRow();
  if (last < 2) return [];
  const cols = Math.max(couponRequests.getLastColumn()||1, 18);
  const vals = couponRequests.getRange(2,1,last-1,cols).getValues();

  const f = opts || {};
  const wantStatus = String(f.status||'').toLowerCase();
  const wantMid = String(f.merchantId||'').trim().toUpperCase();
  const q = String(f.q||'').toLowerCase();

  return vals.map(r => {
    const g = (k)=> r[(_colOrThrow_(idx,k))-1];
    const row = {
      requestId: String(g('RequestId')||''),
      merchantId: String(g('MerchantId')||''),
      code: String(g('Code')||''),
      mode: String(g('Mode')||''),
      type: String(g('Type')||''),
      value: Number(g('Value')||0),
      maxUses: Number(g('MaxUses')||0),
      perMemberLimit: Number(g('PerMemberLimit')||0),
      startIso: String(g('StartIso')||''),
      endIso: String(g('EndIso')||''),
      notes: String(g('Notes')||''),
      status: String(g('Status')||''),
      createdBy: String(g('CreatedBy')||''),
      createdAt: String(g('CreatedAt')||''),
      updatedAt: String(g('UpdatedAt')||''),
      decisionBy: String(couponRequests.getRange(2, _colOrThrow_(idx,'DecisionBy'), 1,1).getValue()||'') // tolerate missing historical
    };
    return row;
  }).filter(x =>
    (!wantStatus || x.status.toLowerCase() === wantStatus) &&
    (!wantMid || x.merchantId.toUpperCase() === wantMid) &&
    (!q || x.code.toLowerCase().includes(q) || x.requestId.toLowerCase().includes(q))
  );
}

function adminApproveCouponRequest(sid, requestId, override){
  const caller = ensureAdminActive_(sid);
  const { coupons, couponRequests, logs } = getUserDb_();

  const r = findRowByColumn_(couponRequests, 1, requestId);
  if (!r) throw new Error('Request not found.');

  const ix = adminHeaderIdx_(couponRequests);
  const cStatus = _colOrThrow_(ix,'Status');

  const status = String(couponRequests.getRange(r, cStatus).getValue()||'').toLowerCase();
  if (status !== 'pending') throw new Error('Only pending requests can be approved.');

  const g = (k)=> couponRequests.getRange(r, _colOrThrow_(ix,k)).getValue();
  const payload = {
    merchantId: String(g('MerchantId')||'').trim().toUpperCase(),
    code: String(g('Code')||'').trim(),
    mode: String(g('Mode')||'').toUpperCase(),
    type: String(g('Type')||'').toUpperCase(),
    value: Number(g('Value')||0),
    maxUses: Number(g('MaxUses')||0),
    perMemberLimit: Number(g('PerMemberLimit')||0),
    startIso: String(g('StartIso')||'').trim(),
    endIso: String(g('EndIso')||'').trim(),
    notes: String(g('Notes')||'')
  };
  if (override && typeof override === 'object'){
    Object.keys(override).forEach(k=>{
      if (typeof override[k] !== 'undefined') payload[k] = override[k];
    });
  }

  // Validate like merchant side
  if (!/^[A-Za-z0-9._-]{3,32}$/.test(payload.code)) throw new Error('Invalid code.');
  if (payload.mode && payload.mode !== 'COLLECT' && payload.mode !== 'REDEEM') throw new Error('Invalid mode.');
  if (payload.type !== 'BONUS' && payload.type !== 'DISCOUNT') throw new Error('Invalid type.');
  if (!(Number(payload.value)>0)) throw new Error('Value must be > 0.');
  if ((payload.startIso && isNaN(Date.parse(payload.startIso))) || (payload.endIso && isNaN(Date.parse(payload.endIso)))) throw new Error('Invalid dates.');

  const wrote = _upsertCouponFromPayload_(coupons, payload);

  const nowIso = new Date().toISOString();
  couponRequests.getRange(r, cStatus).setValue('approved');
  couponRequests.getRange(r, _colOrThrow_(ix,'DecisionBy')).setValue(caller.username);
  couponRequests.getRange(r, _colOrThrow_(ix,'DecisionAt')).setValue(nowIso);
  if (ix['DecisionNotes']) couponRequests.getRange(r, ix['DecisionNotes']).setValue(String((override && override.decisionNotes)||'Approved'));
  couponRequests.getRange(r, _colOrThrow_(ix,'UpdatedAt')).setValue(nowIso);

  try { logs.appendRow([nowIso, 'admin_coupon_approved', JSON.stringify({ requestId, code: payload.code, merchantId: payload.merchantId })]); } catch(_){}
  return { ok:true, requestId, code: payload.code, upserted: wrote };
}

function adminRejectCouponRequest(sid, requestId, reason){
  const caller = ensureAdminActive_(sid);
  const { couponRequests, logs } = getUserDb_();

  const r = findRowByColumn_(couponRequests, 1, requestId);
  if (!r) throw new Error('Request not found.');

  const ix = adminHeaderIdx_(couponRequests);
  const cStatus = _colOrThrow_(ix,'Status');

  const status = String(couponRequests.getRange(r, cStatus).getValue()||'').toLowerCase();
  if (status !== 'pending') throw new Error('Only pending requests can be rejected.');

  const nowIso = new Date().toISOString();
  couponRequests.getRange(r, cStatus).setValue('rejected');
  couponRequests.getRange(r, _colOrThrow_(ix,'DecisionBy')).setValue(caller.username);
  couponRequests.getRange(r, _colOrThrow_(ix,'DecisionAt')).setValue(nowIso);
  if (ix['DecisionNotes']) couponRequests.getRange(r, ix['DecisionNotes']).setValue(String(reason||''));
  couponRequests.getRange(r, _colOrThrow_(ix,'UpdatedAt')).setValue(nowIso);

  try { logs.appendRow([nowIso, 'admin_coupon_rejected', JSON.stringify({ requestId, reason:String(reason||'') })]); } catch(_){}
  return { ok:true, requestId };
}

// Live Coupons upsert (preserve UsedCount and original CreatedAt on updates)
function _upsertCouponFromPayload_(couponsSheet, p){
  const last = couponsSheet.getLastRow();
  const cols = Math.max(13, couponsSheet.getLastColumn()||13);

  // find exact Code+MerchantId match
  let row = 0;
  if (last >= 2){
    const vals = couponsSheet.getRange(2,1,last-1,2).getValues(); // Code, MerchantId
    const targetCode = String(p.code||'');
    const targetMid  = String(p.merchantId||'');
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0])===targetCode && String(vals[i][1])===targetMid){
        row = 2+i; break;
      }
    }
  }

  const nowIso = new Date().toISOString();

  if (row){
    const existing = couponsSheet.getRange(row,1,1,cols).getValues()[0];
    const usedCount = Number(existing[6]||0);
    const createdAt = String(existing[11]||nowIso);
    couponsSheet.getRange(row,1,1,13).setValues([[
      p.code, p.merchantId, p.mode, p.type, Number(p.value||0),
      Math.max(0, Number(p.maxUses||0)), usedCount,
      Math.max(0, Number(p.perMemberLimit||0)),
      String(p.startIso||''), String(p.endIso||''),
      true, createdAt, String(p.notes||'')
    ]]);
    return { updated:true, row };
  } else {
    couponsSheet.appendRow([
      p.code, p.merchantId, p.mode, p.type, Number(p.value||0),
      Math.max(0, Number(p.maxUses||0)), 0,
      Math.max(0, Number(p.perMemberLimit||0)),
      String(p.startIso||''), String(p.endIso||''),
      true, nowIso, String(p.notes||'')
    ]);
    return { created:true, row: couponsSheet.getLastRow() };
  }
}

// ===== COUPONS: core endpoints used by admin/Index.html =====
function adminCreateCoupon(payload) {
  // Expected payload: {
  //   merchantId, title, value, start, end, minSpend, maxRed, maxPerUser, budget, active
  // }
  payload = payload || {};
  const uid = PropertiesService.getScriptProperties().getProperty('USER_DB_ID');
  if (!uid) throw new Error('USER_DB_ID not configured.');
  const ss = SpreadsheetApp.openById(uid);

  // Ensure sheet + header
  const sheetName = 'Coupons';
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  const HEADER = [
    'CouponId',       // A
    'MerchantId',     // B
    'Title',          // C
    'Value',          // D  (number; semantic is up to your app)
    'Start',          // E
    'End',            // F
    'MinSpend',       // G
    'MaxRed',         // H
    'MaxPerUser',     // I
    'Budget',         // J
    'Active',         // K (TRUE/FALSE)
    'CreatedAt',      // L
    'UpdatedAt'       // M
  ];
  // Write header if missing or wrong width
  const width = HEADER.length;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, width).setValues([HEADER]);
  } else {
    // normalize header width
    if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
    sh.getRange(1, 1, 1, width).setValues([HEADER]);
  }

  // Simple validation
  const merchantId = String(payload.merchantId || '').trim();
  const title = String(payload.title || '').trim();
  if (!merchantId || !title) throw new Error('merchantId and title are required.');

  const nowIso = new Date().toISOString();
  const couponId = 'cpn_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);

  const row = [
    couponId,
    merchantId,
    title,
    Number(payload.value || 0),
    String(payload.start || ''),
    String(payload.end || ''),
    Number(payload.minSpend || 0),
    Number(payload.maxRed || 0),
    Number(payload.maxPerUser || 0),
    Number(payload.budget || 0),
    String(payload.active) === 'true' || payload.active === true,
    nowIso,
    nowIso
  ];
  sh.appendRow(row);
  return { ok: true, id: couponId };
}

function adminListCoupons(params) {
  params = params || {};
  const uid = PropertiesService.getScriptProperties().getProperty('USER_DB_ID');
  if (!uid) throw new Error('USER_DB_ID not configured.');
  const ss = SpreadsheetApp.openById(uid);
  const sh = ss.getSheetByName('Coupons');
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  header.forEach((h, i) => idx[String(h)] = i);
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  const q = String(params.search || '').toLowerCase();
  const filterMerchant = String(params.merchantId || '').trim();

  const rows = values.map(r => ({
    id: r[idx['CouponId']] || '',
    merchantId: r[idx['MerchantId']] || '',
    title: r[idx['Title']] || '',
    value: r[idx['Value']] || '',
    start: r[idx['Start']] || '',
    end: r[idx['End']] || '',
    minSpend: r[idx['MinSpend']] || '',
    maxRed: r[idx['MaxRed']] || '',
    maxPerUser: r[idx['MaxPerUser']] || '',
    budget: r[idx['Budget']] || '',
    active: r[idx['Active']] || ''
  })).filter(row => {
    if (filterMerchant && row.merchantId !== filterMerchant) return false;
    if (q && !((row.id||'').toLowerCase().includes(q) || (row.title||'').toLowerCase().includes(q))) return false;
    return true;
  });

  return { ok: true, rows };
}

function adminGetCouponById(couponId) {
  couponId = String(couponId || '').trim();
  if (!couponId) return { ok: false, reason: 'missing id' };

  const uid = PropertiesService.getScriptProperties().getProperty('USER_DB_ID');
  if (!uid) throw new Error('USER_DB_ID not configured.');
  const ss = SpreadsheetApp.openById(uid);
  const sh = ss.getSheetByName('Coupons');
  if (!sh || sh.getLastRow() < 2) return { ok: false, reason: 'not found' };

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  header.forEach((h, i) => idx[String(h)] = i);
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  for (var i = 0; i < values.length; i++) {
    const r = values[i];
    if (String(r[idx['CouponId']]) === couponId) {
      const coupon = {
        id: r[idx['CouponId']] || '',
        merchantId: r[idx['MerchantId']] || '',
        title: r[idx['Title']] || '',
        value: r[idx['Value']] || '',
        start: r[idx['Start']] || '',
        end: r[idx['End']] || '',
        minSpend: r[idx['MinSpend']] || '',
        maxRed: r[idx['MaxRed']] || '',
        maxPerUser: r[idx['MaxPerUser']] || '',
        budget: r[idx['Budget']] || '',
        active: r[idx['Active']] || ''
      };
      return { ok: true, coupon };
    }
  }
  return { ok: false, reason: 'not found' };
}

// ───────────────────────────── Helpers
function getLocalDb_() {
  const ss = SpreadsheetApp.getActive();
  const admins = ensureSheet_(ss, ADMINS_SHEET, [
    'Username','PinHash','CreatedAt',
    'Salt','Hash','Iter','Algo','Peppered','UpdatedAt',
    'Active','Role','MustChangePin'
  ]);
  const asessions = ensureSheet_(ss, A_SESS_SHEET, ['SID','Username','LastSeenMs','CreatedAt']);
  const config = ensureSheet_(ss, 'Config', ['Key','Value']);
  return { ss, admins, asessions, config };
}

function getAppVersion_() {
  try {
    const db = getLocalDb_();
    const config = db.config;
    const last = config.getLastRow();
    if (last < 2) return 'dev';

    const vals = config.getRange(2, 1, last - 1, 2).getValues(); // Key, Value
    for (var i = 0; i < vals.length; i++) {
      var key = String(vals[i][0] || '').trim();
      if (key === 'APP_VERSION') {
        var v = String(vals[i][1] || '').trim();
        return v || 'dev';
      }
    }
    return 'dev';
  } catch (err) {
    return 'dev';
  }
}

function getUserDb_() {
  const id = ASP.getProperty(USER_DB_ID_KEY);
  if (!id) throw new Error('USER_DB_ID is not set. Use the Config panel to set it.');
  const ss = SpreadsheetApp.openById(id);

  const merchants   = ensureSheet_(ss, 'Merchants',    ['MerchantId','Name','Active','Secret','CreatedAt']);
  const tx          = ensureSheet_(ss, 'Transactions', ['TxId','MemberId','MerchantId','Type','Points','AtMs','Staff']);
  const balances    = ensureSheet_(ss, 'Balances',     ['MemberId','Points']);
  // Note: 'Mode' may be blank for neutral QR deep links
  const linkTokens  = ensureSheet_(ss, 'LinkTokens',   ['Token','MemberId','Mode','ExpiresAtMs','Used','CreatedAt']);

  const sessions    = ensureSheet_(ss, 'Sessions',     ['SID','Email','LastSeenMs','CreatedAt']);
  const logs        = ensureSheet_(ss, 'Logs',         ['At','Type','Message']);
  const resetTokens = ensureSheet_(ss, 'ResetTokens',  ['Token','Email','ExpiresAtMs','Used','CreatedAt','VerifiedAtMs']);

  // ── Coupons ──
  const coupons     = ensureSheet_(ss, 'Coupons', [
    'Code','MerchantId','Mode','Type','Value','MaxUses','UsedCount','PerMemberLimit',
    'StartIso','EndIso','Active','CreatedAt','Notes'
  ]);
  const cpnUses     = ensureSheet_(ss, 'CouponUses', ['Code','MemberId','MerchantId','AtMs','Staff','TxId']);
  const couponRequests = ensureSheet_(ss, 'CouponRequests', [
    'RequestId','MerchantId','Code','Mode','Type','Value',
    'MaxUses','PerMemberLimit','StartIso','EndIso','Notes',
    'Status','CreatedBy','CreatedAt','UpdatedAt',
    'DecisionBy','DecisionAt','DecisionNotes'
  ]);  

  // ── NEW: Campaigns analytics ──
  const campaigns   = ensureSheet_(ss, 'Campaigns', [
    'CampaignId','MerchantId','Title','Type','Multiplier',
    'StartIso','EndIso','MinSpend','MaxRedemptions','MaxPerCustomer','BudgetCap',
    'BillingModel','CostPerRedemption','Active','CreatedAt','UpdatedAt',
    'ImageUrl'
  ]);

  const CampaignActivations = ensureSheet_(ss, 'CampaignActivations', [
    'ActivationId','CampaignId','MemberId','ActivatedAtMs','ExpiresAtMs','Source','Notes'
  ]);
  const CampaignRedemptions = ensureSheet_(ss, 'CampaignRedemptions', [
    'RedemptionId','CampaignId','MemberId','MerchantId','TxId','AtMs',
    'BasePoints','Multiplier','BonusPoints','CostAccrued'
  ]);
  // ── NEW: Merchant Campaign Requests ──
  const campaignRequests = ensureSheet_(ss, 'CampaignRequests', [
    'RequestId','MerchantId','RequestType','Title','Multiplier',
    'StartIso','EndIso','MinSpend','MaxRedemptions','MaxPerCustomer','BudgetCap',
    'BillingModel','CostPerRedemption','Notes','ImageUrl',
    'Status','CreatedBy','CreatedAt','UpdatedAt',
    'DecisionBy','DecisionAt','DecisionNotes','LinkedCampaignId'
  ]);

  return {
    ss, merchants, tx, balances, linkTokens, sessions, logs, resetTokens,
    coupons, cpnUses, couponRequests,
    campaigns, CampaignActivations, CampaignRedemptions,
    campaignRequests
  };
}

function getMerchantDb_() {
  const id = ASP.getProperty(MERCHANT_DB_ID_KEY);
  if (!id) throw new Error('MERCHANT_DB_ID is not set.');
  const ss = SpreadsheetApp.openById(id);
  const msessions = ensureSheet_(ss, 'MSessions', ['SID','Username','LastSeenMs','CreatedAt']);
  const staff     = ensureSheet_(ss, 'Staff', [
    'Username','PinHash','MerchantId','Role','Active','CreatedAt',
    'Salt','Hash','Iter','Algo','Peppered','UpdatedAt','MustChangePin'
  ]);
  return { ss, msessions, staff };
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const h = sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn()||headers.length)).getValues()[0];
  if (h.every(v => !v)) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    for (let i=0;i<headers.length;i++){ if (!h[i]) sh.getRange(1,i+1).setValue(headers[i]); }
  }
  return sh;
}
function protectAllSheetHeaders_(ss, warningOnly) {
  const sheets = ss.getSheets();
  const me = Session.getActiveUser().getEmail() || null;
  sheets.forEach(sh => {
    const range = sh.getRange(1,1,1, sh.getMaxColumns());
    // remove existing header protections on row 1
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
      const r = p.getRange();
      if (r.getRow()===1 && r.getNumRows()===1) p.remove();
    });
    const prot = range.protect().setDescription('Header row protected by Admin App');
    prot.setWarningOnly(!!warningOnly);
    if (!warningOnly && me) {
      try {
        prot.removeEditors(prot.getEditors());
        prot.addEditor(me); // keep yourself
      } catch(_) {}
    }
  });
}
function findRowByColumn_(sh, colIndex, value) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, colIndex, last-1, 1).getValues();
  const needle = String(value).toLowerCase();
  for (let i=0;i<vals.length;i++) {
    if (String(vals[i][0]).toLowerCase() === needle) return 2 + i;
  }
  return 0;
}
function csvEscape_(v) {
  const s = String(v==null?'':v);
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g,'""') + '"';
  return s;
}

// List logs with filters
// opts: { q, type, onlyErrors, fromIso, toIso, limit }
function adminListLogsFiltered(sid, opts){
  requireAdmin_(sid);
  const { logs } = getUserDb_();
  const last = logs.getLastRow();
  if (last < 2) return [];

  const o = opts || {};
  const q = String(o.q || '').toLowerCase();
  const wantType = String(o.type || '').trim().toLowerCase();
  const onlyErrors = String(o.onlyErrors || '').toLowerCase() === 'true' || o.onlyErrors === true;

  const fromMs = _parseMs_A(o.fromIso);
  const toMs   = _parseMs_A(o.toIso);
  const limit  = Math.max(1, Math.min(Number(o.limit) || 200, 1000));

  const vals = logs.getRange(2,1,last-1,3).getValues(); // [At, Type, Message]
  const out = [];

  for (let i = vals.length - 1; i >= 0 && out.length < limit; i--){
    const atStr = String(vals[i][0] || '');
    const type  = String(vals[i][1] || '');
    const msg   = String(vals[i][2] || '');

    const atMs = _parseMs_A(atStr);
    if (fromMs && (!atMs || atMs < fromMs)) continue;
    if (toMs && (!atMs || atMs > toMs)) continue;

    if (wantType && type.toLowerCase() !== wantType) continue;
    if (onlyErrors && !/error|fail/i.test(type)) continue;

    if (q){
      const hay = (type + ' ' + msg).toLowerCase();
      if (!hay.includes(q)) continue;
    }

    out.push({ at: atStr, type, msg });
  }
  return out;
}


// ───────────────────────────── Visibility Dashboard (metrics) ─────────────────────────────
// Time helpers (script timezone)
function _startOfTodayMs_A(){ const d=new Date(); d.setHours(0,0,0,0); return d.getTime(); }
function _endOfTodayMs_A(){ const d=new Date(); d.setHours(23,59,59,999); return d.getTime(); }
function _daysAgoStartMs_A(n){ const d=new Date(); d.setHours(0,0,0,0); d.setDate(d.getDate()-Number(n||0)); return d.getTime(); }
function _thisMonthStartMs_A(){ const d=new Date(); d.setHours(0,0,0,0); d.setDate(1); return d.getTime(); }
function _parseMs_A(v){
  if (!v && v!==0) return 0;
  if (typeof v === 'number') return v;
  const s = String(v||'').trim();
  if (/^\d{10,13}$/.test(s)) return Number(s);
  const t = Date.parse(s);
  return isNaN(t)?0:t;
}
function _within_A(ms, start, end){ const t=Number(ms||0); return !!t && t>=start && t<=end; }

// Safe get merchant DB (optional)
function _tryGetMerchantDb_A(){
  try { return getMerchantDb_(); } catch(_){ return null; }
}

// MAIN: Admin visibility payload
function adminGetVisibility(sid) {
  requireAdmin_(sid);
  const { ss, merchants, tx, balances, linkTokens, sessions, logs, resetTokens } = getUserDb_();
  const mdb = _tryGetMerchantDb_A();
  const msessions = mdb ? mdb.msessions : null;

  const today0 = _startOfTodayMs_A();
  const today1 = _endOfTodayMs_A();
  const last7d0 = _daysAgoStartMs_A(6);            // inclusive (today + 6 prior days)
  const last30d0= _daysAgoStartMs_A(29);
  const month0  = _thisMonthStartMs_A();

  // ── Users ──
  const users = ss.getSheetByName('Users');
  let totalUsers=0, newUsersToday=0, newUsers7d=0;
  if (users && users.getLastRow()>=2) {
    const cols = Math.max(users.getLastColumn()||0, 8);
    const hdr = users.getRange(1,1,1,cols).getValues()[0].map(String);
    const idxCreated = hdr.findIndex(h => String(h).toLowerCase()==='createdat');
    const last = users.getLastRow();
    totalUsers = last - 1;
    if (idxCreated>=0){
      const createdCol = users.getRange(2, idxCreated+1, totalUsers, 1).getValues();
      for (let i=0;i<createdCol.length;i++){
        const ms = _parseMs_A(createdCol[i][0]);
        if (_within_A(ms, today0, today1)) newUsersToday++;
        if (_within_A(ms, last7d0, today1)) newUsers7d++;
      }
    }
  }

  // ── Merchants ──
  let totalMerchants=0, activeMerchants=0, inactiveMerchants=0;
  if (merchants && merchants.getLastRow()>=2) {
    const last = merchants.getLastRow();
    const vals = merchants.getRange(2,1,last-1,3).getValues(); // [MerchantId, Name, Active]
    totalMerchants = vals.length;
    vals.forEach(r=>{
      const active = String(r[2]).toUpperCase()!=='FALSE';
      if (active) activeMerchants++; else inactiveMerchants++;
    });
  }

  // ── Active sessions (user + merchant) = lastSeen within 60s ──
  const now = Date.now();
  const freshMs = U_SESSION_TTL_MS; // and for merchant sessions use M_SESSION_TTL_MS
  let activeUserSessions=0, activeMerchantSessions=0;
  if (sessions && sessions.getLastRow()>=2) {
    const last = sessions.getLastRow();
    const vals = sessions.getRange(2,1,last-1,3).getValues(); // [SID, Email, LastSeenMs]
    activeUserSessions = vals.reduce((n,r)=> n + ((now - Number(r[2]||0))<=freshMs ? 1:0), 0);
  }
  if (msessions && msessions.getLastRow()>=2) {
    const last = msessions.getLastRow();
    const vals = msessions.getRange(2,1,last-1,3).getValues(); // [SID, Username, LastSeenMs]
    activeMerchantSessions = vals.reduce((n,r)=> n + ((now - Number(r[2]||0))<=freshMs ? 1:0), 0);
  }

  // ── Transactions ──
  let txToday=0, tx7d=0, ptsToday=0, pts7d=0, avgTx7d=0;
  // Should-have: top5 merchants last 7d, inactive merchants 30d, circulation this month
  const topMerchants = {}; // id -> {count, points}
  const touchedMerchantLast30 = new Set();
  let collectMonth=0, redeemMonth=0;
  // Nice: churn, last activity per member
  const lastSeenMember = {}; // memberId -> lastMs

  if (tx && tx.getLastRow()>=2) {
    const last = tx.getLastRow();
    const cols = Math.max(7, tx.getLastColumn()||7);
    const vals = tx.getRange(2,1,last-1,cols).getValues(); // [TxId,MemberId,MerchantId,Type,Points,AtMs,Staff]
    for (let i=0;i<vals.length;i++){
      const row = vals[i];
      const mId = String(row[2]||'');
      const type = String(row[3]||'').toUpperCase();
      const pts = Number(row[4]||0);
      const at  = Number(row[5]||0);

      // Must-have: counts
      if (_within_A(at, today0, today1)){ txToday++; ptsToday += pts; }
      if (_within_A(at, last7d0, today1)){ tx7d++; pts7d += pts; }

      // Should: top merchants 7d
      if (_within_A(at, last7d0, today1) && mId){
        if (!topMerchants[mId]) topMerchants[mId] = { count:0, points:0 };
        topMerchants[mId].count++; topMerchants[mId].points += pts;
      }
      // Should: inactive merchants 30d
      if (_within_A(at, last30d0, today1) && mId) touchedMerchantLast30.add(mId);

      // Should: circulation (this month)
      if (_within_A(at, month0, today1)){
        if (type==='COLLECT') collectMonth += pts;
        else if (type==='REDEEM') redeemMonth += pts;
      }

      // Nice: churn (last seen per member)
      const mem = String(row[1]||'');
      if (mem && at) {
        const prev = lastSeenMember[mem]||0;
        if (at > prev) lastSeenMember[mem] = at;
      }
    }
  }
  avgTx7d = tx7d ? (pts7d / tx7d) : 0;

  // ── Balances ──
  let nonZeroBalances=0, avgBalance=0;
  if (balances && balances.getLastRow()>=2){
    const last = balances.getLastRow();
    const vals = balances.getRange(2,1,last-1,2).getValues(); // [MemberId, Points]
    let sum=0;
    for (let i=0;i<vals.length;i++){
      const p = Number(vals[i][1]||0);
      if (p>0) nonZeroBalances++;
      sum += p;
    }
    avgBalance = (vals.length>0) ? (sum/vals.length) : 0;
  }

  // ── Errors last 24h ──
  let errors24h=0;
  const since24h = now - 24*3600*1000;
  if (logs && logs.getLastRow()>=2){
    const last = logs.getLastRow();
    const vals = logs.getRange(2,1,last-1,3).getValues(); // [At, Type, Message]
    for (let i=0;i<vals.length;i++){
      const at = _parseMs_A(vals[i][0]);
      const type = String(vals[i][1]||'');
      if (at >= since24h && /error|fail/i.test(type)) errors24h++;
    }
  }

  // ── Storage health ──
  function _sheetSizeWarn_(sh, name){
    if (!sh) return { name, rows:0, cols:0, warn:false };
    const rows = sh.getLastRow()||0, cols = sh.getLastColumn()||0;
    // Heuristics: warn if rows > 20k on Transactions, > 10k on Logs, > 50k cells elsewhere
    let warn=false;
    if (name==='Transactions' && rows>20000) warn=true;
    else if (name==='Logs' && rows>10000) warn=true;
    else if ((rows*cols) > 50000) warn=true;
    return { name, rows, cols, warn };
  }
  const storage = [
    _sheetSizeWarn_(tx, 'Transactions'),
    _sheetSizeWarn_(logs, 'Logs'),
    _sheetSizeWarn_(balances, 'Balances'),
    _sheetSizeWarn_(sessions, 'Sessions'),
    _sheetSizeWarn_(merchants, 'Merchants'),
  ];

  // ── Should: inactive merchants (no tx in 30+ days) ──
  const inactiveMerchants30d = [];
  if (merchants && merchants.getLastRow()>=2){
    const last = merchants.getLastRow();
    const vals = merchants.getRange(2,1,last-1,3).getValues(); // [MerchantId, Name, Active]
    for (let i=0;i<vals.length;i++){
      const id = String(vals[i][0]||'');
      const nm = String(vals[i][1]||'');
      const active = String(vals[i][2]).toUpperCase()!=='FALSE';
      if (!active) continue;
      if (!touchedMerchantLast30.has(id)) inactiveMerchants30d.push({ merchantId:id, name:nm });
    }
  }

  // ── Should: Top 5 merchants (7d) ──
  const top5 = Object.keys(topMerchants).map(id => ({
    merchantId: id,
    count: topMerchants[id].count,
    points: topMerchants[id].points
  })).sort((a,b)=> b.count - a.count || b.points - a.points).slice(0,5);

  // ── Should: circulation + ratio ──
  const redemptionRatio = collectMonth>0 ? (redeemMonth/collectMonth) : 0;

  // ── Should: System health snapshot ──
  // Last admin login: newest CreatedAt in ASessions
  let lastAdminLogin = '';
  try {
    const { asessions } = getLocalDb_();
    if (asessions.getLastRow()>=2){
      const last = asessions.getLastRow();
      const vals = asessions.getRange(2,1,last-1,4).getValues();
      let maxMs = 0;
      for (let i=0;i<vals.length;i++){
        const at = _parseMs_A(vals[i][3]); // CreatedAt ISO
        if (at>maxMs) maxMs=at;
      }
      if (maxMs) lastAdminLogin = new Date(maxMs).toISOString();
    }
  } catch(_){}
  // Last purge run: look for admin_purge_* in Logs (we’ll also add logging in purge fns below)
  let lastPurgeRun = '';
  if (logs && logs.getLastRow()>=2){
    const last = logs.getLastRow();
    const vals = logs.getRange(2,1,last-1,3).getValues();
    let maxMs=0;
    for (let i=0;i<vals.length;i++){
      const at = _parseMs_A(vals[i][0]);
      const type = String(vals[i][1]||'').toLowerCase();
      if (/^admin_purge_/.test(type) && at>maxMs) maxMs=at;
    }
    if (maxMs) lastPurgeRun = new Date(maxMs).toISOString();
  }
  const nextScheduledHousekeeping = ''; // none (no triggers configured yet)

  // ── Nice: Top 10 users by points ──
  let topUsers10 = [];
  if (balances && balances.getLastRow()>=2){
    const last = balances.getLastRow();
    const vals = balances.getRange(2,1,last-1,2).getValues();
    topUsers10 = vals
      .map(r => ({ memberId:String(r[0]||''), points:Number(r[1]||0) }))
      .sort((a,b)=> b.points - a.points)
      .slice(0,10);
  }

  // ── Nice: Most active merchants by member count (last 30d unique members) ──
  const membersByMerchant = {}; // id -> Set
  if (tx && tx.getLastRow()>=2){
    const last = tx.getLastRow();
    const cols = Math.max(7, tx.getLastColumn()||7);
    const vals = tx.getRange(2,1,last-1,cols).getValues();
    for (let i=0;i<vals.length;i++){
      const row = vals[i];
      const mid = String(row[2]||'');
      const mem = String(row[1]||'');
      const at  = Number(row[5]||0);
      if (!_within_A(at, last30d0, today1)) continue;
      if (!mid || !mem) continue;
      if (!membersByMerchant[mid]) membersByMerchant[mid] = new Set();
      membersByMerchant[mid].add(mem);
    }
  }
  const mostActiveMerchantsByMembers = Object.keys(membersByMerchant)
    .map(id=>({ merchantId:id, uniqueMembers: membersByMerchant[id].size }))
    .sort((a,b)=> b.uniqueMembers - a.uniqueMembers)
    .slice(0,5);

  // ── Nice: churn — members with no activity in 60+ days ──
  const cutoff60 = now - 60*24*3600*1000;
  let churnCount60 = 0;
  // If we didn’t see some members in tx loop, we can’t know; we’ll estimate from lastSeenMember map vs Users count
  // Strategy: if we have totalUsers and lastSeenMember size, count as churn those in map older than 60d
  for (const mem in lastSeenMember) if (lastSeenMember[mem] < cutoff60) churnCount60++;

  // ── Nice: geo snapshot (Users.City / Users.Country) ──
  const geo = { cities:{}, countries:{} };
  if (users && users.getLastRow()>=2){
    const last = users.getLastRow();
    const cols = Math.max(users.getLastColumn()||0, 21);
    const hdr = users.getRange(1,1,1,cols).getValues()[0].map(String);
    const iCity = hdr.findIndex(h => String(h).toLowerCase()==='city');
    const iCountry = hdr.findIndex(h => String(h).toLowerCase()==='country');
    if (iCity>=0 || iCountry>=0){
      const vals = users.getRange(2,1,last-1,cols).getValues();
      for (let i=0;i<vals.length;i++){
        if (iCity>=0) {
          const c = String(vals[i][iCity]||'').trim()||'—';
          geo.cities[c] = (geo.cities[c]||0)+1;
        }
        if (iCountry>=0) {
          const c = String(vals[i][iCountry]||'').trim()||'—';
          geo.countries[c] = (geo.countries[c]||0)+1;
        }
      }
    }
  }

  // ── Nice: errors trend (last 7d) ──
  const errTrend = {}; // YYYY-MM-DD -> count
  for (let d=0; d<7; d++){
    const day = new Date(_daysAgoStartMs_A(6-d));
    const y = day.getFullYear(), m=('0'+(day.getMonth()+1)).slice(-2), dd=('0'+day.getDate()).slice(-2);
    errTrend[y+'-'+m+'-'+dd] = 0;
  }
  if (logs && logs.getLastRow()>=2){
    const last = logs.getLastRow();
    const vals = logs.getRange(2,1,last-1,3).getValues();
    for (let i=0;i<vals.length;i++){
      const at = _parseMs_A(vals[i][0]);
      const type = String(vals[i][1]||'');
      if (!/error|fail/i.test(type)) continue;
      if (!_within_A(at, last7d0, today1)) continue;
      const d = new Date(at);
      const y = d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), dd=('0'+d.getDate()).slice(-2);
      const key = y+'-'+m+'-'+dd;
      if (errTrend[key]!==undefined) errTrend[key] += 1;
    }
  }

  // ── Alerts: sheets near limits + expiring reset tokens (next 60 min) ──
  const alerts = [];
  storage.forEach(s=>{
    if (s.warn) alerts.push('Sheet "'+s.name+'" is getting large ('+s.rows+' rows). Consider purging/archiving.');
  });
  // Reset tokens expiring soon
  try {
    if (resetTokens && resetTokens.getLastRow()>=2){
      const last = resetTokens.getLastRow();
      const vals = resetTokens.getRange(2,1,last-1,6).getValues(); // Token, Email, ExpMs, Used, CreatedAt, VerifiedAtMs
      let soon=0;
      for (let i=0;i<vals.length;i++){
        const exp = Number(vals[i][2]||0);
        const used= String(vals[i][3]).toLowerCase()==='true';
        if (used) continue;
        if (exp>now && (exp-now)<=60*60*1000) soon++;
      }
      if (soon>0) alerts.push(soon+' reset link(s) expiring in the next 60 minutes.');
    }
  } catch(_){}

  return {
    must: {
      users: { total: totalUsers, newToday: newUsersToday, new7d: newUsers7d },
      merchants: { total: totalMerchants, active: activeMerchants, inactive: inactiveMerchants },
      sessions: { user: activeUserSessions, merchant: activeMerchantSessions },
      tx: { today: txToday, last7d: tx7d, ptsToday, pts7d, avgTx7d },
      balances: { nonZero: nonZeroBalances, avg: avgBalance },
      errors24h,
      storage
    },
    should: {
      inactiveMerchants30d,
      topMerchants7d: top5,
      avgTx7d,
      circulation: { collectMonth, redeemMonth, redemptionRatio },
      system: { lastAdminLogin, lastPurgeRun, nextScheduledHousekeeping }
    },
    nice: {
      topUsers10,
      mostActiveMerchantsByMembers,
      churn60d: churnCount60,
      geo,
      errorTrend7d: errTrend,
      alerts
    }
  };
}

function adminRecentTxByStaff(sid, staffUsername, limit){
  const caller = ensureAdminActive_(sid);
  const { tx } = getUserDb_();
  const n = Math.max(1, Math.min(Number(limit)||50, 200));
  if (tx.getLastRow()<2) return [];
  const cols = Math.max(7, tx.getLastColumn()||7);
  const last = tx.getLastRow();
  const vals = tx.getRange(2,1,last-1,cols).getValues();

  const uname = sanitizeUsername_(staffUsername||'');
  const out = [];
  for (let i=vals.length-1; i>=0 && out.length<n; i--){
    const [txId, memberId, merchantId, type, points, atMs, staff] = vals[i];
    if (uname && String(staff||'').toLowerCase() !== uname) continue;
    out.push({ txId, memberId, merchantId, type, points:Number(points||0), atMs:Number(atMs||0), staff });
  }
  return out;
}


// ── Secure PIN helpers (PBKDF2-SHA256) ─────────────────────────────────────────
const PIN_ALGO_DEFAULT = 'PBKDF2-SHA256';
const PIN_ITER_DEFAULT = Number(PropertiesService.getScriptProperties().getProperty('PIN_PBKDF2_ITER')) || 75000;
const PIN_PEPPER_B64   = (function(){
  const v = PropertiesService.getScriptProperties().getProperty('PIN_PEPPER');
  return v ? String(v) : '';
})();

function utf8Bytes_(s){ return Utilities.newBlob(String(s)||'').getBytes(); }
function b64enc_(bytes){ return Utilities.base64Encode(bytes); }
function b64dec_(b64){ return Utilities.base64Decode(String(b64)||''); }

function randomSaltB64_(len){
  const seed = Utilities.getUuid() + '|' + Date.now() + '|' + Math.random();
  let bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed);
  if (bytes.length > len) { bytes = bytes.slice(0, len); }
  if (bytes.length < len) {
    const extra = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed + '|extra');
    const out = [];
    for (let i=0;i<len;i++) out.push((bytes[i%bytes.length] ^ extra[i%extra.length]) & 0xff);
    bytes = out;
  }
  return b64enc_(bytes);
}
function hmacSha256_(keyBytes, dataBytes){
  // Apps Script: use computeHmacSignature (SHA-256), returns byte[]
  return Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    dataBytes,
    keyBytes
  );
}

function xorBytes_(a,b){ const out=[]; for (let i=0;i<Math.min(a.length,b.length);i++) out[i]=(a[i]^b[i])&0xff; return out; }
function pbkdf2Sha256_(passwordBytes, saltBytes, iterations, dkLen){
  const hLen=32, l=Math.ceil(dkLen/hLen), r=dkLen-(l-1)*hLen, dk=[];
  for (let i=1;i<=l;i++){
    const blockIndex=[0,0,0,(i&0xff)];
    const u1=hmacSha256_(passwordBytes, saltBytes.concat(blockIndex));
    let t=u1.slice(), u=u1;
    for (let c=2;c<=iterations;c++){ u=hmacSha256_(passwordBytes,u); t=xorBytes_(t,u); }
    const part=(i===l)? t.slice(0,r) : t; for(let j=0;j<part.length;j++) dk.push(part[j]);
  }
  return dk;
}
function derivePinHashB64_(pin, saltB64, iter, usePepper){
  const salt=b64dec_(saltB64), pinBytes=utf8Bytes_(String(pin||'')), pepper=(usePepper&&PIN_PEPPER_B64)? b64dec_(PIN_PEPPER_B64):[];
  const pwd=pinBytes.concat(pepper); const dk=pbkdf2Sha256_(pwd, salt, Math.max(1000,Number(iter)||PIN_ITER_DEFAULT),32);
  return b64enc_(dk);
}
function adminHeaderIdx_(sheet){
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {}; // Handle empty sheet
  const hdr=sheet.getRange(1,1,1,lastCol).getValues()[0].map(String);
  const idx={}; hdr.forEach((h,i)=> idx[h.trim()]=i+1); return idx;
}
function _colOrThrow_(idx, name){
  const c = idx[name];
  if (!c) throw new Error('Missing column "'+name+'". Please ensure the sheet header matches exactly.');
  return c;
}
function makeSaltedAdminRecord_(pin){
  const saltB64=randomSaltB64_(24);
  const iter=PIN_ITER_DEFAULT, algo=PIN_ALGO_DEFAULT, peppered=!!PIN_PEPPER_B64;
  const hashB64=derivePinHashB64_(pin, saltB64, iter, peppered);
  return { saltB64, hashB64, iter, algo, peppered, updatedAt: new Date().toISOString() };
}
function verifyAdminPinAndMigrateIfNeeded_(adminsSheet, row, pin){
  const idx=adminHeaderIdx_(adminsSheet);
  const get=(name)=>{ const c=idx[name]; return c? adminsSheet.getRange(row,c).getValue() : ''; };
  const setRow = (map)=>{
    Object.keys(map).forEach(k=>{ const c=idx[k]; if(c) adminsSheet.getRange(row,c).setValue(map[k]); });
  };
  const algo=String(get('Algo')||''); const saltB64=String(get('Salt')||''); const hashB64=String(get('Hash')||'');
  const iter=Number(get('Iter')||0); const peppered=String(get('Peppered')||'').toLowerCase()==='true';

  if (algo && saltB64 && hashB64 && iter){
    const calc=derivePinHashB64_(pin, saltB64, iter, peppered);
    return { ok:(calc===String(hashB64)), migrated:false };
  }
  // legacy
  const legacyHash=String(get('PinHash')||'');
  if (legacyHash){
    const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin||'')); 
    const hex=bytes.map(b => (b+256)%256).map(b=>('0'+b.toString(16)).slice(-2)).join('');
    const okLegacy=(hex===legacyHash);
    if (okLegacy){
      const rec=makeSaltedAdminRecord_(pin);
      setRow({'Salt':rec.saltB64,'Hash':rec.hashB64,'Iter':rec.iter,'Algo':rec.algo,'Peppered':rec.peppered,'UpdatedAt':rec.updatedAt});
      return { ok:true, migrated:true };
    }
  }
  return { ok:false, migrated:false };
}

// ── Utils ──
function sanitizeUsername_(u) {
  if (!u) return null;
  u = String(u).trim().toLowerCase();
  if (!/^[a-z0-9._]{3,32}$/.test(u)) return null;
  return u;
}
function isValidPin_(pin) { return typeof pin === 'string' && /^[0-9]{6}$/.test(pin); }
function hashPin_(pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin);
  return bytes.map(b => (b + 256) % 256).map(b => ('0' + b.toString(16)).slice(-2)).join('');
}
function genSid_() { return 's_' + Utilities.getUuid().replace(/-/g,''); }



