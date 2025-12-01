// File: Code.gs
const SP = PropertiesService.getScriptProperties();
const DB_ID_KEY = 'MEMBER_APP_DB_ID';
const DB_NAME = 'MemberAppData';
const USERS_SHEET = 'Users';
const SESSIONS_SHEET = 'Sessions';
const MERCHANTS_SHEET = 'Merchants';
const TX_SHEET = 'Transactions';
const BALANCES_SHEET = 'Balances';
const LINK_TOKENS_SHEET = 'LinkTokens';     // short-lived deep-link tokens
const LOCKOUTS_SHEET = 'Lockouts';          // new: auth lockouts
const LOGS_SHEET = 'Logs';                  // new: telemetry logs
const WAITLIST_SHEET = 'Waitlist';          // new: waitlist captures
const CONFIG_SHEET   = 'Config';            // new sheet for key value config

// —— Per-invocation cache to avoid repeated open/ensure overhead ——
// These live only for the duration of a single Apps Script invocation.
let __DB_CACHE = null;



const SESSION_TTL_MS = 60 * 1000;           // 1 minute
// Add near top with other constants
const RESET_TOKENS_SHEET = 'ResetTokens'; // [Token, Email, ExpiresAtMs, Used, CreatedAt]
const RESET_TOKEN_TTL_MS = 30 * 60 * 1000; // 30 minutes
// —— PIN hardening for USERS (PBKDF2 + optional pepper) ——
const PIN_ALGO_DEFAULT = 'PBKDF2-SHA256';
// Optionales Label-Override per Script Property (wie bei Merchant/Admin):
const PIN_HASH_ALGO = String(SP.getProperty('PIN_HASH_ALGO') || PIN_ALGO_DEFAULT);

const PIN_ITER_DEFAULT = Number(SP.getProperty('PIN_PBKDF2_ITER')) || 3000;
const PIN_PEPPER_B64   = (function(){

  // Optional global pepper (base64) set in Script Properties: PIN_PEPPER
  const v = SP.getProperty('PIN_PEPPER');
  return v ? String(v) : '';
})();

// —— Magic link tokens ——
const MAGIC_TOKENS_SHEET = 'MagicTokens';    // [Token, Email, ExpiresAtMs, Used, CreatedAt]
const MAGIC_TOKEN_TTL_MS = 10 * 60 * 1000;   // 10 minutes
const PENDING_REG_SHEET = 'PendingRegs';   // staging new registrations + OTP
const OTP_TTL_MS = 15 * 60 * 1000;         // 15 minutes
// Optional: set in Script Properties to override default
// SP.setProperty('REG_ADMIN_KEY','cashbeik-admin');

// Merchant app base URL (set once by admin via Script Property)
const MERCHANT_APP_BASE_URL_KEY = 'MERCHANT_APP_BASE_URL';

// Token lifetime (e.g., 10 minutes)
const LINK_TOKEN_TTL_MS = 10 * 60 * 1000;

// Lockout policy (progressive)
// After N total consecutive failures, impose the listed cooldown.
// durMin = -1 means permanent lock (admin must reset)
const LOCK_RULES = [
  { fails: 5,  durMin: 15 },    // 15 minutes
  { fails: 7,  durMin: 60 },    // 1 hour
  { fails: 9,  durMin: 1440 },  // 24 hours
  { fails: 10, durMin: -1 }     // permanent until admin reset
];

// ——— Launch flag (pre-launch waitlist mode) ———
const WAITLIST_MODE_KEY = 'WAITLIST_MODE'; // set to 'on' (or 'true') for pre-launch

function isWaitlistMode_() {
  const v = String(SP.getProperty(WAITLIST_MODE_KEY) || '').toLowerCase();
  return v === 'on' || v === 'true' || v === '1';
}

// Minimal flags endpoint for client boot
function getFlags() {
  return { waitlist: isWaitlistMode_() };
}


function doGet(e) {
  // API passthroughs for testing flows (non-HTML)
  if (e && e.parameter && e.parameter.api) {
    const api = String(e.parameter.api || '').toLowerCase();

    if (api === 'reg_peek') {
      const pid = String(e.parameter.pid || '').trim();
      const key = String(e.parameter.key || '').trim();
      const ADMIN_KEY = (SP.getProperty('REG_ADMIN_KEY') || 'cashbeik-admin');
      if (!pid) return ContentService.createTextOutput('missing pid').setMimeType(ContentService.MimeType.TEXT);
      if (!key || key !== ADMIN_KEY) return ContentService.createTextOutput('unauthorized').setMimeType(ContentService.MimeType.TEXT);

      const { pending } = getDb_();
      const r = findRowByColumn_(pending, 1, pid);
      if (!r) return ContentService.createTextOutput('not found').setMimeType(ContentService.MimeType.TEXT);

      const row = pending.getRange(r, 1, 1, pending.getLastColumn()).getValues()[0];
      // [PendingId, CreatedAtMs, Status, Email, PhoneE164, WaE164, PrefComms, OTP, OtpExpiresAtMs, First, Last, BirthYmd, Gender, City, Country, Lang, Referral, OptMarketing, UTM_Source, UTM_Medium, UTM_Campaign, UTM_Term, UTM_Content]
      const pidOut = String(row[0]||'');
      const status = String(row[2]||'');
      const email  = String(row[3]||'');
      const phone  = String(row[4]||'');
      const wa     = String(row[5]||'');
      const pref   = String(row[6]||'');
      const otp    = String(row[7]||'');
      const expMs  = Number(row[8]||0);

      const obj = { ok:true, pid:pidOut, status, email, phone, whatsapp:wa, pref, otp, expiresAtMs: expMs };
      return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
    }

    // Unknown API
    return ContentService.createTextOutput('unknown api').setMimeType(ContentService.MimeType.TEXT);
  }

  // Normal HTML rendering
  const t = HtmlService.createTemplateFromFile('Index');
  t.appVersion = getAppVersion_();

  return t.evaluate()
    .setTitle('Minimal Member App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function warm() {
  // Touch the DB and one cheap read to keep container hot
  const { users } = getDb_();
  // Minimal read to force lazy init
  try { users.getLastRow(); } catch(_) {}
  return { ok:true, at: new Date().toISOString() };
}


// ---- Public API ----
function signup(email, pin) {
  email = sanitizeEmail_(email);
  if (!email) throw new Error('Invalid email.');
  if (!isValidPin_(pin)) throw new Error('PIN must be exactly 6 digits.');
  const { users } = getDb_();
  const row = findRowByColumn_(users, 1, email);
  if (row) throw new Error('Email already registered.');
  const memberId = genMemberId_();
  users.appendRow([email, '', memberId, new Date().toISOString()]); // leave legacy blank
  const newRow = users.getLastRow();
  setUserSecurity_(newRow, pin);
  return { memberId };
}

function getSigninLockInfo(email){
  email = sanitizeEmail_(email);
  if (!email) return { ok:false, reason:'invalid' };
  const { lockouts } = getDb_();
  const key = 'user:' + email;

  // Current status (uses our progressive helpers)
  const st = getLockoutStatus_(lockouts, key); // {locked, permanent, until}
  // Read row if any
  const r = findRowByColumn_(lockouts, 1, key);
  let failCount = 0;
  if (r){
    const [, fc] = lockouts.getRange(r,1,1,2).getValues()[0];
    failCount = Number(fc||0);
  }

  // Determine next threshold to display like "2 / 5"
  let nextTarget = null;
  for (let i=0;i<LOCK_RULES.length;i++){
    if (failCount < LOCK_RULES[i].fails){
      nextTarget = LOCK_RULES[i].fails;
      break;
    }
  }
  // If already beyond last rule, keep showing last as the target
  if (nextTarget === null && LOCK_RULES.length){
    nextTarget = LOCK_RULES[LOCK_RULES.length - 1].fails;
  }

  return {
    ok: true,
    failCount,
    nextTarget,
    locked: !!st.locked,
    permanent: !!st.permanent,
    untilMs: Number(st.until||0)
  };
}

function getLockoutStatus_(lockouts, key){
  const r = findRowByColumn_(lockouts, 1, key);
  if (!r) return { locked: false };
  const [k, failCount, lockUntilMs, updatedAt, firstFailMs, total24h, permLocked] =
    lockouts.getRange(r,1,1,7).getValues()[0];
  const now = Date.now();
  const permanent = String(permLocked).toLowerCase()==='true';
  const until = Number(lockUntilMs||0);
  if (permanent) return { locked:true, permanent:true, until };
  if (until && until > now) return { locked:true, permanent:false, until };
  return { locked:false };
}

function signin(email, pin) {
  email = sanitizeEmail_(email);
  if (!email) throw new Error('Invalid email.');

  if (isWaitlistMode_()) {
    throw new Error('We’re not live yet. Please check back soon.');
  }  


  // lockout check (progressive)
  const { users, sessions, lockouts } = getDb_();
  const key = 'user:' + email;
  const status = getLockoutStatus_(lockouts, key);
  if (status.locked) {
    if (status.permanent) {
      throw new Error('This account is temporarily blocked. Please contact support.');
    }
    const msLeft = Number(status.until||0) - Date.now();
    const human = msToHuman_(msLeft > 0 ? msLeft : 0);
    throw new Error('Too many attempts. Try again in ' + human + '.');
  }

  if (!isValidPin_(pin)) {
    const res = recordSigninFailure_(lockouts, key);
    if (res && res.lockedUntil && res.lockedUntil > Date.now()) {
      const human = msToHuman_(res.lockedUntil - Date.now());
      throw new Error('Too many attempts. Try again in ' + human + '.');
    }
    throw new Error('Invalid PIN.');
  }

  const uRow = findRowByColumn_(users, 1, email);
  if (!uRow) {
    recordSigninFailure_(lockouts, key);
    throw new Error('Not registered.');
  }
  const [e, /*PinHash legacy*/, memberId] = users.getRange(uRow, 1, 1, 3).getValues()[0];

  // PBKDF2 verify (with seamless legacy migration)
  const ok = verifyUserPinAndMigrateIfNeeded_(uRow, pin).ok;
  if (!ok){
    recordSigninFailure_(lockouts, key);
    throw new Error('Wrong PIN.');
  }

  // success → clear lockout record
  clearLockout_(lockouts, key);

  const sid = genSid_();
  const now = Date.now(); // ← define 'now' before using it
  sessions.appendRow([sid, email, now, new Date(now).toISOString()]);
  return { sid, memberId };
}

function signinAttempt(email, pin){
  // Always return a structured result instead of throwing
  try {
    email = sanitizeEmail_(email);
    if (!email) return { ok:false, msg:'Invalid email.' };

    if (isWaitlistMode_()) {
      return { ok:false, msg:'We’re not live yet. Please check back soon.' };
    }

    const { users, sessions, lockouts } = getDb_();
    const key = 'user:' + email;

    // Check lock status first
    const status = getLockoutStatus_(lockouts, key);
    if (status.locked) {
      const msLeft = Number(status.until||0) - Date.now();
      return {
        ok:false,
        msg: status.permanent ? 'This account is temporarily blocked. Please contact support.'
                              : ('Too many attempts. Try again in ' + msToHuman_(msLeft > 0 ? msLeft : 0) + '.'),
        lock: {
          locked:true,
          permanent: !!status.permanent,
          untilMs: Number(status.until||0)
        }
      };
    }

    // Validate PIN format early (cheap)
    if (!isValidPin_(pin)) {
      const rec = recordSigninFailure_(lockouts, key); // increments + sets cooldown if threshold hit
      const lockInfo = getLockoutStatus_(lockouts, key);
      // Compute nextTarget from rules
      let nextTarget = null;
      for (let i=0;i<LOCK_RULES.length;i++){
        if (rec.failCount < LOCK_RULES[i].fails){ nextTarget = LOCK_RULES[i].fails; break; }
      }
      if (nextTarget === null && LOCK_RULES.length){
        nextTarget = LOCK_RULES[LOCK_RULES.length - 1].fails;
      }
      return {
        ok:false,
        msg:'Invalid PIN.',
        lock:{
          locked: !!lockInfo.locked,
          permanent: !!lockInfo.permanent,
          untilMs: Number(lockInfo.until||0),
          failCount: rec.failCount,
          nextTarget: nextTarget
        }
      };
    }

    // Look up user
    const uRow = findRowByColumn_(users, 1, email);
    if (!uRow) {
      const rec = recordSigninFailure_(lockouts, key);
      let nextTarget = null;
      for (let i=0;i<LOCK_RULES.length;i++){
        if (rec.failCount < LOCK_RULES[i].fails){ nextTarget = LOCK_RULES[i].fails; break; }
      }
      if (nextTarget === null && LOCK_RULES.length) nextTarget = LOCK_RULES[LOCK_RULES.length - 1].fails;
      return { ok:false, msg:'Not registered.', lock:{ locked:false, permanent:false, untilMs:0, failCount: rec.failCount, nextTarget } };
    }

    const [ , /*legacy*/, memberId] = users.getRange(uRow, 1, 1, 3).getValues()[0];

    // Verify PIN (PBKDF2, or migrate from legacy)
    const ok = verifyUserPinAndMigrateIfNeeded_(uRow, pin).ok;
    if (!ok){
      const rec = recordSigninFailure_(lockouts, key);
      const lockInfo = getLockoutStatus_(lockouts, key);
      let nextTarget = null;
      for (let i=0;i<LOCK_RULES.length;i++){
        if (rec.failCount < LOCK_RULES[i].fails){ nextTarget = LOCK_RULES[i].fails; break; }
      }
      if (nextTarget === null && LOCK_RULES.length) nextTarget = LOCK_RULES[LOCK_RULES.length - 1].fails;

      return {
        ok:false,
        msg:'Wrong PIN.',
        lock:{
          locked: !!lockInfo.locked,
          permanent: !!lockInfo.permanent,
          untilMs: Number(lockInfo.until||0),
          failCount: rec.failCount,
          nextTarget: nextTarget
        }
      };
    }

    // Success → clear lockout, create session
    clearLockout_(lockouts, key);
    const sid = genSid_();
    const now = Date.now();
    sessions.appendRow([sid, email, now, new Date(now).toISOString()]);
    return { ok:true, sid, memberId };

  } catch (e){
    return { ok:false, msg: (e && e.message) || 'Sign in failed.' };
  }
}


function validateSession(sid) {
  if (!sid) return { ok: false };
  const { users, sessions } = getDb_();
  const sRow = findRowByColumn_(sessions, 1, sid);
  if (!sRow) return { ok: false };
  const [, email, lastSeenMs] = sessions.getRange(sRow, 1, 1, 4).getValues()[0];
  const now = Date.now();
  if (now - Number(lastSeenMs || 0) > SESSION_TTL_MS) {
    sessions.deleteRow(sRow);
    return { ok: false, reason: 'expired' };
  }
  // Touch ONLY — NO SID ROTATION
  sessions.getRange(sRow, 3).setValue(now);

  // Return memberId and SAME sid
  const uRow = findRowByColumn_(users, 1, email);
  const memberId = uRow ? users.getRange(uRow, 3).getValue() : null;
  return { ok: true, memberId, sid };
}

function logout(sid) {
  if (!sid) return { ok: true };
  const { sessions } = getDb_();
  const sRow = findRowByColumn_(sessions, 1, sid);
  if (sRow) sessions.deleteRow(sRow);
  return { ok: true };
}

// ---- Deep-link tokens (Option B) ----
// Client: create short-lived token URL for this member (rate-limited / reuse)
function createLinkToken(sid) {
  const v = validateSession(sid); // does not rotate
  if (!v || !v.ok) throw new Error('Not signed in.');
  const { linkTokens, logs } = getDb_();
  const now = Date.now();

  // Reuse last usable token: same member, not Used, not expired, and created within last 60s
  const prev = findLastTokenForMember_(linkTokens, v.memberId);
  if (prev && !prev.used && Number(prev.expiresAtMs || 0) > now && (now - Number(prev.createdAtMs || 0)) < 60 * 1000) {
    const base = SP.getProperty(MERCHANT_APP_BASE_URL_KEY) || '';
    const urlPrev = base ? (base + '?t=' + encodeURIComponent(prev.token)) : '';
    return { ok: true, url: urlPrev, token: prev.token, expiresAtMs: Number(prev.expiresAtMs) };
  }

  // Create fresh
  const token = 'LT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  const exp = now + LINK_TOKEN_TTL_MS;
  // [Token, MemberId, Mode, ExpiresAtMs, Used, CreatedAt]
  linkTokens.appendRow([token, v.memberId, '', exp, false, new Date(now).toISOString()]);
  log_(logs, 'create_token', 'member=' + v.memberId);
  const base = SP.getProperty(MERCHANT_APP_BASE_URL_KEY) || '';
  const url = base ? (base + '?t=' + encodeURIComponent(token)) : '';
  return { ok: true, url, token, expiresAtMs: exp };
}

// ---- DB / Sheets ----
function getDb_() {
  if (__DB_CACHE) return __DB_CACHE;

  let id = SP.getProperty(DB_ID_KEY);
  let ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create(DB_NAME);
  if (!id) SP.setProperty(DB_ID_KEY, ss.getId());

  // Ensure sheets only once per invocation; reuse handles thereafter
  const users      = ensureSheet_(ss, USERS_SHEET,       ['Email','PinHash','MemberId','CreatedAt']);
  const sessions   = ensureSheet_(ss, SESSIONS_SHEET,    ['SID','Email','LastSeenMs','CreatedAt']);
  const merchants  = ensureSheet_(ss, MERCHANTS_SHEET,   ['MerchantId','Name','Active','Secret','CreatedAt']);
  const tx         = ensureSheet_(ss, TX_SHEET,          ['TxId','MemberId','MerchantId','Type','Points','AtMs','Staff']);
  const balances   = ensureSheet_(ss, BALANCES_SHEET,    ['MemberId','Points']);
  const linkTokens = ensureSheet_(ss, LINK_TOKENS_SHEET, ['Token','MemberId','Mode','ExpiresAtMs','Used','CreatedAt']);
  const lockouts   = ensureSheet_(ss, LOCKOUTS_SHEET,    ['Key','FailCount','LockUntilMs','UpdatedAt','FirstFailMs','TotalFails24h','PermLocked']);
  const logs       = ensureSheet_(ss, LOGS_SHEET,        ['At','Type','Message']);
  const resets     = ensureSheet_(ss, RESET_TOKENS_SHEET,['Token','Email','ExpiresAtMs','Used','CreatedAt','VerifiedAtMs']);
  const magic      = ensureSheet_(ss, MAGIC_TOKENS_SHEET,['Token','Email','ExpiresAtMs','Used','CreatedAt']);
  const waitlist   = ensureSheet_(ss, WAITLIST_SHEET,    ['At','Name','Email','PhoneE164','City','Country','Notes','Source','MemberId']);
  const pending    = ensureSheet_(ss, PENDING_REG_SHEET, [
    'PendingId','CreatedAtMs','Status',
    'Email','PhoneE164','WaE164','PrefComms',
    'OTP','OtpExpiresAtMs',
    'First','Last','BirthYmd','Gender','City','Country','Lang',
    'Referral','OptMarketing',
    'UTM_Source','UTM_Medium','UTM_Campaign','UTM_Term','UTM_Content',
    'CompletedAtMs','MemberId'
  ]);
  const config     = ensureSheet_(ss, CONFIG_SHEET,      ['Key','Value']);

  __DB_CACHE = {
    ss,
    users,
    sessions,
    merchants,
    tx,
    balances,
    linkTokens,
    lockouts,
    logs,
    resets,
    magic,
    waitlist,
    pending,
    config
  };
  return __DB_CACHE;
}

function getAppVersion_() {
  try {
    const { config } = getDb_();
    const last = config.getLastRow();
    if (last < 2) return 'dev';

    const vals = config.getRange(2, 1, last - 1, 2).getValues(); // Key, Value
    for (let i = 0; i < vals.length; i++) {
      const key = String(vals[i][0] || '').trim().toUpperCase();
      if (key === 'APP_VERSION') {
        const v = String(vals[i][1] || '').trim();
        return v || 'dev';
      }
    }
    return 'dev';
  } catch (e) {
    return 'dev';
  }
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  // Ensure header row has at least these headers (don’t shrink if more exist)
  const h = sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn()||headers.length)).getValues()[0];
  if (h.every(v => !v)) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    // fill any missing header cells to the right (non-destructive)
    for (let i=0;i<headers.length;i++){ if (!h[i]) sh.getRange(1, i+1).setValue(headers[i]); }
  }
  return sh;
}

// ---- Utils ----
function sanitizeEmail_(email) {
  if (!email) return null;
  const e = String(email).trim().toLowerCase();
  if (!/@/.test(e)) return null;
  return e;
}
function isValidPin_(pin) { return typeof pin === 'string' && /^[0-9]{6}$/.test(pin); }
function hashPin_(pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin);
  return bytes.map(b => (b + 256) % 256).map(b => ('0' + b.toString(16)).slice(-2)).join('');
}
function genMemberId_() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += alphabet[Math.floor(Math.random() * alphabet.length)];
  return `MBR-${s}`;
}
function genSid_() { return `s_${Utilities.getUuid().replace(/-/g, '')}`; }

function findRowByColumn_(sh, colIndex, value) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, colIndex, last - 1, 1).getValues();
  const needle = String(value).toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === needle) return 2 + i;
  }
  return 0;
}
// last-match by column (for token reuse)
function findLastRowByColumn_(sh, colIndex, value) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, colIndex, last - 1, 1).getValues();
  const needle = String(value).toLowerCase();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).toLowerCase() === needle) return 2 + i;
  }
  return 0;
}
function findLastTokenForMember_(linkTokens, memberId) {
  const r = findLastRowByColumn_(linkTokens, 2, memberId);
  if (!r) return null;
  const [tok, mid, mode, expMs, used, createdAt] = linkTokens.getRange(r,1,1,6).getValues()[0];
  return {
    token: String(tok||''),
    memberId: String(mid||''),
    mode: String(mode||''),
    expiresAtMs: Number(expMs||0),
    used: String(used).toLowerCase() === 'true',
    createdAtMs: Date.parse(createdAt || '') || 0
  };
}
function maskEmail_(email) {
  const e = String(email || '');
  const parts = e.split('@');
  if (parts.length !== 2) return '***';
  const name = parts[0];
  const domain = parts[1];
  const maskedName = name.length <= 2 ? (name[0] || '*') + '*' : name[0] + '***' + name.slice(-1);
  return maskedName + '@' + domain;
}


// ---- Signed QR helpers ----
// Script Properties needed (set in User App):
//   QR_SIGNING_KEY_ID = 'v1'
//   QR_SIGNING_SECRET_B64 = '<base64 random 32-64 bytes>'
// Config sheet (optional override):
//   QR_TTL_SEC = '180'

function base64urlFromBytes_(bytes){
  const b64 = Utilities.base64Encode(bytes);
  return b64.replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
}
function getConfigValue_(key, fallback){
  try{
    const { config } = getDb_();
    const last = config.getLastRow();
    if (last >= 2) {
      const vals = config.getRange(2,1,last-1,2).getValues();
      for (var i=0;i<vals.length;i++){
        if (String(vals[i][0]||'').trim().toUpperCase() === String(key||'').trim().toUpperCase()){
          const v = String(vals[i][1]||'').trim();
          return v || fallback;
        }
      }
    }
  }catch(_){}
  return fallback;
}

// --- ADD: load the same key the Merchant reads (from Config: QR_HMAC_KEY_B64)
function getQrHmacKeyBytes_() {
  const b64 = getConfigValue_('QR_HMAC_KEY_B64', '');
  if (!b64) throw new Error('QR_HMAC_KEY_B64 missing in Config sheet.');
  return Utilities.base64Decode(b64); // STANDARD Base64
}

// --- ADD: build base64url(JSON) payload expected by Merchant resolveSignedQrPayload()
function buildMerchantQrUrl_(memberId, mode /* 'COLLECT'|'REDEEM'|'' */, expMs, webAppExecUrl) {
  const mid = String(memberId || '').toUpperCase();
  if (!/^MBR-[A-Z0-9]{8}$/.test(mid)) throw new Error('Bad memberId format.');

  const m = String(mode || '').toUpperCase();
  const modeNorm = (m === 'COLLECT' || m === 'REDEEM') ? m : '';

  const exp = Number(expMs || (Date.now() + 3 * 60 * 1000)); // default +3 min TTL

  // Canonical string "v|mid|mode|exp"
  const canonical = ['1', mid, modeNorm, String(exp)].join('|');

  // HMAC-SHA256 over canonical using QR_HMAC_KEY_B64
  const keyBytes = getQrHmacKeyBytes_();
  const macBytes = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(canonical).getBytes(),
    keyBytes
  );
  const sigB64u = base64urlFromBytes_(macBytes);

  // JSON payload the Merchant expects
  const obj = { v: 1, mid: mid, mode: modeNorm, exp: exp, sig: sigB64u };
  const jsonBytes = Utilities.newBlob(JSON.stringify(obj)).getBytes();
  const payloadB64u = base64urlFromBytes_(jsonBytes);

  // Build final URL
  const base = SP.getProperty('MERCHANT_APP_BASE_URL') || SP.getProperty(MERCHANT_APP_BASE_URL_KEY) || webAppExecUrl || '';
  if (!base) return payloadB64u; // return raw payload if you want to embed yourself
  const sep = base.includes('?') ? '&' : '?';
  return base + sep + 'p=' + encodeURIComponent(payloadB64u);
}

function getSignedQr(sid){
  const v = validateSession(sid);
  if (!v || !v.ok) throw new Error('Not signed in.');

  // TTL in seconds (defaults to 180), clamp to a sane range just in case
  let ttlSec = Number(getConfigValue_('QR_TTL_SEC','180')) || 180;
  if (!Number.isFinite(ttlSec) || ttlSec <= 0) ttlSec = 180;

  const exp = Date.now() + ttlSec * 1000;

  // IMPORTANT: mode is '' → staff decides on the merchant device
  const url = buildMerchantQrUrl_(v.memberId, '', exp, null);

  // If base URL is set, `url` is a full URL (...?p=payload).
  // If not, `url` is already the raw payload string — the replace() is harmless.
  const qrPayload = url.replace(/^.*\?p=/, '');

  return {
    ok: true,
    url,
    qr: qrPayload,
    ttlSec,
    keyId: 'v1'
  };
}


// —— PBKDF2 helpers (same approach we used for Merchant/Admin) ——
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
  // Apps Script requires the *data first*, then the *key* and the correct method name is ...Signature
  return Utilities.computeHmacSha256Signature(dataBytes, keyBytes);
}

function xorBytes_(a,b){
  const out = new Array(Math.min(a.length,b.length));
  for (let i=0;i<out.length;i++) out[i] = (a[i]^b[i]) & 0xff;
  return out;
}
// PBKDF2-HMAC-SHA256 minimal impl; dkLen fixed to 32 bytes
function pbkdf2Sha256_(passwordBytes, saltBytes, iterations, dkLen){
  const hLen = 32; const l = Math.ceil(dkLen / hLen);
  const r = dkLen - (l-1)*hLen; const dk = [];
  for (let i=1;i<=l;i++){
    const blockIndex = [0,0,0,(i & 0xff)];
    const u1 = hmacSha256_(passwordBytes, saltBytes.concat(blockIndex));
    let t = u1.slice(); let u = u1;
    for (let c=2;c<=iterations;c++){ u = hmacSha256_(passwordBytes, u); t = xorBytes_(t,u); }
    const outPart = (i===l) ? t.slice(0, r) : t;
    for (let j=0;j<outPart.length;j++) dk.push(outPart[j]);
  }
  return dk;
}
function derivePinHashB64_(pin, saltB64, iter, usePepper){
  const salt = b64dec_(saltB64);
  const pinBytes = utf8Bytes_(String(pin||''));
  const pepperBytes = (usePepper && PIN_PEPPER_B64) ? b64dec_(PIN_PEPPER_B64) : [];
  const pwd = pinBytes.concat(pepperBytes);
  const dk = pbkdf2Sha256_(pwd, salt, Math.max(1000, Number(iter)||PIN_ITER_DEFAULT), 32);
  return b64enc_(dk);
}

function nextCooldownFor_(failCount){
  // Return {durMin, permanent} for a given failCount
  let chosen = null;
  for (let i=0;i<LOCK_RULES.length;i++){
    if (failCount >= LOCK_RULES[i].fails) chosen = LOCK_RULES[i];
  }
  if (!chosen) return { durMin: 0, permanent: false };
  return { durMin: chosen.durMin, permanent: (chosen.durMin === -1) };
}
function msToHuman_(ms){
  // Rough humanization for messages
  if (ms <= 0) return 'now';
  const m = Math.round(ms/60000);
  if (m < 60) return m + ' minute' + (m===1?'':'s');
  const h = Math.round(m/60);
  if (h < 48) return h + ' hour' + (h===1?'':'s');
  const d = Math.round(h/24);
  return d + ' day' + (d===1?'':'s');
}

function getAppBaseUrl_() {
  // Optional override via script property; otherwise use the deployed web app URL.
  const prop = SP.getProperty('APP_BASE_URL');
  if (prop) return prop;
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (_) {
    return '';
  }
}
function genPid_(){ return 'PID-' + Utilities.getUuid().replace(/-/g,'').slice(0,20); }
function genOtp_(){ return String(Math.floor(100000 + Math.random()*900000)); } // 6-digit

function isValidLebanonLocal8_(s){ return /^[0-9]{8}$/.test(String(s||'')); }
function toE164Lebanon_(local8){ return '+961' + String(local8||''); }

function normalizeLebanonOrEmpty_(s){
  const t = String(s||'').replace(/\D/g,'');
  if (!t) return '';
  // allow "961xxxxxxxx" or "0xxxxxxxx" or just "xxxxxxxx"
  let last8 = '';
  if (t.length === 8) last8 = t;
  else if (t.length === 9 && t[0] === '0') last8 = t.slice(1);
  else if (t.length === 11 && t.startsWith('961')) last8 = t.slice(3);
  else return ''; // not a Lebanon 8-digit
  return toE164Lebanon_(last8);
}

function isAtLeast16_(birthYmd){
  // birthYmd: 'YYYY-MM-DD'
  const m = String(birthYmd||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return false;
  const d = new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  if (isNaN(d.getTime())) return false;
  const now = new Date();
  const sixteen = new Date(d.getFullYear()+16, d.getMonth(), d.getDate());
  return now >= sixteen;
}

function extendUserRow_(u){
  // Ensure Users sheet has columns we’ll write (it won’t shrink existing)
  const { users } = getDb_();
  const needed = ['Email','PinHash','MemberId','CreatedAt',
    'First','Last','BirthYmd','Gender','City','Country','Lang',
    'PhoneE164','WaE164','PrefComms','Referral','OptMarketing','UTM_Source','UTM_Medium','UTM_Campaign','UTM_Term','UTM_Content'
  ];
  const lastCol = Math.max(needed.length, users.getLastColumn()||needed.length);
  const hdr = users.getRange(1,1,1,lastCol).getValues()[0];
  for (let i=0;i<needed.length;i++){
    if (!hdr[i]) users.getRange(1, i+1).setValue(needed[i]);
  }
}
function usersHeaderIdx_(sheet){
  const lastCol = Math.max(sheet.getLastColumn()||0, 30);
  const hdr = sheet.getRange(1,1,1,lastCol).getValues()[0].map(String);
  const idx = {}; hdr.forEach((h,i)=>{ if (h) idx[h.trim()] = i+1; });
  return idx;
}
function ensureUsersSecurityHeaders_(){
  const { users } = getDb_();
  const needed = ['Salt','Hash','Iter','Algo','Peppered','UpdatedAt']; // non-destructive
  const lastCol = Math.max(users.getLastColumn()||0, needed.length);
  const hdr = users.getRange(1,1,1,lastCol).getValues()[0];
  // append/fill to the right, keeping existing columns intact
  const map = usersHeaderIdx_(users);
  for (let i=0;i<needed.length;i++){
    if (!map[needed[i]]) {
      users.getRange(1, (users.getLastColumn()||hdr.length)+1).setValue(needed[i]);
    }
  }
}
function setUserSecurity_(row, pin){
  const { users } = getDb_();
  ensureUsersSecurityHeaders_();
  const idx = usersHeaderIdx_(users);
  const saltB64 = randomSaltB64_(24);
  const iter = PIN_ITER_DEFAULT;
  const algo = PIN_HASH_ALGO;  // liest ggf. dein Property, sonst Default
  const peppered = !!PIN_PEPPER_B64;
  const hashB64 = derivePinHashB64_(pin, saltB64, iter, peppered);

  if (idx['Salt'])     users.getRange(row, idx['Salt']).setValue(saltB64);
  if (idx['Hash'])     users.getRange(row, idx['Hash']).setValue(hashB64);
  if (idx['Iter'])     users.getRange(row, idx['Iter']).setValue(iter);
  if (idx['Algo'])     users.getRange(row, idx['Algo']).setValue(algo);
  if (idx['Peppered']) users.getRange(row, idx['Peppered']).setValue(peppered);
  if (idx['UpdatedAt'])users.getRange(row, idx['UpdatedAt']).setValue(new Date().toISOString());

  // Optional: clear legacy PinHash for brand-new accounts
  // const PINHASH_COL = 2; // [Email, PinHash, MemberId,...] in your current layout
  // users.getRange(row, PINHASH_COL).clearContent();
}
function verifyUserPinAndMigrateIfNeeded_(row, pin){
  const { users } = getDb_();
  const idx = usersHeaderIdx_(users);

  const get = (name)=> { const c = idx[name]; return c ? users.getRange(row, c).getValue() : ''; };
  const algo = String(get('Algo')||'');
  const saltB64 = String(get('Salt')||'');
  const hashB64 = String(get('Hash')||'');
  const iter = Number(get('Iter')||0);
  const peppered = String(get('Peppered')||'').toLowerCase()==='true';

  if (algo && saltB64 && hashB64 && iter){
    const calc = derivePinHashB64_(pin, saltB64, iter, peppered);
    return { ok: (calc === String(hashB64)), migrated:false };
  }

  // Legacy fallback to plain SHA-256 PinHash (column 2)
  const legacyHash = String(users.getRange(row, 2).getValue()||'');
  if (legacyHash){
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin||''));
    const hex = bytes.map(b => (b + 256) % 256).map(b => ('0' + b.toString(16)).slice(-2)).join('');
    const okLegacy = (hex === legacyHash);
    if (okLegacy){
      // Migrate immediately
      setUserSecurity_(row, pin);
      return { ok:true, migrated:true };
    }
  }
  return { ok:false, migrated:false };
}
function duplicateExists_(email, phoneE164, waE164){
  const { users } = getDb_();
  const last = users.getLastRow();
  if (last < 2) return false;
  const cols = Math.max(users.getLastColumn()||0, 20);
  const vals = users.getRange(2,1,last-1,cols).getValues();
  const emailSet = new Set(), phoneSet = new Set(), waSet = new Set();
  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    const e = String(row[0]||'').toLowerCase();
    if (e) emailSet.add(e);
    const phone = String(row[12]||''); // PhoneE164 (after extendUserRow_)
    if (phone) phoneSet.add(phone);
    const wa = String(row[13]||'');    // WaE164
    if (wa) waSet.add(wa);
  }
  if (email && emailSet.has(String(email).toLowerCase())) return 'email';
  if (phoneE164 && phoneSet.has(phoneE164)) return 'phone';
  if (waE164 && waSet.has(waE164)) return 'whatsapp';
  return false;
}

// Maintenance
function purgeOldSessions(maxAgeMinutes) {
  const ageMs = (Number(maxAgeMinutes) > 0 ? Number(maxAgeMinutes) : 120) * 60 * 1000;
  const { sessions } = getDb_();
  const last = sessions.getLastRow();
  if (last < 2) return 0;
  const now = Date.now();
  const vals = sessions.getRange(2, 1, last - 1, 4).getValues();
  const toDelete = [];
  for (let i = 0; i < vals.length; i++) {
    const lastSeen = Number(vals[i][2]) || 0;
    if (!lastSeen || (now - lastSeen) > ageMs) toDelete.push(2 + i);
  }
  for (let i = toDelete.length - 1; i >= 0; i--) sessions.deleteRow(toDelete[i]);
  return toDelete.length;
}

function purgeOldResetTokens(maxAgeMinutes) {
  // Deletes tokens that are expired for > maxAgeMinutes OR are already Used.
  // Default = 60 minutes.
  const ageMs = (Number(maxAgeMinutes) > 0 ? Number(maxAgeMinutes) : 60) * 60 * 1000;
  const { resets } = getDb_();
  const last = resets.getLastRow();
  if (last < 2) return 0;
  const now = Date.now();
  const vals = resets.getRange(2, 1, last - 1, 6).getValues(); // Token, Email, ExpMs, Used, CreatedAt, VerifiedAtMs
  const toDel = [];
  for (let i = 0; i < vals.length; i++) {
    const expMs = Number(vals[i][2]) || 0;
    const used  = String(vals[i][3]).toLowerCase() === 'true';
    if ((expMs && (now - expMs) > ageMs) || used) toDel.push(2 + i);
  }
  for (let i = toDel.length - 1; i >= 0; i--) resets.deleteRow(toDel[i]);
  return toDel.length;
}

// Helpers for merchants/points
function getOrInitBalance_(balances, memberId) {
  const row = findRowByColumn_(balances, 1, memberId);
  if (row) return { row, points: Number(balances.getRange(row, 2).getValue() || 0) };
  balances.appendRow([memberId, 0]);
  const newRow = balances.getLastRow();
  return { row: newRow, points: 0 };
}
function merchantExistsAndActive_(merchants, merchantId) {
  const r = findRowByColumn_(merchants, 1, merchantId);
  if (!r) return false;
  const active = String(merchants.getRange(r, 3).getValue() || 'TRUE').toUpperCase() !== 'FALSE';
  return active;
}
function appendTx_(txSheet, memberId, merchantId, type, points, staff) {
  const txId = Utilities.getUuid();
  const now = Date.now();
  // include Staff (audit)
  txSheet.appendRow([txId, memberId, merchantId, type, Number(points||0), now, staff || '']);
  return { txId, atMs: now };
}

// Public APIs kept (Forgot Pin/getMemberSummary/collectPoints/redeemPoints)
// ───── Forgot PIN (Users) ─────

// Request a 30-min reset token for an email.
// Returns the token so the client can show/copy a reset link (or email it if you prefer).
function requestPinReset(email) {
  email = sanitizeEmail_(email);
  if (!email) throw new Error('Enter a valid email.');
  if (isWaitlistMode_()) {
    throw new Error('We’re not live yet. Please check back soon.');
  }
  const { users, resets, logs } = getDb_();

  const r = findRowByColumn_(users, 1, email);
  // Always behave the same to avoid account enumeration
  if (!r) {
    log_(logs, 'pinreset_request_unknown', maskEmail_(email));
    return { ok: true };
  }

  // Simple throttle: max 3 active (unexpired+unused) per 30 minutes
  const now = Date.now();
  const last = resets.getLastRow();
  if (last >= 2) {
    const vals = resets.getRange(2, 1, last - 1, 4).getValues(); // Token, Email, ExpiresAtMs, Used
    const active = vals.filter(v =>
      String(v[1]).toLowerCase() === email &&
      Number(v[2] || 0) > now &&
      String(v[3]).toLowerCase() !== 'true'
    ).length;
    if (active >= 3) throw new Error('Too many recent requests. Try again later.');
  }

  // Create token
  const token = 'RT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 24);
  const exp   = now + RESET_TOKEN_TTL_MS;
  resets.appendRow([token, email, exp, false, new Date(now).toISOString(), null]);
  log_(logs, 'pinreset_request_ok', maskEmail_(email));

  // Build reset URL pointing back to THIS web app
  const baseUrl = getAppBaseUrl_();
  const url = baseUrl ? (baseUrl + '?reset=' + encodeURIComponent(token)) : '';

  // Email it (HTML + plain text)
  const subj = 'Your PIN reset link (valid 30 minutes)';
  const text = 'Use the following link to reset your PIN. It expires in 30 minutes:\n\n' + url;
  const html =
    '<p>Click the button below to reset your PIN. This link expires in <b>30 minutes</b>.</p>' +
    '<p><a href="' + url + '" ' +
    'style="display:inline-block;padding:10px 14px;border-radius:6px;background:#1a73e8;color:#fff;text-decoration:none;">' +
    'Reset PIN</a></p>' +
    '<p>If the button doesn’t work, copy and paste this URL:<br>' +
    '<code>' + url + '</code></p>';

  // NOTE: the first time you run this, Apps Script will ask for MailApp authorization.
  try {
    MailApp.sendEmail({
      to: email,
      subject: subj,
      body: text,
      htmlBody: html
    });
  } catch (e) {
    // Fail closed (don’t leak the URL). Still return ok to avoid enumeration.
    log_(logs, 'pinreset_email_fail', maskEmail_(email) + ' :: ' + (e && e.message));
  }

  // Do NOT return the token or URL
  return { ok: true };
}

// —— MAGIC LINK SIGN-IN ——
// 1) Request a magic link (email must exist). Always respond ok to avoid enumeration.
function requestMagicLink(email){
  email = sanitizeEmail_(email);
  if (!email) throw new Error('Enter a valid email.');
  if (isWaitlistMode_()) throw new Error('We’re not live yet. Please check back soon.');

  const { users, magic, logs } = getDb_();
  const r = findRowByColumn_(users, 1, email);
  if (!r) { log_(logs, 'magic_req_unknown', maskEmail_(email)); return { ok:true }; }

  const token = 'ML-' + Utilities.getUuid().replace(/-/g,'').slice(0,24);
  const exp   = Date.now() + MAGIC_TOKEN_TTL_MS;
  magic.appendRow([token, email, exp, false, new Date().toISOString()]);
  log_(logs, 'magic_req_ok', maskEmail_(email));

  // Build URL back to THIS app with ?magic=
  const base = getAppBaseUrl_();
  const url  = base ? (base + '?magic=' + encodeURIComponent(token)) : '';

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Your sign-in link (valid 10 minutes)',
      htmlBody:
        '<p>Click to sign in:</p>' +
        '<p><a href="'+url+'" style="display:inline-block;padding:10px 14px;border-radius:6px;background:#1a73e8;color:#fff;text-decoration:none;">Sign in</a></p>' +
        '<p>This link expires in 10 minutes.</p>',
      body: 'Open this link to sign in (valid 10 minutes):\n\n'+url
    });
  } catch(e){
    log_(logs, 'magic_email_fail', maskEmail_(email) + ' :: ' + (e && e.message));
  }
  return { ok:true };
}

// 2) Complete magic-link sign-in and mint a session
function completeMagicSignin(token){
  token = String(token||'').trim();
  if (!token) throw new Error('Invalid link.');

  const { magic, users, sessions, logs } = getDb_();
  const row = findRowByColumn_(magic, 1, token);
  if (!row) throw new Error('Invalid or expired link.');

  const [tok, email, expMs, used] = magic.getRange(row,1,1,4).getValues()[0];
  const now = Date.now();
  if (String(used).toLowerCase() === 'true') throw new Error('Link already used.');
  if (Number(expMs||0) < now) throw new Error('Link expired.');

  // Locate user (should exist)
  const rUser = findRowByColumn_(users, 1, String(email||''));
  if (!rUser) throw new Error('Account not found.');

  // Mark token used
  magic.getRange(row, 4).setValue(true);

  // Create session
  const sid = genSid_();
  sessions.appendRow([sid, String(email||''), now, new Date(now).toISOString()]);

  // Return memberId
  const memberId = users.getRange(rUser, 3).getValue();
  log_(logs, 'magic_ok', maskEmail_(String(email||'')));
  return { ok:true, sid, memberId };
}

// Housekeeping
function purgeOldMagicTokens(maxAgeMinutes){
  const ageMs = (Number(maxAgeMinutes)>0 ? Number(maxAgeMinutes) : 60) * 60 * 1000;
  const { magic } = getDb_();
  const last = magic.getLastRow();
  if (last < 2) return 0;
  const now = Date.now();
  const vals = magic.getRange(2,1,last-1,4).getValues(); // Token, Email, ExpiresAtMs, Used
  const toDel = [];
  for (let i=0;i<vals.length;i++){
    const expMs = Number(vals[i][2])||0;
    const used  = String(vals[i][3]).toLowerCase()==='true';
    if (used || (expMs && (now - expMs) > ageMs)) toDel.push(2+i);
  }
  for (let i=toDel.length-1;i>=0;i--) magic.deleteRow(toDel[i]);
  return toDel.length;
}

// Verify token (for client to pre-check before showing reset form)
function verifyResetToken(token) {
  token = String(token || '').trim();
  if (!token) return { ok:false };
  const { resets } = getDb_();
  const r = findRowByColumn_(resets, 1, token);
  if (!r) return { ok:false };
  const rowVals = resets.getRange(r,1,1,6).getValues()[0]; // [Token, Email, ExpiresAtMs, Used, CreatedAt, VerifiedAtMs]
  const [, email, expMs, used, , verifiedAtMs] = rowVals;
  const now = Date.now();
  if (String(used).toLowerCase() === 'true') return { ok:false, reason:'used' };
  if (Number(expMs||0) < now) return { ok:false, reason:'expired' };

  // Mark as verified (first time only)
  const already = Number(verifiedAtMs||0);
  if (!already) resets.getRange(r, 6).setValue(now); // VerifiedAtMs

  return { ok:true, email: String(email||'') };
}

// Complete reset: set a new 6-digit PIN and invalidate token + sessions.
function completePinReset(token, newPin) {
  token = String(token || '').trim();
  if (!isValidPin_(newPin)) throw new Error('PIN must be exactly 6 digits.');

  const { users, sessions, resets, logs } = getDb_();
  const rTok = findRowByColumn_(resets, 1, token);
  if (!rTok) throw new Error('Invalid or expired link.');
  const rowVals = resets.getRange(rTok,1,1,6).getValues()[0]; // [Token, Email, ExpiresAtMs, Used, CreatedAt, VerifiedAtMs]
  const [, email, expMs, used, , verifiedAtMs] = rowVals;
  const now = Date.now();
  if (String(used).toLowerCase() === 'true') throw new Error('Link already used.');
  if (Number(expMs||0) < now) throw new Error('Link expired.');

  // Require that the link was "opened" (verified) recently (<= 5 minutes)
  const VERIFY_WINDOW_MS = 5 * 60 * 1000;
  if (!Number(verifiedAtMs || 0) || (now - Number(verifiedAtMs)) > VERIFY_WINDOW_MS) {
    throw new Error('Verification window elapsed. Please reopen your reset link.');
  }

  // Update user's PIN
  const rUser = findRowByColumn_(users, 1, email);
  if (!rUser) throw new Error('Account not found.');
  // Write modern PBKDF2 security fields for the NEW PIN
  setUserSecurity_(rUser, String(newPin));

  // Clear legacy PinHash to prevent any fallback path using the old PIN
  try { users.getRange(rUser, 2).clearContent(); } catch(_){}



  // Invalidate ALL active sessions for this email (force sign-in)
  const sLast = sessions.getLastRow();
  if (sLast >= 2) {
    const sVals = sessions.getRange(2,1,sLast-1,4).getValues(); // SID, Email, LastSeenMs, CreatedAt
    for (let i = sVals.length - 1; i >= 0; i--) {
      if (String(sVals[i][1]).toLowerCase() === String(email).toLowerCase()) {
        sessions.deleteRow(2 + i);
      }
    }
  }

  // Mark token used
  resets.getRange(rTok, 4).setValue(true); // Used = TRUE

  // Clear any sign-in lockout for this account (fresh start)
  const { lockouts } = getDb_();
  try { clearLockout_(lockouts, 'user:' + String(email||'').toLowerCase()); } catch(_){}

  log_(logs, 'pinreset_complete', maskEmail_(email));
  return { ok:true };
}
/**
 * Start staged registration: validates, stores pending row, generates OTP, emails if needed.
 * @param {Object} data - payload from client
 */
function startRegistration(data){
  if (isWaitlistMode_()) {
    throw new Error('We’re not live yet. Please check back soon.');
  }
  const {
    first, last, email, phoneLocal8, waLocal8, prefComms,
    birthYmd, gender, city, country, lang,
    referral, optMarketing,
    pin1, pin2,
    utm_source, utm_medium, utm_campaign, utm_term, utm_content
  } = data || {};

  // Requireds
  const e = sanitizeEmail_(email);
  if (!e) throw new Error('Enter a valid email.');
  if (!first || !last) throw new Error('Enter first and last name.');
  if (!birthYmd || !isAtLeast16_(birthYmd)) throw new Error('You must be at least 16.');
  if (!/^(email|sms|whatsapp)$/i.test(String(prefComms||''))) throw new Error('Choose a communication method.');

  // PIN
  if (!/^[0-9]{6}$/.test(String(pin1||'')) || pin1 !== pin2) throw new Error('PINs must match and be 6 digits.');

  // Phone
  if (!isValidLebanonLocal8_(phoneLocal8)) throw new Error('Phone must be 8 digits (Lebanon).');
  const phoneE164 = toE164Lebanon_(phoneLocal8);

  // WhatsApp (optional, but if provided must be valid 8 digits)
  let waE164 = '';
  if (String(waLocal8||'').trim()){
    if (!isValidLebanonLocal8_(waLocal8)) throw new Error('WhatsApp must be 8 digits (Lebanon).');
    waE164 = toE164Lebanon_(waLocal8);
  }

  // Dedupe against Users
  const dupe = duplicateExists_(e, phoneE164, waE164);
  if (dupe === 'email') throw new Error('This email is already registered.');
  if (dupe === 'phone') throw new Error('This phone number is already registered.');
  if (dupe === 'whatsapp') throw new Error('This WhatsApp number is already registered.');

  // Create Pending
  const pid = genPid_();
  const otp = genOtp_();
  const now = Date.now();
  const exp = now + OTP_TTL_MS;

  const { pending, logs } = getDb_();
  pending.appendRow([
    pid, now, 'pending',
    e, phoneE164, waE164, String(prefComms||'').toLowerCase(),
    otp, exp,
    String(first||''), String(last||''), String(birthYmd||''), String(gender||''),
    String(city||''), String(country||''), String(lang||''),
    String(referral||''), String(optMarketing||'').toLowerCase()==='true',
    String(utm_source||''), String(utm_medium||''), String(utm_campaign||''), String(utm_term||''), String(utm_content||''),
    '', '' // CompletedAtMs, MemberId
  ]);
  log_(logs, 'reg_start', 'pid=' + pid + ' email=' + maskEmail_(e) + ' pref=' + String(prefComms||''));
  // NEW: log coarse geolocation captured at registration
  log_(logs, 'reg_geo', 'pid=' + pid + ' city=' + String(city||'') + ' country=' + String(country||''));


  // Send OTP if email
  if (String(prefComms||'').toLowerCase() === 'email'){
    try {
      MailApp.sendEmail({
        to: e,
        subject: 'Your OTP (valid 15 minutes)',
        htmlBody: '<p>Your OTP code is:</p><p style="font-size:20px;"><b>' + otp + '</b></p><p>This code expires in 15 minutes.</p>',
        body: 'Your OTP: ' + otp + ' (valid 15 minutes)'
      });
      log_(logs, 'otp_sent_email', maskEmail_(e));
    } catch(err){
      log_(logs, 'otp_email_fail', (err && err.message) || 'error');
      throw new Error('Could not send OTP email right now.');
    }
  } else {
    // Not implemented (SMS/WhatsApp); test via reg_peek
    log_(logs, 'otp_placeholder_'+String(prefComms||'').toLowerCase(), 'pid=' + pid);
  }

  return { ok:true, pendingId: pid, method: String(prefComms||'').toLowerCase(), expiresAtMs: exp };
}

function resendRegistrationOtp(pendingId){
  pendingId = String(pendingId||'').trim();
  if (!pendingId) throw new Error('Missing pending ID.');
  const { pending, logs } = getDb_();
  const r = findRowByColumn_(pending, 1, pendingId);
  if (!r) throw new Error('Not found.');
  const row = pending.getRange(r,1,1,pending.getLastColumn()).getValues()[0];
  const status = String(row[2]||'');
  const email = String(row[3]||'');
  const pref  = String(row[6]||'').toLowerCase();
  if (status !== 'pending') throw new Error('Registration is not pending.');

  const otp = genOtp_();
  const exp = Date.now() + OTP_TTL_MS;
  pending.getRange(r, 8, 1, 2).setValues([[otp, exp]]); // OTP, OtpExpiresAtMs
  log_(logs, 'otp_resend', 'pid=' + pendingId);

  if (pref === 'email'){
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Your OTP (valid 15 minutes)',
        htmlBody: '<p>Your new OTP code is:</p><p style="font-size:20px;"><b>' + otp + '</b></p><p>This code expires in 15 minutes.</p>',
        body: 'Your OTP: ' + otp + ' (valid 15 minutes)'
      });
      log_(logs, 'otp_sent_email', maskEmail_(email));
    } catch(err){
      log_(logs, 'otp_email_fail', (err && err.message) || 'error');
      throw new Error('Could not send OTP email right now.');
    }
  }
  return { ok:true, expiresAtMs: exp };
}

function verifyRegistrationOtp(pendingId, otp, newPin){
  pendingId = String(pendingId||'').trim();
  otp = String(otp||'').trim();
  if (!pendingId || !/^\d{6}$/.test(otp)) throw new Error('Invalid request.');
  if (!/^\d{6}$/.test(String(newPin||''))) throw new Error('PIN must be 6 digits.');

  const { pending, users, balances, logs } = getDb_();
  const r = findRowByColumn_(pending, 1, pendingId);
  if (!r) throw new Error('Not found.');

  const cols = pending.getLastColumn();
  const row = pending.getRange(r,1,1,cols).getValues()[0];

  let i=0;
  const pid = row[i++];                // 1
  const createdMs = row[i++];          // 2
  const status = String(row[i++]||''); // 3
  const email = String(row[i++]||'');  // 4
  const phoneE164 = String(row[i++]||''); // 5
  const waE164 = String(row[i++]||''); // 6
  const pref = String(row[i++]||'');   // 7
  const otpStored = String(row[i++]||''); // 8
  const otpExpMs = Number(row[i++]||0);   // 9
  const first = String(row[i++]||'');     // 10
  const last  = String(row[i++]||'');     // 11
  const birthYmd = String(row[i++]||'');  // 12
  const gender   = String(row[i++]||'');  // 13
  const city     = String(row[i++]||'');  // 14
  const country  = String(row[i++]||'');  // 15
  const lang     = String(row[i++]||'');  // 16
  const referral = String(row[i++]||'');  // 17
  const optMarketing = String(row[i++]||''); // 18
  const utm_source = String(row[i++]||'');   // 19
  const utm_medium = String(row[i++]||'');   // 20
  const utm_campaign = String(row[i++]||''); // 21
  const utm_term = String(row[i++]||'');     // 22
  const utm_content = String(row[i++]||'');  // 23

  if (status !== 'pending') throw new Error('Registration is not pending.');
  const now = Date.now();
  if (!otpStored || otp !== otpStored) throw new Error('Wrong code.');
  if (Number(otpExpMs||0) < now) throw new Error('Code expired.');

  // Final dedupe right before write
  const dupe = duplicateExists_(email, phoneE164, waE164);
  if (dupe) throw new Error('This ' + dupe + ' is already registered.');

  const memberId = genMemberId_();
  const pinHash = hashPin_(String(newPin));

  extendUserRow_();
  users.appendRow([
    email, pinHash, memberId, new Date().toISOString(),
    first, last, birthYmd, gender, city, country, lang,
    phoneE164, waE164, pref, referral, String(optMarketing||''),
    utm_source, utm_medium, utm_campaign, utm_term, utm_content
  ]);

  // Immediately write PBKDF2 security fields for this new account
  const newRow = users.getLastRow();
  setUserSecurity_(newRow, String(newPin));
  // Optional: blank legacy PinHash for new accounts (kept commented for now)
  // users.getRange(newRow, 2).clearContent();
  
  // Init balance (0)
  const { row: balRow } = getOrInitBalance_(getDb_().balances, memberId);

  // Mark pending as completed
  pending.getRange(r, 3, 1, 1).setValue('completed');        // Status
  pending.getRange(r, 24, 1, 2).setValues([[now, memberId]]); // CompletedAtMs, MemberId

  // Welcome email
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Welcome to Cashbeik',
      htmlBody: '<p>Welcome, ' + first + '!</p><p>Your Member ID: <b>' + memberId + '</b></p><p><a href="https://cashbeik.com">Go to Cashbeik</a></p>',
      body: 'Welcome!\nMember ID: ' + memberId + '\nVisit: https://cashbeik.com'
    });
  } catch(_) {}

  log_(logs, 'reg_complete', 'member=' + memberId + ' email=' + maskEmail_(email));
  // NEW: log final geo written to Users
  log_(logs, 'user_geo', 'member=' + memberId + ' city=' + String(city||'') + ' country=' + String(country||''));

  return { ok:true, memberId };
}



function getMemberSummary(sid, range) {
  // Validate session (does NOT rotate SID)
  const v = validateSession(sid);
  if (!v || !v.ok) throw new Error('Not signed in.');

  const { balances, tx, merchants } = getDb_();

  // Get or init current balance
  const bal = getOrInitBalance_(balances, v.memberId);

  // Parse range: days (default 5)
  let days = Number(range);
  if (!Number.isFinite(days) || days <= 0) days = 5;
  const now = Date.now();
  const cutoff = now - days * 24 * 60 * 60 * 1000;

  // Build merchantId -> merchantName map
  const merchMap = {};
  const mLast = merchants.getLastRow();
  if (mLast >= 2) {
    const mVals = merchants.getRange(2, 1, mLast - 1, 2).getValues(); // [MerchantId, Name]
    for (let i = 0; i < mVals.length; i++) {
      const mid = String(mVals[i][0] || '');
      const name = String(mVals[i][1] || '');
      if (mid) merchMap[mid] = name;
    }
  }

  // Gather recent TX for this member within the cutoff
  const recent = [];
  const tLast = tx.getLastRow();
  if (tLast >= 2) {
    const cols = Math.max(7, tx.getLastColumn() || 7);
    const vals = tx.getRange(2, 1, tLast - 1, cols).getValues();
    // Walk newest → oldest
    for (let i = vals.length - 1; i >= 0; i--) {
      const [txId, memberId, merchantId, type, points, atMs] = vals[i]; // Staff is col 7 (ignored here)
      if (String(memberId) !== String(v.memberId)) continue;
      const at = Number(atMs || 0);
      if (!at || at < cutoff) continue;
      recent.push({
        txId: String(txId || ''),
        merchantId: String(merchantId || ''),            // kept internally but not shown in UI
        merchantName: merchMap[String(merchantId)] || '',// <-- name for the UI
        type: String(type || ''),
        points: Number(points || 0),
        atMs: at
      });
      // Optional cap so we don't send too much to the client on very large ranges
      if (recent.length >= 200) break;
    }
  }

  return { memberId: v.memberId, balance: bal.points, recent };
}

function collectPoints(sid, merchantId, points) {
  // Policy: customers cannot collect points themselves.
  throw new Error('Points can only be processed by merchant staff devices.');
}

function redeemPoints(sid, merchantId, points) {
  // Policy: customers cannot redeem points themselves.
  throw new Error('Points can only be processed by merchant staff devices.');
}


/**
 * Add an entry to the Waitlist sheet.
 * @param {Object} data
 *   { name, email, phoneLocal8, notes, city, country, source, memberId }
 */
function addToWaitlist(data) {
  const {
    name, email, phoneLocal8, notes,
    city, country, source, memberId
  } = (data || {});

  // Basic validation
  const nm = String(name || '').trim();
  if (!nm) throw new Error('Please enter your name.');
  const em = sanitizeEmail_(email);
  if (!em) throw new Error('Enter a valid email.');

  // Phone: optional — but if provided must be Lebanon 8 digits
  let phoneE164 = '';
  const p0 = String(phoneLocal8 || '').replace(/\D/g, '');
  if (p0) {
    let local8 = '';
    if (/^[0-9]{8}$/.test(p0)) local8 = p0;
    else if (/^[0-9]{7}$/.test(p0) && p0[0] === '3') local8 = '0' + p0; // support 7-digit “3xxxxxx”
    if (!local8 || !isValidLebanonLocal8_(local8)) {
      throw new Error('Phone must be 8 digits (or 7 if it starts with 3).');
    }
    phoneE164 = toE164Lebanon_(local8);
  }

  const citySafe    = String(city || '').trim();
  const countrySafe = String(country || '').trim();
  const notesSafe   = String(notes || '').trim();
  const srcSafe     = String(source || '').trim();    // e.g., 'home' or 'member'
  const midSafe     = String(memberId || '').trim();

  const { waitlist, logs } = getDb_();
  waitlist.appendRow([
    new Date().toISOString(),
    nm,
    em,
    phoneE164,
    citySafe,
    countrySafe,
    notesSafe,
    srcSafe,
    midSafe
  ]);

  log_(logs, 'waitlist_add', `name=${nm} email=${maskEmail_(em)} src=${srcSafe} mid=${midSafe||'-'}`);
  return { ok: true };
}


// ---- Lockout helpers ----
function recordSigninFailure_(lockouts, key) {
  const now = Date.now();
  const hdr = { Key:1, FailCount:2, LockUntilMs:3, UpdatedAt:4, FirstFailMs:5, TotalFails24h:6, PermLocked:7 };
  let r = findRowByColumn_(lockouts, 1, key);

  if (!r) {
    // First failure for this key
    lockouts.appendRow([key, 1, 0, new Date(now).toISOString(), now, 1, false]);
    return { failCount: 1, lockedUntil: 0, permanent: false };
  }

  const vals = lockouts.getRange(r,1,1,7).getValues()[0];
  const failCount    = Number(vals[hdr.FailCount-1]||0) + 1;
  const firstFailMs  = Number(vals[hdr.FirstFailMs-1]||0) || now;
  const prev24h      = Number(vals[hdr.TotalFails24h-1]||0) || 0;
  const permLocked   = String(vals[hdr.PermLocked-1]).toLowerCase()==='true';
  const lockedUntil0 = Number(vals[hdr.LockUntilMs-1]||0);

  // Update rolling 24h counter (reset if more than 24h since first fail in window)
  let firstWindowMs = firstFailMs;
  let total24h = prev24h + 1;
  if (now - firstFailMs > 24*60*60*1000) {
    firstWindowMs = now;
    total24h = 1;
  }

  // Decide cooldown
  const rule = nextCooldownFor_(failCount);
  let lockedUntil = lockedUntil0;
  let permanent = permLocked;

  if (rule.permanent) {
    permanent = true;
    lockedUntil = now + 365*24*60*60*1000; // keep a date for UI, practically "forever"
  } else if (rule.durMin > 0) {
    lockedUntil = now + rule.durMin*60*1000;
  }

  // Save
  lockouts.getRange(r,1,1,7).setValues([[
    key,
    failCount,
    lockedUntil,
    new Date(now).toISOString(),
    firstWindowMs,
    total24h,
    permanent
  ]]);

  // Log
  const { logs } = getDb_();
  log_(logs, 'lock_fail', key + ' :: count=' + failCount);
  if (permanent) {
    log_(logs, 'lock_permanent', key + ' :: after=' + failCount);
  } else if (lockedUntil && lockedUntil > now) {
    log_(logs, 'lock_impose', key + ' :: until=' + new Date(lockedUntil).toISOString());
  }

  return { failCount, lockedUntil, permanent };
}
function clearLockout_(lockouts, key) {
  const r = findRowByColumn_(lockouts, 1, key);
  if (!r) return;
  lockouts.getRange(r,1,1,7).setValues([[key, 0, 0, new Date().toISOString(), 0, 0, false]]);
  const { logs } = getDb_();
  log_(logs, 'lock_clear', key);
}

// ---- Logs ----
function log_(logs, type, message) {
  logs.appendRow([new Date().toISOString(), String(type||''), String(message||'')]);
}

/** client side telemetry from the user app, stored for admin to review */
function userReportError(payload){
  const { logs } = getDb_();
  const safe = payload || {};
  try{
    const msg = JSON.stringify({
      ua: String(safe.ua || ''),
      location: String(safe.location || ''),
      message: String(safe.message || ''),
      stack: String(safe.stack || ''),
      extra: safe.extra || null
    });
    log_(logs, 'user_client_error', msg);
  }catch(e){
    log_(logs, 'user_client_error', 'json_fail');
  }
  return { ok:true };
}

