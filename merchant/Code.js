// ===============================
// File: Code.gs  (Merchant App)
// ===============================

// ── Script Properties ──
const MSP = PropertiesService.getScriptProperties();
const USER_DB_ID_KEY = 'USER_DB_ID';   // Set via setupSetUserDbId(<MemberAppDataID>)

// ── Local (merchant project) sheet names ──
const STAFF_SHEET   = 'Staff';           // [Username, PinHash, MerchantId, Role, Active, CreatedAt]
const MSESS_SHEET   = 'MSessions';       // [SID, Username, LastSeenMs, CreatedAt]
const MLOCK_SHEET   = 'Lockouts';        // staff lockouts
const MLOGS_SHEET   = 'Logs';            // telemetry logs
const MCONFIG_SHEET = 'Config';          // [Key, Value]

// ── Remote (user DB) sheet names ──
const R_MERCHANTS = 'Merchants';       // [MerchantId, Name, Active, Secret, CreatedAt]
const R_TX        = 'Transactions';    // [TxId, MemberId, MerchantId, Type, Points, AtMs, Staff]
const R_BALANCES  = 'Balances';        // [MemberId, Points]
const R_LINKTOK   = 'LinkTokens';      // [Token, MemberId, Mode, ExpiresAtMs, Used, CreatedAt]
const R_SESS      = 'Sessions';        // user app sessions
const R_CONFIG    = 'Config';          // [Key, Value]  <-- NEW (to read QR key/TTL)

// ── Remote (user DB) sheets for coupons ──
const R_COUPONS   = 'Coupons';         // [Code, MerchantId, Mode, Type, Value, MaxUses, UsedCount, PerMemberLimit, StartsAtMs, ExpiresAtMs, Active, CreatedAt, Notes]
const R_CPN_USES  = 'CouponUses';      // [Code, MemberId, MerchantId, AtMs, Staff, TxId]
const R_COUPON_REQUESTS = 'CouponRequests'; // [RequestId, MerchantId, Code, Mode, Type, Value, MaxUses, PerMemberLimit, StartIso, EndIso, Notes, Status, CreatedBy, CreatedAt, UpdatedAt, DecisionBy, DecisionAt, DecisionNotes]


// ── Remote (user DB) sheets for Multiplier Campaigns (Option B) ──
const R_CAMPAIGNS   = 'Campaigns';             // [CampaignId, MerchantId, Title, Type, Multiplier, ...] (Live)
const R_CAMPAIGN_REQUESTS = 'CampaignRequests';    // [RequestId, MerchantId, RequestType, Title, Multiplier, ...] (Requests)
const R_CAMPAIGN_REDEMPTIONS = 'CampaignRedemptions'; // [RedemptionId,CampaignId,MemberId,MerchantId,TxId,AtMs,BasePoints,Multiplier,BonusPoints,CostAccrued]

const M_SESSION_TTL_MS = 60 * 1000;    // 1 minute to match user app

// Lockout policy (staff)
const M_LOCK_MAX_ATTEMPTS = 5;             // cooldown threshold
const M_LOCK_DURATION_MS  = 15 * 60 * 1000;
const M_LOCK_PERMA_ATTEMPTS = 10;          // permanent block threshold

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appVersion = getAppVersion_();

  // NEW: pass raw query parameters + full query string to the client
  t.initParamsJson = JSON.stringify((e && e.parameter) || {});
  t.initQueryString = String((e && e.queryString) || '');  

  const out = t.evaluate()
    .setTitle('Merchant App')
    .addMetaTag('viewport','width=device-width, initial-scale=1');

  return out;
}

// ───────────────────────────────── Admin/setup
function setupSetUserDbId(id) {
  if (!id) throw new Error('Pass the spreadsheet ID from the User app.');

  // 1. Validate ID and open
  let ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Failed to open Spreadsheet. Check the ID and permissions.');
  }

  // 2. Perform the "handshake" - check for critical sheets
  if (!ss.getSheetByName('Merchants') || !ss.getSheetByName('Transactions') || !ss.getSheetByName('Balances')) {
    throw new Error('This does not appear to be the correct User DB. Required sheets (Merchants, Transactions, Balances) were not found.');
  }

  // 3. Save property
  MSP.setProperty(USER_DB_ID_KEY, id);
  return { ok: true, id:ss.getId(), name:ss.getName() };
}

// ───────────────────────────────── Auth (merchant staff)
function adminCreateStaff(merchantId, username, pin, role) {
  merchantId = String(merchantId || '').trim();
  username = sanitizeUsername_(username);
  role = String(role || 'staff').toLowerCase();
  if (!merchantId) throw new Error('merchantId required.');
  if (!username) throw new Error('valid username required.');
  if (!isValidPin_(pin)) throw new Error('PIN must be exactly 6 digits.');
  if (role !== 'staff' && role !== 'manager') throw new Error('role must be staff or manager.');

  const { staff } = getLocalDb_();
  const { merchants } = getUserDb_();

  const mr = findRowByColumn_(merchants, 1, merchantId);
  if (!mr) throw new Error('Merchant not found in User DB.');
  const active = String(merchants.getRange(mr, 3).getValue() || 'TRUE').toUpperCase() !== 'FALSE';
  if (!active) throw new Error('Merchant is inactive.');

  if (findRowByColumn_(staff, 1, username)) throw new Error('Username already exists.');

  const rec = makeSaltedStaffRecord_(pin);
  // Keep legacy PinHash blank for new accounts (soft-migration only needed for old ones)
  staff.appendRow([
    username, '', merchantId, role, true, new Date().toISOString(),
    rec.saltB64, rec.hashB64, rec.iter, PIN_ALGO_DEFAULT, rec.peppered, rec.updatedAt,
    false // <-- ADD THIS LINE
  ]);
  return { ok:true, username, merchantId, role };
}

function signinStaff(username, pin) {
  username = sanitizeUsername_(username);
  if (!username) throw new Error('Invalid credentials.'); // generic

  const now = Date.now();
  const { staff, msessions, mlockouts } = getLocalDb_();

  // lockout check
  const key = 'staff:' + username;
  const lr = findRowByColumn_(mlockouts, 1, key);
  if (lr) {
    const [, , lockUntilMs] = mlockouts.getRange(lr, 1, 1, 3).getValues()[0];
    if (Number(lockUntilMs || 0) > now) throw new Error('Invalid credentials.');
  }

  if (!isValidPin_(pin)) { recordStaffFailure_(mlockouts, key); throw new Error('Invalid credentials.'); }

  const r = findRowByColumn_(staff, 1, username);
  if (!r) { recordStaffFailure_(mlockouts, key); throw new Error('Invalid credentials.'); }

  const idx = staffHeaderIdx_(staff);
  const cols = Math.max(staff.getLastColumn() || 13, 13);
  const vals = staff.getRange(r, 1, 1, cols).getValues()[0];



  const active = String(vals[(idx['Active']||5)-1]).toUpperCase() !== 'FALSE';
  if (!active) { recordStaffFailure_(mlockouts, key); throw new Error('Invalid credentials.'); }

  const ok = verifyStaffPinAndMigrateIfNeeded_(staff, r, pin).ok;
  if (!ok) { recordStaffFailure_(mlockouts, key); throw new Error('Invalid credentials.'); }

  // success → clear lockout
  clearStaffLockout_(mlockouts, key);

  const sid = genSid_();
  msessions.appendRow([sid, username, now, new Date(now).toISOString()]);
  const merchantId = String(vals[(idx['MerchantId']||3)-1]||'');
  const role = String(vals[(idx['Role']||4)-1]||'');
  const mustChange = String(vals[(idx['MustChangePin']||13)-1]||'').toString().toLowerCase() === 'true';
  return { ok:true, sid, merchantId, role, username, mustChange };
}


function signinStaffAttempt(username, pin) {
  username = sanitizeUsername_(username);
  if (!username) return { ok:false, msg:'Invalid credentials.', lock: getStaffLockInfo_(username) };

  const now = Date.now();
  const { staff, msessions, mlockouts } = getLocalDb_();

  // Check lock state first
  const info = getStaffLockInfo_(username);
  if (info.permanent) return { ok:false, msg:'Invalid credentials.', lock: info };
  if (info.locked)    return { ok:false, msg:'Invalid credentials.', lock: info };

  if (!isValidPin_(pin)) {
    recordStaffFailure_(mlockouts, 'staff:' + username);
    return { ok:false, msg:'Invalid credentials.', lock: getStaffLockInfo_(username) };
  }

  const r = findRowByColumn_(staff, 1, username);
  if (!r) {
    recordStaffFailure_(mlockouts, 'staff:' + username);
    return { ok:false, msg:'Invalid credentials.', lock: getStaffLockInfo_(username) };
  }

  const idx = staffHeaderIdx_(staff);
  const cols = Math.max(staff.getLastColumn() || 13, 13);
  const vals = staff.getRange(r, 1, 1, cols).getValues()[0];

  const active = String(vals[(idx['Active']||5)-1]).toUpperCase() !== 'FALSE';
  if (!active) {
    recordStaffFailure_(mlockouts, 'staff:' + username);
    return { ok:false, msg:'Invalid credentials.', lock: getStaffLockInfo_(username) };
  }

  const ok = verifyStaffPinAndMigrateIfNeeded_(staff, r, pin).ok;
  if (!ok) {
    recordStaffFailure_(mlockouts, 'staff:' + username);
    return { ok:false, msg:'Invalid credentials.', lock: getStaffLockInfo_(username) };
  }

  // success → clear lockout
  clearStaffLockout_(mlockouts, 'staff:' + username);

  // Create a new session
  const sid = genSid_();
  msessions.appendRow([sid, username, now, new Date(now).toISOString()]);

  const merchantId = String(vals[(idx['MerchantId']||3)-1]||'');
  const role = String(vals[(idx['Role']||4)-1]||'');
  const mustChange = String(vals[(idx['MustChangePin']||13)-1]||'').toString().toLowerCase() === 'true';

  return { ok:true, sid, merchantId, role, username, mustChange };
}

// Public: return lock/counter info for username
function getStaffLockInfo(username){
  username = sanitizeUsername_(username);
  if (!username) return { ok:true, permanent:false, locked:false, failCount:0, nextTarget:M_LOCK_MAX_ATTEMPTS };
  return getStaffLockInfo_(username);
}

// Internal
function getStaffLockInfo_(username){
  const { mlockouts } = getLocalDb_();
  const key = 'staff:' + String(username||'');
  const r = findRowByColumn_(mlockouts, 1, key);
  const now = Date.now();

  let fail = 0, until = 0, permanent = false;
  if (r){
    const cols = Math.max(5, mlockouts.getLastColumn()||5);
    const row = mlockouts.getRange(r,1,1,cols).getValues()[0]; // ['Key','FailCount','LockUntilMs','Permanent','UpdatedAt']
    fail = Number(row[1]||0);
    until = Number(row[2]||0);
    permanent = String(row[3]||'').toLowerCase() === 'true';
  }
  const locked = !!until && (until > now);

  return {
    ok:true,
    permanent,
    locked,
    untilMs: locked ? until : 0,
    failCount: fail,
    nextTarget: M_LOCK_MAX_ATTEMPTS   // show "X / 5 before a cooldown"
  };
}


function validateMSession(sid) {
  if (!sid) return { ok: false };
  const { msessions, staff } = getLocalDb_();
  const sRow = findRowByColumn_(msessions, 1, sid);
  if (!sRow) return { ok: false };
  const [, username, lastSeenMs] = msessions.getRange(sRow, 1, 1, 4).getValues()[0];

  const now = Date.now();
  if (now - Number(lastSeenMs || 0) > M_SESSION_TTL_MS) {
    msessions.deleteRow(sRow);
    return { ok: false, reason: 'expired' };
  }
  // Touch ONLY — NO SID ROTATION
  msessions.getRange(sRow, 3).setValue(now);

  // Load staff profile
  const r = findRowByColumn_(staff, 1, username);
  if (!r) return { ok: false };
  const idx = staffHeaderIdx_(staff);
  const rowVals = staff.getRange(r, 1, 1, Math.max(staff.getLastColumn()||13, 13)).getValues()[0];
  const merchantId = String(rowVals[(idx['MerchantId']||3)-1]||'');
  const role       = String(rowVals[(idx['Role']||4)-1]||'');
  const active     = String(rowVals[(idx['Active']||5)-1]||'');
  if (String(active).toUpperCase() === 'FALSE') return { ok: false };

  const mustChange = String(rowVals[(idx['MustChangePin']||13)-1]||'').toString().toLowerCase() === 'true';
  return { ok: true, sid, username, merchantId, role, mustChange };
}


function logoutMSession(sid) {
  if (!sid) return { ok: true };
  const { msessions } = getLocalDb_();
  const sRow = findRowByColumn_(msessions, 1, sid);
  if (sRow) msessions.deleteRow(sRow);
  return { ok: true };
}

// ───────────────────────────────── Deep-link resolver (Option B)
function resolveDeepLinkToken(token) {
  token = String(token || '').trim();
  if (!token) return { ok: false, reason: 'missing' };
  const { linkTokens, logs } = getUserDb_();
  const r = findRowByColumn_(linkTokens, 1, token);
  if (!r) { log_(logs, 'resolve_token_fail', 'not_found'); return { ok: false, reason: 'not_found' }; }
  const [tok, memberId, mode, expMs, used] = linkTokens.getRange(r, 1, 1, 5).getValues()[0];
  const now = Date.now();
  // NEW: reject if already used
  if (String(used || '').toLowerCase() === 'true') {
    log_(logs, 'resolve_token_fail', 'already_used');
    return { ok: false, reason: 'already_used' };
  }  
  if (Number(expMs || 0) < now) { log_(logs, 'resolve_token_fail', 'expired'); return { ok: false, reason: 'expired' }; }

  // Optional: mark used (one-time)
  linkTokens.getRange(r, 5).setValue(true);
  log_(logs, 'resolve_token_ok', 'mid=' + memberId);
  return { ok: true, memberId: String(memberId || ''), mode: String(mode || '').toUpperCase() };
}

// ───────────────────────────────── Actions (writes to User DB)
function submitCollectRedeem(sid, memberId, points, mode, couponCode) {
  const v = getMerchantScopeOrThrow_(sid); // scope + active checks

  memberId = String(memberId || '').trim().toUpperCase();
  if (!/^MBR-[A-Z0-9]{8}$/.test(memberId)) throw new Error('Invalid memberId.');

  points = Number(points);
  if (!Number.isFinite(points) || points <= 0) throw new Error('Points must be positive.');

  mode = String(mode || '').toUpperCase();
  if (mode !== 'COLLECT' && mode !== 'REDEEM') throw new Error('Mode must be COLLECT or REDEEM.');

  const coupon = String(couponCode || '').trim(); // preserve case

  // pull all needed sheets (must include campaigns + campaign redemptions)
  const { balances, tx, logs, coupons, cpnUses, campaigns, cpnRed } = getUserDb_();

  const bal = getOrInitBalanceRemote_(balances, memberId);

  // 6.1) Coupon first (affects the base used for multiplier)
  let appliedRow = 0;
  let appliedCoupon = null;
  let effective = Number(points || 0);

  if (coupon) {
    const row = findCouponRow_(coupons, coupon);
    const val = validateAndApplyCoupon_(coupons, cpnUses, v, memberId, mode, points, coupon);
    if (!val.ok) throw new Error('Coupon error: ' + String(val.reason || 'invalid'));
    effective = val.effectivePoints;
    appliedCoupon = val.couponApplied;
    appliedRow = row;
  }

  // 6.2) Campaign multiplier (COLLECT only)
  let bonus = 0;
  let usedCampaign = null;

  if (mode === 'COLLECT') {
    const camp = findActiveCampaignForMerchant_(campaigns, v.merchantId);
    if (camp) {
      const usedTotal = countTotalCampaignRedemptions_(cpnRed, camp.campaignId);
      const underGlobalCap = (Number(camp.maxRedemptions || 0) <= 0) || (usedTotal < Number(camp.maxRedemptions || 0));

      const usedByMember = countMemberCampaignRedemptions_(cpnRed, camp.campaignId, memberId);
      const perMemCap = Number(camp.perMemberRedemptions || 0);
      const underMemberCap = (perMemCap <= 0) || (usedByMember < perMemCap);

      const meetsMin = (Number(camp.minSpend || 0) <= 0) || (Number(effective) >= Number(camp.minSpend || 0));

      if (underGlobalCap && underMemberCap && meetsMin && Number(camp.multiplier || 1) > 1) {
        // raw bonus
        let rawBonus = Math.floor(Number(effective) * (Number(camp.multiplier) - 1));
        // trim against per-member bonus cap (points)
        const cap = Number(camp.perMemberBonusCap || 0);
        if (cap > 0) {
          const soFar = sumMemberCampaignBonusPoints_(cpnRed, camp.campaignId, memberId);
          const remaining = Math.max(0, cap - soFar);
          if (remaining <= 0) {
            rawBonus = 0; // block
          } else if (rawBonus > remaining) {
            rawBonus = remaining; // trim
          }
        }
        bonus = Math.max(0, rawBonus);
        usedCampaign = (bonus > 0) ? camp : camp; // record even if 0, for audit decision below
      }
    }
  }


  // 6.3) Apply balance
  let newBal = bal.points;
  if (mode === 'COLLECT') {
    newBal = bal.points + effective + bonus;
  } else {
    if (bal.points < effective) throw new Error('Insufficient points.');
    newBal = bal.points - effective;
  }
  balances.getRange(bal.row, 2).setValue(newBal);

  // 6.4) Append base TX (without bonus)
  const txId = Utilities.getUuid();
  const now = Date.now();
  tx.appendRow([txId, memberId, v.merchantId, mode, effective, now, v.username]);

  // 6.5) Record coupon use
  if (appliedCoupon && appliedRow) {
    try {
      recordCouponUse_(coupons, cpnUses, appliedRow, appliedCoupon.code, memberId, v.merchantId, v.username, txId);
    } catch (_) {}
  }

  // 6.6) Record campaign redemption (audit/billing)
  if (usedCampaign && bonus > 0) {
    const cost = (String(usedCampaign.billingModel || '') === 'per_redemption')
      ? Number(usedCampaign.costPerRedemption || 0)
      : 0;
    const redId = 'CR-' + Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
    cpnRed.appendRow([
      redId, usedCampaign.campaignId, memberId, v.merchantId, txId, now,
      Number(effective), Number(usedCampaign.multiplier || 1), Number(bonus), Number(cost)
    ]);
  }

  // 6.7) Audit log
  try {
    log_(logs, 'tx', JSON.stringify({
      txId, memberId, merchantId: v.merchantId, mode,
      points: effective, bonusApplied: bonus, staff: v.username,
      campaignId: usedCampaign ? usedCampaign.campaignId : '',
      coupon: appliedCoupon ? appliedCoupon.code : ''
    }));
  } catch (_) {}

  return {
    ok: true,
    balance: newBal,
    txId, memberId, merchantId: v.merchantId, mode,
    couponApplied: appliedCoupon ? { code: appliedCoupon.code, type: appliedCoupon.type, value: appliedCoupon.value } : null,
    originalPoints: Number(points || 0),
    effectivePoints: Number(effective || 0),
    campaign: usedCampaign ? {
      campaignId: usedCampaign.campaignId,
      multiplier: Number(usedCampaign.multiplier || 1),
      bonusPoints: Number(bonus || 0),
      perMemberBonusCap: Number(usedCampaign.perMemberBonusCap || 0),
      perMemberRedemptions: Number(usedCampaign.perMemberRedemptions || 0)
    } : null
  };
}



// ───────────────────────────────── Analytics (manager-only)
function getMerchantStats(sid, range) {
  const v = requireRole_(sid, 'manager'); // { ok, username, merchantId, role }
  const { tx, logs } = getUserDb_();

  // --- Cache (10 minutes) ---
  const cache = CacheService.getScriptCache();
  const key = ['stats', v.merchantId, String(range || '7d')].join(':');
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch(_) { /* fallthrough */ }
  }

  const win = _rangeWindow_(range);
  const startMs = Number(win.startMs || 0);
  const labelDays = Number(win.labelDays || 7);

  // Prepare buckets
  const buckets = {};
  for (let i = labelDays - 1; i >= 0; i--) {
    const dayMs = _startOfTodayMs_() - i * 24 * 3600 * 1000;
    buckets[_ymd(dayMs)] = { collect: 0, redeem: 0 };
  }

  // Totals + Top customers
  let collectPoints = 0, redeemPoints = 0, txCount = 0;
  const members = new Set();
  const perMember = {}; // memberId -> {collect, redeem}

  const last = tx.getLastRow();
  if (last >= 2) {
    const cols = Math.max(7, tx.getLastColumn() || 7);
    const vals = tx.getRange(2, 1, last - 1, cols).getValues(); // [TxId, MemberId, MerchantId, Type, Points, AtMs, Staff]
    for (let i = vals.length - 1; i >= 0; i--) {
      const row = vals[i];
      const mId = String(row[2] || ''); // MerchantId
      if (mId !== String(v.merchantId)) continue;

      const atMs = Number(row[5] || 0);
      if (!atMs || atMs < startMs) continue;

      const type = String(row[3] || '').toUpperCase();
      const pts = Number(row[4] || 0);
      const mem = String(row[1] || '');

      txCount++;
      if (type === 'COLLECT') collectPoints += pts;
      else if (type === 'REDEEM') redeemPoints += pts;

      if (mem) {
        members.add(mem);
        if (!perMember[mem]) perMember[mem] = { collect: 0, redeem: 0 };
        if (type === 'COLLECT') perMember[mem].collect += pts;
        else if (type === 'REDEEM') perMember[mem].redeem += pts;
      }

      const dayKey = _ymd(atMs);
      if (buckets[dayKey]) {
        if (type === 'COLLECT') buckets[dayKey].collect += pts;
        else if (type === 'REDEEM') buckets[dayKey].redeem += pts;
      }
    }
  }

  const series = Object.keys(buckets).sort().map(d => ({
    day: d,
    collect: buckets[d].collect || 0,
    redeem: buckets[d].redeem || 0
  }));

  const topCustomers = Object.keys(perMember).map(mem => {
    const c = perMember[mem].collect || 0;
    const r = perMember[mem].redeem || 0;
    return { memberId: mem, totalCollect: c, totalRedeem: r, net: c - r };
  }).sort((a, b) => (b.totalCollect + b.totalRedeem) - (a.totalCollect + a.totalRedeem)).slice(0, 5);

  const payload = {
    range: String(range || '7d'),
    totals: {
      collectPoints,
      redeemPoints,
      netPoints: collectPoints - redeemPoints,
      txCount,
      uniqueMembers: members.size
    },
    series,
    topCustomers
  };

  try { cache.put(key, JSON.stringify(payload), 600); } catch(_) {}
  try { log_(logs, 'mstats', JSON.stringify({ merchantId: v.merchantId, user: v.username, range: payload.range })); } catch (_){}
  return payload;
}

function getRecentTx(sid, limit) {
  const v = requireRole_(sid, 'manager'); // manager-only (staff can view via future POS modal if needed)
  const n = Math.max(1, Math.min(Number(limit) || 20, 50));
  const { tx, logs } = getUserDb_();

  const rows = [];
  const last = tx.getLastRow();
  if (last >= 2) {
    const cols = Math.max(7, tx.getLastColumn() || 7);
    const vals = tx.getRange(2, 1, last - 1, cols).getValues();
    for (let i = vals.length - 1; i >= 0 && rows.length < n; i--) {
      const [txId, memberId, merchantId, type, points, atMs, staff] = vals[i];
      if (String(merchantId) !== String(v.merchantId)) continue;
      rows.push({
        txId: String(txId || ''),
        memberId: String(memberId || ''),
        type: String(type || '').toUpperCase(),
        points: Number(points || 0),
        atMs: Number(atMs || 0),
        staff: String(staff || '')
      });
    }
  }
  try { log_(logs, 'mstats_recent', JSON.stringify({ merchantId: v.merchantId, user: v.username, count: rows.length })); } catch (_){}
  return rows;
}

// ───────────────────────────────── Analytics (custom range)
function getMerchantStatsRange(sid, startIso, endIso) {
  const v = requireRole_(sid, 'manager');
  const { tx, logs } = getUserDb_();

  const startMs = _startOfDayMs_(startIso);
  const endMs   = _endOfDayMs_(endIso);
  if (!startMs || !endMs || startMs > endMs) throw new Error('Invalid date range.');
  if (endMs - startMs > 370 * 24 * 3600 * 1000) throw new Error('Range too large (max ~1 year).');

  // --- Cache (10 minutes) ---
  const cache = CacheService.getScriptCache();
  const key = ['statsr', v.merchantId, String(startMs), String(endMs)].join(':');
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch(_) { /* fallthrough */ }
  }

  const labelDays = Math.floor((endMs - _startOfDayMs_(new Date(startMs))) / (24*3600*1000)) + 1;
  const buckets = {};
  for (let i = 0; i < labelDays; i++) {
    const day = _ymd(startMs + i * 24 * 3600 * 1000);
    buckets[day] = { collect: 0, redeem: 0 };
  }

  let collectPoints = 0, redeemPoints = 0, txCount = 0;
  const members = new Set();
  const perMember = {};

  const last = tx.getLastRow();
  if (last >= 2) {
    const cols = Math.max(7, tx.getLastColumn() || 7);
    const vals = tx.getRange(2, 1, last - 1, cols).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const row = vals[i];
      const mId = String(row[2] || '');
      if (mId !== String(v.merchantId)) continue;

      const atMs = Number(row[5] || 0);
      if (!atMs || atMs < startMs || atMs > endMs) continue;

      const type = String(row[3] || '').toUpperCase();
      const pts = Number(row[4] || 0);
      const mem = String(row[1] || '');

      txCount++;
      if (type === 'COLLECT') collectPoints += pts;
      else if (type === 'REDEEM') redeemPoints += pts;

      if (mem) {
        members.add(mem);
        if (!perMember[mem]) perMember[mem] = { collect: 0, redeem: 0 };
        if (type === 'COLLECT') perMember[mem].collect += pts;
        else if (type === 'REDEEM') perMember[mem].redeem += pts;
      }

      const dayKey = _ymd(atMs);
      if (buckets[dayKey]) {
        if (type === 'COLLECT') buckets[dayKey].collect += pts;
        else if (type === 'REDEEM') buckets[dayKey].redeem += pts;
      }
    }
  }

  const series = Object.keys(buckets).sort().map(d => ({
    day: d,
    collect: buckets[d].collect || 0,
    redeem: buckets[d].redeem || 0
  }));

  const topCustomers = Object.keys(perMember).map(mem => {
    const c = perMember[mem].collect || 0;
    const r = perMember[mem].redeem || 0;
    return { memberId: mem, totalCollect: c, totalRedeem: r, net: c - r };
  }).sort((a, b) => (b.totalCollect + b.totalRedeem) - (a.totalCollect + a.totalRedeem)).slice(0, 5);

  const payload = {
    range: 'custom',
    startMs, endMs,
    totals: {
      collectPoints,
      redeemPoints,
      netPoints: collectPoints - redeemPoints,
      txCount,
      uniqueMembers: members.size
    },
    series,
    topCustomers
  };

  try { cache.put(key, JSON.stringify(payload), 600); } catch(_) {}
  try { log_(logs, 'mstats_range', JSON.stringify({ merchantId: v.merchantId, user: v.username, startMs, endMs })); } catch (_){}
  return payload;
}


function getRecentTxRange(sid, startIso, endIso, limit) {
  const v = requireRole_(sid, 'manager');
  const startMs = _startOfDayMs_(startIso);
  const endMs   = _endOfDayMs_(endIso);
  if (!startMs || !endMs || startMs > endMs) throw new Error('Invalid date range.');
  const n = Math.max(1, Math.min(Number(limit) || 20, 50));

  const rows = [];
  const { tx } = getUserDb_();
  const last = tx.getLastRow();
  if (last >= 2) {
    const cols = Math.max(7, tx.getLastColumn() || 7);
    const vals = tx.getRange(2, 1, last - 1, cols).getValues();
    for (let i = vals.length - 1; i >= 0 && rows.length < n; i--) {
      const [txId, memberId, merchantId, type, points, atMs, staff] = vals[i];
      if (String(merchantId) !== String(v.merchantId)) continue;
      const when = Number(atMs || 0);
      if (!when || when < startMs || when > endMs) continue;
      rows.push({
        txId: String(txId || ''),
        memberId: String(memberId || ''),
        type: String(type || '').toUpperCase(),
        points: Number(points || 0),
        atMs: when,
        staff: String(staff || '')
      });
    }
  }
  return rows;
}

// day boundaries (script timezone)
function _startOfDayMs_(isoOrDate) {
  const d = new Date(isoOrDate); if (isNaN(d)) return 0;
  d.setHours(0,0,0,0); return d.getTime();
}
function _endOfDayMs_(isoOrDate) {
  const d = new Date(isoOrDate); if (isNaN(d)) return 0;
  d.setHours(23,59,59,999); return d.getTime();
}



// ───────────────────────────────── DB helpers
function getLocalDb_() {
  const ss = SpreadsheetApp.getActive();
  const staff     = ensureSheet_(ss, STAFF_SHEET, [
    'Username','PinHash','MerchantId','Role','Active','CreatedAt',
    'Salt','Hash','Iter','Algo','Peppered','UpdatedAt','MustChangePin'
  ]);

  const msessions = ensureSheet_(ss, MSESS_SHEET, ['SID','Username','LastSeenMs','CreatedAt']);
  const mlockouts = ensureSheet_(ss, MLOCK_SHEET, ['Key','FailCount','LockUntilMs','Permanent','UpdatedAt']);
  const mlogs     = ensureSheet_(ss, MLOGS_SHEET, ['At','Type','Message']);
  const config    = ensureSheet_(ss, MCONFIG_SHEET, ['Key','Value']);
  return { ss, staff, msessions, mlockouts, mlogs, config };
}
function getAppVersion_() {
  try {
    const { config } = getLocalDb_();
    const last = config.getLastRow();
    if (last < 2) return 'dev';

    const rows = config.getRange(2, 1, last - 1, 2).getValues();
    const needle = 'APP_VERSION';

    for (let i = 0; i < rows.length; i++) {
      const key = String(rows[i][0] || '').trim().toUpperCase();
      if (key === needle) {
        const val = String(rows[i][1] || '').trim();
        return val || 'dev';
      }
    }
    return 'dev';
  } catch (e) {
    return 'dev';
  }
}

// --- Replace the existing getUserDb_ function in Merchant App - Code.gs entirely ---
function getUserDb_() {
  const id = MSP.getProperty(USER_DB_ID_KEY);
  if (!id) throw new Error('User DB is not linked. Run setupSetUserDbId(<ID>) once.');
  const ss = SpreadsheetApp.openById(id);
  const merchants  = ensureSheet_(ss, R_MERCHANTS, ['MerchantId','Name','Active','Secret','CreatedAt']);
  const tx         = ensureSheet_(ss, R_TX,        ['TxId','MemberId','MerchantId','Type','Points','AtMs','Staff']);
  const balances   = ensureSheet_(ss, R_BALANCES,  ['MemberId','Points']);
  const linkTokens = ensureSheet_(ss, R_LINKTOK,   ['Token','MemberId','Mode','ExpiresAtMs','Used','CreatedAt']);
  const logs       = ensureSheet_(ss, 'Logs',      ['At','Type','Message']);
  const sessions   = ensureSheet_(ss, R_SESS,      ['SID','Email','LastSeenMs','CreatedAt']);
  const config     = ensureSheet_(ss, R_CONFIG,    ['Key','Value']);
  const coupons = ensureSheet_(ss, R_COUPONS, [
    'Code','MerchantId','Mode','Type','Value','MaxUses','UsedCount','PerMemberLimit',
    'StartIso','EndIso','Active','CreatedAt','Notes'
  ]);
  const cpnUses    = ensureSheet_(ss, R_CPN_USES, ['Code','MemberId','MerchantId','AtMs','Staff','TxId']);
  const couponRequests = ensureSheet_(ss, R_COUPON_REQUESTS, [
    'RequestId','MerchantId',
    'Code','Mode','Type','Value',
    'MaxUses','PerMemberLimit',
    'StartIso','EndIso',
    'Notes',
    'Status','CreatedBy','CreatedAt','UpdatedAt',
    'DecisionBy','DecisionAt','DecisionNotes'
  ]);
  // --- UPDATED CAMPAIGN HEADERS (Problem 1 Fix) ---
  const campaigns  = ensureSheet_(ss, R_CAMPAIGNS, [
    'CampaignId','MerchantId','Title','Type','Multiplier',
    'StartIso','EndIso','MinSpend',
    'MaxRedemptions',
    'MaxPerCustomer',               // Added back for consistency with Admin App
    'PerMemberRedemptions',
    'BudgetCap',
    'BillingModel','CostPerRedemption','Active','CreatedAt','UpdatedAt',
    'ImageUrl',
    'PerMemberBonusCap'
  ]);
  const cpnRed     = ensureSheet_(ss, R_CAMPAIGN_REDEMPTIONS, [
    'RedemptionId','CampaignId','MemberId','MerchantId','TxId','AtMs',
    'BasePoints','Multiplier','BonusPoints','CostAccrued'
  ]);
  const campaignRequests = ensureSheet_(ss, R_CAMPAIGN_REQUESTS, [
    'RequestId','MerchantId','RequestType','Title','Multiplier',
    'StartIso','EndIso','MinSpend',
    'MaxRedemptions',
    'MaxPerCustomer',               // Added back for consistency with Admin App
    'PerMemberRedemptions',
    'BudgetCap',
    'BillingModel','CostPerRedemption','Notes','ImageUrl',
    'Status','CreatedBy','CreatedAt','UpdatedAt',
    'DecisionBy','DecisionAt','DecisionNotes','LinkedCampaignId',
    'PerMemberBonusCap'
  ]);
  // --- END UPDATED CAMPAIGN HEADERS ---
  return {
    ss, merchants, tx, balances, linkTokens, logs, sessions, config,
    coupons, cpnUses, couponRequests,
    campaigns, cpnRed, campaignRequests
  };
}


function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const h = sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn()||headers.length)).getValues()[0];
  if (h.every(v => !v)) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    for (let i=0;i<headers.length;i++){ if (!h[i]) sh.getRange(1, i+1).setValue(headers[i]); }
  }
  return sh;
}
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

function findRowByColumnExact_(sh, colIndex, value){
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, colIndex, last - 1, 1).getValues();
  const needle = String(value);
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0]) === needle) return 2 + i;  // exact/case-sensitive
  }
  return 0;
}

function getOrInitBalanceRemote_(balances, memberId) {
  const row = findRowByColumn_(balances, 1, memberId);
  if (row) return { row, points: Number(balances.getRange(row, 2).getValue() || 0) };
  balances.appendRow([memberId, 0]);
  const newRow = balances.getLastRow();
  return { row: newRow, points: 0 };
}

// ── Secure PIN helpers (PBKDF2-SHA256) ─────────────────────────────────────────
const PIN_ALGO_DEFAULT = 'PBKDF2-SHA256';
const PIN_ITER_DEFAULT = Number(PropertiesService.getScriptProperties().getProperty('PIN_PBKDF2_ITER')) || 75000;
const PIN_PEPPER_B64   = (function(){
  // Optional. Set once via Script Properties, e.g. PIN_PEPPER (base64).
  // If not set, we proceed without pepper (Peppered=false).
  const v = PropertiesService.getScriptProperties().getProperty('PIN_PEPPER');
  return v ? String(v) : '';
})();

function utf8Bytes_(s){ return Utilities.newBlob(String(s)||'').getBytes(); }
function b64enc_(bytes){ return Utilities.base64Encode(bytes); }
function b64dec_(b64){ return Utilities.base64Decode(String(b64)||''); }

function randomSaltB64_(len){
  // Apps Script lacks crypto RNG; emulate with UUID+time hashed a few times.
  // For salt this is acceptable. len ~ 16..32 bytes is fine; we output base64 anyway.
  const seed = Utilities.getUuid() + '|' + Date.now() + '|' + Math.random();
  let bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed);
  // Truncate/extend to len
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


function xorBytes_(a,b){
  const out = new Array(Math.min(a.length,b.length));
  for (let i=0;i<out.length;i++) out[i] = (a[i]^b[i]) & 0xff;
  return out;
}

// PBKDF2-HMAC-SHA256 (RFC 2898) minimal impl; dkLen fixed to 32 for our use.
function pbkdf2Sha256_(passwordBytes, saltBytes, iterations, dkLen){
  const hLen = 32; // SHA256
  const l = Math.ceil(dkLen / hLen);
  const r = dkLen - (l-1)*hLen;
  const dk = [];
  for (let i=1;i<=l;i++){
    // INT(i) big-endian
    const blockIndex = [0,0,0,(i & 0xff)];
    const u1 = hmacSha256_(passwordBytes, saltBytes.concat(blockIndex));
    let t = u1.slice(); // copy
    let u = u1;
    for (let c=2;c<=iterations;c++){
      u = hmacSha256_(passwordBytes, u);
      t = xorBytes_(t, u);
    }
    const outPart = (i===l) ? t.slice(0, r) : t;
    for (let j=0;j<outPart.length;j++) dk.push(outPart[j]);
  }
  return dk;
}

// Derive base64 hash from (pin [+ pepper]) and salt
function derivePinHashB64_(pin, saltB64, iter, usePepper){
  const salt = b64dec_(saltB64);
  const pinBytes = utf8Bytes_(String(pin||''));
  const pepperBytes = (usePepper && PIN_PEPPER_B64) ? b64dec_(PIN_PEPPER_B64) : [];
  const pwd = pinBytes.concat(pepperBytes);
  const dk = pbkdf2Sha256_(pwd, salt, Math.max(1000, Number(iter)||PIN_ITER_DEFAULT), 32);
  return b64enc_(dk);
}

function staffHeaderIdx_(sheet){
  const lastCol = Math.max(sheet.getLastColumn() || 0, 13);
  const hdr = sheet.getRange(1,1,1,lastCol).getValues()[0].map(String);
  const idx = {};
  hdr.forEach((h,i)=> idx[h.trim()] = i+1); // 1-based
  return idx;
}

// Create a salted record payload for a given PIN
function makeSaltedStaffRecord_(pin){
  const saltB64 = randomSaltB64_(24);
  const iter = PIN_ITER_DEFAULT;
  const algo = PIN_ALGO_DEFAULT;
  const peppered = !!PIN_PEPPER_B64;
  const hashB64 = derivePinHashB64_(pin, saltB64, iter, peppered);
  return { saltB64, hashB64, iter, algo, peppered, updatedAt: new Date().toISOString() };
}

// Verify + migrate legacy if needed; returns {ok, migrated}
function verifyStaffPinAndMigrateIfNeeded_(staffSheet, row, pin){
  const idx = staffHeaderIdx_(staffSheet);

  const get = (name)=> {
    const c = idx[name]; if (!c) return '';
    return staffSheet.getRange(row, c).getValue();
  };
  const setRow = (map)=> {
    const updates = {};
    Object.keys(map).forEach(k=> { const c=idx[k]; if (c) updates[c]=map[k]; });
    if (Object.keys(updates).length){
      const cols = Object.keys(updates).map(k=> idx[k]);
      const vals = Object.keys(updates).map(k=> updates[k]);
      // Write each individually to avoid header shifts
      for (let i=0;i<cols.length;i++) staffSheet.getRange(row, cols[i]).setValue(vals[i]);
    }
  };

  const algo = String(get('Algo')||'');
  const saltB64 = String(get('Salt')||'');
  const hashB64 = String(get('Hash')||'');
  const iter = Number(get('Iter')||0);
  const peppered = String(get('Peppered')||'').toLowerCase()==='true';

  if (algo && saltB64 && hashB64 && iter){
    // New scheme present → verify
    const calc = derivePinHashB64_(pin, saltB64, iter, peppered);
    return { ok: (calc===String(hashB64)), migrated:false };
  }

  // Legacy only: verify against PinHash (sha256(pin))
  const legacyHash = String(get('PinHash')||'');
  if (legacyHash){
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin||''));
    const hex = bytes.map(b => (b + 256) % 256).map(b => ('0' + b.toString(16)).slice(-2)).join('');
    const okLegacy = (hex === legacyHash);
    if (okLegacy){
      // Migrate immediately to salted scheme
      const rec = makeSaltedStaffRecord_(pin);
      setRow({
        'Salt': rec.saltB64,
        'Hash': rec.hashB64,
        'Iter': rec.iter,
        'Algo': rec.algo,
        'Peppered': rec.peppered,
        'UpdatedAt': rec.updatedAt
      });
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
function genSid_() { return 's_' + Utilities.getUuid().replace(/-/g, ''); }

// ── Role guard ──
function requireRole_(sid, role) {
  const v = validateMSession(sid);
  if (!v || !v.ok) throw new Error('Not signed in.');
  const r = String(role || '').toLowerCase();
  if (String(v.role || '').toLowerCase() !== r) throw new Error('Insufficient permission.');
  return v; // return session payload (username, merchantId, role)
}

/**
 * Manager resets a staff user's PIN to a temporary one.
 * Forces first-login change and clears that staff's lockouts.
 */
function managerResetStaffPin(sid, staffUsername, tempPin){
  const v = requireRole_(sid, 'manager'); // { ok, username, merchantId, role }
  const targetU = sanitizeUsername_(staffUsername);
  if (!targetU) throw new Error('Invalid staff username.');
  if (!isValidPin_(tempPin)) throw new Error('PIN must be exactly 6 digits.');

  const { staff, mlockouts } = getLocalDb_();
  const r = findRowByColumn_(staff, 1, targetU);
  if (!r) throw new Error('Staff not found.');

  const idx = staffHeaderIdx_(staff);
  const rowVals = staff.getRange(r,1,1,Math.max(staff.getLastColumn()||13,13)).getValues()[0];

  const targetMerchantId = String(rowVals[(idx['MerchantId']||3)-1]||'');
  if (targetMerchantId !== String(v.merchantId)) throw new Error('You can only reset staff under your merchant.');

  // Generate salted record
  const rec = makeSaltedStaffRecord_(tempPin);

  // Persist new PIN + set MustChangePin = TRUE
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

  // Clear lockouts for that staff
  const key = 'staff:' + targetU;
  const lr = findRowByColumn_(mlockouts, 1, key);
  if (lr) mlockouts.deleteRow(lr);

  return { ok:true, username: targetU, merchantId: targetMerchantId };
}

/**
 * Staff completes first-login PIN change (or normal change).
 * Requires current valid session + correct oldPin.
 */
function staffChangePin(sid, oldPin, newPin){
  const v = validateMSession(sid);
  if (!v || !v.ok) throw new Error('Not signed in.');
  if (!isValidPin_(oldPin) || !isValidPin_(newPin)) throw new Error('PIN must be 6 digits.');

  const { staff } = getLocalDb_();
  const r = findRowByColumn_(staff, 1, v.username);
  if (!r) throw new Error('Staff not found.');

  // Verify old pin matches current
  const ok = verifyStaffPinAndMigrateIfNeeded_(staff, r, oldPin).ok;
  if (!ok) throw new Error('Current PIN is incorrect.');

  // Update to new salted pin, clear MustChangePin
  const rec = makeSaltedStaffRecord_(newPin);
  const idx = staffHeaderIdx_(staff);
  const updates = {
    'Salt': rec.saltB64,
    'Hash': rec.hashB64,
    'Iter': rec.iter,
    'Algo': PIN_ALGO_DEFAULT,
    'Peppered': rec.peppered,
    'UpdatedAt': rec.updatedAt,
    'MustChangePin': false
  };
  Object.keys(updates).forEach(k=>{
    const c = idx[k]; if (c) staff.getRange(r, c).setValue(updates[k]);
  });

  return { ok:true };
}

/**
 * (NEW) Allows a manager to list all staff accounts
 * associated with THEIR merchantId.
 */
function managerListMyStaff(sid) {
  // 1. Require 'manager' role. This also returns their merchantId.
  const v = requireRole_(sid, 'manager'); 
  
  const { staff } = getLocalDb_();
  const last = staff.getLastRow();
  if (last < 2) return [];

  const idx = staffHeaderIdx_(staff);
  const myStaff = [];
  
  const vals = staff.getRange(2, 1, last - 1, Math.max(staff.getLastColumn() || 13, 13)).getValues();

  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const staffMerchantId = String(row[(idx['MerchantId'] || 3) - 1] || '');
    
    // 2. Only return staff that match the manager's merchantId
    if (staffMerchantId === v.merchantId) {
      myStaff.push({
        username: String(row[(idx['Username'] || 1) - 1] || ''),
        role: String(row[(idx['Role'] || 4) - 1] || 'staff'),
        active: String(row[(idx['Active'] || 5) - 1]).toString().toLowerCase() !== 'false'
      });
    }
  }
  return myStaff;
}

/**
 * (NEW) Allows a manager to activate or deactivate
 * a staff member within their own merchant.
 */
function managerSetStaffActive(sid, targetUsername, active) {
  // 1. Require 'manager' role.
  const v = requireRole_(sid, 'manager');
  const uname = sanitizeUsername_(targetUsername);
  
  if (uname === v.username) {
    throw new Error('You cannot change your own active status.');
  }

  const { staff } = getLocalDb_();
  const r = findRowByColumn_(staff, 1, uname);
  if (!r) throw new Error('Staff not found.');

  const idx = staffHeaderIdx_(staff);
  
  // 2. Security Check: Verify the target staff is in the manager's merchant.
  const targetMerchantId = String(staff.getRange(r, (idx['MerchantId'] || 3)).getValue() || '');
  if (targetMerchantId !== v.merchantId) {
    throw new Error('Permission denied. Staff is not in your merchant.');
  }

  // 3. Set the active status
  staff.getRange(r, (idx['Active'] || 5)).setValue(!!active);
  
  return { ok: true, username: uname, active: !!active };
}

/**
 * (NEW) Manager: create or upsert a coupon for THIS merchant.
 * type: 'BONUS' (collect bonus points) | 'DISCOUNT' (redeem discount points)
 * mode: '' | 'COLLECT' | 'REDEEM'  (empty = both)
 * value: integer points (>=1)
 * caps: { maxUses, perMemberLimit } (0 = unlimited)
 * window: { startsAtMs, expiresAtMs } (0 = open)
 */
function managerUpsertCoupon(sid, code, mode, type, value, caps, windowObj, notes){
  const v = requireRole_(sid, 'manager');
  const { coupons, config } = getUserDb_();

  // NEW: block direct upserts unless explicitly allowed by policy
  if (!policyAllowDirectCouponUpsert_(config /*, v.merchantId */)) {
    throw new Error('Direct coupon edits are disabled. Please submit a Coupon Request for approval.');
  }

  code = String(code||'').trim(); // preserve case
  if (!/^[A-Za-z0-9._-]{3,32}$/.test(code)) throw new Error('Invalid coupon code.');
  mode = String(mode||'').toUpperCase();
  if (mode && mode !== 'COLLECT' && mode !== 'REDEEM') throw new Error('Mode must be COLLECT, REDEEM or empty.');
  type = String(type||'').toUpperCase();
  if (type !== 'BONUS' && type !== 'DISCOUNT') throw new Error('Type must be BONUS or DISCOUNT.');
  value = Number(value||0);
  if (!Number.isFinite(value) || value <= 0) throw new Error('Value must be a positive integer.');

  const maxUses = Math.max(0, Number((caps && caps.maxUses) || 0));
  const perMemberLimit = Math.max(0, Number((caps && caps.perMemberLimit) || 0));
  const startIso = String(windowObj && windowObj.startIso || '').trim();
  const endIso   = String(windowObj && windowObj.endIso   || '').trim();
  if ((startIso && isNaN(Date.parse(startIso))) || (endIso && isNaN(Date.parse(endIso)))) {
    throw new Error('Invalid dates.');
  }
  const active = true;

  const row = findCouponRow_(coupons, code);
  if (row) {
    coupons.getRange(row,1,1,13).setValues([[
      code, v.merchantId, mode, type, value, 
      maxUses, Number(coupons.getRange(row,7).getValue()||0), // keep UsedCount
      perMemberLimit, startIso, endIso, active, new Date().toISOString(), String(notes||'')
    ]]);
  } else {
    coupons.appendRow([
      code, v.merchantId, mode, type, value, 
      maxUses, 0, perMemberLimit, startIso, endIso, active, new Date().toISOString(), String(notes||'')
    ]);
  }
  return { ok:true, code };
}


/** (NEW) Manager: list my coupons */
function managerListMyCoupons(sid){
  const v = requireRole_(sid, 'manager');
  const { coupons } = getUserDb_();
  const last = coupons.getLastRow();
  if (last < 2) return [];
  const cols = Math.max(13, coupons.getLastColumn()||13);
  const vals = coupons.getRange(2,1,last-1,cols).getValues();
  const out = [];
  for (let i=0;i<vals.length;i++){
    const r = vals[i];
    if (String(r[1]||'') !== String(v.merchantId)) continue;
    out.push({
      code: String(r[0]||''),
      mode: String(r[2]||''),
      type: String(r[3]||''),
      value: Number(r[4]||0),
      maxUses: Number(r[5]||0),
      usedCount: Number(r[6]||0),
      perMemberLimit: Number(r[7]||0),
      startIso: String(r[8]||''),
      endIso: String(r[9]||''),
      active: String(r[10]||'').toUpperCase() !== 'FALSE',
      createdAt: String(r[11]||''),
      notes: String(r[12]||'')
    });
  }
  return out;
}

/** (NEW) Manager: set active flag for a coupon (mine only) */
function managerSetCouponActive(sid, code, active){
  const v = requireRole_(sid, 'manager');
  const { coupons, config } = getUserDb_();

  // NEW: block live active toggles unless explicitly allowed by policy
  if (!policyAllowManagerSetCouponActive_(config /*, v.merchantId */)) {
    throw new Error('Direct coupon activation/deactivation is disabled. Please submit a Coupon Request.');
  }

  const row = findCouponRow_(coupons, String(code||'').trim());
  if (!row) throw new Error('Coupon not found.');
  const merchantId = String(coupons.getRange(row,2).getValue()||'');
  if (merchantId !== String(v.merchantId)) throw new Error('Permission denied.');
  coupons.getRange(row,11).setValue(!!active);
  return { ok:true, code: String(code||'').trim(), active: !!active };
}


function _genCouponRequestId_(){ return 'CPRQ-' + Math.random().toString(36).slice(2, 8).toUpperCase(); }

function managerCreateCouponRequest(sid, payload){
  const v = requireRole_(sid, 'manager');
  const { couponRequests, logs } = getUserDb_();
  const p = payload || {};

  const code = String(p.code||'').trim();
  const mode = String(p.mode||'').toUpperCase();
  const type = String(p.type||'').toUpperCase();
  const value = Math.max(1, Number(p.value||0));
  const maxUses = Math.max(0, Number(p.maxUses||0));
  const perMemberLimit = Math.max(0, Number(p.perMemberLimit||0));
  const startIso = String(p.startIso||'').trim();
  const endIso   = String(p.endIso||'').trim();
  const notes = String(p.notes||'');

  if (!/^[A-Za-z0-9._-]{3,32}$/.test(code)) throw new Error('Invalid code.');
  if (mode && mode !== 'COLLECT' && mode !== 'REDEEM') throw new Error('Invalid mode.');
  if (type !== 'BONUS' && type !== 'DISCOUNT') throw new Error('Invalid type.');
  if ((startIso && isNaN(Date.parse(startIso))) || (endIso && isNaN(Date.parse(endIso)))) throw new Error('Invalid dates.');

  const reqId = _genCouponRequestId_();
  const now = new Date().toISOString();
  couponRequests.appendRow([
    reqId, v.merchantId,
    code, mode, type, value,
    maxUses, perMemberLimit,
    startIso, endIso,
    notes,
    'pending', v.username, now, now,
    '', '', ''
  ]);
  log_(logs, 'm_req_coupon_new', JSON.stringify({ reqId, merchantId:v.merchantId, code }));
  return { ok:true, requestId: reqId };
}

function managerListMyCouponRequests(sid){
  const v = requireRole_(sid, 'manager');
  const { couponRequests } = getUserDb_();
  const last = couponRequests.getLastRow();
  if (last < 2) return [];
  const cols = Math.max(18, couponRequests.getLastColumn()||18);
  const vals = couponRequests.getRange(2,1,last-1,cols).getValues();
  const out = [];
  for (const r of vals){
    if (String(r[1]||'') !== String(v.merchantId)) continue;
    out.push({
      requestId: String(r[0]||''),
      code: String(r[2]||''),
      mode: String(r[3]||''),
      type: String(r[4]||''),
      value: Number(r[5]||0),
      maxUses: Number(r[6]||0),
      perMemberLimit: Number(r[7]||0),
      startIso: String(r[8]||''),
      endIso: String(r[9]||''),
      notes: String(r[10]||''),
      status: String(r[11]||''),
      createdBy: String(r[12]||''),
      createdAt: String(r[13]||'')
    });
  }
  return out;
}

function managerUpdateCouponRequest(sid, requestId, payload){
  const v = requireRole_(sid, 'manager');
  const { couponRequests, logs } = getUserDb_();
  const r = findRowByColumn_(couponRequests, 1, requestId);
  if (!r) throw new Error('Request not found.');
  // columns map (1-based)
  const hdr = headerIndexByName_(couponRequests);
  const mid = String(couponRequests.getRange(r, hdr['MerchantId']).getValue()||'');
  const status = String(couponRequests.getRange(r, hdr['Status']).getValue()||'').toLowerCase();
  if (mid !== v.merchantId) throw new Error('Permission denied.');
  if (status !== 'pending') throw new Error('Only pending requests can be modified.');

  const map = {
    'Code': payload.code,
    'Mode': payload.mode,
    'Type': payload.type,
    'Value': payload.value,
    'MaxUses': payload.maxUses,
    'PerMemberLimit': payload.perMemberLimit,
    'StartIso': payload.startIso,
    'EndIso': payload.endIso,
    'Notes': payload.notes
  };
  Object.keys(map).forEach(k=>{
    if (typeof map[k] !== 'undefined') {
      const c = hdr[k]; if (c) couponRequests.getRange(r, c).setValue(map[k]);
    }
  });
  couponRequests.getRange(r, hdr['UpdatedAt']).setValue(new Date().toISOString());
  log_(logs, 'm_req_coupon_update', JSON.stringify({requestId, merchantId:v.merchantId}));
  return { ok:true, requestId };
}

function managerCancelCouponRequest(sid, requestId){
  const v = requireRole_(sid, 'manager');
  const { couponRequests, logs } = getUserDb_();
  const r = findRowByColumn_(couponRequests, 1, requestId);
  if (!r) throw new Error('Request not found.');
  const hdr = headerIndexByName_(couponRequests);
  const mid = String(couponRequests.getRange(r, hdr['MerchantId']).getValue()||'');
  const status = String(couponRequests.getRange(r, hdr['Status']).getValue()||'').toLowerCase();
  if (mid !== v.merchantId) throw new Error('Permission denied.');
  if (status !== 'pending') throw new Error('Only pending requests can be cancelled.');

  const now = new Date().toISOString();
  couponRequests.getRange(r, hdr['Status']).setValue('cancelled');
  couponRequests.getRange(r, hdr['DecisionBy']).setValue(v.username);
  couponRequests.getRange(r, hdr['DecisionAt']).setValue(now);
  couponRequests.getRange(r, hdr['DecisionNotes']).setValue('Cancelled by merchant.');
  couponRequests.getRange(r, hdr['UpdatedAt']).setValue(now);
  log_(logs, 'm_req_coupon_cancel', JSON.stringify({requestId, merchantId:v.merchantId}));
  return { ok:true, requestId };
}

// ───────────────────────────────── Multiplier Campaign Management (Option B) ─────────────────────────────────

/**
 * Helper to generate a unique Request ID.
 */
function _genCampaignRequestId_(){ return 'CRQ-' + Math.random().toString(36).slice(2, 8).toUpperCase(); }

/**
 * Merchant Manager: Submit a NEW campaign request.
 * @param {string} sid Session ID.
 * @param {object} payload Campaign details.
 */
function merchantCreateCampaignRequest(sid, payload) {
    const v = requireRole_(sid, 'manager');
    const { campaignRequests, logs } = getUserDb_();

    const p = payload || {};
    const title = String(p.title || '').trim();
    const multiplier = Math.max(1, Number(p.multiplier || 1));
    const startIso = String(p.startIso || '').trim();
    const endIso = String(p.endIso || '').trim();
    const minSpend = Math.max(0, Number(p.minSpend || 0));
    const maxRedemptions = Math.max(0, Number(p.maxRedemptions || 0));          // GLOBAL
    const perMemberRedemptions = Math.max(0, Number(p.perMemberRedemptions || 0)); // ← NEW per-member
    const budgetCap = Math.max(0, Number(p.budgetCap || 0));
    const billingModel = String(p.billingModel || 'per_redemption').toLowerCase();
    const perMemberBonusCap = Math.max(0, Number(p.perMemberBonusCap || 0));

    

    if (!title || !startIso || !endIso) throw new Error('Title, start date, and end date are required.');
    if (isNaN(Date.parse(startIso)) || isNaN(Date.parse(endIso))) throw new Error('Invalid date format.');
    if (multiplier < 1) throw new Error('Multiplier must be >= 1.');
    if (billingModel !== 'per_redemption' && billingModel !== 'flat_monthly') throw new Error('Invalid billing model.');

    const requestId = _genCampaignRequestId_();
    const nowIso = new Date().toISOString();
    // Use the generic indexer if you ever need positions (not required for appendRow)
    const headers = headerIndexByName_(campaignRequests);

    // Your row already matches the header order you set in getUserDb_()
    const row = [
      requestId, v.merchantId, 'new', title, multiplier,
      startIso, endIso, minSpend,
      maxRedemptions, perMemberRedemptions,
      budgetCap,
      billingModel, Number(p.costPerRedemption || 0), String(p.notes || ''), 
      String(p.imageUrl || ''), // ImageUrl (optional)
      'pending', v.username, nowIso, nowIso,
      '', '', '', '',
      perMemberBonusCap
    ];

    campaignRequests.appendRow(row);


    log_(logs, 'm_req_new_campaign', JSON.stringify({ requestId, merchantId: v.merchantId, title }));
    return { ok: true, requestId };
}

/**
 * Merchant Manager: List my active campaigns (read-only from Admin's live Campaigns sheet).
 */
function merchantListCampaigns(sid) {
  const v = requireRole_(sid, 'manager'); // { ok, username, merchantId, role }
  const myMid = String(v.merchantId || '').trim().toUpperCase();

  const { campaigns } = getUserDb_();
  const last = campaigns.getLastRow();
  if (last < 2) return [];

  const idx = headerIndexByName_(campaigns); // helper below
  const cols = campaigns.getLastColumn();
  const rows = campaigns.getRange(2, 1, last - 1, cols).getValues();

  const nowMs = Date.now();
  const out = [];

  for (const r of rows) {
    const get = (name) => {
      const c = idx[name];
      return c ? r[c - 1] : '';
    };

    const mid = String(get('MerchantId') || '').trim().toUpperCase();
    if (!mid || mid !== myMid) continue;

    const active = String(get('Active') || 'TRUE').toLowerCase() !== 'false';
    if (!active) continue;

    const startMs = parseToMsSafe_(get('StartIso'));
    const endMs   = parseToMsSafe_(get('EndIso'));
    if (!(startMs && endMs && startMs <= nowMs && nowMs <= endMs)) continue;

    out.push({
      campaignId: String(get('CampaignId') || ''),
      merchantId: mid,
      title: String(get('Title') || ''),
      type: String(get('Type') || ''),
      multiplier: Number(get('Multiplier') || 1),
      startIso: String(get('StartIso') || ''),
      endIso: String(get('EndIso') || ''),
      minSpend: Number(get('MinSpend') || 0),
      maxRedemptions: Number(get('MaxRedemptions') || 0),      // GLOBAL
      perMemberRedemptions: Number(get('PerMemberRedemptions') || 0), // ← NEW
      budgetCap: Number(get('BudgetCap') || 0),
      billingModel: String(get('BillingModel') || 'per_redemption'),
      costPerRedemption: Number(get('CostPerRedemption') || 0),
      active: true,
      createdAt: get('CreatedAt') || '',
      updatedAt: get('UpdatedAt') || '',
      imageUrl: String(get('ImageUrl') || ''),
      perMemberBonusCap: Number(get('PerMemberBonusCap') || 0) // NEW
    });
  }

  // Newest first
  out.sort((a, b) =>
    (Date.parse(b.updatedAt || b.createdAt || 0) || 0) -
    (Date.parse(a.updatedAt || a.createdAt || 0) || 0)
  );

  return out;
}

/** Build a 1-based header index by exact header text. */
function headerIndexByName_(sheet) {
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idx = {};
  for (let i = 0; i < hdr.length; i++) {
    const key = String(hdr[i] || '').trim();
    if (key) idx[key] = i + 1;
  }
  return idx;
}

/** Parse ISO string or serial-like number into ms; returns 0 on failure. */
function parseToMsSafe_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') {
    // Apps Script sometimes returns Date objects as numbers when coerced
    try { return new Date(v).getTime(); } catch (_) { /* ignore */ }
  }
  const ms = Date.parse(String(v));
  return isNaN(ms) ? 0 : ms;
}

function findActiveCampaignForMerchant_(campaignsSheet, merchantId) {
  const last = campaignsSheet.getLastRow();
  if (last < 2) return null;

  const idx = headerIndexByName_(campaignsSheet);
  const cols = campaignsSheet.getLastColumn();
  const rows = campaignsSheet.getRange(2, 1, last - 1, cols).getValues();
  const nowMs = Date.now();
  const mid = String(merchantId || '').trim().toUpperCase();

  const candidates = [];
  for (const r of rows) {
    const get = (name) => { const c = idx[name]; return c ? r[c - 1] : ''; };

    const rowMid = String(get('MerchantId') || '').trim().toUpperCase();
    if (rowMid !== mid) continue;

    const active = String(get('Active') || 'TRUE').toLowerCase() !== 'false';
    if (!active) continue;

    if (String(get('Type') || '').toLowerCase() !== 'multiplier') continue;

    const startMs = parseToMsSafe_(get('StartIso'));
    const endMs   = parseToMsSafe_(get('EndIso'));
    if (!(startMs && endMs && startMs <= nowMs && nowMs <= endMs)) continue;

    candidates.push({
      campaignId: String(get('CampaignId') || ''),
      title: String(get('Title') || ''),
      multiplier: Number(get('Multiplier') || 1),
      minSpend: Number(get('MinSpend') || 0),
      maxRedemptions: Number(get('MaxRedemptions') || 0),             // GLOBAL cap
      perMemberRedemptions: Number(get('PerMemberRedemptions') || 0), // ← NEW per member
      perMemberBonusCap: Number(get('PerMemberBonusCap') || 0),
      billingModel: String(get('BillingModel') || 'per_redemption').toLowerCase(),
      costPerRedemption: Number(get('CostPerRedemption') || 0),
      createdAt: String(get('CreatedAt') || ''),
      updatedAt: String(get('UpdatedAt') || '')
    });
  }

  if (!candidates.length) return null;
  candidates.sort((a, b) =>
    (Date.parse(b.updatedAt || b.createdAt || 0) || 0) -
    (Date.parse(a.updatedAt || a.createdAt || 0) || 0)
  );
  return candidates[0];
}
/**
 * Merchant Manager: List my campaign requests (pending/approved/rejected).
 */
function merchantListCampaignRequests(sid) {
  const v = requireRole_(sid, 'manager');
  const myMid = String(v.merchantId || '').trim().toUpperCase();

  const { campaignRequests } = getUserDb_();
  const last = campaignRequests.getLastRow();
  if (last < 2) return [];

  // Be tolerant to sheet header order/case
  const headers = headerIndexByName_(campaignRequests); // <— use the generic header resolver
  const getVal = (r, name) => {
    const c = headers[name];
    return c ? r[c - 1] : '';
  };

  const colCount = Math.max(campaignRequests.getLastColumn() || 1, Object.keys(headers).length);
  const allRows = campaignRequests.getRange(2, 1, last - 1, colCount).getValues();

  const out = [];
  for (const r of allRows) {
    const mid = String(getVal(r, 'MerchantId') || '').trim().toUpperCase();
    if (mid !== myMid) continue;

    out.push({
      requestId: String(getVal(r, 'RequestId') || ''),
      type: String(getVal(r, 'RequestType') || ''),
      title: String(getVal(r, 'Title') || ''),
      multiplier: Number(getVal(r, 'Multiplier') || 0),
      status: String(getVal(r, 'Status') || ''),
      createdBy: String(getVal(r, 'CreatedBy') || ''),
      createdAt: String(getVal(r, 'CreatedAt') || ''),
      linkedCampaignId: String(getVal(r, 'LinkedCampaignId') || ''),
      statusNote: String(getVal(r, 'DecisionNotes') || ''),
      startIso: String(getVal(r, 'StartIso') || ''),
      endIso: String(getVal(r, 'EndIso') || ''),
      minSpend: Number(getVal(r, 'MinSpend') || 0),
      maxRedemptions: Number(getVal(r, 'MaxRedemptions') || 0),
      budgetCap: Number(getVal(r, 'BudgetCap') || 0),
      billingModel: String(getVal(r, 'BillingModel') || ''),
      costPerRedemption: Number(getVal(r, 'CostPerRedemption') || 0),
      imageUrl: String(getVal(r, 'ImageUrl') || ''),
      perMemberBonusCap: Number(getVal(r, 'PerMemberBonusCap') || 0)
    });
  }
  return out;
}


/**
 * Merchant Manager: Update a PENDING campaign request.
 */
function merchantUpdateCampaignRequest(sid, requestId, updatePayload) {
    const v = requireRole_(sid, 'manager');
    const { campaignRequests, logs } = getUserDb_();

    const r = findRowByColumn_(campaignRequests, 1, requestId);
    if (!r) throw new Error('Request not found.');

    const headers = staffHeaderIdx_(campaignRequests);
    const merchantId = String(campaignRequests.getRange(r, headers['MerchantId']).getValue() || '');
    const status = String(campaignRequests.getRange(r, headers['Status']).getValue() || '').toLowerCase();

    if (merchantId !== v.merchantId) throw new Error('Permission denied. Cannot edit another merchant\'s request.');
    if (status !== 'pending') throw new Error('Only pending requests can be modified.');
    
    const nowIso = new Date().toISOString();
    const map = {};

    // Map allowed editable fields (subset of 'new' fields)
    const editableFields = {
      'title': 'Title', 'multiplier': 'Multiplier', 'startIso': 'StartIso',
      'endIso': 'EndIso', 'minSpend': 'MinSpend', 'maxRedemptions': 'MaxRedemptions',
      'budgetCap': 'BudgetCap', 'billingModel': 'BillingModel', 'costPerRedemption': 'CostPerRedemption',
      'notes': 'Notes', 'imageUrl': 'ImageUrl', 'requestType': 'RequestType',
      'perMemberBonusCap': 'PerMemberBonusCap' // NEW
    };


    Object.keys(updatePayload).forEach(key => {
        const header = editableFields[key];
        if (header) {
            const value = updatePayload[key];
            map[header] = typeof value === 'string' ? value.trim() : value;
        }
    });
    
    // Validation checks for core fields if updated
    if ('Multiplier' in map) map['Multiplier'] = Math.max(1, Number(map['Multiplier'] || 1));
    if ('StartIso' in map || 'EndIso' in map) {
        const startIso = 'StartIso' in map ? map['StartIso'] : campaignRequests.getRange(r, headers['StartIso']).getValue();
        const endIso = 'EndIso' in map ? map['EndIso'] : campaignRequests.getRange(r, headers['EndIso']).getValue();
        if (isNaN(Date.parse(startIso)) || isNaN(Date.parse(endIso))) throw new Error('Invalid date format in update.');
    }

    Object.keys(map).forEach(k => {
        const c = headers[k];
        if (c) campaignRequests.getRange(r, c).setValue(map[k]);
    });

    // Update timestamp
    campaignRequests.getRange(r, headers['UpdatedAt']).setValue(nowIso);

    log_(logs, 'm_req_update_campaign', JSON.stringify({ requestId, merchantId: v.merchantId, user: v.username }));
    return { ok: true, requestId };
}


/**
 * Merchant Manager: Cancel a PENDING campaign request.
 */
function merchantCancelCampaignRequest(sid, requestId) {
    const v = requireRole_(sid, 'manager');
    const { campaignRequests, logs } = getUserDb_();

    const r = findRowByColumn_(campaignRequests, 1, requestId);
    if (!r) throw new Error('Request not found.');

    const headers = staffHeaderIdx_(campaignRequests);
    const merchantId = String(campaignRequests.getRange(r, headers['MerchantId']).getValue() || '');
    const status = String(campaignRequests.getRange(r, headers['Status']).getValue() || '').toLowerCase();

    if (merchantId !== v.merchantId) throw new Error('Permission denied. Cannot cancel another merchant\'s request.');
    if (status !== 'pending') throw new Error('Only pending requests can be cancelled.');

    const nowIso = new Date().toISOString();

    // Update status to 'cancelled'
    campaignRequests.getRange(r, headers['Status']).setValue('cancelled');
    campaignRequests.getRange(r, headers['DecisionBy']).setValue(v.username); // Recorded as self-cancelled
    campaignRequests.getRange(r, headers['DecisionAt']).setValue(nowIso);
    campaignRequests.getRange(r, headers['DecisionNotes']).setValue('Cancelled by merchant.');
    campaignRequests.getRange(r, headers['UpdatedAt']).setValue(nowIso);

    log_(logs, 'm_req_cancel_campaign', JSON.stringify({ requestId, merchantId: v.merchantId, user: v.username }));
    return { ok: true, requestId };
}
// ───────────────────────────────── End of Multiplier Campaign Management ─────────────────────────────────

// ── Scope guard (WRITE BARRIER) ──
// Always call this before any write to user DB (TX, balances, etc).
// It validates session, resolves the merchant row in the remote DB, and enforces Active=TRUE.
function getMerchantScopeOrThrow_(sid) {
  const v = validateMSession(sid);
  if (!v || !v.ok) throw new Error('Not signed in.');

  const { merchants } = getUserDb_();
  const mr = findRowByColumn_(merchants, 1, v.merchantId);
  if (!mr) throw new Error('Merchant not found.');
  const active = String(merchants.getRange(mr, 3).getValue() || 'TRUE').toUpperCase() !== 'FALSE';
  if (!active) throw new Error('Merchant inactive.');

  // NEW: enforce first-login change barrier on the server for *all* roles
  const { staff } = getLocalDb_();
  const sr = findRowByColumn_(staff, 1, v.username);
  if (sr) {
    const idx = staffHeaderIdx_(staff);
    const mustChange = String(staff.getRange(sr, (idx['MustChangePin']||13)).getValue()||'')
      .toString().toLowerCase() === 'true';
    if (mustChange) {
      throw new Error('PIN change required. Please set a new PIN before using POS.');
    }
  }

  // v: { ok, sid, username, merchantId, role }
  return v;
}


// ── Time helpers ──
function _ymd(ms) {
  const d = new Date(ms); // Apps Script uses script timezone
  const y = d.getFullYear();
  const m = ('0' + (d.getMonth() + 1)).slice(-2);
  const day = ('0' + d.getDate()).slice(-2);
  return y + '-' + m + '-' + day;
}
function _startOfTodayMs_() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d.getTime();
}
function _rangeWindow_(range) {
  const r = String(range || '7d').toLowerCase();
  const today0 = _startOfTodayMs_();
  if (r === 'today') return { startMs: today0, labelDays: 1 };
  if (r === '30d')   return { startMs: today0 - 29 * 24 * 3600 * 1000, labelDays: 30 };
  // default 7d
  return { startMs: today0 - 6 * 24 * 3600 * 1000, labelDays: 7 };
}


// ── Lockouts ──
function recordStaffFailure_(mlockouts, key) {
  const now = Date.now();
  let r = findRowByColumn_(mlockouts, 1, key);

  if (!r) {
    // Key, FailCount, LockUntilMs, Permanent, UpdatedAt
    mlockouts.appendRow([key, 1, 0, false, new Date(now).toISOString()]);
    return;
  }

  const cols = Math.max(5, mlockouts.getLastColumn()||5);
  const [k, count, lockUntil, permanent] = mlockouts.getRange(r,1,1,cols).getValues()[0];
  const newCount = Number(count||0) + 1;

  let newLock = Number(lockUntil||0);
  let perma = String(permanent||'').toLowerCase() === 'true';

  // Escalation:
  if (!perma && newCount >= M_LOCK_PERMA_ATTEMPTS) {
    perma = true;
    newLock = 0; // no countdown; needs manager/admin to clear
  } else if (newCount >= M_LOCK_MAX_ATTEMPTS) {
    // cooldown lock if not yet permanent
    if (!perma) newLock = now + M_LOCK_DURATION_MS;
  }

  mlockouts.getRange(r,1,1,5).setValues([[key, newCount, newLock, perma, new Date(now).toISOString()]]);
}

function clearStaffLockout_(mlockouts, key) {
  const r = findRowByColumn_(mlockouts, 1, key);
  if (r) mlockouts.deleteRow(r);
}

// ───────────────────────────────── Config helpers (remote)
function getConfigValueRemote_(configSheet, key){
  if (!configSheet || !key) return '';
  const last = configSheet.getLastRow();
  if (last < 2) return '';
  const vals = configSheet.getRange(2,1,last-1,2).getValues();
  const needle = String(key||'').trim().toUpperCase();
  for (let i=0;i<vals.length;i++){
    const k = String(vals[i][0]||'').trim().toUpperCase();
    if (k === needle) return String(vals[i][1]||'').trim();
  }
  return '';
}

// ───────────────────────────────── Policy flags (remote Config)
// Keys live in the User DB → "Config" sheet. Default = false.
function getBooleanConfigRemote_(configSheet, key, defVal){
  const v = String(getConfigValueRemote_(configSheet, key) || '').trim().toLowerCase();
  if (v === 'true')  return true;
  if (v === 'false') return false;
  return !!defVal;
}

// If you ever need per-merchant overrides later, you can extend these to read a keyed variant
// like ALLOW_DIRECT_COUPON_UPSERT:MID, falling back to global.
function policyAllowDirectCouponUpsert_(configSheet /*, merchantId */){
  return getBooleanConfigRemote_(configSheet, 'ALLOW_DIRECT_COUPON_UPSERT', false);
}
function policyAllowManagerSetCouponActive_(configSheet /*, merchantId */){
  return getBooleanConfigRemote_(configSheet, 'ALLOW_MANAGER_SET_COUPON_ACTIVE', false);
}


// ───────────────────────────────── Coupons helpers (remote)

const COUPON_TYPES = {
  // Collect: adds fixed bonus points to COLLECT
  BONUS: 'BONUS',
  // Redeem: reduces the points to deduct by a fixed amount on REDEEM
  DISCOUNT: 'DISCOUNT'
};
// Mode: 'COLLECT' | 'REDEEM' | '' (empty means both)

function findCouponRow_(coupons, code){
  return findRowByColumnExact_(coupons, 1, String(code||'').trim()); // exact/case-sensitive
}


function getCouponRecord_(coupons, row){
  const cols = Math.max(13, coupons.getLastColumn()||13);
  const v = coupons.getRange(row,1,1,cols).getValues()[0];
  return {
    code: String(v[0]||''), // keep exact case
    merchantId: String(v[1]||''),
    mode: String(v[2]||'').toUpperCase(),
    type: String(v[3]||'').toUpperCase(),
    value: Number(v[4]||0),
    maxUses: Number(v[5]||0),
    usedCount: Number(v[6]||0),
    perMemberLimit: Number(v[7]||0),
    startIso: String(v[8]||''),
    endIso:   String(v[9]||''),
    active: String(v[10]||'').toUpperCase() !== 'FALSE',
    createdAt: String(v[11]||''),
    notes: String(v[12]||'')
  };
}

function countMemberCouponUses_(cpnUses, code, memberId){
  const last = cpnUses.getLastRow();
  if (last < 2) return 0;
  const cols = Math.max(6, cpnUses.getLastColumn()||6);
  const vals = cpnUses.getRange(2,1,last-1,cols).getValues();
  let n = 0;
  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    if (String(row[0]||'') === String(code||'') &&
        String(row[1]||'').toUpperCase() === String(memberId||'').toUpperCase()){
      n++;
    }
  }
  return n;
}

/**
 * Validate coupon for the merchant/member/mode/now and compute effect.
 * Returns { ok, reason?, effectivePoints, originalPoints, couponApplied? }
 *
 * Rules:
 * - Mode must match (or coupon mode empty = both)
 * - Active + date window
 * - MaxUses + PerMemberLimit
 * - Type BONUS (collect only): adds fixed value to points
 * - Type DISCOUNT (redeem only): reduces points to deduct by fixed value (floored at 0)
 */
function validateAndApplyCoupon_(coupons, cpnUses, vScope, memberId, mode, points, code){
  if (!code) return { ok:true, effectivePoints: points, originalPoints: points, couponApplied: null };
  const now = Date.now();
  const r = findCouponRow_(coupons, code);
  if (!r) return { ok:false, reason:'coupon_not_found' };

  const c = getCouponRecord_(coupons, r);

  // Merchant scope
  if (String(c.merchantId) !== String(vScope.merchantId)) return { ok:false, reason:'coupon_wrong_merchant' };

  // Active window (ISO-based)
  if (!c.active) return { ok:false, reason:'coupon_inactive' };

  const startMs = c.startIso ? Date.parse(c.startIso) : 0;
  const endMs   = c.endIso   ? Date.parse(c.endIso)   : 0;

  if (startMs && now < startMs) return { ok:false, reason:'coupon_not_started' };
  if (endMs && now > endMs)     return { ok:false, reason:'coupon_expired' };


  // Mode
  const m = String(mode||'').toUpperCase();
  if (c.mode && c.mode !== m) return { ok:false, reason:'coupon_wrong_mode' };

  // Usage caps
  if (c.maxUses > 0 && c.usedCount >= c.maxUses) return { ok:false, reason:'coupon_exhausted' };
  if (c.perMemberLimit > 0) {
    const usedByMember = countMemberCouponUses_(cpnUses, c.code, memberId);
    if (usedByMember >= c.perMemberLimit) return { ok:false, reason:'coupon_member_limit' };
  }

  // Effect
  let effective = Number(points||0);
  if (c.type === COUPON_TYPES.BONUS) {
    if (m !== 'COLLECT') return { ok:false, reason:'coupon_type_mismatch' };
    effective = Math.max(0, effective + Number(c.value||0));
  } else if (c.type === COUPON_TYPES.DISCOUNT) {
    if (m !== 'REDEEM') return { ok:false, reason:'coupon_type_mismatch' };
    effective = Math.max(0, effective - Number(c.value||0));
  } else {
    return { ok:false, reason:'coupon_unknown_type' };
  }

  return { ok:true, effectivePoints: effective, originalPoints: Number(points||0), couponApplied: c };
}

function countMemberCampaignRedemptions_(cpnRedSheet, campaignId, memberId) {
  const last = cpnRedSheet.getLastRow();
  if (last < 2) return 0;
  const cols = Math.max(9, cpnRedSheet.getLastColumn() || 9);
  const vals = cpnRedSheet.getRange(2, 1, last - 1, cols).getValues();
  let n = 0;
  for (const r of vals) {
    if (String(r[1] || '') === String(campaignId || '') &&
        String(r[2] || '').toUpperCase() === String(memberId || '').toUpperCase()) {
      n++;
    }
  }
  return n;
}

function countTotalCampaignRedemptions_(cpnRedSheet, campaignId) {
  const last = cpnRedSheet.getLastRow();
  if (last < 2) return 0;
  const cols = Math.max(9, cpnRedSheet.getLastColumn() || 9);
  const vals = cpnRedSheet.getRange(2, 1, last - 1, cols).getValues();
  let n = 0;
  for (const r of vals) {
    if (String(r[1] || '') === String(campaignId || '')) n++;
  }
  return n;
}

function sumMemberCampaignBonusPoints_(cpnRedSheet, campaignId, memberId) {
  const last = cpnRedSheet.getLastRow();
  if (last < 2) return 0;
  const cols = Math.max(9, cpnRedSheet.getLastColumn() || 9);
  const vals = cpnRedSheet.getRange(2, 1, last - 1, cols).getValues();
  let s = 0;
  for (const r of vals) {
    if (String(r[1] || '') === String(campaignId || '') &&
        String(r[2] || '').toUpperCase() === String(memberId || '').toUpperCase()) {
      s += Number(r[8] || 0); // BonusPoints col
    }
  }
  return s;
}

/**
 * Mark a coupon use (append to uses + increment UsedCount).
 * Best effort; throws if ref row is missing.
 */
function recordCouponUse_(coupons, cpnUses, couponRow, couponCode, memberId, merchantId, staff, txId){
  const now = Date.now();
  cpnUses.appendRow([String(couponCode||''), String(memberId||''), String(merchantId||''), now, String(staff||''), String(txId||'')]);

  // bump UsedCount
  const usedCol = 7; // UsedCount index in Coupons (1-based)
  const prev = Number(coupons.getRange(couponRow, usedCol).getValue()||0);
  coupons.getRange(couponRow, usedCol).setValue(prev + 1);
}


// ───────────────────────────────── Base64url helpers
function b64urlToBytes_(b64u){
  let b = String(b64u||'').replace(/-/g,'+').replace(/_/g,'/');
  while (b.length % 4) b += '=';
  return b64dec_(b);
}
function bytesToB64url_(bytes){
  return b64enc_(bytes).replace(/\=+$/,'').replace(/\+/g,'-').replace(/\//g,'_');
}
function timingSafeEq_(a,b){
  if (!a || !b) return false;
  if (a.length !== b.length) return false;
  let out = 0;
  for (let i=0;i<a.length;i++) out |= (a[i]^b[i]);
  return out === 0;
}

// ───────────────────────────────── Signed QR verifier (Option A)
function resolveSignedQrPayload(payloadB64u){
  const { config, logs } = getUserDb_();

  const keyB64 = getConfigValueRemote_(config, 'QR_HMAC_KEY_B64'); // set in User DB Config
  const ttlSec = Number(getConfigValueRemote_(config, 'QR_TTL_SEC') || 180);
  if (!keyB64) {
    log_(logs, 'qr_verify_fail', 'missing_key');
    return { ok:false, reason:'server_not_configured' };
  }
  const keyBytes = b64dec_(keyB64);

  // --- Path A: compact "CB1" format: CB1.<MID>.<EXP_MS>[.<NONCE>].<SIG> ---
  // Examples we see: CB1.MBR-XXXXXXXX.1763586....<nonceOrShort>.<sig>
  try {
    const raw = String(payloadB64u || '');
    if (raw.startsWith('CB1.')) {
      const parts = raw.split('.');
      // Allow 4 or 5 parts to be tolerant: CB1.MID.EXP.SIG  OR  CB1.MID.EXP.NONCE.SIG
      if (parts.length === 4 || parts.length === 5) {
        const vStr  = parts[0];              // "CB1"
        const mid   = String(parts[1] || '');
        const exp   = Number(parts[2] || 0);
        const maybeNonce = (parts.length === 5) ? String(parts[3] || '') : '';
        const sigB64u    = String(parts[parts.length - 1] || '');

        // Basic field validation
        if (!/^MBR-[A-Z0-9]{8}$/.test(mid)) { 
          log_(logs, 'qr_verify_fail', 'cb1_bad_mid'); 
          return { ok:false, reason:'invalid_fields' }; 
        }
        if (!exp) { 
          log_(logs, 'qr_verify_fail', 'cb1_bad_exp'); 
          return { ok:false, reason:'invalid_fields' }; 
        }

        const now = Date.now();
        if (exp <= now) {
          log_(logs, 'qr_verify_fail', 'cb1_expired');
          return { ok:false, reason:'expired' };
        }
        if (!_qrTtlCapOk_(exp, now, ttlSec)) {
          log_(logs, 'qr_verify_fail', 'cb1_ttl_exceeded');
          return { ok:false, reason:'ttl_exceeded' };
        }

        // Canonical string (same as JSON flow): "v|mid|mode|exp"
        // v = 1 (from "CB1"); mode is not carried in compact format → use empty string
        const canonical = ['1', mid, '', String(exp)].join('|');
        const mac       = hmacSha256_(keyBytes, utf8Bytes_(canonical));
        const calcB64u  = bytesToB64url_(mac);

        // timing-safe compare
        const sigBytes  = b64urlToBytes_(sigB64u);
        const calcBytes = b64urlToBytes_(calcB64u);
        if (!timingSafeEq_(sigBytes, calcBytes)) {
          log_(logs, 'qr_verify_fail', 'cb1_bad_signature');
          return { ok:false, reason:'bad_signature' };
        }

        log_(logs, 'qr_verify_ok', 'cb1 mid='+mid);
        return { ok:true, memberId: mid, mode: '', exp };
      }
      // If it started with CB1. but parts don’t match, fall through to generic failure
    }
  } catch (e) {
    // Fall through to JSON attempt below
  }

  // --- Path B: original base64url-of-JSON format ---
  let jsonStr = '';
  try{
    const raw = b64urlToBytes_(payloadB64u);
    jsonStr = Utilities.newBlob(raw).getDataAsString(); // UTF-8
  }catch(e){
    log_(logs, 'qr_verify_fail', 'bad_base64url');
    return { ok:false, reason:'bad_payload' };
  }

  let obj;
  try { obj = JSON.parse(jsonStr); } catch(e){
    log_(logs, 'qr_verify_fail', 'bad_json');
    return { ok:false, reason:'bad_payload' };
  }

  const v   = Number(obj && obj.v);
  const mid = String(obj && obj.mid || '');
  const mode = String(obj && obj.mode || '').toUpperCase(); // optional
  const exp = Number(obj && obj.exp || 0);
  const sig = String(obj && obj.sig || '');

  if (v !== 1 || !/^MBR-[A-Z0-9]{8}$/.test(mid) || !exp || !sig) {
    log_(logs, 'qr_verify_fail', 'invalid_fields');
    return { ok:false, reason:'invalid_fields' };
  }

  const now = Date.now();
  if (exp <= now) { 
    log_(logs, 'qr_verify_fail', 'expired'); 
    return { ok:false, reason:'expired' }; 
  }
  if (!_qrTtlCapOk_(exp, now, ttlSec)) {
    log_(logs, 'qr_verify_fail', 'ttl_exceeded');
    return { ok:false, reason:'ttl_exceeded' };
  }

  const canonical = [String(v), mid, mode, String(exp)].join('|');
  const mac       = hmacSha256_(keyBytes, utf8Bytes_(canonical));
  const calcB64u  = bytesToB64url_(mac);

  const sigBytes  = b64urlToBytes_(sig);
  const calcBytes = b64urlToBytes_(calcB64u);

  if (!timingSafeEq_(sigBytes, calcBytes)) {
    log_(logs, 'qr_verify_fail', 'bad_signature');
    return { ok:false, reason:'bad_signature' };
  }

  log_(logs, 'qr_verify_ok', 'mid='+mid+(mode?(' mode='+mode):''));
  return { ok:true, memberId: mid, mode: (mode==='COLLECT'||mode==='REDEEM') ? mode : '' , exp };
}



// ===============================
// QR Prefill Staging (new)
// ===============================
const PREFILL_TTL_SEC = 5 * 60; // 5 minutes

/**
 * Verify signed QR (payloadB64u) or deep-link token (token) and stage
 * a normalized prefill for redemption after staff signs in.
 */
function stagePrefill(payloadB64u, token) {
  let res = null;
  if (payloadB64u) {
    res = resolveSignedQrPayload(String(payloadB64u || ''));
  } else if (token) {
    res = resolveDeepLinkToken(String(token || ''));
  } else {
    return { ok: false, reason: 'missing_input' };
  }

  if (!res || !res.ok) {
    return { ok: false, reason: (res && res.reason) || 'invalid' };
  }

  const clean = {
    memberId: String(res.memberId || ''),
    mode: String(res.mode || ''), // '' | COLLECT | REDEEM
    exp: Number(res.exp || 0)
  };

  const stageId = 'pf_' + Utilities.getUuid().replace(/-/g, '');
  const cache = CacheService.getScriptCache();
  cache.put(stageId, JSON.stringify(clean), PREFILL_TTL_SEC);

  return { ok: true, stageId, memberId: clean.memberId, mode: clean.mode };
}

/**
 * One-time redemption of a staged prefill after sign-in.
 */
function redeemPrefill(stageId) {
  stageId = String(stageId || '').trim();
  if (!stageId) return { ok: false, reason: 'missing_stage' };

  const cache = CacheService.getScriptCache();
  const raw = cache.get(stageId);
  if (!raw) return { ok: false, reason: 'not_found_or_expired' };
  cache.remove(stageId);

  try {
    const obj = JSON.parse(raw);
    return { ok: true, memberId: String(obj.memberId || ''), mode: String(obj.mode || '') };
  } catch (_) {
    return { ok: false, reason: 'corrupt' };
  }
}

function _qrTtlCapOk_(exp, now, ttlSec) {
  const maxFuture = (Math.max(30, Number(ttlSec) || 180) * 2) * 1000 + 30 * 1000;
  return (exp - now) <= maxFuture;
}

// ── Logs (remote) ──
function log_(logsSheet, type, message) {
  logsSheet.appendRow([new Date().toISOString(), String(type||''), String(message||'')]);
}

/**
 * Fire and forget sink for Merchant client errors.
 * Accepts any shape and records username and merchantId when available.
 */
function merchantReportError(payload){
  try{
    const v = validateMSession((payload && payload.sid) || '');
    const who = v && v.ok ? (v.username + '@' + v.merchantId + ' (' + v.role + ')') : 'anon';
    const info = {
      who,
      ua: String((payload && payload.ua) || ''),
      atMs: Date.now(),
      location: String((payload && payload.location) || ''),
      message: String((payload && payload.message) || ''),
      stack: String((payload && payload.stack) || ''),
      version: String((payload && payload.version) || ''),
      extra: payload && payload.extra ? payload.extra : {}
    };
    const { logs } = getUserDb_();
    log_(logs, 'mclient_error', JSON.stringify(info));
    return { ok:true };
  }catch(e){
    // best effort. do not throw to client
    return { ok:false };
  }
}


function testResolveQr() {
  const p = 'PASTE-YOUR-p-PAYLOAD-HERE';
  const r = resolveSignedQrPayload(p);
  Logger.log(JSON.stringify(r));
}

function testCouponDryRun(memberId, mode, points, code){
  const sid = ''; // paste a live manager/staff SID here if you want to test scope
  const v = { username:'test', merchantId:'TEST', role:'manager' }; // or use getMerchantScopeOrThrow_(sid)
  const { ss, coupons, cpnUses } = getUserDb_();
  const res = validateAndApplyCoupon_(coupons, cpnUses, v, String(memberId||''), String(mode||''), Number(points||0), String(code||''));
  Logger.log(JSON.stringify(res));
}

function checkQrKeyPresence() {
  const { config } = getUserDb_(); // same path resolveSignedQrPayload uses
  const val = getConfigValueRemote_(config, 'QR_HMAC_KEY_B64');
  Logger.log('QR_HMAC_KEY_B64 raw = "%s"', val);
}


function validateQrKeyBase64() {
  const { config } = getUserDb_();
  const keyB64 = getConfigValueRemote_(config, 'QR_HMAC_KEY_B64');
  if (!keyB64) { Logger.log('Missing QR_HMAC_KEY_B64'); return; }

  try {
    const bytes = Utilities.base64Decode(keyB64); // standard Base64 only
    Logger.log('Decoded key length (bytes) = %s', bytes.length);
    // Common choices: 32 bytes (256-bit), 48, or 64 — anything >16 is fine.
    const roundTrip = Utilities.base64Encode(bytes);
    Logger.log('Round-trip equal? %s', roundTrip === keyB64 ? 'YES' : 'NO (whitespace/padding differences?)');
  } catch (e) {
    Logger.log('Base64 decode failed: %s', e);
  }
}

function generateNewQrKey() {
  // 32 random bytes via digest over UUID+time (good enough for a shared secret in Apps Script)
  const seed = Utilities.getUuid() + '|' + Date.now() + '|' + Math.random();
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed); // 32 bytes
  const keyB64 = Utilities.base64Encode(raw); // STANDARD Base64

  Logger.log('Put this into User DB → Config:');
  Logger.log('Key   = QR_HMAC_KEY_B64');
  Logger.log('Value = %s', keyB64);
}

function whoIsUserDb() {
  const id = PropertiesService.getScriptProperties().getProperty('USER_DB_ID');
  const ss = SpreadsheetApp.openById(id);
  Logger.log('USER_DB_ID = %s  |  Name = %s', id, ss.getName());
}



