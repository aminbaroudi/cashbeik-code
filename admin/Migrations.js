/**
 * Normalize ALL rows in User DB -> Campaigns to the canonical header order.
 * Fixes the common "shift" where column D had the string 'multiplier'
 * and the numeric multiplier sat in column E, pushing dates rightward.
 * Also trims any trailing columns (e.g., duplicate PerMemberBonusCap).
 *
 * Header we enforce:
 * A: CampaignId
 * B: MerchantId
 * C: Title
 * D: Multiplier
 * E: Start
 * F: End
 * G: MinSpend
 * H: MaxRed
 * I: MaxPerCustomer
 * J: Budget
 * K: BillingModel
 * L: CostPerRedemption
 * M: Active
 * N: ImageFileId
 * O: ImagePublicUrl
 * P: CreatedAt
 * Q: UpdatedAt
 */
function AdminMigrateCampaignRows() {
  const HEADER = [
    'CampaignId','MerchantId','Title','Multiplier','Start','End',
    'MinSpend','MaxRed','MaxPerCustomer','Budget','BillingModel',
    'CostPerRedemption','Active','ImageFileId','ImagePublicUrl',
    'CreatedAt','UpdatedAt'
  ];

  const uid = PropertiesService.getScriptProperties().getProperty('USER_DB_ID');
  if (!uid) throw new Error('USER_DB_ID not configured.');
  const ss = SpreadsheetApp.openById(uid);
  const sh = ss.getSheetByName('Campaigns') || ss.insertSheet('Campaigns');

  // Ensure enough columns & write header
  if (sh.getMaxColumns() < HEADER.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), HEADER.length - sh.getMaxColumns());
  }
  sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok:true, rows:0 };

  const lastCol = sh.getLastColumn();
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const norm = [];
  for (var i = 0; i < values.length; i++) {
    const r = values[i];

    // Read defensively by index; many sheets had an extra token at D/E.
    // Common bad pattern seen:
    // [A]id, [B]merchant, [C]title,
    // [D]='multiplier', [E]=<num>, [F]=Start, [G]=End, [H]=Min, [I]=MaxRed, [J]=MaxPerCust, ...
    //
    // We normalize to:
    // D=Number multiplier, E=Start, F=End, G=MinSpend, H=MaxRed, I=MaxPerCustomer, etc.

    var id   = getCell(r, 0);
    var mid  = getCell(r, 1);
    var tit  = getCell(r, 2);

    var dRaw = getCell(r, 3);
    var eRaw = getCell(r, 4);

    var multiplier;
    var start, end, minSpend, maxRed, maxPer, budget, billing, cpr, active, imgId, imgUrl, createdAt, updatedAt;

    if (String(dRaw).toLowerCase() === 'multiplier') {
      // Shift-left fix
      multiplier = toNumber(eRaw);
      start      = getCell(r, 5);
      end        = getCell(r, 6);
      minSpend   = toNumber(getCell(r, 7));
      maxRed     = toNumber(getCell(r, 8));
      maxPer     = toNumber(getCell(r, 9));
      budget     = toNumber(getCell(r,10));
      billing    = getCell(r,11);
      cpr        = toNumber(getCell(r,12));
      active     = toBool( getCell(r,13) );
      imgId      = getCell(r,14);
      imgUrl     = getCell(r,15);
      createdAt  = getCell(r,16);
      updatedAt  = getCell(r,17);
    } else {
      // Already in the right place (or at least D is numeric)
      multiplier = toNumber(dRaw);
      start      = getCell(r, 4);
      end        = getCell(r, 5);
      minSpend   = toNumber(getCell(r, 6));
      maxRed     = toNumber(getCell(r, 7));
      maxPer     = toNumber(getCell(r, 8));
      budget     = toNumber(getCell(r, 9));
      billing    = getCell(r,10);
      cpr        = toNumber(getCell(r,11));
      active     = toBool( getCell(r,12) );
      imgId      = getCell(r,13);
      imgUrl     = getCell(r,14);
      createdAt  = getCell(r,15);
      updatedAt  = getCell(r,16);
    }

    norm.push([
      id, mid, tit, multiplier, start, end,
      minSpend, maxRed, maxPer, budget, billing,
      cpr, active, imgId, imgUrl, createdAt, updatedAt
    ]);
  }

  // Overwrite normalized rows to exactly HEADER length.
  sh.getRange(2, 1, norm.length, HEADER.length).setValues(norm);

  // If there are extra columns beyond Q, delete them.
  while (sh.getMaxColumns() > HEADER.length) {
    sh.deleteColumn(sh.getMaxColumns());
  }

  return { ok:true, rows: norm.length };

  // ---- helpers ----
  function getCell(arr, i){ return (i < arr.length) ? arr[i] : ''; }
  function toNumber(v){ var n = Number(v); return isNaN(n) ? 0 : n; }
  function toBool(v){ var s = String(v).toLowerCase(); return s === 'true' || s === 'yes' || s === '1'; }
}
