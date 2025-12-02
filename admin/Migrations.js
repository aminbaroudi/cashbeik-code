function adminMigrateCampaignMultiplierColumn() {
  const uid = PropertiesService.getScriptProperties().getProperty('USER_DB_ID');
  if (!uid) throw new Error('USER_DB_ID not configured.');
  const ss = SpreadsheetApp.openById(uid);
  const sh = ss.getSheetByName('Campaigns');
  if (!sh) throw new Error('Campaigns sheet not found.');

  const HEADER = [
    'CampaignId','MerchantId','Title','Multiplier','Start','End',
    'MinSpend','MaxRed','MaxPerCustomer','Budget','BillingModel',
    'CostPerRedemption','Active','ImageFileId','ImagePublicUrl','CreatedAt','UpdatedAt'
  ];
  sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);

  const last = sh.getLastRow();
  if (last < 2) return { fixed: 0 };
  const cols = Math.max(sh.getLastColumn(), HEADER.length);
  const vals = sh.getRange(2,1,last-1,cols).getValues();

  let fixed = 0;
  for (let i = 0; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[3]).toLowerCase() === 'multiplier') {
      r[3] = r[4]; // D <- numeric multiplier from E
      // left-shift E..Q by one
      for (let j = 4; j < HEADER.length; j++) r[j] = (r[j+1] !== undefined) ? r[j+1] : '';
      fixed++;
    }
    vals[i] = r.slice(0, HEADER.length);
  }
  sh.getRange(2,1,vals.length,HEADER.length).setValues(vals);
  while (sh.getMaxColumns() > HEADER.length) sh.deleteColumn(sh.getMaxColumns());
  return { fixed };
}
