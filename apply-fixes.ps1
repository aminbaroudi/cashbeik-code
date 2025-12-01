Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Read-Text($p){ Get-Content -Raw -Encoding UTF8 $p }
function Write-Text($p,$c){ Set-Content -Path $p -Value $c -Encoding UTF8 }

function Insert-AfterLineMatch($text,$pattern,$insert){
  $lines = $text -split "`r?`n"
  for($i=0;$i -lt $lines.Count;$i++){
    if($lines[$i] -match $pattern){
      $before = ""
      if($i -ge 0){ $before = [string]::Join("`n", $lines[0..$i]) }
      $after = ""
      if($i+1 -le $lines.Count-1){ $after = [string]::Join("`n", $lines[($i+1)..($lines.Count-1)]) }
      return ($before + "`n" + $insert + "`n" + $after)
    }
  }
  return $text
}

function Insert-BeforeLineMatch($text,$pattern,$insert){
  $lines = $text -split "`r?`n"
  for($i=0;$i -lt $lines.Count;$i++){
    if($lines[$i] -match $pattern){
      $before = ""
      if($i -gt 0){ $before = [string]::Join("`n", $lines[0..($i-1)]) }
      $after = [string]::Join("`n", $lines[$i..($lines.Count-1)])
      return ($before + $insert + "`n" + $after)
    }
  }
  return $text
}

$repo = (Get-Location).Path
$adminFile = Join-Path $repo 'admin/Code.js'
$userFile  = Join-Path $repo 'user/Index.html.html'

if(!(Test-Path $adminFile)){ throw "Missing file: $adminFile" }
if(!(Test-Path $userFile)){  throw "Missing file: $userFile"  }

# -------- ADMIN PATCHES --------
$admin = Read-Text $adminFile

$healer = @"
/**
 * Inserts missing "MaxPerCustomer" after "MaxRedemptions" and standardizes headers.
 * WHY: Prevents positional misalignment when writing with appendRow.
 */
function ensureCampaignsShape_() {
  const { ss } = getUserDb_();
  const sh = ss.getSheetByName('Campaigns');
  if (!sh) return;

  var CANON = [
    'CampaignId','MerchantId','Title','Type','Multiplier',
    'StartIso','EndIso','MinSpend','MaxRedemptions','MaxPerCustomer',
    'BudgetCap','BillingModel','CostPerRedemption','Active',
    'CreatedAt','UpdatedAt','ImageUrl'
  ];

  var lastCol = Math.max(sh.getLastColumn(), CANON.length);
  var hdr = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);

  var idxMaxRed = hdr.indexOf('MaxRedemptions');
  var hasMaxPer = hdr.indexOf('MaxPerCustomer') !== -1;
  if (idxMaxRed !== -1 && !hasMaxPer) {
    sh.insertColumnAfter(idxMaxRed + 1); // 1-based
  }

  sh.getRange(1,1,1,CANON.length).setValues([CANON]);

  var curLast = sh.getLastColumn();
  if (curLast > CANON.length) {
    sh.deleteColumns(CANON.length + 1, curLast - CANON.length);
  }

  try {
    var prots = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i=0;i<prots.length;i++){
      var p = prots[i];
      var r = p.getRange();
      if (r.getRow()===1 && r.getNumRows()===1) p.remove();
    }
    var pr = sh.getRange(1,1,1,sh.getLastColumn()).protect()
      .setDescription('Header row protected by Admin App');
    pr.setWarningOnly(true);
  } catch(e) {}
}
"@

# 1) Add healer if missing
if($admin -notmatch 'function\s+ensureCampaignsShape_\s*\('){
  $admin = $admin.TrimEnd() + "`n`n" + $healer + "`n"
}

# 2) Heal right after ensuring Campaigns in getUserDb_()
$admin = Insert-AfterLineMatch $admin "ensureSheet_\(\s*ss\s*,\s*'Campaigns'" "ensureCampaignsShape_();"

# 3) Heal before campaigns.appendRow(...)
$admin = Insert-BeforeLineMatch $admin "campaigns\.appendRow\s*\(" "ensureCampaignsShape_();"

# 4) Heal at start of adminListCampaigns(...) after ensureAdminActive_(sid);
$admin = Insert-AfterLineMatch $admin "ensureAdminActive_\(\s*sid\s*\)\s*;" "ensureCampaignsShape_();"

Write-Text $adminFile $admin

# -------- USER PATCHES --------
$user = Read-Text $userFile

# 1) Hide Logout by default (CSS)
if($user -match '</style>'){
  if($user -notmatch '#btn-logout'){
    $user = $user -replace '</style>', "#btn-logout { display: none; }`n</style>"
  }
} elseif($user -match '<head[^>]*>'){
  if($user -notmatch '#btn-logout'){
    $user = Insert-AfterLineMatch $user '<head[^>]*>' "<style>`n#btn-logout { display: none; }`n</style>"
  }
}

# 2) Toggle Logout in navigateGuarded_()
$user = Insert-AfterLineMatch $user "show\(\s*viewIfOk\s*\)\s*;" "document.getElementById('btn-logout').style.display = 'inline-block';"
$user = Insert-BeforeLineMatch $user "show\(\s*'signin'\s*\)\s*;" "document.getElementById('btn-logout').style.display = 'none';"

# 3) Boot fail-safe after getFlags()
$bootFallback = @"
setTimeout(function(){
  try {
    var logoutBtn = document.getElementById('btn-logout');
    var loggedIn = (logoutBtn && logoutBtn.style.display !== 'none');
    var memberView = document.getElementById('view-member');
    var memberVisible = (memberView && !memberView.classList.contains('hidden'));
    if (!loggedIn && !memberVisible) { clearSid(); navigateGuarded_('signin'); }
  } catch(_) {}
}, 8000);
"@
$user = Insert-AfterLineMatch $user "\.getFlags\(\)\s*;" $bootFallback

Write-Text $userFile $user

# -------- git branch, commit, push --------
try { git rev-parse --abbrev-ref HEAD | Out-Null } catch { }
$cur = ""
try { $cur = (git rev-parse --abbrev-ref HEAD).Trim() } catch { $cur = "" }
if ($cur -ne 'fix/campaign-header-and-user-boot') {
  git checkout -b fix/campaign-header-and-user-boot
}

git add -A
git commit -m "fix(admin): heal Campaigns header; fix(user): boot timeout + gated logout"
git push -u origin fix/campaign-header-and-user-boot

Write-Host "`nAll set. Open this to create the Pull Request:"
Write-Host "https://github.com/aminbaroudi/cashbeik-code/compare/fix/campaign-header-and-user-boot?expand=1"
