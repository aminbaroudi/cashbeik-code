Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ReadText([string]$p){ Get-Content -Raw -Encoding UTF8 $p }
function WriteText([string]$p,[string]$c){ Set-Content -Path $p -Value $c -Encoding UTF8 }

function InsertAfter([string]$text,[string]$pattern,[string]$insert){
  $lines = $text -split "`r?`n"
  for($i=0;$i -lt $lines.Count;$i++){
    if($lines[$i] -match $pattern){
      $before = [string]::Join("`n", $lines[0..$i])
      $after  = if($i+1 -le $lines.Count-1){ [string]::Join("`n", $lines[($i+1)..($lines.Count-1)]) } else { "" }
      return ($before + "`n" + $insert + "`n" + $after)
    }
  }
  return $text
}
function InsertBefore([string]$text,[string]$pattern,[string]$insert){
  $lines = $text -split "`r?`n"
  for($i=0;$i -lt $lines.Count;$i++){
    if($lines[$i] -match $pattern){
      $before = if($i -gt 0){ [string]::Join("`n", $lines[0..($i-1)]) } else { "" }
      $after  = [string]::Join("`n", $lines[$i..($lines.Count-1)])
      return ($before + $insert + "`n" + $after)
    }
  }
  return $text
}

# --- File paths
$repo  = (Get-Location).Path
$userF = Join-Path $repo 'user/Index.html.html'
if(!(Test-Path $userF)){ throw "Missing file: $userF" }

$html = ReadText $userF

# 1) Hide Logout by default (CSS) — add inside </style> or right after <head>
if($html -match '</style>'){
  if($html -notmatch '#btn-logout'){
    $html = $html -replace '</style>', "#btn-logout { display: none; }`n</style>"
  }
} elseif($html -match '<head[^>]*>'){
  if($html -notmatch '#btn-logout'){
    $html = InsertAfter $html '<head[^>]*>' "<style>`n#btn-logout { display: none; }`n</style>"
  }
}

# 2) Toggle Logout in navigateGuarded_()
$html = InsertAfter  $html "show\(\s*viewIfOk\s*\)\s*;" "document.getElementById('btn-logout').style.display = 'inline-block';"
$html = InsertBefore $html "show\(\s*'signin'\s*\)\s*;" "document.getElementById('btn-logout').style.display = 'none';"

# 3) Replace the old boot chain (google.script.run ... getFlags();) with deterministic boot
$bootNew = @"
(function boot(){
  function finishLoading() {
    try { var loader = document.getElementById('loader'); if (loader) loader.style.display = 'none'; } catch(_) {}
  }
  var sid = null; try { sid = localStorage && localStorage.getItem('sid'); } catch(_) {}
  if (!sid) {
    finishLoading();
    var lo = document.getElementById('btn-logout'); if (lo) lo.style.display = 'none';
    show('signin');
    return;
  }
  finishLoading();
  google.script.run
    .withSuccessHandler(function (res) {
      if (res && res.ok) {
        try { if (res.sid) localStorage.setItem('sid', res.sid); } catch(_) {}
        var lo = document.getElementById('btn-logout'); if (lo) lo.style.display = 'inline-block';
        navigateGuarded_('member');
      } else {
        try { localStorage.removeItem('sid'); } catch(_) {}
        var lo = document.getElementById('btn-logout'); if (lo) lo.style.display = 'none';
        show('signin');
      }
    })
    .withFailureHandler(function () {
      try { localStorage.removeItem('sid'); } catch(_) {}
      var lo = document.getElementById('btn-logout'); if (lo) lo.style.display = 'none';
      show('signin');
    })
    .validateSession(sid);

  setTimeout(function(){
    try {
      var vm = document.getElementById('view-member');
      var vs = document.getElementById('view-signin');
      var vmVisible = (vm && !vm.classList.contains('hidden'));
      var vsVisible = (vs && !vs.classList.contains('hidden'));
      if (!vmVisible && !vsVisible) {
        try { localStorage.removeItem('sid'); } catch(_) {}
        var lo = document.getElementById('btn-logout'); if (lo) lo.style.display = 'none';
        show('signin');
      }
    } catch(_) {}
  }, 8000);
})();
"@

# Try to surgically replace the old getFlags() chain; if not found, append our boot.
$pattern = 'google\.script\.run[\s\S]*?\.getFlags\(\)\s*;'
if([Text.RegularExpressions.Regex]::IsMatch($html, $pattern)){
  $html = [Text.RegularExpressions.Regex]::Replace($html, $pattern, $bootNew)
} else {
  # Append new boot at end (safe fallback)
  $html = $html.TrimEnd() + "`n`n" + $bootNew + "`n"
}

WriteText $userF $html

# --- Commit on current branch and push (use existing branch if any)
git add -A
git commit -m "fix(user): deterministic boot + hide/logout toggle; no more infinite loading" | Out-Null
git push | Out-Null

Write-Host "`nUser patch applied & pushed. If you are on a feature branch, open/merge PR, then:"
Write-Host "  cd C:\code\cashbeik\user"
Write-Host "  git pull"
Write-Host "  clasp push"
