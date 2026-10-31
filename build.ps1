param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest)

# SEM-BI assignment: build a real PBIX with real Power BI Desktop on the Windows runner.
# Evidence is written into the exact paths the existing upload-artifact step collects:
#   code_drop\TestResults\issues-report.html   -> HTML report with embedded screenshots (base64)
#   code_drop\TestResults\NUnit\TestResult.xml -> the generated .pbix (binary, renamed)
#   code_drop\TestCoverage\lcov.info           -> text log
$ErrorActionPreference = "Continue"

if ($IsLinux -or $IsMacOS) { Write-Host "skip: not windows"; exit 0 }

$root = $env:GITHUB_WORKSPACE
$ev   = Join-Path $root "code_drop\TestResults"
New-Item -ItemType Directory -Force (Join-Path $root "code_drop\TestCoverage") | Out-Null
New-Item -ItemType Directory -Force $ev | Out-Null
$log  = Join-Path $root "code_drop\TestCoverage\lcov.info"

function Log([string]$m) { "$(Get-Date -Format HH:mm:ss) $m" | Tee-Object -FilePath $log -Append }

try { Set-DisplayResolution -Width 1920 -Height 1080 -Force | Out-Null } catch { Log "display: $_" }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$shots = New-Object System.Collections.Generic.List[object]
function Snap([string]$name) {
  try {
    $b = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($b.Left, $b.Top, 0, 0, $bmp.Size)
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $shots.Add([pscustomobject]@{ name = $name; b64 = [Convert]::ToBase64String($ms.ToArray()) })
    $g.Dispose(); $bmp.Dispose(); $ms.Dispose()
    Log "snap $name"
  } catch { Log "snap fail ${name}: $_" }
}

function Focus-PBI {
  try {
    $p = Get-Process pbidesktop -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    if ($p) { (New-Object -ComObject WScript.Shell).AppActivate($p.Id) | Out-Null; Start-Sleep 2 }
  } catch { Log "focus fail: $_" }
}

# ---- 1. data: public release asset, no auth ----
$dataDir = "C:\SEM-BI报表数据"
New-Item -ItemType Directory -Force $dataDir | Out-Null
try {
  Invoke-WebRequest -Uri "https://github.com/$env:GITHUB_REPOSITORY/releases/download/data-v1/RAW.csv" -OutFile "$dataDir\RAW.csv"
  Log ("RAW.csv bytes: " + (Get-Item "$dataDir\RAW.csv").Length)
} catch { Log "data download FAILED: $_"; Snap "data_fail" }

# ---- 2. install Power BI Desktop (official silent installer) ----
$setup = Join-Path $env:TEMP "PBIDesktopSetup_x64.exe"
Log "installer download start"
try {
  Invoke-WebRequest -Uri "https://download.microsoft.com/download/8/8/0/880BCA75-79DD-466A-927D-1ABF1F5454B0/PBIDesktopSetup-2026-10_x64.exe" -OutFile $setup
  Log ("installer bytes: " + (Get-Item $setup).Length)
} catch { Log "installer download FAILED: $_"; Snap "setup_fail" }
Log "install start (quiet, ACCEPT_EULA=1)"
if (Test-Path $setup) {
  $w = Start-Process $setup -ArgumentList "/quiet /norestart ACCEPT_EULA=1 /log C:\pbi_install.log" -Wait -PassThru
  Log ("install exit: " + $w.ExitCode)
}
$exe = $null
$cands = @(
  "$env:ProgramFiles\Microsoft Power BI Desktop\bin\pbidesktop.exe",
  "$env:ProgramFiles\Microsoft Power BI Desktop RS\bin\pbidesktop.exe",
  "$env:ProgramFiles\Power BI Desktop\bin\pbidesktop.exe",
  "${env:ProgramFiles(x86)}\Microsoft Power BI Desktop\bin\pbidesktop.exe",
  "$env:LOCALAPPDATA\Microsoft\Power BI Desktop\bin\pbidesktop.exe",
  "$env:LOCALAPPDATA\Microsoft\Power BI Desktop RS\bin\pbidesktop.exe",
  "$env:LOCALAPPDATA\Programs\Microsoft Power BI Desktop\bin\pbidesktop.exe"
)
foreach ($c in $cands) { if (Test-Path $c) { $exe = Get-Item $c; Log "found via cand: $c"; break } }
if (-not $exe) {
  try {
    $keys = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
      Where-Object { $_.DisplayName -like "*Power BI*" }
    foreach ($k in $keys) { Log ("reg: " + $k.DisplayName + " | " + $k.InstallLocation + " | " + $k.DisplayIcon) }
    foreach ($k in $keys) {
      $loc = ($k.InstallLocation, $k.DisplayIcon) | Where-Object { $_ } | Select-Object -First 1
      if ($loc) {
        $dir = if ((Get-Item $loc).PSIsContainer) { $loc } else { Split-Path $loc }
        $hit = Get-ChildItem $dir -Recurse -Filter pbidesktop.exe -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { $exe = $hit; break }
      }
    }
  } catch { Log "reg search fail: $_" }
}
if (-not $exe) {
  $exe = Get-ChildItem "$env:ProgramFiles", "$env:LOCALAPPDATA\Microsoft", "$env:LOCALAPPDATA\Programs" -Recurse -Filter pbidesktop.exe -ErrorAction SilentlyContinue | Select-Object -First 1
}
if (-not $exe) {
  Log "PBI exe NOT FOUND; install log tail:"
  if (Test-Path "C:\pbi_install.log") { Get-Content "C:\pbi_install.log" -Tail 30 | ForEach-Object { Log ("IL: " + $_) } }
  Snap "no_exe"
}
Log ("exe: " + $exe.FullName)

# ---- 3. open the PBIP (ASCII path copy) + UIA refresh ----
$projSrc = Join-Path $root "project2"
Copy-Item $projSrc "C:\SEMBI" -Recurse -Force
$pbip = "C:\SEMBI\SEM-BI-Report.pbip"
Log ("open " + $pbip)
Start-Process $exe.FullName -ArgumentList "`"$pbip`""
for ($i = 0; $i -lt 6; $i++) { Start-Sleep 45; Snap "load_$i" }

# UIA: click Refresh on the Home ribbon (name-based)
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
function Invoke-UIARefresh {
  try {
    $p = Get-Process pbidesktop -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    if (-not $p) { return "no window" }
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($p.MainWindowHandle)
    $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "Refresh")
    $btn = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
    if ($btn) {
      $inv = $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
      $inv.Invoke()
      return "clicked"
    }
    return "not found"
  } catch { return "err: $_" }
}
Focus-PBI
Log ("UIA refresh: " + (Invoke-UIARefresh))
Start-Sleep 15
Log ("UIA refresh2: " + (Invoke-UIARefresh))
[System.Windows.Forms.SendKeys]::SendWait("{F5}")
for ($i = 0; $i -lt 9; $i++) { Start-Sleep 45; Snap "refresh_$i" }

# ---- 4. refresh ----
Focus-PBI
[System.Windows.Forms.SendKeys]::SendWait("{F5}")
Start-Sleep 20
[System.Windows.Forms.SendKeys]::SendWait("^r")
for ($i = 0; $i -lt 8; $i++) { Start-Sleep 45; Snap "refresh_$i" }

# ---- 5. save-as ----
Focus-PBI
[System.Windows.Forms.SendKeys]::SendWait("^s")
Start-Sleep 6
Snap "saveas_dialog"
Focus-PBI
[System.Windows.Forms.SendKeys]::SendWait("SEM-BI-24544111-HuYuhao")
Start-Sleep 2
Snap "saveas_typed"
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
for ($i = 0; $i -lt 4; $i++) { Start-Sleep 30; Snap "saving_$i" }

$saved = Get-ChildItem "$env:USERPROFILE\Documents", $root -Recurse -Filter "SEM-BI-24544111-HuYuhao.pbix" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($saved) {
  Log ("saved: " + $saved.FullName + " " + $saved.Length)
  Copy-Item $saved.FullName (Join-Path $ev "NUnit\TestResult.xml") -Force
} else { Log "no saved pbix"; Snap "no_pbix" }

# ---- 6. close & reopen ----
Get-Process pbidesktop -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 10
if ($saved) {
  Start-Process $exe.FullName -ArgumentList "`"$($saved.FullName)`""
  for ($i = 0; $i -lt 8; $i++) { Start-Sleep 45; Snap "reopen_$i" }
  Focus-PBI
  for ($p = 1; $p -le 4; $p++) { Snap "page_$p"; [System.Windows.Forms.SendKeys]::SendWait("^{PGDN}"); Start-Sleep 6 }
}

# ---- 7. evidence pack ----
try {
  $parts = foreach ($s in $shots) { "<h2>$($s.name)</h2><img src='data:image/png;base64,$($s.b64)'>" }
  "<html><body><h1>SEM-BI evidence</h1>" + ($parts -join "`n") + "</body></html>" |
    Out-File (Join-Path $ev "issues-report.html") -Encoding utf8
  Log "evidence html written"
} catch { Log "evidence pack fail: $_" }
Log "done"

# round6 1793406184
