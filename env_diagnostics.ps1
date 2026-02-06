# diagnostics_uipath_env.ps1
# Straight script (no custom functions). Config at the top.

# =========================
# CONFIG
# =========================
$OutDir   = "$env:TEMP"
$FileTag  = "uipath_env_diag"
$TimeTag  = (Get-Date).ToString("yyyyMMdd_HHmmss")
$OutPath  = Join-Path $OutDir "$FileTag`_$TimeTag.json"

# If you want a fixed path instead, uncomment:
# $OutPath = "C:\Temp\uipath_env_diag.json"

# =========================
# DISPLAY / SCREEN INFO
# =========================
Add-Type -AssemblyName System.Windows.Forms

$monitors = @()
foreach ($m in [System.Windows.Forms.Screen]::AllScreens) {
  $monitors += [pscustomobject]@{
    DeviceName  = $m.DeviceName
    Primary     = $m.Primary
    Bounds      = [pscustomobject]@{ X=$m.Bounds.X; Y=$m.Bounds.Y; Width=$m.Bounds.Width; Height=$m.Bounds.Height }
    WorkingArea = [pscustomobject]@{ X=$m.WorkingArea.X; Y=$m.WorkingArea.Y; Width=$m.WorkingArea.Width; Height=$m.WorkingArea.Height }
  }
}

$vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
$virtualScreen = [pscustomobject]@{
  X      = $vs.X
  Y      = $vs.Y
  Width  = $vs.Width
  Height = $vs.Height
}

# =========================
# DPI / SCALING (HKCU)
# =========================
$appliedDPI = $null
try {
  $appliedDPI = (Get-ItemProperty "HKCU:\Control Panel\Desktop\WindowMetrics" -ErrorAction Stop).AppliedDPI
} catch {
  $appliedDPI = $null
}

# Optional: Windows 10/11 per-user scale (sometimes present)
$logPixels = $null
try {
  $logPixels = (Get-ItemProperty "HKCU:\Control Panel\Desktop" -ErrorAction Stop).LogPixels
} catch {
  $logPixels = $null
}

# =========================
# SESSION / INTERACTIVE HINTS
# =========================
$sessionName = $env:SESSIONNAME

$qwinstaRaw = $null
try {
  $qwinstaRaw = (qwinsta.exe) 2>$null
} catch {
  $qwinstaRaw = $null
}

$lockedHint = $false
try {
  $lockedHint = @(Get-Process LogonUI -ErrorAction SilentlyContinue).Count -gt 0
} catch {
  $lockedHint = $false
}

$explorerRunning = $false
try {
  $explorerRunning = @(Get-Process explorer -ErrorAction SilentlyContinue).Count -gt 0
} catch {
  $explorerRunning = $false
}

# =========================
# BASIC SYSTEM CONTEXT
# =========================
$info = [pscustomobject]@{
  Timestamp = (Get-Date).ToString("o")
  Output    = [pscustomobject]@{
    OutPath = $OutPath
  }

  Machine = [pscustomobject]@{
    Name = $env:COMPUTERNAME
  }

  User = [pscustomobject]@{
    Domain   = $env:USERDOMAIN
    Username = $env:USERNAME
    Full     = "$($env:USERDOMAIN)\$($env:USERNAME)"
    Profile  = $env:USERPROFILE
  }

  Session = [pscustomobject]@{
    EnvSessionName = $sessionName
    QwinstaRaw     = $qwinstaRaw
    LockedHint     = $lockedHint
    ExplorerRunning= $explorerRunning
  }

  Display = [pscustomobject]@{
    Monitors      = $monitors
    VirtualScreen = $virtualScreen
    AppliedDPI_HKCU_WindowMetrics = $appliedDPI
    LogPixels_HKCU_Desktop        = $logPixels
  }

  OS = [pscustomobject]@{
    VersionString = [System.Environment]::OSVersion.VersionString
  }

  UiPathHints = [pscustomobject]@{
    Temp = $env:TEMP
  }
}

# =========================
# OUTPUT JSON
# =========================
$info | ConvertTo-Json -Depth 8 | Out-File -FilePath $OutPath -Encoding UTF8

# Write the output path so UiPath can capture it if desired
Write-Output $OutPath
