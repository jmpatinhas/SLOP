# diagnostics_uipath_env.ps1
param(
  [string]$OutPath = "$env:TEMP\uipath_env_diag_$((Get-Date).ToString('yyyyMMdd_HHmmss')).json"
)

function Get-ActiveSessionInfo {
  $session = $null
  try {
    $q = (qwinsta.exe) 2>$null
    # crude parse but usually enough
    $active = $q | Where-Object { $_ -match "Active" } | Select-Object -First 1
    $session = $active
  } catch {}
  return $session
}

Add-Type -AssemblyName System.Windows.Forms

$monitors = @()
foreach ($m in [System.Windows.Forms.Screen]::AllScreens) {
  $monitors += [pscustomobject]@{
    DeviceName = $m.DeviceName
    Primary    = $m.Primary
    Bounds     = @{ X=$m.Bounds.X; Y=$m.Bounds.Y; Width=$m.Bounds.Width; Height=$m.Bounds.Height }
    WorkingArea= @{ X=$m.WorkingArea.X; Y=$m.WorkingArea.Y; Width=$m.WorkingArea.Width; Height=$m.WorkingArea.Height }
  }
}

$virtual = [System.Windows.Forms.SystemInformation]::VirtualScreen
$virtualObj = @{
  X=$virtual.X; Y=$virtual.Y; Width=$virtual.Width; Height=$virtual.Height
}

# DPI (quick signal: registry + one API-like approximation)
$dpiReg = $null
try {
  $dpiReg = (Get-ItemProperty "HKCU:\Control Panel\Desktop\WindowMetrics" -ErrorAction Stop).AppliedDPI
} catch {}

# RDP / session-ish hints
$sessionName = $env:SESSIONNAME
$activeSession = Get-ActiveSessionInfo

# Lock state: presence of LogonUI is a decent hint (not perfect)
$lockedHint = $false
try { $lockedHint = @(Get-Process LogonUI -ErrorAction SilentlyContinue).Count -gt 0 } catch {}

$info = [pscustomobject]@{
  Timestamp = (Get-Date).ToString("o")
  Machine   = $env:COMPUTERNAME
  User      = "$env:USERDOMAIN\$env:USERNAME"
  Session   = @{
    EnvSessionName   = $sessionName
    QwinstaActiveLine= $activeSession
    LockedHint       = $lockedHint
  }
  Display = @{
    Monitors      = $monitors
    VirtualScreen = $virtualObj
    AppliedDPI_HKCU = $dpiReg
  }
  UiPath = @{
    RobotUserProfile = $env:USERPROFILE
    Temp             = $env:TEMP
  }
  OS = @{
    Version = [System.Environment]::OSVersion.VersionString
  }
  Processes = @{
    ExplorerRunning = @(Get-Process explorer -ErrorAction SilentlyContinue).Count -gt 0
  }
}

$info | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutPath -Encoding UTF8
Write-Output $OutPath
