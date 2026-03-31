<#
.SYNOPSIS
    Downloads and runs the Windows Update Assistant to upgrade Windows 10/11 to the latest feature update.
    Tracks state across runs so it can verify if the upgrade succeeded on subsequent executions.
.DESCRIPTION
    Designed for RMM scheduled deployment (e.g. daily). Uses a state file to track upgrade progress:
      - IDLE:      No upgrade in progress. Checks if OS is already current. If not, kicks off upgrade.
      - PENDING:   Upgrade was launched. Waiting for completion/reboot. Checks setup logs for errors.
      - FAILED:    Upgrade failed. Reports why. Resets after a cooldown period.
    Writes status to Ninja custom field for dashboard visibility.
.NOTES
    Runs as SYSTEM via RMM. Schedule daily.
    Exit Codes:
        0 = Success (upgrade launched, verified, or already current)
        1 = Unsupported OS
        2 = Insufficient disk space
        3 = Download failed
        4 = Upgrade assistant failed to start
        5 = Upgrade failed (verified on follow-up)
#>

# ── Configuration ──────────────────────────────────────────────────────────────
$WorkingDir           = "$env:SystemDrive\ProgramData\WindowsUpgrade"
$StateFile            = "$WorkingDir\upgrade_state.json"
$InstallerPath        = "$WorkingDir\WindowsUpgradeAssistant.exe"
$DownloadTimeoutSec   = 1800
$LogRetentionDays     = 30
$CleanMgrTimeoutSec   = 300 # 5 min max for cleanmgr before we kill it

# Ninja custom field name — change this to match your environment
$NinjaCustomField     = 'windowsUpgradeStatus'

# Disk space thresholds (in GB) based on real-world testing and community data:
#   Win11 → Win11 feature update: MS says 6-11GB, community says 10-20GB. We use 15GB.
#   Win10 → Win11 major upgrade:  Creates Windows.old (~20-25GB) + download + temp. We use 30GB.
$MinFreeSpaceGB_FeatureUpdate = 15
$MinFreeSpaceGB_MajorUpgrade  = 30

# Known latest builds — update these as new versions ship
# Used to skip machines that are already current
$LatestBuilds = @{
    'Windows 11' = '26100'   # 24H2/25H2
    'Windows 10' = '19045'   # 22H2 (final)
}

# ── Helper Functions ───────────────────────────────────────────────────────────

function Write-Status {
    param([string]$Message)
    # Only writes to Ninja custom field — caller handles Write-Output to avoid duplicates
    try {
        Ninja-Property-Set $NinjaCustomField $Message 2>$null
    } catch {
        # Not running in Ninja or field doesn't exist
    }
}

function Get-State {
    if (Test-Path $StateFile) {
        try {
            return Get-Content $StateFile -Raw | ConvertFrom-Json
        } catch {
            Write-Output "WARNING: Corrupt state file, resetting."
            return $null
        }
    }
    return $null
}

function Set-State {
    param(
        [string]$Status,
        [string]$Detail,
        [string]$OsBuildBefore = '',
        [string]$OsVersionBefore = ''
    )
    $state = @{
        Status          = $Status
        Detail          = $Detail
        Timestamp       = (Get-Date).ToString('o')
        OsBuildBefore   = $OsBuildBefore
        OsVersionBefore = $OsVersionBefore
    }
    $state | ConvertTo-Json | Set-Content $StateFile -Force
}

function Get-OsBuild {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    return @{
        Caption = $os.Caption
        Build   = $os.BuildNumber
        Version = $os.Version
    }
}

function Get-SetupErrors {
    param(
        [DateTime]$Since = (Get-Date).AddHours(-48)  # Only look at errors from the last 48 hours
    )

    $pantherLog = "$env:SystemDrive\`$WINDOWS.~BT\Sources\Panther\setuperr.log"
    $pantherLog2 = "$env:SystemDrive\Windows\Panther\setuperr.log"

    $errors = @()
    foreach ($log in @($pantherLog, $pantherLog2)) {
        if (Test-Path $log) {
            # Only consider this log if it was modified recently
            $lastWrite = (Get-Item $log -Force -ErrorAction SilentlyContinue).LastWriteTime
            if ($lastWrite -lt $Since) { continue }

            $content = Get-Content $log -Tail 20 -ErrorAction SilentlyContinue
            if ($content) {
                # Filter lines to only include recent timestamps
                $recentLines = $content | Where-Object {
                    if ($_ -match '^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}') {
                        try {
                            $lineDate = [DateTime]::ParseExact($Matches[0], 'yyyy-MM-dd HH:mm:ss', $null)
                            return $lineDate -ge $Since
                        } catch { return $false }
                    }
                    return $false  # Skip lines without timestamps
                }
                if ($recentLines) {
                    $errors += "--- $log ---"
                    $errors += $recentLines
                }
            }
        }
    }

    $setupact = "$env:SystemDrive\`$WINDOWS.~BT\Sources\Panther\setupact.log"
    if ((Test-Path $setupact) -and (Get-Item $setupact -Force -ErrorAction SilentlyContinue).LastWriteTime -ge $Since) {
        $diskFullHits = Select-String -Path $setupact -Pattern '0x80070070|disk full|out of disk space' -ErrorAction SilentlyContinue
        if ($diskFullHits) {
            $errors += "--- Disk space errors found in setupact.log ---"
            $errors += ($diskFullHits | Select-Object -First 5 | ForEach-Object { $_.Line })
        }
    }

    if ($errors.Count -gt 0) {
        return ($errors -join "`n")
    }
    return $null
}

function Get-UpgradeFailureReason {
    param([string]$ErrorLog)

    # Known Windows Setup error codes → actionable tech-friendly messages
    $knownCodes = @{
        '0xC1900200' = 'COMPATIBILITY BLOCK — Machine does not meet minimum requirements. Check: TPM 2.0 enabled, Secure Boot on, 4GB+ RAM, DiagTrack service running.'
        '0xC1900201' = 'COMPATIBILITY BLOCK — System does not meet minimum requirements for this version.'
        '0xC1900202' = 'COMPATIBILITY BLOCK — Machine is not supported for this upgrade.'
        '0x80070070' = 'DISK SPACE — Ran out of disk space during upgrade.'
        '0x800705BB' = 'SETUP ABORTED — Upgrade was cancelled or interrupted during install phase.'
        '0xC1900204' = 'MIGRATION CHOICE — Upgrade option selected is not available.'
        '0x80240020' = 'REBOOT REQUIRED — Pending reboot is blocking the upgrade.'
        '0xC1900208' = 'INCOMPATIBLE APP — An installed application is blocking the upgrade. Check setup logs for which app.'
        '0xC1900209' = 'INCOMPATIBLE DRIVER — A driver is incompatible with the target OS version.'
        '0xC190020E' = 'NOT ENOUGH SPACE — Insufficient disk space. Need more free space on the system drive.'
        '0x80070490' = 'NOT FOUND — Required component or element not found.'
        '0x80073712' = 'CORRUPT FILES — Component store is corrupted. Run DISM /Online /Cleanup-Image /RestoreHealth.'
        '0x800F0922' = 'CONNECTION FAILED — Cannot connect to update servers, or System Reserved partition is too small.'
        '0x800704D3' = 'USER CANCELLED — Operation was cancelled (may indicate service interruption).'
    }

    # Check if DiagTrack is mentioned
    $diagTrack = $ErrorLog -match 'DiagTrack.*not available'

    # Find all matching error codes in the log
    $foundReasons = @()
    foreach ($code in $knownCodes.Keys) {
        if ($ErrorLog -match [regex]::Escape($code)) {
            $foundReasons += $knownCodes[$code]
        }
    }

    if ($diagTrack) {
        $foundReasons += 'DIAGTRACK SERVICE DISABLED — Enable "Connected User Experiences and Telemetry" service (DiagTrack) and retry.'
    }

    if ($foundReasons.Count -gt 0) {
        # Return the most specific/actionable reason (prioritize compatibility and disk over generic)
        return ($foundReasons | Select-Object -Unique) -join ' | '
    }

    return 'UNKNOWN — Check setup logs manually for error details.'
}

function Test-HardwareCompatibility {
    # Pre-flight check against Microsoft's official Windows 11 requirements
    # Returns $null if compatible, or a string describing all failures
    $failures = @()

    # 1. Total disk size — must be 64GB or larger (NOT free space — total capacity)
    $systemDrive = $env:SystemDrive.TrimEnd('\')
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction SilentlyContinue
    if ($disk) {
        $totalGB = [math]::Round($disk.Size / 1GB, 1)
        if ($totalGB -lt 64) {
            $failures += "DISK TOO SMALL — System drive is ${totalGB} GB, must be 64 GB or larger."
        }
    }

    # 2. RAM — must be 4GB or more
    $ram = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    if ($ram) {
        $ramGB = [math]::Round($ram.TotalPhysicalMemory / 1GB, 1)
        if ($ramGB -lt 3.5) {  # Use 3.5 to account for hardware-reserved memory on 4GB systems
            $failures += "INSUFFICIENT RAM — ${ramGB} GB detected, must be 4 GB or more."
        }
    }

    # 3. UEFI / Secure Boot
    try {
        $secureBoot = Confirm-SecureBootUEFI -ErrorAction Stop
        if (-not $secureBoot) {
            $failures += "SECURE BOOT DISABLED — Secure Boot is supported but not enabled. Enable it in UEFI/BIOS."
        }
    } catch {
        # If the cmdlet fails, we're either not on UEFI or Secure Boot isn't supported
        $firmware = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State' -Name 'UEFISecureBootEnabled' -ErrorAction SilentlyContinue
        if (-not $firmware -or $firmware.UEFISecureBootEnabled -ne 1) {
            $failures += "SECURE BOOT — Cannot confirm Secure Boot is enabled. Check UEFI/BIOS settings."
        }
    }

    # 4. TPM 2.0
    try {
        $tpm = Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop
        if ($tpm) {
            $tpmVersion = $tpm.SpecVersion
            if ($tpmVersion -and -not ($tpmVersion -match '^2\.')) {
                $failures += "TPM VERSION — TPM version is $tpmVersion, must be 2.0. Check BIOS/VM settings."
            }
        } else {
            $failures += "NO TPM — No TPM detected. Enable TPM 2.0 in BIOS or add virtual TPM to VM."
        }
    } catch {
        $failures += "NO TPM — Cannot query TPM. Enable TPM 2.0 in BIOS or add virtual TPM to VM."
    }

    # 5. Processor — must be 64-bit with 2+ cores
    $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cpu) {
        if ($cpu.NumberOfCores -lt 2) {
            $failures += "CPU CORES — Only $($cpu.NumberOfCores) core(s) detected, must be 2 or more."
        }
        if ($cpu.AddressWidth -ne 64) {
            $failures += "CPU ARCHITECTURE — Not a 64-bit processor. Windows 11 requires 64-bit."
        }
    }

    # 6. DiagTrack service — not a hard MS requirement but causes 0xC1900200 when disabled
    #    This one we can fix automatically
    $diagTrack = Get-Service -Name 'DiagTrack' -ErrorAction SilentlyContinue
    if ($diagTrack -and $diagTrack.StartType -eq 'Disabled') {
        Write-Host "  DiagTrack service is disabled — enabling and starting it..."
        try {
            Set-Service -Name 'DiagTrack' -StartupType Automatic -ErrorAction Stop
            Start-Service -Name 'DiagTrack' -ErrorAction Stop
            Write-Host "  DiagTrack service enabled and started."
        } catch {
            $failures += "DIAGTRACK — Failed to enable DiagTrack service: $_. Enable it manually."
        }
    }

    # 7. Generation 1 VM check (Hyper-V) — Gen 1 VMs cannot upgrade to Win11
    $msvm = Get-CimInstance -Namespace 'root\virtualization\v2' -ClassName Msvm_VirtualSystemSettingData -ErrorAction SilentlyContinue 2>$null
    # Not applicable for most cases — skip silently if not Hyper-V

    if ($failures.Count -gt 0) {
        return $failures -join ' | '
    }
    return $null
}

function Test-PendingReboot {
    $pending = $false
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pending = $true }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pending = $true }
    $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if ($pfro) { $pending = $true }
    return $pending
}

function Test-UpgradeAssistantRunning {
    $upgradeProcesses = @(
        'Windows10UpgraderApp',
        'WindowsUpdateAssistant',
        'SetupHost',
        'setupprep',
        'DismHost'
    )
    $running = Get-Process -Name $upgradeProcesses -ErrorAction SilentlyContinue
    if ($running) {
        return ($running | Select-Object -ExpandProperty Name -Unique) -join ', '
    }
    return $null
}

function Invoke-UpgradeCleanup {
    $reclaimedMB = 0

    function Remove-UpgradeDir {
        param([string]$Path, [string]$Label)
        if (Test-Path $Path) {
            # Quick size estimate — just top-level to avoid slow recursive scan on big dirs
            Write-Output "  Found $Label ($Path). Removing..."
            try {
                Start-Process -FilePath 'cmd.exe' -ArgumentList "/c takeown /f `"$Path`" /r /d y >nul 2>&1 && icacls `"$Path`" /grant administrators:F /t /q >nul 2>&1 && rd /s /q `"$Path`"" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                if (-not (Test-Path $Path)) {
                    Write-Output "  Removed $Label."
                } else {
                    Write-Output "  WARNING: $Label partially removed (some files may be locked)."
                }
            } catch {
                Write-Output "  WARNING: Could not remove ${Label}: $_"
            }
        }
    }

    Write-Output "Scanning for upgrade artifacts..."

    # 1. Windows.old — previous OS installation backup (15-25GB)
    Remove-UpgradeDir "$env:SystemDrive\Windows.old" 'Windows.old'

    # 2. $WINDOWS.~BT — upgrade staging/download directory (4-8GB)
    Remove-UpgradeDir "$env:SystemDrive\`$WINDOWS.~BT" '$WINDOWS.~BT'

    # 3. $WINDOWS.~WS — upgrade working directory (2-6GB)
    Remove-UpgradeDir "$env:SystemDrive\`$WINDOWS.~WS" '$WINDOWS.~WS'

    # 4. Windows10Upgrade folder — Update Assistant install directory
    Remove-UpgradeDir "$env:SystemDrive\Windows10Upgrade" 'Windows10Upgrade'

    # 5. ESD staging folder
    Remove-UpgradeDir "$env:SystemDrive\`$Windows.~Q" '$Windows.~Q'

    # 6. Windows\SoftwareDistribution\Download — cached update packages
    $swDistDl = "$env:SystemRoot\SoftwareDistribution\Download"
    if (Test-Path $swDistDl) {
        $sizeMB = [math]::Round((Get-ChildItem $swDistDl -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1MB, 0)
        if ($sizeMB -gt 100) {
            Write-Output "  Found SoftwareDistribution\Download cache — ${sizeMB} MB"
            try {
                Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
                Remove-Item -Path "$swDistDl\*" -Recurse -Force -ErrorAction SilentlyContinue
                Start-Service -Name wuauserv -ErrorAction SilentlyContinue
                Write-Output "  SoftwareDistribution cache cleared."
            } catch {
                Write-Output "  WARNING: Could not clear SoftwareDistribution cache: $_"
                Start-Service -Name wuauserv -ErrorAction SilentlyContinue
            }
        }
    }

    # 7. Windows\Temp — system temp files
    $winTemp = "$env:SystemRoot\Temp"
    if (Test-Path $winTemp) {
        $sizeMB = [math]::Round((Get-ChildItem $winTemp -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1MB, 0)
        if ($sizeMB -gt 100) {
            Write-Output "  Found Windows\Temp — ${sizeMB} MB"
            Remove-Item -Path "$winTemp\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Output "  Windows\Temp cleared."
        }
    }

    # 8. DISM component store cleanup
    Write-Output "  Running DISM component cleanup..."
    try {
        & dism /Online /Cleanup-Image /StartComponentCleanup 2>$null
        Write-Output "  DISM cleanup completed."
    } catch {
        Write-Output "  WARNING: DISM cleanup failed: $_"
    }

    # 9. Clean up our own installer
    if (Test-Path $InstallerPath) {
        Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
        Write-Output "  Removed cached Update Assistant executable."
    }

    # 10. Uninstall the Update Assistant if it registered itself
    $uninstallPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($regPath in $uninstallPaths) {
        $entry = Get-ItemProperty $regPath -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -match 'Windows.*Update Assistant|Windows.*Upgrade Assistant' } |
            Select-Object -First 1
        if ($entry -and $entry.UninstallString) {
            Write-Output "  Uninstalling Update Assistant via registry..."
            try {
                $uninstCmd = $entry.UninstallString -replace '/I', '/X'
                Start-Process -FilePath 'cmd.exe' -ArgumentList "/c $uninstCmd /quiet /norestart" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                Write-Output "  Update Assistant uninstalled."
            } catch {
                Write-Output "  WARNING: Could not uninstall Update Assistant: $_"
            }
        }
    }

    # Measure how much we actually freed
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive.TrimEnd('\'))'" -ErrorAction SilentlyContinue
    $nowFreeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    Write-Output "Cleanup complete. Current free space: ${nowFreeGB} GB"
}

# ── Bootstrap ──────────────────────────────────────────────────────────────────
[Net.ServicePointManager]::SecurityProtocol = [Enum]::ToObject([Net.SecurityProtocolType], 3072)

if (-not (Test-Path $WorkingDir)) {
    New-Item -Path $WorkingDir -ItemType Directory -Force | Out-Null
}

# Clean up old logs
Get-ChildItem -Path $WorkingDir -Filter 'Upgrade_*.log' -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
    Remove-Item -Force -ErrorAction SilentlyContinue

# Start logging
$LogFile = "$WorkingDir\Upgrade_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $LogFile -Append -Force
Write-Output "================================================================"
Write-Output "Windows Upgrade Script - $(Get-Date)"
Write-Output "================================================================"

# ── Get current OS info ────────────────────────────────────────────────────────
try {
    $currentOs = Get-OsBuild
    Write-Output "Current OS: $($currentOs.Caption) | Build: $($currentOs.Build) | Version: $($currentOs.Version)"
} catch {
    $msg = "ERROR: Failed to query OS. $_"
    Write-Output $msg
    Write-Status $msg
    Stop-Transcript
    exit 1
}

# ── State Machine ──────────────────────────────────────────────────────────────
$state = Get-State

# ─── PENDING: Check on a previously launched upgrade ──────────────────────────
if ($state -and $state.Status -eq 'PENDING') {
    $launchedAt = [DateTime]::Parse($state.Timestamp)
    $hoursSince = [math]::Round(((Get-Date) - $launchedAt).TotalHours, 1)
    Write-Output "Previous upgrade launched $hoursSince hours ago (build was $($state.OsBuildBefore))."

    # Did the build number change? That means the upgrade worked.
    if ($currentOs.Build -ne $state.OsBuildBefore -or $currentOs.Version -ne $state.OsVersionBefore) {
        $msg = "UPGRADE SUCCESSFUL | Was build $($state.OsBuildBefore) ($($state.OsVersionBefore)), now $($currentOs.Build) ($($currentOs.Version)) | Completed in ~${hoursSince}h"
        Write-Output $msg

        Write-Output "Running post-upgrade cleanup..."
        Invoke-UpgradeCleanup

        Write-Status $msg
        Set-State -Status 'IDLE' -Detail 'Upgrade verified successful'
        Stop-Transcript
        exit 0
    }

    # Is the upgrade assistant still actively running?
    $runningProcs = Test-UpgradeAssistantRunning
    if ($runningProcs) {
        $msg = "UPGRADE ACTIVELY RUNNING | Launched ${hoursSince}h ago | Processes: $runningProcs"
        Write-Output $msg
        Write-Status $msg
        Stop-Transcript
        exit 0
    }

    # Build didn't change — is a reboot pending or is the upgrade staged?
    $upgradeStaged = Test-Path "$env:SystemDrive\`$WINDOWS.~BT\Sources"
    if (Test-PendingReboot -or $upgradeStaged) {
        $msg = "UPGRADE COMPLETE — NEEDS REBOOT | Upgrade was successful, just waiting on a reboot to finalize."
        Write-Output $msg
        Write-Status $msg
        Stop-Transcript
        exit 0
    }

    # Check setup logs for errors (BEFORE cleanup so we can read them)
    $setupErrors = Get-SetupErrors
    if ($setupErrors) {
        Write-Output "Setup errors found:"
        Write-Output $setupErrors

        # Parse error codes into actionable message for the custom field
        $failureReason = Get-UpgradeFailureReason -ErrorLog $setupErrors

        Write-Output "Cleaning up failed upgrade artifacts..."
        Invoke-UpgradeCleanup

        $msg = "UPGRADE FAILED | $failureReason"
        Write-Output $msg
        Write-Status $msg
        Set-State -Status 'FAILED' -Detail $msg
        Stop-Transcript
        exit 5
    }

    # No build change, no reboot pending, no errors — might still be in progress
    if ($hoursSince -lt 6) {
        $msg = "UPGRADE IN PROGRESS | Launched ${hoursSince}h ago, no reboot yet, no errors. Waiting."
        Write-Output $msg
        Write-Status $msg
        Stop-Transcript
        exit 0
    }

    # Over 6 hours with no change — something probably went wrong silently
    if ($hoursSince -ge 6 -and $hoursSince -lt 48) {
        $msg = "UPGRADE STALLED | Launched ${hoursSince}h ago, build unchanged, no pending reboot. May need manual review."
        Write-Output $msg
        Write-Status $msg
        Stop-Transcript
        exit 0
    }

    # Over 48 hours — declare failure and allow retry after cooldown
    Write-Output "Cleaning up stalled upgrade artifacts..."
    Invoke-UpgradeCleanup
    $msg = "UPGRADE FAILED | Launched ${hoursSince}h ago, build never changed. Resetting for retry."
    Write-Output $msg
    Write-Status $msg
    Set-State -Status 'FAILED' -Detail $msg
    Stop-Transcript
    exit 5
}

# ─── FAILED: Previous attempt failed — just log it and retry ─────────────────
if ($state -and $state.Status -eq 'FAILED') {
    $failedAt = [DateTime]::Parse($state.Timestamp)
    $hoursSince = [math]::Round(((Get-Date) - $failedAt).TotalHours, 1)
    Write-Output "Previous attempt failed ${hoursSince}h ago. Retrying."
}

# ─── IDLE: Determine if upgrade is needed and launch it ──────────────────────

# Check if upgrade assistant is already running
$runningProcs = Test-UpgradeAssistantRunning
if ($runningProcs) {
    $msg = "UPGRADE ALREADY RUNNING | Processes: $runningProcs | Skipping launch."
    Write-Output $msg
    Write-Status $msg
    Stop-Transcript
    exit 0
}

# Don't launch a new upgrade if a reboot is already pending or an upgrade is staged
$upgradeStaged = Test-Path "$env:SystemDrive\`$WINDOWS.~BT\Sources"
if (Test-PendingReboot -or $upgradeStaged) {
    $msg = "REBOOT NEEDED | An upgrade or update is staged and waiting for a reboot. Reboot the machine first."
    Write-Output $msg
    Write-Status $msg
    Stop-Transcript
    exit 0
}

# Check if already on latest known build
$matchedOs = $LatestBuilds.Keys | Where-Object { $currentOs.Caption -match $_ } | Select-Object -First 1
if ($matchedOs -and [int]$currentOs.Build -ge [int]$LatestBuilds[$matchedOs]) {
    $msg = "ALREADY CURRENT | $($currentOs.Caption) build $($currentOs.Build) meets or exceeds target ($($LatestBuilds[$matchedOs]))"
    Write-Output $msg
    Write-Status $msg
    Stop-Transcript
    exit 0
}

# ── Hardware Compatibility Pre-Flight ──────────────────────────────────────────
Write-Output "Running hardware compatibility check..."
$compatFailures = Test-HardwareCompatibility
if ($compatFailures) {
    $msg = "NOT COMPATIBLE | $compatFailures"
    Write-Output $msg
    Write-Status $msg
    Stop-Transcript
    exit 1
}
Write-Output "Hardware compatibility check passed."

# Check for leftover artifacts from a previously failed upgrade
$staleArtifacts = @(
    "$env:SystemDrive\Windows.old",
    "$env:SystemDrive\`$WINDOWS.~BT",
    "$env:SystemDrive\`$WINDOWS.~WS",
    "$env:SystemDrive\`$Windows.~Q",
    "$env:SystemDrive\Windows10Upgrade"
)
$foundStale = $staleArtifacts | Where-Object { Test-Path $_ }
if ($foundStale) {
    Write-Output "Found stale upgrade artifacts from a previous failed/incomplete upgrade:"
    foreach ($path in $foundStale) { Write-Output "  $path" }
    Write-Output "Cleaning up before proceeding..."
    Invoke-UpgradeCleanup
}

# Detect OS and set thresholds
switch -Wildcard ($currentOs.Caption) {
    '*Windows 11*' {
        $url = 'https://go.microsoft.com/fwlink/?linkid=2171764'
        $MinFreeSpaceGB = $MinFreeSpaceGB_FeatureUpdate
        Write-Output "Target: Windows 11 Update Assistant (feature update - ${MinFreeSpaceGB} GB required)"
    }
    '*Windows 10*' {
        $url = 'https://go.microsoft.com/fwlink/?LinkID=799445'
        $MinFreeSpaceGB = $MinFreeSpaceGB_MajorUpgrade
        Write-Output "Target: Windows 10 Update Assistant (major upgrade - ${MinFreeSpaceGB} GB required)"
    }
    default {
        $msg = "ERROR: Unsupported OS - '$($currentOs.Caption)'"
        Write-Output $msg
        Write-Status $msg
        Stop-Transcript
        exit 1
    }
}

# ── Disk Space Check ───────────────────────────────────────────────────────────
$systemDrive = $env:SystemDrive.TrimEnd('\')
$disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction SilentlyContinue
$freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
Write-Output "System drive ($systemDrive) free space: ${freeGB} GB (minimum required: ${MinFreeSpaceGB} GB)"

if ($freeGB -lt $MinFreeSpaceGB) {
    Write-Output "Insufficient disk space. Running full cleanup..."

    # Full upgrade artifact cleanup
    Invoke-UpgradeCleanup

    # Also run cleanmgr with a timeout so it can't hang forever
    try {
        $volCaches = Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches' -ErrorAction SilentlyContinue
        foreach ($key in $volCaches) {
            Set-ItemProperty -Path $key.PSPath -Name 'StateFlags0001' -Value 2 -ErrorAction SilentlyContinue
        }
        $cleanMgr = Start-Process -FilePath 'cleanmgr.exe' -ArgumentList '/sagerun:1' -PassThru -NoNewWindow -ErrorAction SilentlyContinue
        if ($cleanMgr -and -not $cleanMgr.WaitForExit($CleanMgrTimeoutSec * 1000)) {
            Write-Output "WARNING: cleanmgr hung for $CleanMgrTimeoutSec seconds, killing it."
            $cleanMgr.Kill()
        }
        Write-Output "Windows Disk Cleanup completed."
    } catch {
        Write-Output "WARNING: Disk Cleanup failed: $_"
    }

    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction SilentlyContinue
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    Write-Output "Free space after all cleanup: ${freeGB} GB"

    if ($freeGB -lt $MinFreeSpaceGB) {
        $msg = "DISK SPACE INSUFFICIENT | Need ${MinFreeSpaceGB} GB, have ${freeGB} GB after full cleanup"
        Write-Output $msg
        Write-Status $msg
        Stop-Transcript
        exit 2
    }
    Write-Output "Cleanup freed enough space. Continuing."
}

# ── Download Update Assistant ──────────────────────────────────────────────────
if (Test-Path $InstallerPath) {
    Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
}

Write-Output "Downloading Update Assistant from: $url"
try {
    Start-BitsTransfer -Source $url -Destination $InstallerPath -Priority Normal -ErrorAction Stop
    Write-Output "BITS download completed."
} catch {
    Write-Output "BITS failed ($_). Falling back to Invoke-WebRequest..."
    try {
        Invoke-WebRequest -Uri $url -OutFile $InstallerPath -UseBasicParsing -TimeoutSec $DownloadTimeoutSec -ErrorAction Stop
        Write-Output "WebRequest download completed."
    } catch {
        $msg = "DOWNLOAD FAILED | Both BITS and WebRequest failed. $_"
        Write-Output $msg
        Write-Status $msg
        Stop-Transcript
        exit 3
    }
}

# ── Validate Download ──────────────────────────────────────────────────────────
if (-not (Test-Path $InstallerPath)) {
    $msg = "DOWNLOAD FAILED | File not found after download"
    Write-Output $msg
    Write-Status $msg
    Stop-Transcript
    exit 3
}

$fileSize = (Get-Item $InstallerPath).Length
$fileSizeMB = [math]::Round($fileSize / 1MB, 2)
Write-Output "Downloaded file size: ${fileSizeMB} MB"

if ($fileSize -lt 1MB) {
    $msg = "DOWNLOAD FAILED | File too small (${fileSizeMB} MB), likely corrupt"
    Write-Output $msg
    Write-Status $msg
    Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
    Stop-Transcript
    exit 3
}

$header = [System.IO.File]::ReadAllBytes($InstallerPath)[0..1]
if ($header[0] -ne 0x4D -or $header[1] -ne 0x5A) {
    $msg = "DOWNLOAD FAILED | Not a valid PE executable"
    Write-Output $msg
    Write-Status $msg
    Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
    Stop-Transcript
    exit 3
}
Write-Output "File validation passed."

# ── Launch Upgrade ─────────────────────────────────────────────────────────────
Write-Output "Starting Windows Update Assistant..."
$upgradeArgs = '/quietinstall /skipeula /auto upgrade /copylogs "{0}"' -f $WorkingDir

try {
    $process = Start-Process -FilePath $InstallerPath -ArgumentList $upgradeArgs -PassThru -ErrorAction Stop
    Write-Output "Update Assistant launched. PID: $($process.Id)"

    Start-Sleep -Seconds 15

    if ($process.HasExited -and $process.ExitCode -ne 0) {
        $msg = "LAUNCH FAILED | Update Assistant exited with code $($process.ExitCode)"
        Write-Output $msg
        Write-Status $msg
        Set-State -Status 'FAILED' -Detail $msg
        Stop-Transcript
        exit 4
    }

    # Record state for follow-up verification
    Set-State -Status 'PENDING' -Detail 'Upgrade assistant launched' -OsBuildBefore $currentOs.Build -OsVersionBefore $currentOs.Version

    $msg = "UPGRADE LAUNCHED | Build $($currentOs.Build) ($($currentOs.Version)) | Will verify on next run"
    Write-Output $msg
    Write-Status $msg

} catch {
    $msg = "LAUNCH FAILED | $_"
    Write-Output $msg
    Write-Status $msg
    Set-State -Status 'FAILED' -Detail $msg
    Stop-Transcript
    exit 4
}

Write-Output "================================================================"
Write-Output "Script completed at $(Get-Date)"
Write-Output "Log location: $LogFile"
Write-Output "================================================================"
Stop-Transcript
exit 0
