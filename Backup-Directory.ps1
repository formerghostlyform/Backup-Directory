<#
.SYNOPSIS
    Backs up a source directory to a zip file and cleans up old backups.

.DESCRIPTION
    Creates a timestamped zip archive of the source directory in the destination
    directory, then applies a tiered retention policy:
      - Past 90 days  : keep all backups
      - Past 12 months: keep the last backup of each calendar month
      - Past 5 years  : keep the last backup of each calendar year
      - Older         : delete

.PARAMETER SourcePath
    Path to the directory to back up.

.PARAMETER DestinationPath
    Path to the directory where zip files are stored.

.PARAMETER BrowseDestination
    Opens a folder picker dialog to choose the destination directory.

.PARAMETER LogDirectory
    Directory where run log files are written. Defaults to the current directory.

.PARAMETER SendNotification
    Sends a Windows notification when the script completes (success or failure).

.PARAMETER Help
    Displays this help message and exits.

.PARAMETER WhatIf
    Shows what would be deleted during cleanup without actually deleting anything.

.EXAMPLE
    .\Backup-Directory.ps1 C:\Projects\MyApp D:\Backups
    .\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups
    .\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -BrowseDestination
    .\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups -LogDirectory C:\Logs
    .\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups -SendNotification
    .\Backup-Directory.ps1 -SourcePath C:\Projects\MyApp -DestinationPath D:\Backups -WhatIf
    .\Backup-Directory.ps1 -Help

.NOTES
    Retention policy applied during cleanup:

      0 – 90 days       Keep ALL backups
            90 days – ~15 mo  Keep the last backup of each calendar month
            ~15 mo – 5 years  Keep the last backup of each calendar year
      > 5 years         Delete

    Only zip files matching the pattern <SourceDirName>_yyyyMMdd_HHmmss.zip are
    considered; other files in the destination directory are left untouched.

    Reliability features:
      Atomic write    - Compression writes to a .tmp file; it is renamed to the
                        final name only after validation passes. A failed or
                        interrupted backup leaves no corrupt zip behind.
            Checksum manifest- A SHA-256 checksum manifest is written alongside each
                                                backup zip after successful validation.
            Concurrency lock- A per-source+destination named mutex prevents overlapping
                                                runs of the same backup job.
      VSS snapshot    - When run as Administrator, a Volume Shadow Copy is taken
                        before compression so files held open by other processes
                        are captured consistently. Gracefully skipped otherwise.
      Free-space check- The destination drive is checked for sufficient space
                        (source size x 1.1) before compression begins.
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Position = 0)]
    [string]$SourcePath = '',

    [Parameter(Position = 1)]
    [string]$DestinationPath = '',

    [switch]$BrowseDestination,

    [string]$LogDirectory = '.',

    [switch]$SendNotification,

    [switch]$Help
)

function Send-BackupNotification {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [bool]$IsSuccess
    )

    # Prefer BurntToast if available (true Action Center toast notifications).
    try {
        if (Get-Module -ListAvailable -Name BurntToast) {
            Import-Module BurntToast -ErrorAction Stop | Out-Null
            New-BurntToastNotification -Text $Title, $Message | Out-Null
            return
        }
    }
    catch {
        # Fall through to balloon-tip fallback.
    }

    # Fallback for systems without BurntToast.
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = if ($IsSuccess) { [System.Drawing.SystemIcons]::Information } else { [System.Drawing.SystemIcons]::Error }
        $notify.BalloonTipIcon = if ($IsSuccess) { [System.Windows.Forms.ToolTipIcon]::Info } else { [System.Windows.Forms.ToolTipIcon]::Error }
        $notify.BalloonTipTitle = $Title
        $notify.BalloonTipText = $Message
        $notify.Visible = $true
        $notify.ShowBalloonTip(10000)

        Start-Sleep -Milliseconds 4000
        $notify.Dispose()
    }
    catch {
        Write-Warning "Unable to send Windows notification: $($_.Exception.Message)"
    }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$transcriptStarted = $false
$logFilePath = $null
$runMutex = $null
$runLockAcquired = $false
$runSucceeded = $false
$notificationTitle = 'Backup Failed'
$notificationMessage = 'Backup did not complete.'
$originalWhatIfPreference = $WhatIfPreference

try {

# ---------------------------------------------------------------------------
# Show help when -Help is passed or no parameters are supplied
# ---------------------------------------------------------------------------
if ($Help -or (-not $SourcePath -and -not $DestinationPath)) {
    Get-Help -Full $MyInvocation.MyCommand.Path
    exit 0
}

# ---------------------------------------------------------------------------
# Validate inputs
# ---------------------------------------------------------------------------
if (-not $SourcePath) {
    Write-Error 'SourcePath is required. Run the script with -Help for usage.'
    exit 1
}

if ($BrowseDestination) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'Select backup destination folder'
        $dialog.ShowNewFolderButton = $true

        if ($DestinationPath -and (Test-Path -LiteralPath $DestinationPath -PathType Container)) {
            $dialog.SelectedPath = (Resolve-Path -Path $DestinationPath).Path
        }

        $dialogResult = $dialog.ShowDialog()
        if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK -or -not $dialog.SelectedPath) {
            Write-Error 'Destination folder selection was canceled.'
            exit 1
        }

        $DestinationPath = $dialog.SelectedPath
        Write-Host "Selected destination: $DestinationPath"
    }
    catch {
        Write-Error "Unable to open destination folder picker: $($_.Exception.Message)"
        exit 1
    }
    finally {
        if ($dialog) {
            $dialog.Dispose()
        }
    }
}

if (-not $DestinationPath) {
    Write-Error 'DestinationPath is required. Run the script with -Help for usage.'
    exit 1
}

$resolvedSource = Resolve-Path -Path $SourcePath -ErrorAction SilentlyContinue
if (-not $resolvedSource -or -not (Test-Path -LiteralPath $resolvedSource -PathType Container)) {
    Write-Error "Source directory not found: '$SourcePath'"
    exit 1
}
$SourcePath = $resolvedSource.Path
$sourceDirName = (Get-Item -LiteralPath $SourcePath).Name
$dateCode      = Get-Date -Format 'yyyyMMdd_HHmmss'

if (-not (Test-Path -LiteralPath $DestinationPath)) {
    Write-Host "Destination directory does not exist; creating: $DestinationPath"
    New-Item -ItemType Directory -Path $DestinationPath -Force -WhatIf:$false | Out-Null
}
$DestinationPath = (Resolve-Path -Path $DestinationPath).Path

# ---------------------------------------------------------------------------
# Concurrency protection
# ---------------------------------------------------------------------------
$mutexKey = ('{0}|{1}|{2}' -f
    $MyInvocation.MyCommand.Path.ToLowerInvariant(),
    $SourcePath.ToLowerInvariant(),
    $DestinationPath.ToLowerInvariant())

$sha256 = [System.Security.Cryptography.SHA256]::Create()
try {
    $hashBytes = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($mutexKey))
}
finally {
    $sha256.Dispose()
}

$mutexHash = [System.BitConverter]::ToString($hashBytes).Replace('-', '')
$mutexName = 'Local\BackupDirectory_' + $mutexHash
$createdNew = $false
$runMutex = [System.Threading.Mutex]::new($false, $mutexName, [ref]$createdNew)

try {
    $runLockAcquired = $runMutex.WaitOne(0)
}
catch [System.Threading.AbandonedMutexException] {
    # An abandoned mutex means prior owner crashed; lock is now acquired by this process.
    Write-Warning 'Recovered an abandoned backup lock from a previous failed run.'
    $runLockAcquired = $true
}

if (-not $runLockAcquired) {
    Write-Error ("Another backup run is already active for source '$SourcePath' " +
        "and destination '$DestinationPath'.")
    exit 1
}

Write-Host 'Execution lock acquired for this backup job.'

if (-not (Test-Path -LiteralPath $LogDirectory)) {
    Write-Host "Log directory does not exist; creating: $LogDirectory"
    New-Item -ItemType Directory -Path $LogDirectory -Force -WhatIf:$false | Out-Null
}
$LogDirectory = (Resolve-Path -Path $LogDirectory).Path

$logFileName = "${sourceDirName}_${dateCode}.log"
$logFilePath = Join-Path $LogDirectory $logFileName
if (-not $WhatIfPreference) {
    try {
        Start-Transcript -Path $logFilePath -Force | Out-Null
        $transcriptStarted = $true
        Write-Host "Logging to '$logFilePath'"
    }
    catch {
        Write-Warning "Unable to start transcript logging at '$logFilePath': $($_.Exception.Message)"
    }
}
else {
    Write-Host "WhatIf: transcript logging skipped."
}

# ---------------------------------------------------------------------------
# VSS shadow copy (requires elevation; falls back gracefully if unavailable)
# ---------------------------------------------------------------------------
$shadowObj      = $null
$compressSource = $SourcePath

$currentIdentity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdmin) {
    try {
        Write-Host 'Creating VSS shadow copy ...'
        $sourceVolume = (Split-Path -Qualifier $SourcePath) + '\'
        $shadowResult = Invoke-CimMethod -ClassName Win32_ShadowCopy -MethodName Create `
            -Arguments @{ Volume = $sourceVolume; Context = 'ClientAccessible' }

        if ($shadowResult.ReturnValue -eq 0) {
            $shadowObj = Get-CimInstance -ClassName Win32_ShadowCopy `
                -Filter "ID='$($shadowResult.ShadowID)'"
            $relPath        = (Split-Path -NoQualifier $SourcePath).TrimStart([char]92)
            $compressSource = Join-Path ($shadowObj.DeviceName + '\') $relPath
            Write-Host "Shadow copy created: $($shadowObj.DeviceName)"
        } else {
            Write-Warning "VSS creation returned code $($shadowResult.ReturnValue); compressing live files."
        }
    }
    catch {
        Write-Warning "VSS shadow copy failed: $_  Compressing live files."
        $shadowObj = $null
    }
} else {
    Write-Warning ('Not running as Administrator; VSS shadow copy skipped. ' +
        'Files held open by other processes may be missed or cause errors.')
}

# ---------------------------------------------------------------------------
# Build source inventory once and run free-space pre-check
# ---------------------------------------------------------------------------
Write-Host 'Indexing source files ...'
$sourceFiles = @(Get-ChildItem -LiteralPath $compressSource -Recurse -File | Sort-Object -Property FullName)
$sourceSize = ($sourceFiles | Measure-Object -Property Length -Sum).Sum
if ($null -eq $sourceSize) { $sourceSize = 0 }
$largestSourceFile = ($sourceFiles | Measure-Object -Property Length -Maximum).Maximum
if ($null -eq $largestSourceFile) { $largestSourceFile = 0 }

Write-Host 'Checking available disk space ...'
$destDrive = Split-Path -Qualifier $DestinationPath
$freeSpace = (Get-PSDrive -Name $destDrive.TrimEnd(':') -ErrorAction SilentlyContinue).Free
if (-not $freeSpace) {
    $disk      = Get-CimInstance -ClassName Win32_LogicalDisk `
                     -Filter "DeviceID='$destDrive'" -ErrorAction SilentlyContinue
    $freeSpace = if ($disk) { $disk.FreeSpace } else { $null }
}

if ($null -ne $freeSpace) {
    $requiredSpace = [long]($sourceSize * 1.1)
    if ($freeSpace -lt $requiredSpace) {
        Write-Error ("Insufficient disk space on destination. " +
            "Required: {0:N0} MB, Available: {1:N0} MB." -f
            [math]::Ceiling($requiredSpace / 1MB), [math]::Floor($freeSpace / 1MB))
        exit 1
    }
    Write-Host ("Space check passed. Required ~{0:N0} MB, available {1:N0} MB." -f
        [math]::Ceiling($requiredSpace / 1MB), [math]::Floor($freeSpace / 1MB))
} else {
    Write-Warning 'Could not determine available disk space; skipping space check.'
}

# ---------------------------------------------------------------------------
# Create backup (atomic: write to .tmp, rename to final name after validation)
# ---------------------------------------------------------------------------
# Keep WhatIf scoped to cleanup/delete operations; backup creation should still run.
$WhatIfPreference = $false

$zipFileName   = "${sourceDirName}_${dateCode}.zip"
$zipFilePath   = Join-Path $DestinationPath $zipFileName
# Compress-Archive on Windows PowerShell 5.1 requires a .zip extension.
$tempZipPath   = Join-Path $DestinationPath ("${sourceDirName}_${dateCode}.tmp.zip")
$manifestFileName = "${sourceDirName}_${dateCode}.manifest.sha256"
$manifestFilePath = Join-Path $DestinationPath $manifestFileName
$tempManifestPath = $manifestFilePath + '.tmp'

Write-Host "Backing up '$SourcePath' -> '$zipFilePath' ..."

$backupJob = $null
$oversizeThresholdBytes = [int64]2GB
$useDotNetZip = ($largestSourceFile -ge $oversizeThresholdBytes)
if ($useDotNetZip) {
    Write-Host ("Large source file detected ({0:N0} bytes). " -f $largestSourceFile) +
        'Using Zip64-capable .NET compression engine.'
}

try {
    # Run compression in a background job so we can show progress on the foreground thread.
    $backupJob = Start-Job -ScriptBlock {
        param($src, $dest, $preferDotNet)

        function Invoke-DotNetZip {
            param(
                [string]$InSource,
                [string]$InDest
            )

            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::CreateFromDirectory(
                $InSource,
                $InDest,
                [System.IO.Compression.CompressionLevel]::Optimal,
                $true
            )
        }

        if ($preferDotNet) {
            Invoke-DotNetZip -InSource $src -InDest $dest
            return
        }

        try {
            Compress-Archive -Path $src -DestinationPath $dest -CompressionLevel Optimal
        }
        catch {
            $message = $_.Exception.Message
            if ($message -match 'Stream was too long') {
                Invoke-DotNetZip -InSource $src -InDest $dest
            }
            else {
                throw
            }
        }
    } -ArgumentList $compressSource, $tempZipPath, $useDotNetZip

$spinnerFrames = @('|', '/', '-', '\')
$frame         = 0
$stopwatch     = [System.Diagnostics.Stopwatch]::StartNew()

while ($backupJob.State -eq 'Running') {
    $elapsed = $stopwatch.Elapsed
    $status  = 'Elapsed: {0:mm\:ss}' -f $elapsed
    Write-Progress -Activity "Creating backup  $($spinnerFrames[$frame % 4])" `
                   -Status $status -PercentComplete -1
    $frame++
    Start-Sleep -Milliseconds 150
}

$stopwatch.Stop()
Write-Progress -Activity 'Creating backup' -Completed

# Surface any errors from the background job
$backupJob | Receive-Job -Wait -AutoRemoveJob -ErrorAction Stop

Write-Host ("Compression complete: {0:mm\:ss} elapsed" -f $stopwatch.Elapsed)

# -------------------------------------------------------------------------
# Validate backup
# -------------------------------------------------------------------------
Write-Host 'Validating backup ...'

# 1. Temp zip must exist and have a non-zero size
if (-not (Test-Path -LiteralPath $tempZipPath)) {
    throw "Validation failed: temporary zip file not found at '$tempZipPath'"
}
$zipSize = (Get-Item -LiteralPath $tempZipPath).Length
if ($zipSize -eq 0) {
    throw 'Validation failed: zip file is empty.'
}

# 2. Zip must open without errors; 3. Compare entry sizes to source files
Add-Type -AssemblyName System.IO.Compression.FileSystem

$archive = [System.IO.Compression.ZipFile]::OpenRead($tempZipPath)
try {
    $entryTable = @{}
    foreach ($entry in $archive.Entries) {
        if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) { continue }
        $entryTable[$entry.FullName.Replace('\', '/')] = $entry.Length
    }

    $mismatches   = [System.Collections.Generic.List[string]]::new()
    $missingFiles = [System.Collections.Generic.List[string]]::new()

    foreach ($file in $sourceFiles) {
        $relative = $file.FullName.Substring($compressSource.Length).TrimStart([char]92, [char]'/')
        $entryKey = "$sourceDirName/$($relative.Replace('\', '/'))"

        if (-not $entryTable.ContainsKey($entryKey)) {
            $missingFiles.Add($relative)
        }
        elseif ($entryTable[$entryKey] -ne $file.Length) {
            $mismatches.Add("$relative  (source: $($file.Length) bytes, zip: $($entryTable[$entryKey]) bytes)")
        }
    }

    $valErrors = $missingFiles.Count + $mismatches.Count
    if ($valErrors -gt 0) {
        if ($missingFiles.Count -gt 0) {
            Write-Warning "Validation: $($missingFiles.Count) file(s) missing from zip:"
            $missingFiles | ForEach-Object { Write-Warning "  Missing : $_" }
        }
        if ($mismatches.Count -gt 0) {
            Write-Warning "Validation: $($mismatches.Count) file(s) with size mismatch:"
            $mismatches | ForEach-Object { Write-Warning "  Mismatch: $_" }
        }
        throw "Validation failed: $valErrors issue(s) found. The zip may be incomplete."
    }

    $entryCount = $entryTable.Count
}
finally {
    $archive.Dispose()
}

# Atomic rename: only promote to final name after successful validation
Move-Item -LiteralPath $tempZipPath -Destination $zipFilePath -WhatIf:$false
Write-Host ("Backup complete: $zipFileName  ({0} file(s), {1:N0} bytes)" -f $entryCount, $zipSize)

# Create checksum manifest alongside the zip file
Write-Host "Creating checksum manifest: $manifestFileName"
$zipHash = (Get-FileHash -LiteralPath $zipFilePath -Algorithm SHA256).Hash
$manifestLines = [System.Collections.Generic.List[string]]::new()
$manifestLines.Add("# Backup checksum manifest")
$manifestLines.Add("ArchiveFile=$zipFileName")
$manifestLines.Add("ArchiveSHA256=$zipHash")
$manifestLines.Add("ArchiveSizeBytes=$zipSize")
$manifestLines.Add("GeneratedUtc=$((Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))")
$manifestLines.Add("SourcePath=$SourcePath")
$manifestLines.Add('')
$manifestLines.Add('SHA256  SizeBytes  RelativePath')

$sourceFiles |
    ForEach-Object {
        $relativePath = $_.FullName.Substring($compressSource.Length).TrimStart([char]92, [char]'/')
        $fileHash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        $manifestLines.Add("$fileHash  $($_.Length)  $($relativePath.Replace('\', '/'))")
    }

Set-Content -LiteralPath $tempManifestPath -Value $manifestLines -Encoding utf8 -WhatIf:$false
Move-Item -LiteralPath $tempManifestPath -Destination $manifestFilePath -Force -WhatIf:$false
Write-Host "Manifest complete: $manifestFileName"
}
catch {
    # Remove the partial temp file so it cannot be mistaken for a valid backup
    if (Test-Path -LiteralPath $tempZipPath) {
        Write-Warning "Removing incomplete temporary file: $(Split-Path $tempZipPath -Leaf)"
        Remove-Item -LiteralPath $tempZipPath -Force -ErrorAction SilentlyContinue -WhatIf:$false
    }
    if (Test-Path -LiteralPath $tempManifestPath) {
        Write-Warning "Removing incomplete temporary manifest: $(Split-Path $tempManifestPath -Leaf)"
        Remove-Item -LiteralPath $tempManifestPath -Force -ErrorAction SilentlyContinue -WhatIf:$false
    }
    throw
}
finally {
    if ($null -ne $backupJob) {
        try {
            Stop-Job -Job $backupJob -ErrorAction SilentlyContinue -WhatIf:$false
            Remove-Job -Job $backupJob -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }
        catch {
            Write-Warning "Failed to clean up background backup job: $_"
        }
    }

    # Always release the VSS shadow copy regardless of success or failure
    if ($null -ne $shadowObj) {
        try {
            Write-Host 'Removing VSS shadow copy ...'
            $shadowObj | Remove-CimInstance
        }
        catch {
            Write-Warning "Failed to remove VSS shadow copy '$($shadowObj.ID)': $_"
        }
    }
}

$WhatIfPreference = $originalWhatIfPreference

# ---------------------------------------------------------------------------
# Retention policy helpers
# ---------------------------------------------------------------------------
$now           = Get-Date
$cutoffRecent  = $now.AddDays(-90)          # keep ALL backups newer than this
$cutoffMonthly = $cutoffRecent.AddMonths(-12) # keep one-per-MONTH between here and cutoffRecent
$cutoffYearly  = $now.AddYears(-5)           # keep one-per-YEAR between here and cutoffMonthly
                                              # delete anything older than cutoffYearly

# Match files produced by this script for the same source directory name.
# Expected pattern: <DirName>_yyyyMMdd_HHmmss.zip
$escapedSourceDirName = [regex]::Escape($sourceDirName)
$dateRegex   = "^${escapedSourceDirName}_(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})\.zip$"

# Build a list of backup objects with parsed dates
$allBackups = Get-ChildItem -LiteralPath $DestinationPath -Filter "*.zip" |
    Where-Object { $_.Name -match $dateRegex } |
    ForEach-Object {
        $m = [regex]::Match($_.Name, $dateRegex)
        $fileDate = [datetime]::new(
            [int]$m.Groups[1].Value,  # year
            [int]$m.Groups[2].Value,  # month
            [int]$m.Groups[3].Value,  # day
            [int]$m.Groups[4].Value,  # hour
            [int]$m.Groups[5].Value,  # minute
            [int]$m.Groups[6].Value   # second
        )
        [PSCustomObject]@{
            File = $_
            Date = $fileDate
        }
    } |
    Sort-Object -Property Date

# ---------------------------------------------------------------------------
# Determine which files to keep
# ---------------------------------------------------------------------------
$keepPaths = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

# Tier 1 – last 90 days: keep everything
$allBackups |
    Where-Object { $_.Date -ge $cutoffRecent } |
    ForEach-Object { $keepPaths.Add($_.File.FullName) | Out-Null }

# Tier 2 – 90 days to 12 months before the 90-day mark: keep last per calendar month
$allBackups |
    Where-Object { $_.Date -ge $cutoffMonthly -and $_.Date -lt $cutoffRecent } |
    Group-Object { $_.Date.ToString('yyyy-MM') } |
    ForEach-Object {
        $last = $_.Group | Sort-Object Date | Select-Object -Last 1
        $keepPaths.Add($last.File.FullName) | Out-Null
    }

# Tier 3 – beyond 12+3 months back to 5 years: keep last per calendar year
$allBackups |
    Where-Object { $_.Date -ge $cutoffYearly -and $_.Date -lt $cutoffMonthly } |
    Group-Object { $_.Date.Year } |
    ForEach-Object {
        $last = $_.Group | Sort-Object Date | Select-Object -Last 1
        $keepPaths.Add($last.File.FullName) | Out-Null
    }

# Tier 4 – older than 5 years: delete (not added to keepPaths)

# ---------------------------------------------------------------------------
# Cleanup
# ---------------------------------------------------------------------------
$deletedCount = 0
foreach ($backup in $allBackups) {
    if (-not $keepPaths.Contains($backup.File.FullName)) {
        if ($PSCmdlet.ShouldProcess($backup.File.Name, 'Remove old backup')) {
            try {
                Remove-Item -LiteralPath $backup.File.FullName -Force
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($backup.File.Name)
                $matchingManifests = Get-ChildItem -LiteralPath $DestinationPath -Filter "${baseName}.manifest.sha256" -ErrorAction SilentlyContinue
                foreach ($manifest in $matchingManifests) {
                    try {
                        Remove-Item -LiteralPath $manifest.FullName -Force
                    }
                    catch {
                        Write-Warning "Failed to remove manifest '$($manifest.Name)': $_"
                    }
                }
            }
            catch {
                Write-Warning "Failed to remove old backup '$($backup.File.Name)': $_"
                continue
            }
        }
        $deletedCount++
    }
}

if ($WhatIfPreference) {
    Write-Host "WhatIf: $deletedCount backup(s) would be removed."
} else {
    Write-Host "Cleanup complete. Removed $deletedCount old backup(s)."
}

# ---------------------------------------------------------------------------
# Log retention (same policy as backup files)
# ---------------------------------------------------------------------------
$logDateRegex = "^${escapedSourceDirName}_(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})\.log$"

$allLogs = Get-ChildItem -LiteralPath $LogDirectory -Filter "*.log" |
    Where-Object { $_.Name -match $logDateRegex } |
    ForEach-Object {
        $m = [regex]::Match($_.Name, $logDateRegex)
        $fileDate = [datetime]::new(
            [int]$m.Groups[1].Value,
            [int]$m.Groups[2].Value,
            [int]$m.Groups[3].Value,
            [int]$m.Groups[4].Value,
            [int]$m.Groups[5].Value,
            [int]$m.Groups[6].Value
        )
        [PSCustomObject]@{
            File = $_
            Date = $fileDate
        }
    } |
    Sort-Object -Property Date

$keepLogPaths = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

$allLogs |
    Where-Object { $_.Date -ge $cutoffRecent } |
    ForEach-Object { $keepLogPaths.Add($_.File.FullName) | Out-Null }

$allLogs |
    Where-Object { $_.Date -ge $cutoffMonthly -and $_.Date -lt $cutoffRecent } |
    Group-Object { $_.Date.ToString('yyyy-MM') } |
    ForEach-Object {
        $last = $_.Group | Sort-Object Date | Select-Object -Last 1
        $keepLogPaths.Add($last.File.FullName) | Out-Null
    }

$allLogs |
    Where-Object { $_.Date -ge $cutoffYearly -and $_.Date -lt $cutoffMonthly } |
    Group-Object { $_.Date.Year } |
    ForEach-Object {
        $last = $_.Group | Sort-Object Date | Select-Object -Last 1
        $keepLogPaths.Add($last.File.FullName) | Out-Null
    }

$deletedLogCount = 0
foreach ($logFile in $allLogs) {
    if (-not $keepLogPaths.Contains($logFile.File.FullName)) {
        if ($PSCmdlet.ShouldProcess($logFile.File.Name, 'Remove old log file')) {
            try {
                Remove-Item -LiteralPath $logFile.File.FullName -Force
            }
            catch {
                Write-Warning "Failed to remove old log file '$($logFile.File.Name)': $_"
                continue
            }
        }
        $deletedLogCount++
    }
}

if ($WhatIfPreference) {
    Write-Host "WhatIf: $deletedLogCount log file(s) would be removed."
} else {
    Write-Host "Log cleanup complete. Removed $deletedLogCount old log file(s)."
}

$runSucceeded = $true
$notificationTitle = 'Backup Completed Successfully'
$notificationMessage =
    "Archive: $zipFileName`nValidated files: $entryCount`nRemoved backups: $deletedCount`nRemoved logs: $deletedLogCount"

}
catch {
    $runSucceeded = $false
    $notificationTitle = 'Backup Failed'
    $notificationMessage = $_.Exception.Message
    Write-Error "Backup failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($null -ne $runMutex) {
        try {
            if ($runLockAcquired) {
                $runMutex.ReleaseMutex() | Out-Null
                $runLockAcquired = $false
            }
            $runMutex.Dispose()
        }
        catch {
            Write-Warning "Failed to release execution lock: $_"
        }
    }

    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }

    if ($SendNotification) {
        Send-BackupNotification -Title $notificationTitle -Message $notificationMessage -IsSuccess $runSucceeded
    }
}
