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

function Get-VssCreateReturnMessage {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Code
    )

    switch ($Code) {
        0 { 'Success' }
        1 { 'Access denied' }
        2 { 'Invalid argument' }
        3 { 'Specified volume not found' }
        4 { 'Specified volume not supported' }
        5 { 'Unsupported shadow copy context' }
        6 { 'Insufficient storage' }
        7 { 'Volume is in use' }
        8 { 'Maximum number of shadow copies reached' }
        9 { 'Another shadow copy operation is already in progress' }
        10 { 'Shadow copy provider vetoed the operation' }
        11 { 'Shadow copy provider is not registered' }
        12 { 'Provider failure (unknown error)' }
        default { 'Unknown error code' }
    }
}

function Try-ParseBackupTimestampFromName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Parameter(Mandatory = $true)]
        [string]$Pattern,

        [Parameter(Mandatory = $true)]
        [string]$Kind
    )

    $m = [regex]::Match($FileName, $Pattern)
    if (-not $m.Success) {
        return $null
    }

    $stamp = '{0}{1}{2}_{3}{4}{5}' -f
        $m.Groups[1].Value,
        $m.Groups[2].Value,
        $m.Groups[3].Value,
        $m.Groups[4].Value,
        $m.Groups[5].Value,
        $m.Groups[6].Value

    $parsed = [datetime]::MinValue
    $ok = [datetime]::TryParseExact(
        $stamp,
        'yyyyMMdd_HHmmss',
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$parsed
    )

    if (-not $ok) {
        Write-Warning "Skipping $Kind file with invalid timestamp format: '$FileName'"
        return $null
    }

    return $parsed
}

function Remove-StaleStagingDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prefix,

        [Parameter(Mandatory = $true)]
        [datetime]$OlderThan
    )

    $tempRoot = [System.IO.Path]::GetTempPath()
    $removed = 0
    $failed = 0

    $candidates = Get-ChildItem -LiteralPath $tempRoot -Directory -Filter ($Prefix + '*') -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $OlderThan }

    foreach ($dir in $candidates) {
        try {
            Remove-Item -LiteralPath $dir.FullName -Recurse -Force -ErrorAction Stop -WhatIf:$false
            $removed++
        }
        catch {
            $failed++
            Write-Warning "Failed to remove stale staging directory '$($dir.FullName)': $($_.Exception.Message)"
        }
    }

    if ($removed -gt 0 -or $failed -gt 0) {
        Write-Host ("Stale staging cleanup: removed {0}, failed {1}." -f $removed, $failed)
    }
}

function Get-CanonicalDirectoryPath {
    param([string]$Path)

    # Resolve existing directories through Windows handles. This handles short
    # (8.3) names and directory junctions that lexical path comparison misses.
    if (-not ('BackupDirectory.NativePaths' -as [type])) {
        $nativeDefinition = @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace BackupDirectory {
    public static class NativePaths {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(string name, uint access,
            uint share, IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFinalPathNameByHandleW(SafeFileHandle handle,
            StringBuilder path, uint size, uint flags);
        public static string Resolve(string path) {
            using (SafeFileHandle handle = CreateFileW(path, 0, 7, IntPtr.Zero,
                3, 0x02000000, IntPtr.Zero)) {
                if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
                StringBuilder result = new StringBuilder(32768);
                uint length = GetFinalPathNameByHandleW(handle, result, (uint)result.Capacity, 0);
                if (length == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
                if (length >= result.Capacity) throw new System.IO.PathTooLongException();
                string resolved = result.ToString();
                if (resolved.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
                    return @"\\" + resolved.Substring(8);
                if (resolved.StartsWith(@"\\?\", StringComparison.Ordinal))
                    return resolved.Substring(4);
                return resolved;
            }
        }
    }
}
'@
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            # Resolve temp junctions before invoking .NET Framework's compiler,
            # which cannot create its scratch files through a junction path.
            $compilerParent = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
            $ancestor = $compilerParent
            $redirects = 0
            while ($ancestor) {
                $ancestorItem = Get-Item -LiteralPath $ancestor -Force
                if (($ancestorItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
                    $ancestorItem.PSObject.Properties['Target'] -and $ancestorItem.Target) {
                    if (++$redirects -gt 40) { throw 'Unable to resolve compiler temp directory links.' }
                    $target = @($ancestorItem.Target)[0]
                    if (-not [System.IO.Path]::IsPathRooted($target)) {
                        $target = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($ancestor), $target)
                    }
                    $suffix = $compilerParent.Substring($ancestor.Length).TrimStart([char]92, [char]'/')
                    $compilerParent = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($target, $suffix))
                    $ancestor = $compilerParent
                    continue
                }
                $ancestor = [System.IO.Path]::GetDirectoryName($ancestor.TrimEnd([char]92, [char]'/'))
            }
            $compilerDirectory = Join-Path $compilerParent ('BackupDirectoryCompiler_' + [guid]::NewGuid().ToString('N'))
            [System.IO.Directory]::CreateDirectory($compilerDirectory) | Out-Null
            $previousCompilerTemp = $env:TEMP
            $previousCompilerTmp = $env:TMP
            try {
                $env:TEMP = $compilerDirectory
                $env:TMP = $compilerDirectory
                $compilerParameters = [System.CodeDom.Compiler.CompilerParameters]::new()
                $compilerParameters.GenerateInMemory = $true
                $compilerParameters.ReferencedAssemblies.Add([System.ComponentModel.Win32Exception].Assembly.Location) | Out-Null
                $compilerParameters.TempFiles = [System.CodeDom.Compiler.TempFileCollection]::new($compilerDirectory)
                Add-Type -TypeDefinition $nativeDefinition -CompilerParameters $compilerParameters
            }
            finally {
                $env:TEMP = $previousCompilerTemp
                $env:TMP = $previousCompilerTmp
                if ([System.IO.Directory]::Exists($compilerDirectory)) {
                    [System.IO.Directory]::Delete($compilerDirectory, $true)
                }
            }
        }
        else {
            Add-Type -TypeDefinition $nativeDefinition
        }
    }

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $suffix = [System.Collections.Generic.List[string]]::new()
    while (-not [System.IO.Directory]::Exists($fullPath)) {
        if ([System.IO.File]::Exists($fullPath)) { throw "Expected a directory: '$fullPath'." }
        $leaf = [System.IO.Path]::GetFileName($fullPath.TrimEnd([char]92, [char]'/'))
        $parent = [System.IO.Path]::GetDirectoryName($fullPath.TrimEnd([char]92, [char]'/'))
        if (-not $parent) { throw "Cannot resolve directory path: '$Path'." }
        $suffix.Insert(0, $leaf)
        $fullPath = $parent
    }
    $canonical = [BackupDirectory.NativePaths]::Resolve($fullPath)
    foreach ($leaf in $suffix) { $canonical = [System.IO.Path]::Combine($canonical, $leaf) }
    return $canonical
}

function Test-PathWithinDirectory {
    param([string]$Path, [string]$Directory)

    $normalizedPath = (Get-CanonicalDirectoryPath $Path).TrimEnd([char]92, [char]'/')
    $normalizedDirectory = (Get-CanonicalDirectoryPath $Directory).TrimEnd([char]92, [char]'/')
    return $normalizedPath.Equals($normalizedDirectory, [System.StringComparison]::OrdinalIgnoreCase) -or
        $normalizedPath.StartsWith($normalizedDirectory + [System.IO.Path]::DirectorySeparatorChar,
            [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-BackupArchive {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectoryName,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.IO.FileInfo[]]$SourceFiles
    )

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $entryTable = @{}
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) { continue }
            $key = $entry.FullName.Replace('\', '/')
            if ($entryTable.ContainsKey($key)) {
                throw "Validation failed: duplicate archive entry '$key'."
            }
            $entryTable[$key] = $entry
        }

        $validatedFiles = [System.Collections.Generic.List[object]]::new()
        foreach ($file in $SourceFiles) {
            $relative = $file.FullName.Substring($SourceDirectory.Length).TrimStart([char]92, [char]'/').Replace('\', '/')
            $entryKey = "$SourceDirectoryName/$relative"
            if (-not $entryTable.ContainsKey($entryKey)) {
                throw "Validation failed: file missing from archive: '$relative'."
            }

            $entry = $entryTable[$entryKey]
            if ($entry.Length -ne $file.Length) {
                throw "Validation failed: size mismatch for '$relative'."
            }
            $sourceHash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            $entryStream = $entry.Open()
            $hasher = [System.Security.Cryptography.SHA256]::Create()
            try {
                # Reading the entire entry also detects decompression/read failures.
                $entryHash = [System.BitConverter]::ToString($hasher.ComputeHash($entryStream)).Replace('-', '')
            }
            finally {
                $hasher.Dispose()
                $entryStream.Dispose()
            }
            if ($entryHash -ne $sourceHash) {
                throw "Validation failed: SHA-256 mismatch for '$relative'."
            }
            $validatedFiles.Add([pscustomobject]@{
                RelativePath = $relative
                SizeBytes = $file.Length
                SHA256 = $entryHash
            })
            $entryTable.Remove($entryKey)
        }

        if ($entryTable.Count -ne 0) {
            throw "Validation failed: archive contains unexpected files: $($entryTable.Keys -join ', ')."
        }
        return [pscustomobject]@{ EntryCount = $validatedFiles.Count; Files = $validatedFiles }
    }
    finally {
        $archive.Dispose()
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
$stagingPrefix = 'BackupDirectory_'
$staleStagingRetentionDays = 2
$stagingRoot = $null
$stagingSource = $null
$shadowObj = $null
$shadowIdForCleanup = $null
$backupJob = $null
$tempZipPath = $null
$tempManifestPath = $null

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
$sourceDirName = (Get-Item -LiteralPath $SourcePath -Force).Name
$dateCode      = Get-Date -Format 'yyyyMMdd_HHmmss'

# Reject overlaps before creating any output inside the source tree.
$destinationFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath)
if (Test-PathWithinDirectory -Path $destinationFullPath -Directory $SourcePath) {
    throw 'DestinationPath must be outside the source directory to prevent backing up previous backups.'
}

if (-not (Test-Path -LiteralPath $DestinationPath)) {
    Write-Host "Destination directory does not exist; creating: $DestinationPath"
    New-Item -ItemType Directory -Path $DestinationPath -Force -WhatIf:$false | Out-Null
}
$DestinationPath = (Resolve-Path -Path $DestinationPath).Path

$stagingParent = [System.IO.Path]::GetTempPath()
if (Test-PathWithinDirectory -Path $stagingParent -Directory $SourcePath) {
    # For example, a user-profile backup contains that user's system temp folder.
    $stagingParent = $DestinationPath
    Write-Host 'System temp is inside the source; staging in the backup destination instead.'
}

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

$staleCutoff = (Get-Date).AddDays(-$staleStagingRetentionDays)
Remove-StaleStagingDirectories -Prefix $stagingPrefix -OlderThan $staleCutoff

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
$compressSource = $SourcePath

$currentIdentity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdmin) {
    try {
        Write-Host 'Creating VSS shadow copy ...'
        $sourceVolume = (Split-Path -Qualifier $SourcePath) + '\'
        $contextAttempts = @('ClientAccessibleWriters', 'ClientAccessible', $null)
        $shadowCreated = $false

        foreach ($context in $contextAttempts) {
            $createArgs = @{ Volume = $sourceVolume }
            if ($null -ne $context) {
                $createArgs.Context = $context
            }

            $shadowResult = Invoke-CimMethod -ClassName Win32_ShadowCopy -MethodName Create -Arguments $createArgs
            $returnCode = [int]$shadowResult.ReturnValue

            if ($returnCode -ne 0) {
                $contextLabel = if ($null -ne $context) { $context } else { 'Default' }
                Write-Warning ("VSS create failed for context '{0}' with code {1} ({2})." -f
                    $contextLabel, $returnCode, (Get-VssCreateReturnMessage -Code $returnCode))
                continue
            }

            $shadowId = if ($shadowResult.PSObject.Properties['ShadowID']) {
                [string]$shadowResult.ShadowID
            }
            elseif ($shadowResult.PSObject.Properties['ShadowId']) {
                [string]$shadowResult.ShadowId
            }
            else {
                $null
            }

            if (-not $shadowId) {
                Write-Warning 'VSS reported success but did not return a ShadowID. Compressing live files.'
                break
            }

            $shadowIdForCleanup = $shadowId

            $shadowObj = Get-CimInstance -ClassName Win32_ShadowCopy | Where-Object { $_.ID -eq $shadowId } | Select-Object -First 1
            if ($null -eq $shadowObj -or -not $shadowObj.DeviceObject) {
                Write-Warning "VSS created shadow copy '$shadowId' but it could not be resolved for use. Compressing live files."
                if ($null -ne $shadowObj) {
                    try { $shadowObj | Remove-CimInstance } catch {}
                }
                $shadowObj = $null
                break
            }

            $relPath = (Split-Path -NoQualifier $SourcePath).TrimStart([char]92)
            $compressSource = Join-Path ($shadowObj.DeviceObject + '\') $relPath
            Write-Host "Shadow copy created: $($shadowObj.ID)"
            Write-Host "Using shadow path: $compressSource"
            $shadowCreated = $true
            break
        }

        if (-not $shadowCreated) {
            Write-Warning 'Unable to create a usable VSS snapshot; compressing live files.'
        }
    }
    catch {
        Write-Warning "VSS shadow copy failed: $($_.Exception.Message)  Compressing live files."
        $shadowObj = $null
        $compressSource = $SourcePath
    }
} else {
    Write-Warning ('Not running as Administrator; VSS shadow copy skipped. ' +
        'Files held open by other processes may be missed or cause errors.')
}

# ---------------------------------------------------------------------------
# Build staging copy outside the source for more reliable reads from live trees
# ---------------------------------------------------------------------------
$stagingRoot = Join-Path $stagingParent ($stagingPrefix + [guid]::NewGuid().ToString('N'))
$stagingSource = Join-Path $stagingRoot $sourceDirName

Write-Host "Preparing staging directory: $stagingRoot"
New-Item -ItemType Directory -Path $stagingSource -Force -WhatIf:$false | Out-Null

Write-Host 'Staging source files ...'
$robocopyArgs = @(
    $compressSource,
    $stagingSource,
    '/E',
    '/COPY:DAT',
    '/DCOPY:DAT',
    '/R:3',
    '/W:2',
    '/Z',
    '/FFT',
    '/XJ',
    '/NP',
    '/NFL',
    '/NDL',
    '/NJH',
    '/NJS'
)

& robocopy @robocopyArgs | Out-Null
$robocopyExitCode = $LASTEXITCODE
if ($robocopyExitCode -ge 8) {
    throw "Staging copy failed (Robocopy exit code $robocopyExitCode)."
}

$compressSource = $stagingSource
Write-Host "Staging complete (Robocopy exit code $robocopyExitCode)."

# ---------------------------------------------------------------------------
# Build source inventory once and run free-space pre-check
# ---------------------------------------------------------------------------
Write-Host 'Indexing source files ...'
$sourceFiles = @(Get-ChildItem -LiteralPath $compressSource -Recurse -File -Force | Sort-Object -Property FullName)
$sourceSize = 0L
if ($sourceFiles.Count -gt 0) {
    $sourceSize = ($sourceFiles | Measure-Object -Property Length -Sum).Sum
}

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
$tempZipPath   = Join-Path $DestinationPath ("${sourceDirName}_${dateCode}.tmp.zip")
$manifestFileName = "${sourceDirName}_${dateCode}.manifest.sha256"
$manifestFilePath = Join-Path $DestinationPath $manifestFileName
$tempManifestPath = $manifestFilePath + '.tmp'

Write-Host "Backing up '$SourcePath' -> '$zipFilePath' ..."

# Run compression in a background job so we can show progress on the foreground thread.
$backupJob = Start-Job -ScriptBlock {
    param($src, $dest)

    $ErrorActionPreference = 'Stop'
    # .NET includes hidden files and empty directories and supports Zip64.
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $src, $dest, [System.IO.Compression.CompressionLevel]::Optimal, $true)
} -ArgumentList $compressSource, $tempZipPath

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

# Fully read and hash every archived file before promoting the backup.
$validation = Test-BackupArchive -ArchivePath $tempZipPath -SourceDirectory $compressSource `
    -SourceDirectoryName $sourceDirName -SourceFiles $sourceFiles
$entryCount = $validation.EntryCount

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

$validation.Files |
    ForEach-Object {
        $manifestLines.Add("$($_.SHA256)  $($_.SizeBytes)  $($_.RelativePath)")
    }

Set-Content -LiteralPath $tempManifestPath -Value $manifestLines -Encoding utf8 -WhatIf:$false
Move-Item -LiteralPath $tempManifestPath -Destination $manifestFilePath -Force -WhatIf:$false
Write-Host "Manifest complete: $manifestFileName"
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
        $fileDate = Try-ParseBackupTimestampFromName -FileName $_.Name -Pattern $dateRegex -Kind 'backup'
        if ($null -eq $fileDate) {
            return
        }
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
        $fileDate = Try-ParseBackupTimestampFromName -FileName $_.Name -Pattern $logDateRegex -Kind 'log'
        if ($null -eq $fileDate) {
            return
        }
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
    Write-Error "Backup failed: $($_.Exception.Message)" -ErrorAction Continue
    exit 1
}
finally {
    # These resources can be allocated well before compression starts.
    # Stop the writer before deleting its temporary archive or staging inputs.
    if ($null -ne $backupJob) {
        try {
            Stop-Job -Job $backupJob -ErrorAction SilentlyContinue -WhatIf:$false
            Remove-Job -Job $backupJob -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }
        catch {
            Write-Warning "Failed to clean up background backup job: $_"
        }
    }

    foreach ($temporaryFile in @($tempZipPath, $tempManifestPath)) {
        if ($temporaryFile -and (Test-Path -LiteralPath $temporaryFile)) {
            Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }
    }

    if ($null -ne $shadowObj) {
        try {
            Write-Host 'Removing VSS shadow copy ...'
            $shadowObj | Remove-CimInstance -WhatIf:$false -Confirm:$false
            $shadowIdForCleanup = $null
        }
        catch {
            Write-Warning "Failed to remove VSS shadow copy '$($shadowObj.ID)': $_"
        }
    }
    if ($shadowIdForCleanup) {
        try {
            $orphanShadow = Get-CimInstance -ClassName Win32_ShadowCopy -Filter "ID='$shadowIdForCleanup'" -ErrorAction SilentlyContinue
            if ($null -ne $orphanShadow) {
                Write-Host 'Removing VSS shadow copy (fallback by ID) ...'
                $orphanShadow | Remove-CimInstance -WhatIf:$false -Confirm:$false
            }
        }
        catch {
            Write-Warning "Failed to remove VSS shadow copy '$shadowIdForCleanup': $_"
        }
    }

    if ($stagingRoot -and (Test-Path -LiteralPath $stagingRoot)) {
        try {
            Write-Host "Removing staging directory: $stagingRoot"
            Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction Stop -WhatIf:$false
        }
        catch {
            Write-Warning "Failed to remove staging directory '$stagingRoot': $($_.Exception.Message)"
        }
    }
    $WhatIfPreference = $originalWhatIfPreference

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
