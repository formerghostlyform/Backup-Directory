[CmdletBinding()]
param([string]$PowerShellExecutable = (Get-Process -Id $PID).Path)

# Dependency-free regression suite; run with Windows PowerShell 5.1 or PowerShell 7.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$backupScript = Join-Path (Split-Path $PSScriptRoot) 'Backup-Directory.ps1'
$tempParent = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$testRoot = Join-Path $tempParent ('BackupDirectoryTests_' + [guid]::NewGuid().ToString('N'))
$passed = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function New-Fixture {
    param([string]$Name)
    $root = Join-Path $testRoot $Name
    foreach ($directory in @('Source', 'Destination', 'Logs', 'Temp')) {
        [System.IO.Directory]::CreateDirectory((Join-Path $root $directory)) | Out-Null
    }
    return [pscustomobject]@{
        Root = $root
        Source = Join-Path $root 'Source'
        Destination = Join-Path $root 'Destination'
        Logs = Join-Path $root 'Logs'
        Temp = Join-Path $root 'Temp'
        Marker = Join-Path $root 'snapshot-cleanup.txt'
    }
}

function Invoke-Fixture {
    param($Fixture, [string]$Fault = 'None', [string]$Destination = '', [string]$Temp = '')
    if (-not $Destination) { $Destination = $Fixture.Destination }
    if (-not $Temp) { $Temp = $Fixture.Temp }
    $previousPreference = $ErrorActionPreference
    try {
        # Expected failures write to stderr in Windows PowerShell.
        $ErrorActionPreference = 'Continue'
        $output = & $PowerShellExecutable -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $wrapper `
            -BackupScript $backupScript -Source $Fixture.Source -Destination $Destination `
            -Logs $Fixture.Logs -Temp $Temp -Marker $Fixture.Marker -Fault $Fault 2>&1
        $code = $LASTEXITCODE
    }
    finally { $ErrorActionPreference = $previousPreference }
    return [pscustomobject]@{ ExitCode = $code; Output = ($output | Out-String) }
}

function Assert-NoResources {
    param($Fixture, $Result, [bool]$ExpectSnapshot = $false)
    $stagingDirectories = @(Get-ChildItem -LiteralPath $Fixture.Root -Directory -Recurse -Force |
        Where-Object { $_.Name -like 'BackupDirectory_*' -or $_.Name -like 'BackupDirectoryCompiler_*' })
    Assert-True ($stagingDirectories.Count -eq 0) "Temporary directories leaked: $($Result.Output)"
    $partialFiles = @(Get-ChildItem -LiteralPath $Fixture.Destination -File -Force |
        Where-Object { $_.Name -like '*.tmp*' })
    Assert-True ($partialFiles.Count -eq 0) "Partial output leaked: $($Result.Output)"
    if ($ExpectSnapshot) {
        Assert-True (Test-Path -LiteralPath $Fixture.Marker) "Snapshot was not released: $($Result.Output)"
    }
}

function Get-FixtureArchive {
    param($Fixture, $Result)
    Assert-True ($Result.ExitCode -eq 0) "Backup failed: $($Result.Output)"
    $zips = @(Get-ChildItem -LiteralPath $Fixture.Destination -Filter '*.zip' -File -Force)
    Assert-True ($zips.Count -eq 1) 'Expected exactly one completed ZIP.'
    return [System.IO.Compression.ZipFile]::OpenRead($zips[0].FullName)
}

function Get-FixtureEntry {
    param($Archive, [string]$Name)
    # .NET Framework may write backslashes; modern .NET writes forward slashes.
    return $Archive.Entries | Where-Object { $_.FullName.Replace('\', '/') -eq $Name } | Select-Object -First 1
}

function Pass {
    param([string]$Name)
    $script:passed++
    Write-Host "PASS: $Name"
}

try {
    [System.IO.Directory]::CreateDirectory($testRoot) | Out-Null
    $wrapper = Join-Path $testRoot 'Invoke-Fixture.ps1'
    # All VSS calls are mocked. The suite never creates or deletes real snapshots.
    @'
param($BackupScript, $Source, $Destination, $Logs, $Temp, $Marker, $Fault)
$ErrorActionPreference = 'Stop'
$env:TEMP = $Temp
$env:TMP = $Temp
$fakeShadow = [pscustomobject]@{ ID = 'fixture-snapshot'; DeviceObject = (Split-Path -Qualifier $Source) }
function New-Object {
    [CmdletBinding()]
    param([Parameter(Position=0)][string]$TypeName, [Parameter(Position=1)][object[]]$ArgumentList)
    if ($TypeName -eq 'Security.Principal.WindowsPrincipal') {
        $principal = [pscustomobject]@{}
        $principal | Add-Member -MemberType ScriptMethod -Name IsInRole -Value { return $true }
        return $principal
    }
    Microsoft.PowerShell.Utility\New-Object @PSBoundParameters
}
function Invoke-CimMethod {
    [CmdletBinding(SupportsShouldProcess)]
    param($ClassName, $MethodName, $Arguments)
    if ($ClassName -ne 'Win32_ShadowCopy') { throw 'Unexpected CIM method in fixture.' }
    return [pscustomobject]@{ ReturnValue = 0; ShadowID = $fakeShadow.ID }
}
function Get-CimInstance {
    [CmdletBinding()]
    param($ClassName, $Filter)
    if ($ClassName -eq 'Win32_ShadowCopy') {
        if ($Fault -eq 'UnresolvedSnapshot' -and -not $Filter) { throw 'Injected snapshot lookup failure' }
        return $fakeShadow
    }
    CimCmdlets\Get-CimInstance @PSBoundParameters
}
function Remove-CimInstance {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(ValueFromPipeline)]$InputObject)
    process {
        if ($InputObject.ID -ne 'fixture-snapshot') { throw 'Refusing to remove a non-fixture snapshot.' }
        if ($Fault -eq 'SnapshotRemoveFailure' -and -not (Get-Variable -Name fixtureRemovalAttempted -Scope Global -ErrorAction SilentlyContinue)) {
            $global:fixtureRemovalAttempted = $true
            throw 'Injected snapshot removal failure'
        }
        [System.IO.File]::AppendAllText($Marker, "Removed fixture snapshot`n")
    }
}
# Suppress unrelated cleanup of the machine's existing staging directories.
function Get-ChildItem {
    [CmdletBinding()]
    param($LiteralPath, $Filter, [switch]$Directory, [switch]$File, [switch]$Recurse, [switch]$Force)
    if ($Fault -eq 'InventoryFailure' -and $File -and $Recurse) { throw 'Injected inventory failure' }
    Microsoft.PowerShell.Management\Get-ChildItem @PSBoundParameters
}
if ($Fault -eq 'StagingFailure') {
    function robocopy { $global:LASTEXITCODE = 8 }
}
if ($Fault -eq 'NoSpace') {
    function Get-PSDrive { [CmdletBinding()] param($Name) [pscustomobject]@{ Free = 1L } }
}
if ($Fault -eq 'CompressionFailure') {
    function Start-Job {
        [CmdletBinding()]
        param($ScriptBlock, $ArgumentList)
        [System.IO.File]::WriteAllBytes($ArgumentList[1], [byte[]]@(1,2,3))
        throw 'Injected compression failure'
    }
}
if ($Fault -eq 'ArchiveMismatch') {
    function Receive-Job {
        [CmdletBinding()]
        param([Parameter(ValueFromPipeline)]$Job, [switch]$Wait, [switch]$AutoRemoveJob)
        process {
            Microsoft.PowerShell.Core\Receive-Job @PSBoundParameters
            Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
            $zipPath = @(Microsoft.PowerShell.Management\Get-ChildItem -LiteralPath $Destination -Filter '*.tmp.zip')[0].FullName
            $zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Update)
            try {
                $entry = $zip.Entries | Where-Object { $_.FullName.Replace('\', '/') -eq 'Source/payload.txt' } | Select-Object -First 1
                $stream = $entry.Open()
                try {
                    $stream.Position = 0
                    $stream.WriteByte([byte][char]'X')
                }
                finally { $stream.Dispose() }
            }
            finally { $zip.Dispose() }
        }
    }
}
& $BackupScript -SourcePath $Source -DestinationPath $Destination -LogDirectory $Logs
if ($?) { exit 0 } else { exit 1 }
'@ | Set-Content -LiteralPath $wrapper -Encoding UTF8

    $fixture = New-Fixture 'hidden-files'
    [System.IO.File]::WriteAllText((Join-Path $fixture.Source 'visible.txt'), 'visible payload')
    $hiddenFile = Join-Path $fixture.Source 'hidden.txt'
    [System.IO.File]::WriteAllText($hiddenFile, 'hidden payload')
    [System.IO.File]::SetAttributes($hiddenFile, [System.IO.FileAttributes]::Hidden)
    $hiddenDirectory = Join-Path $fixture.Source 'secret'
    [System.IO.Directory]::CreateDirectory($hiddenDirectory) | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $hiddenDirectory 'nested.txt'), 'nested hidden payload')
    [System.IO.File]::SetAttributes($hiddenDirectory, [System.IO.FileAttributes]::Hidden)
    $result = Invoke-Fixture $fixture
    $archive = Get-FixtureArchive $fixture $result
    try {
        foreach ($entryName in @('Source/visible.txt', 'Source/hidden.txt', 'Source/secret/nested.txt')) {
            Assert-True ($null -ne (Get-FixtureEntry $archive $entryName)) "Missing ZIP entry ${entryName}: $($archive.Entries.FullName -join ', ')"
        }
    }
    finally { $archive.Dispose() }
    $manifest = Get-Content -LiteralPath @(Get-ChildItem -LiteralPath $fixture.Destination -Filter '*.manifest.sha256')[0].FullName
    foreach ($relative in @('visible.txt', 'hidden.txt', 'secret/nested.txt')) {
        $hash = (Get-FileHash -LiteralPath (Join-Path $fixture.Source $relative) -Algorithm SHA256).Hash
        Assert-True ([bool]($manifest -match "^$hash  \d+  $([regex]::Escape($relative))$")) "Manifest missing verified hash for $relative"
    }
    Assert-NoResources $fixture $result $true
    Pass 'Hidden files and hidden subdirectories appear in ZIP and checksum manifest'

    $fixture = New-Fixture 'hidden-only'
    $hiddenFile = Join-Path $fixture.Source 'only.txt'
    [System.IO.File]::WriteAllText($hiddenFile, 'hidden only')
    [System.IO.File]::SetAttributes($hiddenFile, [System.IO.FileAttributes]::Hidden)
    [System.IO.File]::SetAttributes($fixture.Source, [System.IO.FileAttributes]::Hidden)
    $result = Invoke-Fixture $fixture
    $archive = Get-FixtureArchive $fixture $result
    try { Assert-True ($null -ne (Get-FixtureEntry $archive 'Source/only.txt')) 'Hidden-only backup lost its file.' }
    finally { $archive.Dispose() }
    Assert-NoResources $fixture $result $true
    Pass 'Hidden source root containing only hidden files'

    foreach ($case in @('empty', 'empty-subdirectories')) {
        $fixture = New-Fixture $case
        if ($case -eq 'empty-subdirectories') {
            [System.IO.Directory]::CreateDirectory((Join-Path $fixture.Source 'one\two')) | Out-Null
        }
        $result = Invoke-Fixture $fixture
        $archive = Get-FixtureArchive $fixture $result
        try {
            $expected = if ($case -eq 'empty') { 'Source/' } else { 'Source/one/two/' }
            Assert-True ($null -ne (Get-FixtureEntry $archive $expected)) "Empty directory entry $expected missing."
        }
        finally { $archive.Dispose() }
        Assert-NoResources $fixture $result $true
        Pass "Successful backup of $case"
    }

    foreach ($destinationCase in @('equal', 'child')) {
        $fixture = New-Fixture "destination-$destinationCase"
        $destination = if ($destinationCase -eq 'equal') { $fixture.Source } else { Join-Path $fixture.Source 'Backups' }
        $result = Invoke-Fixture $fixture -Destination $destination
        Assert-True ($result.ExitCode -ne 0 -and $result.Output -match 'DestinationPath must be outside') "Overlapping destination was accepted: $($result.Output)"
        if ($destinationCase -eq 'child') { Assert-True (-not (Test-Path -LiteralPath $destination)) 'Rejected destination was created.' }
        Assert-NoResources $fixture $result
        Pass "Reject source/$destinationCase destination before creating resources"
    }

    $fixture = New-Fixture 'temp-overlap'
    $insideTemp = Join-Path $fixture.Source 'Temp'
    [System.IO.Directory]::CreateDirectory($insideTemp) | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $fixture.Source 'file.txt'), 'payload')
    $result = Invoke-Fixture $fixture -Temp $insideTemp
    $archive = Get-FixtureArchive $fixture $result
    try { Assert-True ($null -ne (Get-FixtureEntry $archive 'Source/file.txt')) 'Fallback staging lost the source file.' }
    finally { $archive.Dispose() }
    Assert-True ($result.Output -match 'staging in the backup destination') 'Expected safe staging fallback.'
    Assert-NoResources $fixture $result $true
    Pass 'Temp inside source uses destination staging without recursion'

    $fixture = New-Fixture 'destination-junction'
    $insideDestination = Join-Path $fixture.Source 'Backups'
    [System.IO.Directory]::CreateDirectory($insideDestination) | Out-Null
    $alias = Join-Path $fixture.Root 'DestinationAlias'
    New-Item -ItemType Junction -Path $alias -Value $insideDestination | Out-Null
    try {
        $result = Invoke-Fixture $fixture -Destination (Join-Path $alias 'NewChild')
        Assert-True ($result.ExitCode -ne 0 -and $result.Output -match 'DestinationPath must be outside') "Junction overlap was accepted: $($result.Output)"
        Assert-True (-not (Test-Path -LiteralPath (Join-Path $insideDestination 'NewChild'))) 'Rejected aliased destination was created.'
        Assert-NoResources $fixture $result
    }
    finally { [System.IO.Directory]::Delete($alias) }
    Pass 'Reject destination junction pointing inside source, including a nonexistent child'

    $fixture = New-Fixture 'temp-junction'
    $insideTemp = Join-Path $fixture.Source 'Temp'
    [System.IO.Directory]::CreateDirectory($insideTemp) | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $fixture.Source 'file.txt'), 'payload')
    $alias = Join-Path $fixture.Root 'TempAlias'
    New-Item -ItemType Junction -Path $alias -Value $insideTemp | Out-Null
    try {
        $result = Invoke-Fixture $fixture -Temp $alias
        $archive = Get-FixtureArchive $fixture $result
        $archive.Dispose()
        Assert-True ($result.Output -match 'staging in the backup destination') 'Temp junction overlap did not trigger safe staging.'
        Assert-NoResources $fixture $result $true
    }
    finally { [System.IO.Directory]::Delete($alias) }
    Pass 'Temp junction pointing inside source uses safe destination staging'

    foreach ($fault in @('StagingFailure', 'InventoryFailure', 'NoSpace', 'CompressionFailure', 'ArchiveMismatch', 'UnresolvedSnapshot', 'SnapshotRemoveFailure')) {
        $fixture = New-Fixture $fault
        [System.IO.File]::WriteAllText((Join-Path $fixture.Source 'payload.txt'), 'original payload that requires more than one byte of disk space')
        $result = Invoke-Fixture $fixture -Fault $fault
        if ($fault -eq 'UnresolvedSnapshot' -or $fault -eq 'SnapshotRemoveFailure') {
            Assert-True ($result.ExitCode -eq 0) "Live fallback failed: $($result.Output)"
            Assert-True ($result.Output -match 'fallback by ID') 'ID-based snapshot cleanup was not exercised.'
        }
        else {
            Assert-True ($result.ExitCode -ne 0) "Injected $fault unexpectedly succeeded."
            $expectedMessage = switch ($fault) {
                'StagingFailure' { 'Staging copy failed' }
                'InventoryFailure' { 'Injected inventory failure' }
                'NoSpace' { 'Insufficient disk space' }
                'CompressionFailure' { 'Injected compression failure' }
                'ArchiveMismatch' { 'SHA-256 mismatch' }
            }
            Assert-True ($result.Output -match $expectedMessage) "Wrong failure for ${fault}: $($result.Output)"
            Assert-True (@(Get-ChildItem -LiteralPath $fixture.Destination -Filter '*.zip' -File).Count -eq 0) 'Failed backup was promoted to final ZIP.'
        }
        Assert-NoResources $fixture $result $true
        Pass "$fault releases staging, temporary output and snapshot"
    }

    # Load only the validation helper, never the script's executable body.
    $tokens = $null; $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($backupScript, [ref]$tokens, [ref]$parseErrors)
    Assert-True ($parseErrors.Count -eq 0) 'Production script has syntax errors.'
    $helper = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-BackupArchive' }, $true)
    . ([scriptblock]::Create($helper.Extent.Text))
    $fixture = New-Fixture 'validation'
    $filePath = Join-Path $fixture.Source 'file.txt'
    [System.IO.File]::WriteAllText($filePath, 'HELLO')
    $files = @(Get-ChildItem -LiteralPath $fixture.Source -File -Force)
    foreach ($badCase in @('missing', 'unexpected', 'duplicate', 'size-mismatch', 'damaged-payload', 'invalid-zip')) {
        $zipPath = Join-Path $fixture.Destination "$badCase.zip"
        if ($badCase -eq 'invalid-zip') {
            [System.IO.File]::WriteAllText($zipPath, 'not a ZIP')
        }
        else {
            $zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                $entryNames = switch ($badCase) {
                    'missing' { @() }
                    'unexpected' { @('Source/file.txt', 'Source/extra.txt') }
                    'duplicate' { @('Source/file.txt', 'Source/file.txt') }
                    'size-mismatch' { @('Source/file.txt') }
                    'damaged-payload' { @('Source/file.txt') }
                }
                foreach ($entryName in $entryNames) {
                    $entry = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::NoCompression)
                    $stream = $entry.Open()
                    try {
                        $payload = if ($badCase -eq 'size-mismatch') { 'FOUR' } else { 'HELLO' }
                        $bytes = [System.Text.Encoding]::ASCII.GetBytes($payload)
                        $stream.Write($bytes, 0, $bytes.Length)
                    }
                    finally { $stream.Dispose() }
                }
            }
            finally { $zip.Dispose() }
            if ($badCase -eq 'damaged-payload') {
                # Change raw payload without updating CRC, sizes or directory metadata.
                $bytes = [System.IO.File]::ReadAllBytes($zipPath)
                $offset = [System.Text.Encoding]::ASCII.GetString($bytes).IndexOf('HELLO', [System.StringComparison]::Ordinal)
                Assert-True ($offset -ge 0) 'Uncompressed test payload not found.'
                $bytes[$offset] = [byte][char]'J'
                [System.IO.File]::WriteAllBytes($zipPath, $bytes)
            }
        }
        $rejected = $false
        try { Test-BackupArchive -ArchivePath $zipPath -SourceDirectory $fixture.Source -SourceDirectoryName 'Source' -SourceFiles $files | Out-Null }
        catch { $rejected = $true }
        Assert-True $rejected "Validation accepted $badCase."
        Pass "Validation rejects $badCase"
    }

    Write-Host "$passed regression cases passed with $($PSVersionTable.PSVersion)."
}
finally {
    # Recursive deletion is limited to the unique temp root created by this suite.
    $resolvedRoot = [System.IO.Path]::GetFullPath($testRoot)
    $expectedParent = $tempParent.TrimEnd([char]92, [char]'/') + [System.IO.Path]::DirectorySeparatorChar
    if (-not $resolvedRoot.StartsWith($expectedParent, [System.StringComparison]::OrdinalIgnoreCase) -or
        [System.IO.Path]::GetFileName($resolvedRoot) -notmatch '^BackupDirectoryTests_[0-9a-f]{32}$') {
        throw 'Refusing cleanup outside the verified test temp root.'
    }
    if (Test-Path -LiteralPath $resolvedRoot) { Remove-Item -LiteralPath $resolvedRoot -Recurse -Force }
}
