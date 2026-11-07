$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $projectRoot 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'GUI script has syntax errors' }

# Load only pure configuration helpers and the copy routine, never the GUI or AD audit.
$functionNames = @('ConvertTo-TargetFolders', 'Get-MissingTargetFolders', 'Load-Config', 'Save-Config', 'Sync-UserFiles', 'Invoke-UserFileSync', 'Assert-ProfileTargetPath')
foreach ($node in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $functionNames) { . ([scriptblock]::Create($node.Extent.Text)) }
}
Add-Type -AssemblyName System.Windows.Forms
function Assert ($Condition, $Message) { if (-not $Condition) { throw $Message } }

$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('PM_TargetFolders_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $fixtureRoot -ItemType Directory)
$sourceRoot = Join-Path $fixtureRoot 'source'
$destinationRoot = Join-Path $fixtureRoot 'destination'
$fakeSource = '\\PM-TEST\C$\Users\tester'

# Redirect only this fake source to a local fixture; do not access a server or real config.
function Test-Path {
    param($Path, $LiteralPath, $PathType = 'Any')
    $candidate = if ($PSBoundParameters.ContainsKey('LiteralPath')) { $LiteralPath } else { $Path }
    if ($candidate -eq 'C:\Migrate\MigrateConfig.json') { throw 'Legacy configuration must never be accessed' }
    if ($candidate.StartsWith($fakeSource)) { $candidate = $candidate.Replace($fakeSource, $sourceRoot) }
    Microsoft.PowerShell.Management\Test-Path -LiteralPath $candidate -PathType $PathType
}
function Get-Item {
    [CmdletBinding()]
    param([string]$LiteralPath,[switch]$Force)
    if ($LiteralPath.StartsWith($fakeSource)) { $LiteralPath = $LiteralPath.Replace($fakeSource,$sourceRoot) }
    Microsoft.PowerShell.Management\Get-Item -LiteralPath $LiteralPath -Force:$Force -ErrorAction Stop
}
function Write-GuiLog ($msg, $type) { $script:messages += "$type $msg" }
function Apply-SingleACL ($UserName, $UserFolder) { $script:aclCalls++ }
function Assert-ProfileOffline ($Profile) {}
function Get-UserProfileContext ($UserName, [switch]$Create, [switch]$AdoptExisting) {
    Assert ($Create -and $AdoptExisting.IsPresent -eq (-not $isDelta)) 'Only step 2 may adopt an existing unregistered directory'
    if ($script:failProfile) { throw 'Profile registration failed' }
    $path = Join-Path $destinationRoot ($UserName + '.DOMAIN')
    [void](New-Item -Path $path -ItemType Directory -Force)
    return [pscustomobject]@{ Sid = 'S-1-5-21-111-222-333-1001'; Path = $path; UserName = $UserName }
}
function Invoke-CliCommand ($ExePath, [string[]]$ArgsList) {
    Assert ($ExePath -eq 'robocopy.exe') 'Unexpected process requested'
    $script:copyCalls += ,$ArgsList
    if ($script:forceCopyFailure) { return 8 }
    # Local test files need no backup privilege; production keeps /ZB.
    $localArgs = @($ArgsList | ForEach-Object { if ($_ -eq '/ZB') { '/Z' } else { $_ } })
    $localArgs[0] = $localArgs[0].Replace($fakeSource, $sourceRoot)
    $copyOutput = & robocopy.exe @localArgs
    $code = $LASTEXITCODE
    if ($code -ge 8) {
        Write-Host ($copyOutput -join "`n")
        foreach ($arg in $localArgs) {
            if ($arg.StartsWith('/LOG:')) { Write-Host (Get-Content -LiteralPath $arg.Substring(5) -Raw) }
        }
    }
    return $code
}

try {
    $Global:ConfigFile = Join-Path $fixtureRoot 'pm_config.json'
    $defaultConfig = $ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$Global:Config'
    }, $true)
    . ([scriptblock]::Create($defaultConfig.Extent.Text))
    $Global:Config.TargetOU = 'DC=test,DC=local'
    $Global:Config.LogsDir = Join-Path $fixtureRoot 'logs'
    # First run: no configuration and no working/log directory must be required.
    Load-Config
    $expectedDefaults = 'Desktop,Documents,Downloads'
    Assert (($Global:Config.TargetFolders -join ',') -eq $expectedDefaults) 'Missing config did not select default folders'
    Assert (Test-Path -LiteralPath $Global:ConfigFile) 'First launch did not create pm_config.json'
    $generated = Get-Content -LiteralPath $Global:ConfigFile -Raw | ConvertFrom-Json
    Assert (($generated.TargetFolders -join ',') -eq $expectedDefaults) 'Generated folder defaults incorrect'
    Assert ($generated.SourceServer -eq 'SRV-SOURCE' -and $generated.InactiveDaysLimit -eq 90) 'Generated server/activity defaults incorrect'
    Assert ($generated.ProfilesRoot -eq 'C:\Users' -and $generated.WorkDir -eq 'C:\Migrate') 'Generated profile/work defaults incorrect'
    foreach ($key in $Global:Config.Keys) { Assert ($generated.PSObject.Properties.Name -contains $key) "Missing generated setting: $key" }
    Assert (-not (Test-Path -LiteralPath $Global:Config.LogsDir)) 'Config loading unexpectedly required/created logs'
    Set-Content -LiteralPath $Global:ConfigFile -Value '{}' -Encoding UTF8
    $Global:Config.TargetFolders = @('PreviousCustomFolder')
    Load-Config
    Assert (($Global:Config.TargetFolders -join ',') -eq $expectedDefaults) 'Absent TargetFolders did not restore default folders'
    Set-Content -LiteralPath $Global:ConfigFile -Value '{"SourceServer":"PM-TEST","TargetFolders":["WorkData"],"CustomSetting":"keep me"}' -Encoding UTF8
    $existingHash = (Get-FileHash -LiteralPath $Global:ConfigFile).Hash
    Load-Config
    Assert ($Global:Config.TargetFolders.Count -eq 1 -and $Global:Config.TargetFolders[0] -eq 'WorkData') 'Explicit folder list must replace defaults'
    Assert ((Get-FileHash -LiteralPath $Global:ConfigFile).Hash -eq $existingHash) 'Startup overwrote existing configuration'
    # Keep fixtures independent of the user's editable pm_config.json.
    @{ TargetFolders = @('Desktop', 'Documents', 'Downloads', 'Favorites', 'Pictures', 'AppData\Roaming\Thunderbird', 'AppData\Local\Crypto Pro') } |
        ConvertTo-Json | Set-Content -LiteralPath $Global:ConfigFile -Encoding UTF8
    Load-Config
    Assert ($Global:Config.TargetFolders.Count -eq 7) 'Custom configuration did not load all seven folders'
    Assert ($Global:Config.SourceServer -eq 'PM-TEST') 'Folders-only config overwrote server settings'
    Save-Config
    $Global:Config.TargetFolders = @('Desktop')
    Load-Config
    Assert ($Global:Config.TargetFolders -contains 'AppData\Local\Crypto Pro') 'Save/load lost Crypto Pro path'

    $normalized = @(ConvertTo-TargetFolders @('AppData/Roaming/Thunderbird', 'appdata\roaming\thunderbird'))
    Assert ($normalized.Count -eq 1 -and $normalized[0] -eq 'AppData\Roaming\Thunderbird') 'Normalization failed'
    foreach ($bad in @('..\Windows', 'C:\Data', '\\server\share', 'AppData\..\Data', 'AppData\*', 'AppData\', 'AppData\folder.', '')) {
        $rejected = $false
        try { ConvertTo-TargetFolders @($bad) | Out-Null } catch { $rejected = $true }
        Assert $rejected "Invalid path accepted: $bad"
    }
    foreach ($badJson in @('{', '{"TargetFolders":[]}', '{"TargetFolders":"Desktop"}', '{"TargetFolders":null}')) {
        Set-Content -LiteralPath $Global:ConfigFile -Value $badJson -Encoding UTF8
        $rejected = $false
        try { Load-Config } catch { $rejected = $true }
        Assert $rejected "Invalid config accepted: $badJson"
        Assert ((Get-Content -LiteralPath $Global:ConfigFile -Raw).Trim() -eq $badJson) 'Invalid config was overwritten'
    }

    $Global:Config.TargetFolders = @('Desktop', 'AppData\Roaming\Thunderbird', 'AppData\Local\Crypto Pro', 'WorkData\Custom', 'WorkData_Custom')
    foreach ($folder in $Global:Config.TargetFolders) {
        $directory = Join-Path $sourceRoot $folder
        [void](New-Item -Path $directory -ItemType Directory -Force)
        Set-Content -LiteralPath (Join-Path $directory 'sample.txt') -Value "Fixture: $folder" -Encoding UTF8
    }
    $txtSrc = [pscustomobject]@{ Text = 'PM-TEST' }
    $txtDst = [pscustomobject]@{ Text = $destinationRoot }
    $Global:CancelRequested = $false
    $script:messages = @(); $script:copyCalls = @(); $script:aclCalls = 0
    $script:forceCopyFailure = $false
    $initialCopy = Sync-UserFiles 'tester' $false
    Assert $initialCopy ("Initial copy failed: " + ($script:messages -join '; '))
    $targetProfile = Join-Path $destinationRoot 'tester.DOMAIN'
    Assert (-not (Test-Path -LiteralPath (Join-Path $destinationRoot 'tester'))) 'Copy ignored registered profile path'
    Assert (@(Get-MissingTargetFolders $targetProfile).Count -eq 0) 'Validation missed copied folders'
    foreach ($folder in $Global:Config.TargetFolders) {
        $srcFile = Join-Path (Join-Path $sourceRoot $folder) 'sample.txt'
        $dstFile = Join-Path (Join-Path $targetProfile $folder) 'sample.txt'
        Assert ((Get-FileHash -LiteralPath $srcFile).Hash -eq (Get-FileHash -LiteralPath $dstFile).Hash) "Content mismatch: $folder"
    }
    Assert ((Get-ChildItem -LiteralPath $Global:Config.LogsDir -Filter '*.log').Count -eq 5) 'Nested log filenames collided or failed'
    $script:copyCalls = @()
    Assert (Sync-UserFiles 'tester' $true) 'Delta copy failed'
    foreach ($call in $script:copyCalls) { Assert ($call -contains '/XO') 'Delta flag missing' }
    $script:copyCalls = @(); $script:failProfile = $true
    Assert (-not (Sync-UserFiles 'tester' $false)) 'Registration failure reported success'
    Assert ($script:copyCalls.Count -eq 0) 'Files copied after profile registration failure'
    $script:failProfile = $false
    $script:forceCopyFailure = $true
    Assert (-not (Sync-UserFiles 'tester' $false)) 'Robocopy error reported success'
    $script:forceCopyFailure = $false
    $Global:Config.TargetFolders += 'MissingApp'
    $missing = @(Get-MissingTargetFolders $targetProfile)
    Assert ($missing.Count -eq 1 -and $missing[0] -eq 'MissingApp') 'Missing folder not reported'
    foreach ($delta in @($false,$true)) {
        $script:messages = @()
        Assert (Sync-UserFiles 'tester' $delta) 'Absent source folder failed the operation'
        Assert (@($script:messages -like 'INFO *MissingApp*').Count -eq 1) 'Missing folder was not logged as information'
    }
    $Global:CancelRequested = $true
    Assert (-not (Sync-UserFiles 'tester' $false)) 'Cancellation reported success'
    Write-Output 'PASS: config round-trip, path validation, nested copy, file hashes, logs, delta, failures and cancellation.'
} finally {
    $resolvedFixture = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolvedFixture.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or
        [IO.Path]::GetFileName($resolvedFixture) -notlike 'PM_TargetFolders_*') { throw 'Unsafe fixture cleanup path' }
    Remove-Item -LiteralPath $resolvedFixture -Recurse -Force
}
