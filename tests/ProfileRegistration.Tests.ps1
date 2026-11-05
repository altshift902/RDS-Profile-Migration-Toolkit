$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw 'GUI syntax errors' }
$names = @('Resolve-ProfileIdentity', 'ConvertTo-LocalProfilePath', 'Assert-ProfileLocation', 'Find-RegisteredProfile', 'Assert-ProfileOffline',
    'Assert-ProfileDirectory', 'Assert-ProfileFiles', 'Initialize-ProfileApi', 'Get-UserProfileContext', 'Assert-ProfileTargetPath',
    'Get-ProfileReadiness', 'Get-MissingTargetFolders', 'Test-ProfilePermissions', 'Apply-SingleACL', 'Invoke-ProfileAclCommand')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $names) { . ([scriptblock]::Create($node.Extent.Text)) }
}
function Assert ($Condition, $Message) { if (-not $Condition) { throw $Message } }
function Assert-Throws ([scriptblock]$Action, $Message) {
    $thrown = $false
    try { & $Action | Out-Null } catch { $thrown = $true }
    Assert $thrown $Message
}

$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('PM_Profiles_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $fixtureRoot -ItemType Directory)
$sid = 'S-1-5-21-111-222-333-1001'
$otherSid = 'S-1-5-21-111-222-333-1002'
$systemRoot = Join-Path $fixtureRoot 'Users'
$profilePath = Join-Path $systemRoot 'tester.DOMAIN'
$script:state = $null

# These doubles isolate every registry, domain, profile creation and ACL write.
function Import-Module { param($Name, $ErrorAction) Assert ($Name -eq 'ActiveDirectory') 'Unexpected module' }
function Get-ADUser {
    param($Identity, $Properties, $ErrorAction)
    if ($script:state.AdFailure) { throw 'AD unavailable' }
    return [pscustomobject]@{ SamAccountName = 'tester'; SID = [pscustomobject]@{ Value = $sid } }
}
function Get-ProfileRegistrySnapshot { return $script:state.Snapshot }
function Get-CimInstance {
    param($ClassName, $Filter, $ErrorAction)
    Assert ($ClassName -eq 'Win32_UserProfile' -and $Filter -eq "SID='$sid'") 'Wrong profile query'
    if ($script:state.CimFailure) { throw 'CIM unavailable' }
    return [pscustomobject]@{ Loaded = $script:state.Loaded; Special = $false; Status = 0; LocalPath = $profilePath }
}
function Write-GuiLog ($Message, $Kind) {}
function Invoke-CreateWindowsProfile ($Identity) {
    $script:state.CreateCalls++
    Assert ($Identity.Sid -eq $sid) 'CreateProfile received wrong SID'
    if ($script:state.Result -eq 0 -or $script:state.Result -eq -2147024713) {
        [void](New-Item -Path $profilePath -ItemType Directory -Force)
        $hive = Join-Path $profilePath 'NTUSER.DAT'
        if (-not (Test-Path -LiteralPath $hive)) { Set-Content -LiteralPath $hive -Value 'target-default-hive' }
        if (-not $script:state.SkipRegistration) {
            $script:state.Snapshot.Entries = @([pscustomobject]@{ Sid = $sid; Path = $profilePath })
        }
    }
    return [pscustomobject]@{ HResult = $script:state.Result; Path = $script:state.ReturnPath }
}
function Invoke-CliCommand ($ExePath, [string[]]$ArgsList) {
    Assert ($ExePath -eq 'icacls.exe') 'Unexpected native call'
    $script:state.AclCalls += ,$ArgsList
    return $script:state.AclResult
}
function Get-Acl {
    param($LiteralPath, $ErrorAction)
    $rules = @(foreach ($principal in @($script:sid, 'S-1-5-18', 'S-1-5-32-544')) {
        if ($script:state.OmitUserAllow -and $principal -eq $script:sid) { continue }
        [pscustomobject]@{
            IdentityReference = [Security.Principal.SecurityIdentifier]$principal
            FileSystemRights = [Security.AccessControl.FileSystemRights]::FullControl
            AccessControlType = 'Allow'; PropagationFlags = [Security.AccessControl.PropagationFlags]::None
        }
    })
    if ($script:state.Deny) {
        $rules += [pscustomobject]@{
            IdentityReference = [Security.Principal.SecurityIdentifier]'S-1-1-0'
            FileSystemRights = [Security.AccessControl.FileSystemRights]::Read
            AccessControlType = 'Deny'; PropagationFlags = [Security.AccessControl.PropagationFlags]::None
        }
    }
    $acl = [pscustomobject]@{ Rules = $rules }
    $acl | Add-Member -MemberType ScriptMethod -Name GetAccessRules -Value { param($a, $b, $t) return $this.Rules }
    return $acl
}
function Reset-State {
    $script:state = @{
        Snapshot = [pscustomobject]@{ Root = $systemRoot; Entries = @() }
        CreateCalls = 0; Result = 0; ReturnPath = $profilePath; SkipRegistration = $false
        Loaded = $false; AdFailure = $false; CimFailure = $false
        AclCalls = @(); AclResult = 0; Deny = $false; OmitUserAllow = $false
    }
    $Global:CancelRequested = $false
    $script:txtDst = [pscustomobject]@{ Text = $systemRoot }
    $Global:Config = @{ TargetFolders = @('Desktop', 'Documents', 'Downloads') }
}

try {
    # Compile and inspect the real interop declaration, but never invoke the Windows API.
    Initialize-ProfileApi
    $method = [ProfileMigrator.NativeProfile].GetMethod('CreateProfile')
    Assert ($method.ReturnType -eq [int] -and $method.GetParameters().Count -eq 4) 'Invalid interop signature'
    Reset-State
    $profile = Get-UserProfileContext 'tester' -Create
    Assert ($profile.Path -eq $profilePath -and $profile.Sid -eq $sid) 'Creation did not use registered path/SID'
    Assert ($state.CreateCalls -eq 1) 'New profile was not created exactly once'
    $hivePath = Join-Path $profilePath 'NTUSER.DAT'
    $hiveHash = (Get-FileHash -LiteralPath $hivePath).Hash
    # An existing registration takes priority over the target folder field.
    $txtDst.Text = Join-Path $fixtureRoot 'AnotherRoot'
    $again = Get-UserProfileContext 'tester' -Create
    Assert ($again.Path -eq $profilePath -and $state.CreateCalls -eq 1) 'Existing profile was recreated'
    Assert ((Get-FileHash -LiteralPath $hivePath).Hash -eq $hiveHash) 'Existing hive changed'
    $identity = [pscustomobject]@{ UserName = 'tester'; Sid = $sid }
    $state.Snapshot.Entries += [pscustomobject]@{ Sid = "$sid.bak"; Path = $profilePath }
    Assert-Throws { Get-UserProfileContext 'tester' -Create } '.bak was ignored'
    $state.Snapshot.Entries = @([pscustomobject]@{ Sid = $sid; Path = $profilePath }, [pscustomobject]@{ Sid = $otherSid; Path = $profilePath })
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'Shared profile path was accepted'

    Reset-State
    $state.Snapshot.Entries = @([pscustomobject]@{ Sid = $sid; Path = $profilePath })
    $state.Loaded = $true
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'Loaded profile was accepted for writing'
    Assert-Throws { Get-ProfileReadiness 'tester' } 'Loaded profile was reported ready'
    $state.Loaded = $false; $state.CimFailure = $true
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'CIM failure was ignored'
    $state.CimFailure = $false; $state.AdFailure = $true
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'AD failure was ignored'
    $state.AdFailure = $false
    $state.Snapshot.Entries[0].Path = Join-Path $systemRoot 'missing-hive'
    [void](New-Item -Path $state.Snapshot.Entries[0].Path -ItemType Directory)
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'Missing hive was silently replaced'
    Assert ($state.CreateCalls -eq 0) 'Broken existing profile triggered recreation'

    Reset-State
    $txtDst.Text = Join-Path $fixtureRoot 'WrongRoot'
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'Unsupported creation root was accepted'
    Assert ($state.CreateCalls -eq 0) 'CreateProfile called before root validation'
    $txtDst.Text = $systemRoot; $Global:CancelRequested = $true
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'Cancellation ignored before creation'
    Assert ($state.CreateCalls -eq 0) 'Profile created after cancellation'
    Reset-State
    $state.Result = -2147024891
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'Failed HRESULT was ignored'
    Reset-State
    $state.Result = -2147024713
    Assert ((Get-UserProfileContext 'tester' -Create).Path -eq $profilePath) 'Already-exists race was not resolved through registry'
    Reset-State
    $state.SkipRegistration = $true
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'Unregistered API result was accepted'
    Reset-State
    $state.ReturnPath = Join-Path $systemRoot 'wrong-user'
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'API/registry mismatch was accepted'
    Reset-State
    Assert-Throws { Get-ProfileReadiness 'tester' } 'Unregistered profile reported ready'
    Assert ($state.CreateCalls -eq 0) 'Read-only validation created a profile'
    $profile = Get-UserProfileContext 'tester' -Create
    $check = Get-ProfileReadiness 'tester'
    Assert (-not $check.Ready -and $check.Issues.Count -gt 0) 'Missing folders not detected'
    foreach ($folder in $Global:Config.TargetFolders) { [void](New-Item -Path (Join-Path $profilePath $folder) -ItemType Directory -Force) }
    Assert ((Get-ProfileReadiness 'tester').Ready) 'Valid profile not ready'
    $state.OmitUserAllow = $true
    Assert (-not (Get-ProfileReadiness 'tester').Ready) 'Missing user ACL not detected'
    $state.OmitUserAllow = $false; $state.Deny = $true
    Assert (-not (Get-ProfileReadiness 'tester').Ready) 'Deny ACL not detected'
    $state.Deny = $false
    Apply-SingleACL $profile
    $grant = @($state.AclCalls | Where-Object { $_ -contains '/grant' })
    Assert ($grant.Count -gt 0) 'No ACL grants'
    Assert (($state.AclCalls | ForEach-Object { $_ } | Where-Object { $_ -eq "*${sid}:(OI)(CI)F" }).Count -gt 0) 'ACL did not use AD SID'
    $state.AclResult = 5
    Assert-Throws { Apply-SingleACL $profile } 'ACL error was reported as success'
    Assert ((Get-FileHash -LiteralPath $hivePath).Hash -eq $hiveHash) 'Profile checks changed existing hive'
    Write-Output 'PASS: native signature, new/existing profiles, SID/path conflicts, .bak, loaded/missing hives, API errors/races, cancellation, validation and ACL errors.'
} finally {
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notlike 'PM_Profiles_*') { throw 'Unsafe cleanup path' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
