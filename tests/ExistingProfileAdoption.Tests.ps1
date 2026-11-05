$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$names = @('ConvertTo-LocalProfilePath', 'Assert-ProfileLocation', 'Find-RegisteredProfile', 'Assert-ProfileOffline',
    'Assert-ProfileDirectory', 'Assert-ProfileFiles', 'Assert-ProfileTargetPath',
    'Assert-UnregisteredProfile', 'Copy-MissingProfileTemplate', 'Register-ExistingProfileDirectory', 'Get-UserProfileContext')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $names) { . ([scriptblock]::Create($node.Extent.Text)) }
}
function Assert ($Ok, $Message) { if (-not $Ok) { throw $Message } }
function Assert-Throws ([scriptblock]$Action, $Message) {
    $caught = $false
    try { & $Action | Out-Null } catch { $caught = $true }
    Assert $caught $Message
}
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('PM_Adopt_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $fixture -ItemType Directory)
$script:caseIndex = 0
$sid = 'S-1-5-21-111-222-333-1001'
$otherSid = 'S-1-5-21-111-222-333-1002'
$identity = [pscustomobject]@{ UserName = 'tester'; Sid = $sid }

# No real AD, ProfileList, profile creation, or user ACL writes are performed.
function Resolve-ProfileIdentity ($UserName) { Assert ($UserName -eq 'tester') 'Wrong user'; return $identity }
function Get-ProfileRegistrySnapshot { return $script:state.Snapshot }
function Get-CimInstance {
    param($ClassName, $Filter, $ErrorAction)
    Assert ($ClassName -eq 'Win32_UserProfile') 'Unexpected CIM query'
    if ($script:state.CimFailure) { throw 'CIM unavailable' }
    if ($Filter) { return @($script:state.Cim | Where-Object { $Filter -eq "SID='$($_.SID)'" }) }
    return $script:state.Cim
}
function Write-GuiLog ($Message, $Type) {
    if ($script:state.CancelOnLog) { $Global:CancelRequested = $true }
}
function Apply-SingleACL ($Profile) {
    $script:state.AclCalls++
    Assert-ProfileFiles $Profile
    if ($script:state.AclFailure) { throw 'ACL failure' }
    if ($script:state.Race) { $script:state.Snapshot.Entries += [pscustomobject]@{ Sid = $sid; Path = $script:profilePath } }
}
function Write-NewProfileRegistration ($Profile) {
    $script:state.RegisterCalls++
    Assert ($Profile.Sid -eq $sid -and $Profile.Path -eq $script:profilePath) 'Registration moved the profile'
    Assert ($script:state.AclCalls -gt 0) 'Registration happened before ACL preparation'
    Assert-Throws { [IO.File]::OpenWrite((Join-Path $Profile.Path 'NTUSER.DAT')).Dispose() } 'Hive is not protected during publication'
    if ($script:state.RegisterFailure) { throw 'Registry write failed' }
    $script:state.Snapshot.Entries += [pscustomobject]@{ Sid = $Profile.Sid; Path = $Profile.Path }
}
function Ensure-ProfileFullProfile ($Profile) { $script:state.EnsureCalls++ }
function Invoke-CreateWindowsProfile ($Identity) {
    $script:state.CreateCalls++
    throw 'CreateProfile must not run for an existing directory'
}
function Put-File ($Path, $Text) {
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $Path))
    [IO.File]::WriteAllText($Path, $Text)
}
function Reset-Scenario {
    $script:caseIndex++
    $root = Join-Path $fixture ("case$script:caseIndex\Users")
    $script:profilePath = Join-Path $root 'tester'
    $script:defaultPath = Join-Path $root 'Default'
    Put-File (Join-Path $script:profilePath 'Documents\keep.txt') 'user-only-data'
    Put-File (Join-Path $script:profilePath 'AppData\Roaming\Vendor\settings.ini') 'user-settings'
    Put-File (Join-Path $script:defaultPath 'NTUSER.DAT') 'default-target-hive-fixture'
    Put-File (Join-Path $script:defaultPath 'NTUSER.DAT.LOG1') 'default-target-log-fixture'
    Put-File (Join-Path $script:defaultPath 'AppData\Roaming\Vendor\settings.ini') 'default-settings'
    Put-File (Join-Path $script:defaultPath 'AppData\Roaming\Vendor\new.ini') 'new-default-setting'
    Put-File (Join-Path $script:defaultPath 'AppData\Local\Microsoft\Windows\UsrClass.dat') 'default-class-hive'
    Put-File (Join-Path $script:defaultPath 'AppData\Local\Microsoft\Windows\UsrClass.dat.LOG1') 'default-class-log'
    $script:state = @{
        Snapshot = [pscustomobject]@{ Root = $root; DefaultPath = $script:defaultPath; Entries = @() }
        Cim = @(); CimFailure = $false; CreateCalls = 0; RegisterCalls = 0; AclCalls = 0; EnsureCalls = 0
        AclFailure = $false; RegisterFailure = $false; Race = $false; CancelOnLog = $false
    }
    $script:txtDst = [pscustomobject]@{ Text = $root }
    $Global:CancelRequested = $false
}
function Assert-UserData {
    Assert ([IO.File]::ReadAllText((Join-Path $script:profilePath 'Documents\keep.txt')) -ceq 'user-only-data') 'User-only file changed'
    Assert ([IO.File]::ReadAllText((Join-Path $script:profilePath 'AppData\Roaming\Vendor\settings.ini')) -ceq 'user-settings') 'User settings overwritten by Default'
    Assert ($script:state.CreateCalls -eq 0) 'Unexpected native profile creation'
}

try {
    Reset-Scenario
    $profile = Get-UserProfileContext 'tester' -Create -AdoptExisting
    Assert ($profile.Path -eq $script:profilePath -and $state.RegisterCalls -eq 1) 'Folder not adopted in place'
    Assert-UserData
    Assert ((Get-FileHash (Join-Path $script:profilePath 'NTUSER.DAT')).Hash -eq (Get-FileHash (Join-Path $script:defaultPath 'NTUSER.DAT')).Hash) 'Missing hive did not come from local Default'
    Assert ((Get-ChildItem -LiteralPath $state.Snapshot.Root -Directory).Count -eq 2) 'A duplicate profile directory was created'
    $hiveHash = (Get-FileHash (Join-Path $script:profilePath 'NTUSER.DAT')).Hash
    [void](Get-UserProfileContext 'tester' -Create -AdoptExisting)
    Assert ($state.RegisterCalls -eq 1 -and (Get-FileHash (Join-Path $script:profilePath 'NTUSER.DAT')).Hash -eq $hiveHash) 'Repeat changed registration or hive'
    Assert ($state.EnsureCalls -eq 1) 'Repeat step 2 did not check FullProfile in the existing registration'

    Reset-Scenario
    Put-File (Join-Path $script:profilePath 'NTUSER.DAT') 'existing-user-hive'
    Put-File (Join-Path $script:profilePath 'NTUSER.DAT.LOG2') 'existing-user-log'
    Put-File (Join-Path $script:profilePath 'AppData\Local\Microsoft\Windows\UsrClass.dat') 'existing-class-hive'
    $hiveHash = (Get-FileHash (Join-Path $script:profilePath 'NTUSER.DAT')).Hash
    [void](Get-UserProfileContext 'tester' -Create -AdoptExisting)
    Assert-UserData
    Assert ((Get-FileHash (Join-Path $script:profilePath 'NTUSER.DAT')).Hash -eq $hiveHash) 'Existing NTUSER.DAT changed'
    Assert (-not (Test-Path -LiteralPath (Join-Path $script:profilePath 'NTUSER.DAT.LOG1'))) 'Default logs mixed with existing NTUSER.DAT'
    Assert (-not (Test-Path -LiteralPath (Join-Path $script:profilePath 'AppData\Local\Microsoft\Windows\UsrClass.dat.LOG1'))) 'Default logs mixed with existing UsrClass.dat'

    foreach ($kind in @('otherSid', 'bak', 'pending', 'loaded', 'foreignCim', 'cimFailure', 'emptyHive', 'orphanLogs', 'missingDefault', 'cancel', 'cancelOnLog', 'aclFailure', 'registerFailure', 'race')) {
        Reset-Scenario
        switch ($kind) {
            'otherSid' { $state.Snapshot.Entries += [pscustomobject]@{ Sid = $otherSid; Path = $script:profilePath } }
            'bak' { $state.Snapshot.Entries += [pscustomobject]@{ Sid = "$sid.bak"; Path = $script:profilePath } }
            'pending' { $state.Snapshot.Entries += [pscustomobject]@{ Sid = $sid; Path = $script:profilePath; PendingRegistration = 'incomplete' } }
            'loaded' { $state.Cim = @([pscustomobject]@{ SID = $sid; LocalPath = $script:profilePath; Loaded = $true }) }
            'foreignCim' { $state.Cim = @([pscustomobject]@{ SID = $otherSid; LocalPath = $script:profilePath; Loaded = $false }) }
            'cimFailure' { $state.CimFailure = $true }
            'emptyHive' { Put-File (Join-Path $script:profilePath 'NTUSER.DAT') '' }
            'orphanLogs' { Put-File (Join-Path $script:profilePath 'NTUSER.DAT.LOG1') 'unmatched-old-log' }
            'missingDefault' { $state.Snapshot.DefaultPath = Join-Path $fixture 'absent-default' }
            'cancel' { $Global:CancelRequested = $true }
            'cancelOnLog' { $state.CancelOnLog = $true }
            'aclFailure' { $state.AclFailure = $true }
            'registerFailure' { $state.RegisterFailure = $true }
            'race' { $state.Race = $true }
        }
        $originalEntries = @($state.Snapshot.Entries).Count
        Assert-Throws { Get-UserProfileContext 'tester' -Create -AdoptExisting } "Unsafe adoption allowed: $kind"
        Assert-UserData
        if ($kind -ne 'race') { Assert ($state.Snapshot.Entries.Count -eq $originalEntries) "Failed adoption published a registration: $kind" }
        if ($kind -in @('aclFailure', 'registerFailure')) {
            $state.AclFailure = $false; $state.RegisterFailure = $false
            [void](Get-UserProfileContext 'tester' -Create -AdoptExisting)
            Assert-UserData
            Assert ($state.Snapshot.Entries.Count -eq 1) 'Retry did not recover safely'
        }
    }

    Reset-Scenario
    Assert-Throws { Get-UserProfileContext 'tester' } 'Read-only operation adopted a folder'
    Assert-Throws { Get-UserProfileContext 'tester' -Create } 'Delta silently created a duplicate or adopted a folder'
    Assert ($state.RegisterCalls -eq 0 -and $state.CreateCalls -eq 0) 'Non-step-2 operation modified registration'

    # Simulate reparse metadata instead of creating OS junctions on the host.
    function Get-ChildItem {
        param($LiteralPath, [switch]$Force, [switch]$File, $ErrorAction)
        Microsoft.PowerShell.Management\Get-ChildItem @PSBoundParameters
        if ($LiteralPath -eq $script:defaultPath -and -not $File) {
            [pscustomobject]@{ Name = 'template-junction'; FullName = (Join-Path $script:defaultPath 'template-junction');
                PSIsContainer = $true; Attributes = [IO.FileAttributes]::ReparsePoint }
        }
    }
    Reset-Scenario
    [void](Get-UserProfileContext 'tester' -Create -AdoptExisting)
    Assert (-not (Test-Path -LiteralPath (Join-Path $script:profilePath 'template-junction'))) 'Default junction copied'
    Assert-UserData

    Reset-Scenario
    $alias = Join-Path $state.Snapshot.Root 'alias'
    function Get-Item {
        param($LiteralPath, [switch]$Force, $ErrorAction)
        if ($LiteralPath -eq $script:alias) {
            return [pscustomobject]@{ FullName = $LiteralPath; PSIsContainer = $true; Attributes = [IO.FileAttributes]::ReparsePoint }
        }
        Microsoft.PowerShell.Management\Get-Item @PSBoundParameters
    }
    Assert-Throws { Register-ExistingProfileDirectory $identity $alias $state.Snapshot } 'Reparse-point target adopted'
    Assert ($state.RegisterCalls -eq 0) 'Reparse target registered'

    Write-Output 'PASS: in-place adoption, preserved user files/hives, no duplicates, Default template, idempotence, conflicts, cancellation, failure/retry, reparse points. Registry, AD and ACL writes are mocked.'
} finally {
    $resolved = [IO.Path]::GetFullPath($fixture)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notmatch '^PM_Adopt_[0-9a-f]{32}$') { throw 'Unsafe test directory cleanup' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
