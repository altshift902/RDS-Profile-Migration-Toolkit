$ErrorActionPreference = 'Stop'
$path = Join-Path (Split-Path $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$errors)
if ($errors) { throw ($errors -join ' ') }
# Import only file-copy and validation functions. Never mount a hive in these tests.
foreach ($name in @('Get-OutlookSourceProfiles','Select-OutlookSourceProfile','Copy-MissingOutlookPst','Assert-ProfileTargetPath','ConvertTo-LocalProfilePath','Test-OutlookMigrationPlan')) {
    $fn = $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    . ([scriptblock]::Create($fn.Extent.Text))
}
function Write-GuiLog ($Message,$Level) {}
function New-TestKey ($default,$names) {
    $key = [pscustomobject]@{ Default=$default; Names=$names; Children=@{} }
    $key | Add-Member ScriptMethod OpenSubKey { param($path) $this.Children[$path] }
    $key | Add-Member ScriptMethod GetValue { param($name) if ($name -eq 'DefaultProfile') { $this.Default } }
    $key | Add-Member ScriptMethod GetSubKeyNames { $this.Names }
    $key | Add-Member ScriptMethod Dispose {}
    return $key
}
$hive = New-TestKey
foreach ($version in @('15.0','16.0')) {
    $parent = New-TestKey 'Work'
    $parent.Children['Profiles'] = New-TestKey '' @('Archive','Work')
    $hive.Children["Software\Microsoft\Office\$version\Outlook"] = $parent
    $found = @(Get-OutlookSourceProfiles $hive | Where-Object Branch -eq "Office $version")
    if ($found.Count -ne 2 -or (Select-OutlookSourceProfile $found 'tester').Name -ne 'Work') { throw "Default profile discovery failed: $version" }
}
$legacy = New-TestKey
$legacy.Children['Profiles'] = New-TestKey 'Old' @('Old')
$hive.Children['Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem'] = $legacy
$found = @(Get-OutlookSourceProfiles $hive)
if ($found.Count -ne 5 -or @($found | Where-Object IsDefault).Count -ne 3) { throw 'Mixed Office installations not discovered' }
Add-Type -AssemblyName System.Windows.Forms
$root = Join-Path ([IO.Path]::GetTempPath()) ('PM_OutlookTest_' + [guid]::NewGuid().ToString('N'))
$source = Join-Path $root 'src'
$userFolder = Join-Path $root 'dst'
$dest = Join-Path $userFolder 'Documents'
try {
    foreach ($dir in @($source,$dest)) { [void][IO.Directory]::CreateDirectory($dir) }
    [IO.File]::WriteAllText("$source\mail.pst",'source')
    [IO.File]::WriteAllText("$source\archive.pst",'archive')
    [IO.File]::WriteAllText("$dest\mail.pst",'already in use')
    Copy-MissingOutlookPst $source $dest $userFolder
    Copy-MissingOutlookPst $source $dest $userFolder
    if ([IO.File]::ReadAllText("$dest\mail.pst") -ne 'already in use' -or [IO.File]::ReadAllText("$dest\archive.pst") -ne 'archive') { throw 'PST preservation failed' }
    $account = @{Path='9375CFF0413111d3B88A00104B2A6676\00000001';Values=@(@{Name='POP3 Server'})}
    $store = @{Path='store';Values=@(@{Name='001f6700';Data=[Text.Encoding]::Unicode.GetBytes("$dest\mail.pst" + [char]0)})}
    $Global:Config = @{TargetFolders=@('Documents')}
    $target = @{Path=$userFolder}
    if ((Test-OutlookMigrationPlan @($account,$store) $target $userFolder) -ne 1) { throw 'POP3/PST rejected' }
    $account.Values[0].Name = 'IMAP Server'
    if ((Test-OutlookMigrationPlan @($account,$store) $target $userFolder) -ne 1) { throw 'IMAP rejected' }
    $aux = @{Path='9375CFF0413111d3B88A00104B2A6676\00000002';Values=@(@{Name='Service Name';Data=''})}
    foreach ($field in @('Account Name','clsid','MAPI Provider','Mini UID','Preferences UID','Service UID')) {
        $aux.Values += @{Name=$field;Data=''}
    }
    $aux.Values[0].Data = 'OMSAB'
    if ((Test-OutlookMigrationPlan @($account,$store,$aux) $target $userFolder) -ne 1) { throw 'OMSAB blocked the IMAP profile' }
    foreach ($service in @('CONTAB','MSPST MS','MSUPST MS','OMSAB')) {
        $aux.Values[0].Data = [Text.Encoding]::Unicode.GetBytes($service + [char]0)
        if ((Test-OutlookMigrationPlan @($account,$store,$aux) $target $userFolder) -ne 1) { throw "Auxiliary service $service counted as account" }
    }
    $aux.Values[0].Data = 'MSEMS'
    $reason = ''
    try { Test-OutlookMigrationPlan @($account,$store,$aux) $target $userFolder | Out-Null } catch { $reason = $_.Exception.Message }
    if ($reason -notlike '*00000002*MSEMS*Service Name*') { throw 'Unknown/mixed profile was not identified' }
    $Global:Config.TargetFolders=@('Desktop')
    if ((Test-OutlookMigrationPlan @($account,$store) $target $userFolder) -ne 1) { throw 'Existing PST in unselected Documents rejected' }
    $cache = @{Path='imap';Values=@(@{Name='001f6700';Data="$dest\cache.ost"})}
    if ((Test-OutlookMigrationPlan @($account,$cache) $target $userFolder) -ne 1) { throw 'IMAP cache treated as required PST' }
    foreach ($case in @('missing','unsupported')) {
        if ($case -eq 'missing') { $store.Values[0].Data = "$dest\absent.pst" }
        if ($case -eq 'unsupported') { $Global:Config.TargetFolders=@('Documents'); $account.Values[0].Name='Other' }
        $rejected = $false
        try { Test-OutlookMigrationPlan @($account,$store) $target $userFolder | Out-Null } catch { $rejected=$true }
        if (-not $rejected) { throw "Unexpectedly accepted $case" }
    }
    'Outlook offline file and profile-plan tests passed.'
} finally {
    if ((Split-Path $root -Parent).TrimEnd('\') -eq [IO.Path]::GetTempPath().TrimEnd('\') -and (Split-Path $root -Leaf) -match '^PM_OutlookTest_[a-f0-9]{32}$') {
        [IO.Directory]::Delete($root,$true)
    }
}
