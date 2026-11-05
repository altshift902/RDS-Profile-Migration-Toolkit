$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$names = @('Get-ShortcutTarget', 'Get-BrokenUserShortcuts', 'Invoke-ConfirmedShortcutChanges','Find-SystemYandexBrowser', 'Assert-ProfileTargetPath')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $names) { . ([scriptblock]::Create($node.Extent.Text)) }
}

# Execute the real click handlers. Only the blocking empty-selection message is replaced.
Add-Type 'public static class SelectionNotice { public static int Count; public static void Show(string a, string b, string c, string d) { Count++; } }'
$handlers = @{}
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.InvokeMemberExpressionAst] }, $true)) {
    if ($node.Member.Value -eq 'Add_Click' -and $node.Expression -is [Management.Automation.Language.VariableExpressionAst]) {
        $button = $node.Expression.VariablePath.UserPath
        if ($button -in @('btnValidate', 'btnShortcuts')) {
            $body = $node.Arguments[0].ScriptBlock.Extent.Text
            $body = $body.Substring(1, $body.Length - 2).Replace('[System.Windows.Forms.MessageBox]', '[SelectionNotice]')
            $handlers[$button] = [scriptblock]::Create($body)
        }
    }
}
function Assert ($Condition, $Message) { if (-not $Condition) { throw $Message } }
Assert ($handlers.Count -eq 2) 'Could not locate both handlers'
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('PM_Selected_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $fixtureRoot -ItemType Directory)
$gridUsers = New-Object Windows.Forms.DataGridView
$gridUsers.AllowUserToAddRows = $false
[void]$gridUsers.Columns.Add((New-Object Windows.Forms.DataGridViewCheckBoxColumn))
1..6 | ForEach-Object { [void]$gridUsers.Columns.Add("c$_", "c$_") }
$shell = New-Object -ComObject WScript.Shell
$script:profiles = @{}
$script:links = @{}
$Global:pbar = [pscustomobject]@{ Value = 0 }

# Domain, profile state, log and UI boundaries are isolated; only temporary .lnk files are real.
function Get-UserProfileContext ($UserName) {
    $script:contextCalls += $UserName
    if ($UserName -eq 'unchecked') { throw 'An unchecked user was accessed' }
    return $script:profiles[$UserName]
}
function Assert-ProfileOffline ($Profile, [switch]$AllowLoaded) {
    if ($script:loaded -and -not $AllowLoaded) { throw 'Profile is now loaded' }
}
function Get-ProfileReadiness ($UserName) {
    $script:validated += $UserName
    return [pscustomobject]@{ Ready = $true; Issues = @(); Profile = $script:profiles[$UserName] }
}
function Write-GuiLog ($Message, $Type) { $script:logs += $Message }
function Update-Progress ($Value, $Status) {
    $Global:pbar.Value = $Value
    $script:progressCalls++
    if ($script:cancelAtProgress -eq $script:progressCalls) { $Global:CancelRequested = $true }
}
function Set-UiBusy ($Busy, $Text) {
    $script:busy = $Busy
    $script:busyCalls++
    if ($Busy) { $Global:CancelRequested = $false }
}
function Show-ShortcutCleanupConfirmation ($Shortcuts, [switch]$Repair) {
    $script:confirmCalls++
    $script:preview = @($Shortcuts)
    foreach ($item in $Shortcuts) {
        Assert ([IO.File]::Exists($item.ShortcutPath)) 'Shortcut was deleted before confirmation'
        Assert ($item.UserName -in @('checkedA', 'checkedB')) 'Unchecked user in preview'
        Assert (-not [string]::IsNullOrEmpty($item.TargetPath)) 'Missing preview target'
    }
    foreach ($item in $Shortcuts) { if (-not $item.NewTarget) { $item.Action = 'Удалить' } }
    if ($script:onConfirm) { & $script:onConfirm }
    return $script:approve
}
function New-TestShortcut ($Path, $Target) {
    $sc = $shell.CreateShortcut($Path)
    try { $sc.TargetPath = $Target; $sc.Save() }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sc) }
}
function Reset-Scenario {
    $script:contextCalls = @(); $script:validated = @(); $script:logs = @(); $script:preview = @()
    $script:confirmCalls = 0; $script:busyCalls = 0; $script:progressCalls = 0
    $script:cancelAtProgress = -1; $script:approve = $false; $script:loaded = $false
    $script:onConfirm = $null; $script:busy = $false; $Global:CancelRequested = $false
    $gridUsers.Rows.Clear()
    foreach ($user in @('checkedA', 'unchecked', 'checkedB')) {
        $path = Join-Path $fixtureRoot $user
        $desk = Join-Path $path 'Desktop\nested'
        [void](New-Item -Path $desk -ItemType Directory -Force)
        $script:profiles[$user] = [pscustomobject]@{ UserName = $user; Sid = "sid-$user"; Path = $path }
        $script:links[$user] = Join-Path $desk 'broken [1].lnk'
        New-TestShortcut $script:links[$user] (Join-Path $fixtureRoot "$user-missing.exe")
        New-TestShortcut (Join-Path $desk 'working.lnk') $fixtureRoot
        [void]$gridUsers.Rows.Add(($user -ne 'unchecked'), $user, '', '', '', 'unchanged', 'unchanged')
    }
    # A highlighted row is not necessarily checked.
    $gridUsers.ClearSelection()
    $gridUsers.Rows[1].Selected = $true
}

try {
    Reset-Scenario
    & $handlers.btnValidate
    Assert (($script:validated -join ',') -eq 'checkedA,checkedB') 'Validation ignored checkbox selection'
    Assert ($gridUsers.Rows[1].Cells[5].Value -eq 'unchanged' -and $gridUsers.Rows[1].Cells[6].Value -eq 'unchanged') 'Unchecked row was modified'
    Assert (-not $script:busy) 'Validation left UI busy'

    Reset-Scenario
    foreach ($row in $gridUsers.Rows) { $row.Cells[0].Value = $false }
    & $handlers.btnValidate
    & $handlers.btnShortcuts
    Assert ([SelectionNotice]::Count -eq 2) 'Empty selection should display a notice for each command'
    Assert ($script:busyCalls -eq 0 -and $script:contextCalls.Count -eq 0 -and $script:validated.Count -eq 0) 'Work started for empty selection'

    Reset-Scenario
    & $handlers.btnShortcuts
    Assert ($script:confirmCalls -eq 1 -and $script:preview.Count -eq 2) 'Wrong preview after scan'
    Assert (($script:contextCalls -join ',') -eq 'checkedA,checkedB') 'Cleanup scanned an unchecked profile'
    foreach ($path in $script:links.Values) { Assert ([IO.File]::Exists($path)) 'Cancel deleted a shortcut' }
    Assert (-not $script:busy) 'Cancel left UI busy'

    Reset-Scenario
    $script:approve = $true
    & $handlers.btnShortcuts
    Assert (-not [IO.File]::Exists($script:links.checkedA) -and -not [IO.File]::Exists($script:links.checkedB)) 'Approved broken shortcuts were not removed'
    Assert ([IO.File]::Exists($script:links.unchecked)) 'Unchecked shortcut was removed'
    foreach ($user in @('checkedA', 'checkedB')) {
        Assert ([IO.File]::Exists((Join-Path $script:profiles[$user].Path 'Desktop\nested\working.lnk'))) 'Working shortcut was removed'
    }
    Assert (-not $script:busy) 'Successful cleanup left UI busy'

    Reset-Scenario
    $script:cancelAtProgress = 3 # Cancel after scanning the first selected user.
    $script:approve = $true
    & $handlers.btnShortcuts
    Assert ($Global:CancelRequested -and $script:confirmCalls -eq 0) 'Interrupted scan displayed a partial deletion list'
    foreach ($path in $script:links.Values) { Assert ([IO.File]::Exists($path)) 'Interrupted scan deleted a shortcut' }

    Reset-Scenario
    $script:approve = $true
    $script:onConfirm = {
        Set-Content -LiteralPath (Join-Path $fixtureRoot 'checkedA-missing.exe') -Value 'now available'
        New-TestShortcut $script:links.checkedB (Join-Path $fixtureRoot 'changed-target.exe')
    }
    & $handlers.btnShortcuts
    Assert ($script:confirmCalls -eq 1) 'Preview was not shown'
    Assert ([IO.File]::Exists($script:links.checkedA)) 'Shortcut to a restored target was removed'
    Assert ([IO.File]::Exists($script:links.checkedB)) 'Changed shortcut was removed'
    Remove-Item -LiteralPath (Join-Path $fixtureRoot 'checkedA-missing.exe')

    Reset-Scenario
    $script:approve = $true
    $script:loaded = $true
    & $handlers.btnShortcuts
    Assert ($script:preview.Count -eq 2) 'Loaded profiles were not scanned'
    Assert (-not [IO.File]::Exists($script:links.checkedA) -and -not [IO.File]::Exists($script:links.checkedB)) 'Confirmed cleanup failed for loaded profiles'

    Reset-Scenario
    $script:approve = $true
    $script:onConfirm = { $script:cancelAtProgress = $script:progressCalls + 1 }
    & $handlers.btnShortcuts
    Assert ($Global:CancelRequested) 'Deletion cancellation did not trigger'
    foreach ($path in $script:links.Values) { Assert ([IO.File]::Exists($path)) 'Cancellation before first deletion was ignored' }

    Reset-Scenario
    foreach ($user in @('checkedA', 'checkedB')) { New-TestShortcut $script:links[$user] $fixtureRoot }
    & $handlers.btnShortcuts
    Assert ($script:confirmCalls -eq 0 -and -not $script:busy) 'Empty candidate list prompted or left UI busy'
    Reset-Scenario
    $oldPF = $env:ProgramW6432
    try {
        $env:ProgramW6432 = Join-Path $fixtureRoot 'Program Files'
        $systemExe = Join-Path $env:ProgramW6432 'Yandex\YandexBrowser\Application\browser.exe'
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($systemExe))
        [IO.File]::WriteAllText($systemExe,'fixture')
        foreach ($user in @('checkedA','checkedB','unchecked')) {
            $oldExe = Join-Path $script:profiles[$user].Path 'AppData\Local\Yandex\YandexBrowser\Application\browser.exe'
            if ($user -eq 'checkedA') {
                [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($oldExe))
                [IO.File]::WriteAllText($oldExe,'fixture')
            }
            $lnk = $shell.CreateShortcut($script:links[$user])
            $lnk.TargetPath = $oldExe
            $lnk.Arguments = '--profile-directory="Profile 2" https://example.org'
            $lnk.Save()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($lnk)
        }
        & $handlers.btnShortcuts
        Assert ($script:preview.Count -eq 2 -and $script:preview[0].NewTarget -eq $systemExe) 'Repair preview missing'
        Assert ((Get-ShortcutTarget $shell $script:links.checkedA) -ne $systemExe) 'Cancellation changed shortcut'
        $script:approve = $true
        & $handlers.btnShortcuts
        foreach ($user in @('checkedA','checkedB')) {
            $lnk = $shell.CreateShortcut($script:links[$user])
            Assert ($lnk.TargetPath -eq $systemExe -and $lnk.Arguments -ceq '--profile-directory="Profile 2" https://example.org') 'Repair lost target or arguments'
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($lnk)
        }
        Assert ((Get-ShortcutTarget $shell $script:links.unchecked) -ne $systemExe) 'Unchecked browser link modified'
    } finally { $env:ProgramW6432 = $oldPF }
    Reset-Scenario
    $relativeExe = 'AppData\Local\SkbKontur\Kontur.Admin.Abonent.exe'
    $oldExe = Join-Path (Join-Path $fixtureRoot 'Support') $relativeExe
    $ownerExe = Join-Path $script:profiles.checkedA.Path $relativeExe
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($ownerExe))
    [IO.File]::WriteAllText($ownerExe,'fixture')
    New-TestShortcut $script:links.checkedA $oldExe
    New-TestShortcut $script:links.checkedB $fixtureRoot
    $script:approve = $true
    & $handlers.btnShortcuts
    Assert ((Get-ShortcutTarget $shell $script:links.checkedA) -eq $ownerExe) 'Shortcut was not mapped to its owner'
    Reset-Scenario
    $script:approve = $true
    $script:onConfirm = { $a = $script:preview | Where-Object UserName -eq checkedA; $a.NewTarget = $ownerExe; $a.Action = 'Заменить'; ($script:preview | Where-Object UserName -eq checkedB).Action = 'Пропустить' }
    & $handlers.btnShortcuts
    Assert ((Get-ShortcutTarget $shell $script:links.checkedA) -eq $ownerExe) 'Manual replacement failed'
    Assert ((Get-ShortcutTarget $shell $script:links.checkedB) -eq (Join-Path $fixtureRoot 'checkedB-missing.exe')) 'Manual repair affected another row'
    Reset-Scenario
    New-TestShortcut $script:links.checkedA $oldExe
    $script:approve = $true
    & $handlers.btnShortcuts
    Assert ($script:confirmCalls -eq 1 -and $script:preview.Count -eq 2) 'Mixed candidates must share one preview'
    Assert ((Get-ShortcutTarget $shell $script:links.checkedA) -eq $ownerExe -and -not [IO.File]::Exists($script:links.checkedB)) 'Mixed repair and deletion failed'
    Write-Output 'PASS: selected-only validation and cleanup; empty selection; preview before deletion; cancellation; changed/restored targets; loaded profiles; literal paths.'
} finally {
    $gridUsers.Dispose()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'PM_Selected_*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
    }
}
