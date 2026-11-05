$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in @('ConvertTo-TargetFolders', 'Select-TargetFolders', 'Load-Config', 'Save-Config')) {
        $code = $node.Extent.Text.Replace('$dialog.ShowDialog($Owner)', '(Invoke-TestDialog $dialog)')
        . ([scriptblock]::Create($code))
    }
}
function Assert ($Ok, $Message) { if (-not $Ok) { throw $Message } }
function Invoke-TestDialog ($Window) {
    $Window.Show()
    [Windows.Forms.Application]::DoEvents()
    $list = $Window.Controls['FolderList']
    $custom = $Window.Controls['CustomFolder']
    $add = $Window.Controls['AddFolder']
    $apply = $Window.Controls['ApplyFolders']
    if ($script:scenario -eq 'cancel') {
        $list.SetItemChecked(0, $false)
        return [Windows.Forms.DialogResult]::Cancel
    }
    if ($script:scenario -eq 'reload') {
        Assert (@($list.CheckedItems).Count -eq 3) 'Saved checked items not restored'
        Assert ($list.CheckedItems -contains 'WorkData\Custom') 'Saved custom folder missing'
        $apply.PerformClick()
        return $Window.DialogResult
    }
    for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) }
    $apply.PerformClick()
    Assert ($Window.DialogResult -ne 'OK' -and $Window.Controls['ValidationError'].Text) 'Empty selection accepted'
    $custom.Text = '..\Windows'; $add.PerformClick()
    Assert ($Window.Controls['ValidationError'].Text -and $custom.Text) 'Invalid path accepted'
    $custom.Text = 'WorkData/Custom'; $add.PerformClick()
    $count = $list.Items.Count
    $custom.Text = 'workdata\custom'; $add.PerformClick()
    Assert ($list.Items.Count -eq $count) 'Duplicate folder added'
    $list.SetItemChecked($list.Items.IndexOf('Desktop'), $true)
    Assert ($list.Items -notcontains 'AppData\Roaming\Thunderbird') 'Thunderbird offered by default'
    $custom.Text = 'AppData\Roaming\Thunderbird'; $add.PerformClick()
    $custom.Text = 'Unsubmitted'
    $apply.PerformClick()
    Assert ($Window.DialogResult -ne 'OK') 'Unsubmitted custom input silently discarded'
    $custom.Clear()
    $bitmap = New-Object Drawing.Bitmap($Window.Width, $Window.Height)
    try {
        $Window.DrawToBitmap($bitmap, (New-Object Drawing.Rectangle(0, 0, $Window.Width, $Window.Height)))
        $bitmap.Save((Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\target-folders-preview.png'))
    } finally { $bitmap.Dispose() }
    $apply.PerformClick()
    return $Window.DialogResult
}
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('PM_FolderDialog_' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixture)
try {
    $Global:ConfigFile = Join-Path $fixture 'pm_config.json'
    $Global:Config = @{ SourceServer = 'fixture-server'; TargetOU = 'OU=Fixture,DC=example,DC=test'; ProfilesRoot = 'C:\Users'; InactiveDaysLimit = 90; TargetFolders = @('Desktop', 'Documents', 'Downloads'); LogsDir = (Join-Path $fixture 'logs') }
    $script:scenario = 'cancel'
    $cancelled = @(Select-TargetFolders $null $Global:Config.TargetFolders)
    Assert ($cancelled.Count -eq 0 -and $Global:Config.TargetFolders.Count -eq 3) 'Cancel changed configuration'
    $script:scenario = 'edit'
    $Global:Config.TargetFolders = @(Select-TargetFolders $null $Global:Config.TargetFolders)
    Assert ($Global:Config.TargetFolders.Count -eq 3 -and $Global:Config.TargetFolders -contains 'WorkData\Custom') 'Applied selection incorrect'
    Save-Config
    $Global:Config = @{ TargetOU = 'fixture'; TargetFolders = @('Documents') }
    Load-Config
    Assert ($Global:Config.SourceServer -eq 'fixture-server' -and $Global:Config.TargetOU -eq 'OU=Fixture,DC=example,DC=test' -and $Global:Config.ProfilesRoot -eq 'C:\Users' -and $Global:Config.InactiveDaysLimit -eq 90) 'Other settings lost on restart'
    $script:scenario = 'reload'
    $again = @(Select-TargetFolders $null $Global:Config.TargetFolders)
    Assert (($again -join '|') -ceq ($Global:Config.TargetFolders -join '|')) 'Reloaded selection changed'
    Write-Output 'PASS: dialog selection, custom folders, duplicate/invalid/empty rejection, cancel, apply, JSON persistence and reload. No registry or real config access.'
} finally {
    $resolved = [IO.Path]::GetFullPath($fixture)
    if (-not $resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notmatch '^PM_FolderDialog_[a-f0-9]{32}$') { throw 'Unsafe cleanup path' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
