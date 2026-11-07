<#
.SYNOPSIS
    ProfileMigrator GUI - Оконное графическое приложение (WPF/WinForms) для миграции профилей RDS.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:AppVersion = '1.2.0-dev.1'

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$Global:ConfigFile = if ($env:PROFILEMIGRATOR_CONFIG_PATH) {
    $env:PROFILEMIGRATOR_CONFIG_PATH
} else {
    Join-Path $PSScriptRoot 'pm_config.json'
}
$Global:Config = @{
    SourceServer      = "SRV-SOURCE"
    TargetOU          = ""
    ProfilesRoot      = "C:\Users"
    WorkDir           = "C:\Migrate"
    ActiveUsersFile   = "C:\Migrate\ActiveUsersList.txt"
    LogsDir           = "C:\Migrate\Logs"
    TargetFolders     = @("Desktop", "Documents", "Downloads")
    InactiveDaysLimit = 90
}

function ConvertTo-TargetFolders ($Folders) {
    if ($Folders -isnot [array] -or $Folders.Count -eq 0) {
        throw 'TargetFolders должен быть непустым JSON-массивом относительных путей.'
    }
    $result = @()
    foreach ($folder in $Folders) {
        if ($folder -isnot [string] -or [string]::IsNullOrWhiteSpace($folder)) {
            throw 'TargetFolders содержит пустой путь или значение, отличное от строки.'
        }
        $relative = $folder.Replace('/', '\')
        if ([IO.Path]::IsPathRooted($relative) -or $relative -match '[:*?"<>|\[\]\x00-\x1f]') {
            throw "TargetFolders: укажите путь внутри профиля без диска, UNC и масок: $folder"
        }
        foreach ($part in $relative.Split('\')) {
            if ([string]::IsNullOrWhiteSpace($part) -or $part -in @('.', '..') -or $part -match '[ .]$') {
                throw "TargetFolders: недопустимый сегмент пути: $folder"
            }
        }
        if ($result -notcontains $relative) { $result += $relative }
    }
    return $result
}

function Select-TargetFolders ($Owner, $Folders) {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Папки переноса (TargetFolders)'
    if ($Owner -and $Owner.Icon) { $dialog.Icon = $Owner.Icon }
    $dialog.ClientSize = New-Object Drawing.Size(650, 490)
    $dialog.MinimumSize = New-Object Drawing.Size(666, 529)
    $dialog.StartPosition = 'CenterParent'
    $dialog.Font = New-Object Drawing.Font('Segoe UI', 9)
    $dialog.BackColor = [Drawing.Color]::FromArgb(245, 247, 250)
    $dialog.MinimizeBox = $false
    $dialog.ShowInTaskbar = $false
    $hint = New-Object Windows.Forms.Label
    $hint.Text = "Отметьте папки относительно профиля пользователя.`r`nПосле применения нажмите «Сохранить параметры» в основном окне."
    $hint.SetBounds(16, 12, 618, 44)
    $hint.Anchor = 'Top, Left, Right'
    $dialog.Controls.Add($hint)
    $list = New-Object Windows.Forms.CheckedListBox
    $list.Name = 'FolderList'
    $list.CheckOnClick = $true
    $list.HorizontalScrollbar = $true
    $list.SetBounds(16, 62, 618, 260)
    $list.Anchor = 'Top, Bottom, Left, Right'
    $dialog.Controls.Add($list)
    $choices = @(
        'Desktop',
        'Documents',
        'Downloads',
        'AppData\Local\Microsoft\Outlook',
        'AppData\Local\Crypto Pro',
        'AppData\Local\Google',
        'AppData\Local\Yandex\YandexBrowser\User Data',
        'Favorites',
        'Pictures',
        'Music',
        'Videos'
    )
    foreach ($folder in $Folders) { if ($choices -notcontains $folder) { $choices += $folder } }
    foreach ($folder in $choices) { [void]$list.Items.Add($folder, ($Folders -contains $folder)) }
    $customLabel = New-Object Windows.Forms.Label
    $customLabel.Text = 'Другая папка, например AppData\Roaming\Vendor\Product:'
    $customLabel.SetBounds(16, 333, 618, 22)
    $customLabel.Anchor = 'Bottom, Left, Right'
    $dialog.Controls.Add($customLabel)
    $custom = New-Object Windows.Forms.TextBox
    $custom.Name = 'CustomFolder'
    $custom.SetBounds(16, 359, 498, 25)
    $custom.Anchor = 'Bottom, Left, Right'
    $dialog.Controls.Add($custom)
    $add = New-Object Windows.Forms.Button
    $add.Name = 'AddFolder'
    $add.Text = 'Добавить'
    $add.SetBounds(524, 357, 110, 28)
    $add.Anchor = 'Bottom, Right'
    $dialog.Controls.Add($add)
    $errorLabel = New-Object Windows.Forms.Label
    $errorLabel.Name = 'ValidationError'
    $errorLabel.ForeColor = [Drawing.Color]::Firebrick
    $errorLabel.SetBounds(16, 394, 618, 42)
    $errorLabel.Anchor = 'Bottom, Left, Right'
    $dialog.Controls.Add($errorLabel)
    $apply = New-Object Windows.Forms.Button
    $apply.Name = 'ApplyFolders'
    $apply.Text = 'Применить'
    $apply.SetBounds(394, 446, 115, 30)
    $apply.Anchor = 'Bottom, Right'
    $apply.BackColor = [Drawing.Color]::FromArgb(24, 43, 73)
    $apply.ForeColor = [Drawing.Color]::White
    $apply.FlatStyle = 'Flat'
    $dialog.Controls.Add($apply)
    $cancel = New-Object Windows.Forms.Button
    $cancel.Text = 'Отмена'
    $cancel.SetBounds(519, 446, 115, 30)
    $cancel.Anchor = 'Bottom, Right'
    $cancel.DialogResult = [Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($cancel)
    $dialog.CancelButton = $cancel
    $dialog.AcceptButton = $apply
    $add.Add_Click({
        try {
            $folder = @(ConvertTo-TargetFolders @($custom.Text.Trim()))[0]
            $index = -1
            for ($i = 0; $i -lt $list.Items.Count; $i++) { if ($list.Items[$i] -eq $folder) { $index = $i; break } }
            if ($index -lt 0) { $index = $list.Items.Add($folder) }
            $list.SetItemChecked($index, $true)
            $custom.Clear()
            $errorLabel.Text = ''
        } catch { $errorLabel.Text = $_.Exception.Message }
    })
    $apply.Add_Click({
        try {
            if (-not [string]::IsNullOrWhiteSpace($custom.Text)) { throw 'Нажмите «Добавить» для введенной папки или очистите поле.' }
            $dialog.Tag = @(ConvertTo-TargetFolders @($list.CheckedItems | ForEach-Object { [string]$_ }))
            $dialog.DialogResult = [Windows.Forms.DialogResult]::OK
        } catch { $errorLabel.Text = $_.Exception.Message }
    })
    try {
        if ($dialog.ShowDialog($Owner) -eq [Windows.Forms.DialogResult]::OK) { return $dialog.Tag }
    } finally { $dialog.Dispose() }
}

function Get-MissingTargetFolders ([string]$ProfilePath) {
    foreach ($folder in $Global:Config.TargetFolders) {
        if (-not (Test-Path -LiteralPath (Join-Path $ProfilePath $folder) -PathType Container)) {
            $folder
        }
    }
}

function Load-Config {
    $Global:Config.TargetFolders = @('Desktop', 'Documents', 'Downloads')
    $readPath = $Global:ConfigFile
    if (Test-Path -LiteralPath $readPath) {
        try {
            $Loaded = Get-Content -LiteralPath $readPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($null -eq $Loaded -or $Loaded -isnot [pscustomobject]) { throw 'Ожидается JSON-объект настроек.' }
            if ($Loaded.PSObject.Properties.Name -contains 'TargetFolders') {
                $Global:Config.TargetFolders = @(ConvertTo-TargetFolders $Loaded.TargetFolders)
            }
            if ($Loaded.SourceServer)      { $Global:Config.SourceServer = $Loaded.SourceServer }
            if ($Loaded.TargetOU)          { $Global:Config.TargetOU = $Loaded.TargetOU }
            if ($Loaded.ProfilesRoot)      { $Global:Config.ProfilesRoot = $Loaded.ProfilesRoot }
            if ($Loaded.WorkDir)           { $Global:Config.WorkDir = $Loaded.WorkDir }
            if ($Loaded.ActiveUsersFile)   { $Global:Config.ActiveUsersFile = $Loaded.ActiveUsersFile }
            if ($Loaded.LogsDir)           { $Global:Config.LogsDir = $Loaded.LogsDir }
            if ($Loaded.InactiveDaysLimit) { $Global:Config.InactiveDaysLimit = [int]$Loaded.InactiveDaysLimit }
        } catch { throw "Ошибка конфигурации '$readPath': $($_.Exception.Message)" }
    }
    if ([string]::IsNullOrWhiteSpace($Global:Config.TargetOU)) {
        try {
            $DomainObj = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $Global:Config.TargetOU = ($DomainObj.Name.Split('.') | ForEach-Object { "DC=$_" }) -join ','
        } catch {
            $Global:Config.TargetOU = "DC=domain,DC=local"
        }
    }
    if (-not (Test-Path -LiteralPath $Global:ConfigFile)) {
        Save-Config
    }
}

function Save-Config {
    try {
        $Global:Config.TargetFolders = @(ConvertTo-TargetFolders $Global:Config.TargetFolders)
        $configDir = Split-Path -Parent $Global:ConfigFile
        if (-not (Test-Path -LiteralPath $configDir)) {
            New-Item -Path $configDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        $Global:Config | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Global:ConfigFile -Encoding UTF8 -ErrorAction Stop
    } catch { throw "Не удалось сохранить настройки: $($_.Exception.Message)" }
}

Load-Config

# --- Создание графического окна ---
if ($env:PROFILEMIGRATOR_HOST_EXE) {
    # Give the hosted GUI its own taskbar group before creating any window.
    Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
namespace ProfileMigrator {
    public static class Taskbar {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern int SetCurrentProcessExplicitAppUserModelID(string appId);
    }
}
'@
    [void][ProfileMigrator.Taskbar]::SetCurrentProcessExplicitAppUserModelID('ProfileMigrator.RDS.Toolkit')
}
[System.Windows.Forms.Application]::EnableVisualStyles()
$form = New-Object System.Windows.Forms.Form
$form.Text = "RDS Profile Migration Toolkit — версия $script:AppVersion"
$form.Size = New-Object System.Drawing.Size(1040, 750)
$form.MinimumSize = New-Object System.Drawing.Size(1040, 750)
$form.ShowInTaskbar = $true
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.AutoScroll = $false

# Встроенная иконка приложения
try {
    $icoBase64 = "AAABAAEAMDAQAAAAAADoBwAAFgAAACgAAAAwAAAAYAAAAAEABAAAAAAAAAYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAACAAAAAgIAAgAAAAIAAgACAgAAAgICAAMDAwAAAAP8AAP8AAAD//wD/AAAA/wD/AP//AAD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHeIiIiIiIiIiIiIiIiIiIdwAAAAAAAAZ2AAAAAAAAAAAAAAAAAAAABHYAAAAAAGcAAAAAAAAAAAAAAAAAAAAAAAZgAAAAAHAAAAAAAAAAAAAAAAAAAAAAAABwAAAABmAAAAAAAAAAAAAAAAAAAAAAAABmAAAAB0AAAAAAAAAAAAAAAAAAAAAAAABHAAAACEAAAAAAAAAAAAAAAAAA9wAAAAAIAAAACAAAAAAAAAAAAAAAAAAA+HAAAAAIAAAACERERnd3d3d3d3RERERE9ohERERIAAAACERERoiIiIiIiHRERERE9miERERIAAAACERERIiIiIiIiET/////9maIRERIAAAACERERHiIiIiIh0T2ZmZmZmZnhERIAAAACEREZmeIiIiIhmb2ZmZmZmZmb3RIAAAACERERER4iIiHRET2ZmZmZmZmZ4RIAAAACERERERHiIh0RET2ZmZmZmZmiERIAAAACERERERIiIiHRET2ZmZmZmZohERIAAAACERERER4iIiHRESP////9maIRERIAAAACERERER4iIiHRERERERE9mh0RERIAAAACERERERIiIiGRERERERE9odERERIAAAACERERERHiIh0RERERERE/3RERERIAAAACERERERER3REREREREREhkRERERIAAAACERERERERERERERERERERERERERIAAAACERERERERERERERERERERERERERIAAAACEREZmZmZmZmZmZmZmZmZmZmZERIAAAACERERERERERERERERERERERERERIAAAACERERERERERERERERERERERERERIAAAACERERERERERERERERERERERERERIAAAACEREZkRGZERERERERERERERERERIAAAACERGZmRu5kRERERERERERERERERIAAAAB2RGZmRu5kRERERERERERERERERnAAAABmREZkRGZERERERERERERERERERmAAAAAHZERERERERERERERERERERERERwAAAAAEdkRERERERERERERERERERERGdAAAAAAAB3ZkRERERERERERERERERGZnQAAAAAAAAGZ4iIiIiIiIiIiIiIiIiHdgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////////AAD///////8AAP///////wAA////////AAD///////8AAP///////wAA////////AAD/wAAAAf8AAP8AAAAAfwAA/gAAAAA/AAD8AAAAAB8AAPgAAAAADwAA+AAAAAAPAAD4AAAAAAcAAPAAAAAABwAA8AAAAAAHAADwAAAAAAcAAPAAAAAABwAA8AAAAAAHAADwAAAAAAcAAPAAAAAABwAA8AAAAAAHAADwAAAAAAcAAPAAAAAABwAA8AAAAAAHAADwAAAAAAcAAPAAAAAABwAA8AAAAAAHAADwAAAAAAcAAPAAAAAABwAA8AAAAAAHAADwAAAAAAcAAPAAAAAABwAA8AAAAAAHAADwAAAAAAcAAPgAAAAABwAA+AAAAAAPAAD8AAAAAA8AAPwAAAAAHwAA/gAAAAA/AAD/gAAAAP8AAP/gAAAD/wAA////////AAD///////8AAP///////wAA////////AAD///////8AAP///////wAA"
    $icoBytes = [System.Convert]::FromBase64String($icoBase64)
    $msIco = New-Object System.IO.MemoryStream(,$icoBytes)
    $appIcon = New-Object System.Drawing.Icon($msIco)
    $form.Icon = $appIcon
    if ($env:PROFILEMIGRATOR_HOST_EXE) {
        $hostIcon = [System.Drawing.Icon]::ExtractAssociatedIcon($env:PROFILEMIGRATOR_HOST_EXE)
        if ($hostIcon) { $form.Icon = $hostIcon }
    }
} catch {}

# Всплывающие подсказки (ToolTip)
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.AutoPopDelay = 10000
$tooltip.InitialDelay = 300
$tooltip.ReshowDelay = 200
$tooltip.ShowAlways = $true

# Верхняя панель (Header)
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = "Top"
$pnlHeader.Height = 65
$pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(24, 43, 73)
$form.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "RDS Profile Migration Toolkit"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Location = New-Object System.Drawing.Point(20, 10)
$lblTitle.AutoSize = $true
$pnlHeader.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Миграция, изоляция ACL и валидация профилей пользователей"
$lblSub.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$lblSub.ForeColor = [System.Drawing.Color]::FromArgb(180, 200, 230)
$lblSub.Location = New-Object System.Drawing.Point(22, 38)
$lblSub.AutoSize = $true
$pnlHeader.Controls.Add($lblSub)

# Кнопка СПРАВКА в шапке
$lblHelp = New-Object System.Windows.Forms.Button
$btnHelp = New-Object System.Windows.Forms.Button
$btnHelp.Text = "Справка / Помощь (?)"
$btnHelp.Size = New-Object System.Drawing.Size(170, 30)
$btnHelp.Location = New-Object System.Drawing.Point(830, 17)
$btnHelp.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
function Set-MainButtonStyle ([Windows.Forms.Button]$Button) {
    $Button.Font = $btnHelp.Font
    $Button.UseVisualStyleBackColor = $false
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = [Drawing.Color]::FromArgb(41,128,185)
    $Button.ForeColor = [Drawing.Color]::White
    $Button.Add_EnabledChanged({
        param($sender,$eventArgs)
        $sender.BackColor = if ($sender.Enabled) { [Drawing.Color]::FromArgb(41,128,185) } else { [Drawing.Color]::FromArgb(202,212,222) }
    })
    if (-not $Button.Enabled) { $Button.BackColor = [Drawing.Color]::FromArgb(202,212,222) }
}
$btnHelp.Font = New-Object Drawing.Font("Segoe UI",9,[Drawing.FontStyle]::Bold)
Set-MainButtonStyle $btnHelp
$tooltip.SetToolTip($btnHelp, "Открыть подробное руководство по каждому шагу и кнопке миграции")
$pnlHeader.Controls.Add($btnHelp)

$btnAbout = New-Object System.Windows.Forms.Button
$btnAbout.Text = "О программе"
$btnAbout.Size = New-Object System.Drawing.Size(130, 30)
$btnAbout.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnAbout.BackColor = [System.Drawing.Color]::FromArgb(51, 73, 103)
$btnAbout.ForeColor = [System.Drawing.Color]::White
$btnAbout.FlatStyle = "Flat"
$btnAbout.FlatAppearance.BorderSize = 0
$tooltip.SetToolTip($btnAbout, "Назначение, совместимость и системные требования")
$pnlHeader.Controls.Add($btnAbout)

# Блок настроек (GroupBox)
$grpSettings = New-Object System.Windows.Forms.GroupBox
$grpSettings.Text = "Параметры окружения и Active Directory"
$grpSettings.Location = New-Object System.Drawing.Point(15, 75)
$grpSettings.Size = New-Object System.Drawing.Size(995, 95)
$grpSettings.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grpSettings.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$form.Controls.Add($grpSettings)

# Сервер-источник
$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Text = "Сервер-источник:"
$lblSrc.Location = New-Object System.Drawing.Point(15, 25)
$lblSrc.AutoSize = $true
$grpSettings.Controls.Add($lblSrc)

$txtSrc = New-Object System.Windows.Forms.TextBox
$txtSrc.Text = $Global:Config.SourceServer
$txtSrc.Location = New-Object System.Drawing.Point(135, 22)
$txtSrc.Size = New-Object System.Drawing.Size(140, 23)
$grpSettings.Controls.Add($txtSrc)

# Целевой каталог
$lblDst = New-Object System.Windows.Forms.Label
$lblDst.Text = "Целевая папка:"
$lblDst.Location = New-Object System.Drawing.Point(295, 25)
$lblDst.AutoSize = $true
$grpSettings.Controls.Add($lblDst)

$txtDst = New-Object System.Windows.Forms.TextBox
$txtDst.Text = $Global:Config.ProfilesRoot
$txtDst.Location = New-Object System.Drawing.Point(395, 22)
$txtDst.Size = New-Object System.Drawing.Size(140, 23)
$grpSettings.Controls.Add($txtDst)

# Лимит неактивности
$lblLimit = New-Object System.Windows.Forms.Label
$lblLimit.Text = "Лимит неакт. (дн):"
$lblLimit.Location = New-Object System.Drawing.Point(555, 25)
$lblLimit.AutoSize = $true
$grpSettings.Controls.Add($lblLimit)

$numLimit = New-Object System.Windows.Forms.NumericUpDown
$numLimit.Value = $Global:Config.InactiveDaysLimit
$numLimit.Maximum = 9999
$numLimit.Location = New-Object System.Drawing.Point(670, 22)
$numLimit.Size = New-Object System.Drawing.Size(65, 23)
$grpSettings.Controls.Add($numLimit)

# Кнопка сохранения настроек
$btnSaveCfg = New-Object System.Windows.Forms.Button
$btnSaveCfg.Text = "Сохранить параметры"
$btnSaveCfg.Size = New-Object System.Drawing.Size(195, 26)
Set-MainButtonStyle $btnSaveCfg
$grpSettings.Controls.Add($btnSaveCfg)

# Целевое OU
$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Text = "Целевое OU (SearchBase):"
$lblOU.Location = New-Object System.Drawing.Point(15, 60)
$lblOU.AutoSize = $true
$grpSettings.Controls.Add($lblOU)

$txtOU = New-Object System.Windows.Forms.TextBox
$txtOU.Text = $Global:Config.TargetOU
$txtOU.Location = New-Object System.Drawing.Point(180, 57)
$grpSettings.Controls.Add($txtOU)

$btnFolders = New-Object System.Windows.Forms.Button
$btnFolders.Text = "Папки переноса ($($Global:Config.TargetFolders.Count))"
$btnFolders.Size = New-Object Drawing.Size(195, 26)
Set-MainButtonStyle $btnFolders
$grpSettings.Controls.Add($btnFolders)
$tooltip.SetToolTip($btnFolders, ($Global:Config.TargetFolders -join '; '))
$btnFolders.Add_Click({
    $selected = @(Select-TargetFolders $form $Global:Config.TargetFolders)
    if ($selected.Count) {
        $Global:Config.TargetFolders = $selected
        $btnFolders.Text = "Папки переноса ($($selected.Count))*"
        $tooltip.SetToolTip($btnFolders, (($selected -join '; ') + ' — нажмите «Сохранить параметры»'))
    }
})

# Динамическое выравнивание правых элементов в блоке настроек и шапке
$UpdateLayoutFunc = {
    $btnSaveCfg.Left = [Math]::Max(740, ($grpSettings.ClientSize.Width - $btnSaveCfg.Width - 15))
    $btnSaveCfg.Top = 20
    $btnFolders.Left = $btnSaveCfg.Left
    $btnFolders.Top = 55
    $txtOU.Width = [Math]::Max(200, ($btnFolders.Left - $txtOU.Left - 12))
    $btnHelp.Left = [Math]::Max(600, ($pnlHeader.ClientSize.Width - $btnHelp.Width - 20))
    $btnAbout.Location = New-Object System.Drawing.Point(($btnHelp.Left - $btnAbout.Width - 10), 17)
}
$grpSettings.Add_Resize($UpdateLayoutFunc)
$pnlHeader.Add_Resize($UpdateLayoutFunc)
&$UpdateLayoutFunc

# Центральная панель с таблицей пользователей
$pnlData = New-Object System.Windows.Forms.Panel
$pnlData.Location = New-Object System.Drawing.Point(15, 175)
$pnlData.Size = New-Object System.Drawing.Size(995, 184)
$pnlData.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pnlData)

$gridUsers = New-Object System.Windows.Forms.DataGridView
$gridUsers.Dock = "Fill"
$gridUsers.BackgroundColor = [System.Drawing.Color]::White
$gridUsers.AllowUserToAddRows = $false
$gridUsers.SelectionMode = "FullRowSelect"
$gridUsers.MultiSelect = $true
$gridUsers.RowHeadersVisible = $false
$pnlData.Controls.Add($gridUsers)

# Колонки таблицы
$colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colCheck.HeaderText = "Выбор"
$colCheck.Width = 50
$gridUsers.Columns.Add($colCheck) | Out-Null

$colLogin = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colLogin.HeaderText = "Логин"
$colLogin.Width = 140
$gridUsers.Columns.Add($colLogin) | Out-Null

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.HeaderText = "ФИО"
$colName.Width = 220
$gridUsers.Columns.Add($colName) | Out-Null

$colLast = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colLast.HeaderText = "Последний вход"
$colLast.Width = 150
$gridUsers.Columns.Add($colLast) | Out-Null

$colDays = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDays.HeaderText = "Дней назад"
$colDays.Width = 90
$gridUsers.Columns.Add($colDays) | Out-Null

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.HeaderText = "Готовность"
$colStatus.Width = 110
$gridUsers.Columns.Add($colStatus) | Out-Null

$colNotes = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNotes.HeaderText = "Примечание / Базы 1C"
$colNotes.AutoSizeMode = "Fill"
$gridUsers.Columns.Add($colNotes) | Out-Null

$colSize = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSize.HeaderText = "Размер"
$colSize.Width = 95
$gridUsers.Columns.Add($colSize) | Out-Null
$colSize.DisplayIndex = 6

$gridUsers.Add_SortCompare({
    param($s, $e)
    if ($e.Column.Index -eq 7) {
        $v1 = if ($s.Rows[$e.RowIndex1].Tag -and $s.Rows[$e.RowIndex1].Tag.PSObject.Properties['EstimatedBytes']) { [int64]$s.Rows[$e.RowIndex1].Tag.EstimatedBytes } else { [int64]-1 }
        $v2 = if ($s.Rows[$e.RowIndex2].Tag -and $s.Rows[$e.RowIndex2].Tag.PSObject.Properties['EstimatedBytes']) { [int64]$s.Rows[$e.RowIndex2].Tag.EstimatedBytes } else { [int64]-1 }
        $e.SortResult = $v1.CompareTo($v2)
        $e.Handled = $true
    }
})

$gridUsers.Add_CellDoubleClick({
    param($s, $e)
    if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 7) {
        $row = $s.Rows[$e.RowIndex]
        if ($row.Tag -and $row.Tag.PSObject.Properties['Details'] -and $row.Tag.Details.Count -gt 0) {
            $user = $row.Cells[1].Value
            $lines = @("Детализация папок пользователя [$user]:`r`n")
            foreach ($key in $row.Tag.Details.Keys) {
                $d = $row.Tag.Details[$key]
                $lines += "• $($key): $(Format-ByteSize $d.Bytes) ($($d.Files) файлов)"
            }
            $lines += "`r`nВсего: $(Format-ByteSize $row.Tag.EstimatedBytes) ($($row.Tag.EstimatedFiles) файлов)"
            [System.Windows.Forms.MessageBox]::Show($form, ($lines -join "`r`n"), "Детализация объема: $user", "OK", "Information")
        }
    }
})

# Кнопки быстрых действий над списком
$pnlGridTools = New-Object System.Windows.Forms.Panel
$pnlGridTools.Location = New-Object System.Drawing.Point(15, 371)
$pnlGridTools.Size = New-Object System.Drawing.Size(995, 30)
$pnlGridTools.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pnlGridTools)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Выбрать все"
$btnSelectAll.Size = New-Object System.Drawing.Size(95, 26)
$btnSelectAll.Location = New-Object System.Drawing.Point(0, 2)
Set-MainButtonStyle $btnSelectAll
$pnlGridTools.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Снять выбор"
$btnDeselectAll.Size = New-Object System.Drawing.Size(95, 26)
$btnDeselectAll.Location = New-Object System.Drawing.Point(100, 2)
Set-MainButtonStyle $btnDeselectAll
$pnlGridTools.Controls.Add($btnDeselectAll)

# Фильтр "Показать только выбранные"
$chkOnlySelected = New-Object System.Windows.Forms.CheckBox
$chkOnlySelected.Text = "Только выбранные"
$chkOnlySelected.AutoSize = $true
$chkOnlySelected.Location = New-Object System.Drawing.Point(210, 6)
$chkOnlySelected.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$pnlGridTools.Controls.Add($chkOnlySelected)

# Быстрый поиск по логину/ФИО
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Поиск:"
$lblSearch.AutoSize = $true
$lblSearch.Location = New-Object System.Drawing.Point(360, 7)
$lblSearch.ForeColor = [System.Drawing.Color]::DarkSlateGray
$pnlGridTools.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Size = New-Object System.Drawing.Size(130, 23)
$txtSearch.Location = New-Object System.Drawing.Point(405, 4)
$pnlGridTools.Controls.Add($txtSearch)

# Счетчик пользователей и выбранных
$lblCounter = New-Object System.Windows.Forms.Label
$lblCounter.Text = "Всего: 0 | Выбрано: 0"
$lblCounter.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblCounter.ForeColor = [System.Drawing.Color]::FromArgb(24, 43, 73)
$lblCounter.AutoSize = $true
$lblCounter.Location = New-Object System.Drawing.Point(555, 7)
$pnlGridTools.Controls.Add($lblCounter)

$btnSaveList = New-Object Windows.Forms.Button
$btnSaveList.Text = 'Сохранить список'
$btnSaveList.SetBounds(733,2,125,26)
$btnSaveList.Anchor = 'Top,Right'
$btnSaveList.Enabled = $false
Set-MainButtonStyle $btnSaveList
$pnlGridTools.Controls.Add($btnSaveList)
$btnLoadList = New-Object Windows.Forms.Button
$btnLoadList.Text = 'Загрузить список'
$btnLoadList.SetBounds(865,2,125,26)
$btnLoadList.Anchor = 'Top,Right'
Set-MainButtonStyle $btnLoadList
$pnlGridTools.Controls.Add($btnLoadList)

function Format-ByteSize ([int64]$Bytes) {
    if ($Bytes -lt 0) { return "—" }
    if ($Bytes -ge 1TB) { return ("{0:N2} ТБ" -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ("{0:N2} ГБ" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N1} МБ" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N0} КБ" -f ($Bytes / 1KB)) }
    return "$Bytes Б"
}

function Get-DriveFreeSpace ([string]$Path) {
    try {
        if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
        $root = [System.IO.Path]::GetPathRoot($Path)
        if ([string]::IsNullOrWhiteSpace($root)) { return $null }
        $drive = [System.IO.DriveInfo]::new($root)
        if ($drive.IsReady) { return [int64]$drive.AvailableFreeSpace }
    } catch {}
    return $null
}

function Save-AuditList ($Grid, $Context, [string]$Path) {
    if (-not $Context -or -not $Grid.Rows.Count) { throw 'Нет списка для сохранения.' }
    [void]$Grid.EndEdit()
    $selectedRows = @(foreach ($row in $Grid.Rows) { if ([bool]$row.Cells[0].Value) { $row } })
    if ($selectedRows.Count -eq 0) { throw 'Нет отмеченных пользователей для сохранения. Отметьте нужные учетные записи галочками.' }
    $rows = @(foreach ($row in $selectedRows) {
        $sizeVal = if ($row.Cells.Count -gt 7 -and $row.Cells[7].Value) { [string]$row.Cells[7].Value } else { '—' }
        [ordered]@{
            Selected = [bool]$row.Cells[0].Value
            Login = [string]$row.Cells[1].Value
            Name = [string]$row.Cells[2].Value
            LastLogon = [string]$row.Cells[3].Value
            Days = [string]$row.Cells[4].Value
            Status = [string]$row.Cells[5].Value
            Notes = [string]$row.Cells[6].Value
            Size = $sizeVal
            Inactive = $row.DefaultCellStyle.ForeColor -eq [Drawing.Color]::Gray
        }
    })
    $snapshot = [ordered]@{ Version = 1; Server = $Context.Server; OU = $Context.OU; AuditTime = $Context.AuditTime; Rows = $rows }
    $json = ConvertTo-Json -InputObject $snapshot -Depth 4
    $dir = [IO.Path]::GetDirectoryName($Path)
    [void][IO.Directory]::CreateDirectory($dir)
    $tmp = Join-Path $dir ([guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllText($tmp,$json,[Text.Encoding]::UTF8)
        if ([IO.File]::Exists($Path)) { [IO.File]::Replace($tmp,$Path,[System.Management.Automation.Language.NullString]::Value) }
        else { [IO.File]::Move($tmp,$Path) }
    } finally { if ([IO.File]::Exists($tmp)) { [IO.File]::Delete($tmp) } }
}

function Read-AuditList ([string]$Path, [string]$Server, [string]$OU) {
    if (-not [IO.File]::Exists($Path)) { throw "Сохраненный список не найден: $Path" }
    $data = [IO.File]::ReadAllText($Path) | ConvertFrom-Json -ErrorAction Stop
    if ($data.Version -ne 1 -or $data.Rows -isnot [array] -or -not $data.Rows.Count) { throw 'Неверный формат списка аудита.' }
    if ($data.Server -ne $Server -or $data.OU -ne $OU) { throw 'Источник или OU сохраненного списка отличается от текущих настроек.' }
    $seen = @{}
    foreach ($r in $data.Rows) {
        if ($r.Login -isnot [string] -or [string]::IsNullOrWhiteSpace($r.Login) -or $r.Login -match '[\\/:*?"<>|\x00-\x1f]|[ .]$' -or
            $seen.ContainsKey($r.Login) -or $r.Selected -isnot [bool] -or $r.Inactive -isnot [bool]) { throw 'Некорректная строка в списке аудита.' }
        foreach ($field in @('Name','LastLogon','Days','Status','Notes')) {
            if ($r.$field -isnot [string]) { throw "Некорректное поле списка: $field" }
        }
        if ($r.PSObject.Properties['Size'] -and $r.Size -isnot [string]) {
            throw "Некорректное поле списка: Size"
        }
        if ($r.Status -eq 'Готов') { $r.Status = 'Готов' }
        elseif ($r.Status -in @('Внимание','Требует внимания')) { $r.Status = 'Внимание' }
        $seen[$r.Login] = $true
    }
    return $data
}

$btnSaveList.Add_Click({
    try {
        $path = Join-Path $Global:Config.WorkDir 'AuditList.json'
        Save-AuditList $gridUsers $script:auditListContext $path
        $savedCount = @($gridUsers.Rows | Where-Object { [bool]$_.Cells[0].Value }).Count
        Write-GuiLog "Сохранен список отмеченных пользователей ($savedCount): $path" 'SUCCESS'
    } catch { Write-GuiLog "Сохранение списка: $($_.Exception.Message)" 'ERROR' }
})

$btnLoadList.Add_Click({
    try {
        $path = Join-Path $Global:Config.WorkDir 'AuditList.json'
        $data = Read-AuditList $path $txtSrc.Text.Trim() $txtOU.Text.Trim()
        if ($gridUsers.Rows.Count -and [Windows.Forms.MessageBox]::Show($form,'Заменить текущую таблицу и галочки сохраненным списком?','Загрузить список','YesNo','Question') -ne 'Yes') { return }
        $gridUsers.Rows.Clear()
        foreach ($r in $data.Rows) {
            $sizeVal = if ($r.PSObject.Properties['Size'] -and -not [string]::IsNullOrWhiteSpace($r.Size)) { [string]$r.Size } else { '—' }
            $i = $gridUsers.Rows.Add($r.Selected,$r.Login,$r.Name,$r.LastLogon,$r.Days,$r.Status,$r.Notes,$sizeVal)
            if ($r.Inactive) { $gridUsers.Rows[$i].DefaultCellStyle.ForeColor = [Drawing.Color]::Gray }
        }
        $script:auditListContext = @{ Server=$data.Server; OU=$data.OU; AuditTime=$data.AuditTime }
        Update-CountersAndFilter
        $btnSaveList.Enabled = $true
        Write-GuiLog "Загружено пользователей: $($data.Rows.Count). Данные аудита от $($data.AuditTime)." 'SUCCESS'
    } catch { Write-GuiLog "Загрузка списка: $($_.Exception.Message)" 'ERROR' }
})

function Update-CountersAndFilter {
    $total = $gridUsers.Rows.Count
    $selectedCount = 0
    $selectedBytes = [int64]0
    $measuredCount = 0
    $filterText = $txtSearch.Text.Trim().ToLower()
    $onlyChecked = $chkOnlySelected.Checked

    foreach ($row in $gridUsers.Rows) {
        $isChecked = [bool]$row.Cells[0].Value
        if ($isChecked) {
            $selectedCount++
            if ($row.Tag -and $row.Tag.PSObject.Properties['EstimatedBytes'] -and $row.Tag.EstimatedBytes -ge 0) {
                $selectedBytes += [int64]$row.Tag.EstimatedBytes
                $measuredCount++
            }
        }

        $login = "$($row.Cells[1].Value)".ToLower()
        $fio = "$($row.Cells[2].Value)".ToLower()

        $matchesSearch = [string]::IsNullOrEmpty($filterText) -or ($login.Contains($filterText) -or $fio.Contains($filterText))
        $matchesCheck = (-not $onlyChecked) -or $isChecked

        $row.Visible = ($matchesSearch -and $matchesCheck)
    }

    $targetRoot = [System.IO.Path]::GetPathRoot($Global:Config.ProfilesRoot)
    $freeBytes = Get-DriveFreeSpace $targetRoot

    $sizeText = ""
    if ($measuredCount -gt 0) {
        $sizeText = " | Объем: $(Format-ByteSize $selectedBytes)"
    }

    if ($measuredCount -gt 0 -and $freeBytes -ne $null -and $selectedBytes -gt $freeBytes) {
        $lblCounter.ForeColor = [System.Drawing.Color]::DarkRed
        $lblCounter.Text = "Всего: $total | Выбрано: $selectedCount$sizeText [МЕСТА МАЛО: $(Format-ByteSize $freeBytes)]"
    } else {
        $lblCounter.ForeColor = [System.Drawing.Color]::FromArgb(24, 43, 73)
        $lblCounter.Text = "Всего: $total | Выбрано: $selectedCount$sizeText"
    }
}

# Панель основных кнопок действий (FlowLayoutPanel: по возможности в один ряд, если не помещаются — на второй с выравниванием влево)
$pnlActions = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlActions.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$pnlActions.WrapContents = $true
$pnlActions.AutoScroll = $false
$pnlActions.Location = New-Object System.Drawing.Point(15, 412)
$pnlActions.Size = New-Object System.Drawing.Size(995, 80)
$pnlActions.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pnlActions)

# 1. Интерактивная плашка подсказок под кнопками действий (аккуратный информационный блок)
$pnlHintBox = New-Object System.Windows.Forms.Panel
$pnlHintBox.Location = New-Object System.Drawing.Point(15, 500)
$pnlHintBox.Size = New-Object System.Drawing.Size(995, 23)
$pnlHintBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlHintBox.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
$pnlHintBox.BorderStyle = "FixedSingle"
$form.Controls.Add($pnlHintBox)

$Global:lblHintText = New-Object System.Windows.Forms.Label
$Global:lblHintText.Text = "Подсказка: Наведите курсор на любую кнопку выше для просмотра её назначения"
$Global:lblHintText.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$Global:lblHintText.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$Global:lblHintText.Dock = "Fill"
$Global:lblHintText.TextAlign = "MiddleLeft"
$Global:lblHintText.Padding = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$pnlHintBox.Controls.Add($Global:lblHintText)

function Create-ActionButton ($Title, $BgColor, $TipText) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Title
    $btn.BackColor = $BgColor
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 8)
    $btn.Tag = $TipText
    
    $btn.Add_MouseEnter({
        param($s, $e)
        $Global:lblHintText.Text = [string]$s.Tag
        $Global:lblHintText.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    })
    $btn.Add_MouseLeave({
        $Global:lblHintText.Text = "Подсказка: Наведите курсор на любую кнопку выше для просмотра её назначения"
        $Global:lblHintText.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    })

    return $btn
}

$stepReadyColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
$stepBusyColor = [System.Drawing.Color]::FromArgb(206, 222, 234)
$btnAudit = Create-ActionButton "1. Аудит OU (Поиск)" $stepReadyColor "Шаг 1: Поиск пользователей в OU; lastLogon по ответившим DC. Отмечает включенные учетные записи в пределах лимита."
$btnEstimate = Create-ActionButton "Оценка объема" $stepReadyColor "Оценка объема переносимых данных для отмеченных пользователей по сети (без копирования). Проверяет свободное место на целевом диске."
$btnSync = Create-ActionButton "2. Миграция выбранных" $stepReadyColor "Шаг 2: Для отмеченных пользователей готовит профиль, копирует TargetFolders и настройки запуска 1С, применяет ACL."
$btnDelta = Create-ActionButton "3. Докачка дельты (/XO)" $stepReadyColor "Шаг 3: Повторное копирование для отмеченных пользователей; /XO пропускает файлы источника старее целевых."
$btnValidate = Create-ActionButton "4. Валидация профилей" $stepReadyColor "Шаг 4: Проверка отмеченных профилей: SID, путь, NTUSER.DAT, TargetFolders, состояние загрузки и базовые ACL."
$btnShortcuts = Create-ActionButton "5. Проверка ярлыков" $stepReadyColor "Шаг 5: Проверка ярлыков на рабочем столе отмеченных пользователей. Выберите действие для каждого ярлыка в едином списке."
$btnStop = Create-ActionButton "Остановить процесс" ([System.Drawing.Color]::FromArgb(160, 170, 180)) "Запрашивает остановку операции и завершает запущенный ею Robocopy или icacls. Уже выполненные изменения сохраняются."
$btnStop.Enabled = $false

$actionButtons = @($btnAudit, $btnEstimate, $btnSync, $btnDelta, $btnValidate, $btnShortcuts, $btnStop)
$maxActionTextWidth = 0
foreach ($b in $actionButtons) {
    $sz = [System.Windows.Forms.TextRenderer]::MeasureText($b.Text, $b.Font)
    if ($sz.Width -gt $maxActionTextWidth) { $maxActionTextWidth = $sz.Width }
}
$actionButtonWidth = $maxActionTextWidth + 40
foreach ($b in $actionButtons) {
    $b.Size = New-Object System.Drawing.Size($actionButtonWidth, 35)
    $pnlActions.Controls.Add($b)
}

# 2. Выделенный блок статуса и прогресса операции с отдельным заголовком и процентами
$pnlProgressHeader = New-Object System.Windows.Forms.Panel
$pnlProgressHeader.Location = New-Object System.Drawing.Point(15, 530)
$pnlProgressHeader.Size = New-Object System.Drawing.Size(995, 20)
$pnlProgressHeader.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pnlProgressHeader)

$Global:lblProgressTitle = New-Object System.Windows.Forms.Label
$Global:lblProgressTitle.Text = "Прогресс операции: Ожидание запуска"
$Global:lblProgressTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$Global:lblProgressTitle.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$Global:lblProgressTitle.Location = New-Object System.Drawing.Point(0, 1)
$Global:lblProgressTitle.Size = New-Object System.Drawing.Size(850, 18)
$pnlProgressHeader.Controls.Add($Global:lblProgressTitle)

$Global:lblProgressPercent = New-Object System.Windows.Forms.Label
$Global:lblProgressPercent.Text = "0%"
$Global:lblProgressPercent.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$Global:lblProgressPercent.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
$Global:lblProgressPercent.TextAlign = "MiddleRight"
$Global:lblProgressPercent.Size = New-Object System.Drawing.Size(120, 18)
$Global:lblProgressPercent.Location = New-Object System.Drawing.Point(875, 0)
$Global:lblProgressPercent.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlProgressHeader.Controls.Add($Global:lblProgressPercent)

# Прогресс бар (высота 18px с четким обрамлением)
if (-not ('ProfileMigrator.ToolkitProgressBar' -as [type])) {
    Add-Type -ReferencedAssemblies System.Windows.Forms,System.Drawing -TypeDefinition @'
using System;
using System.Drawing;
using System.Windows.Forms;
namespace ProfileMigrator {
    public sealed class ToolkitProgressBar : ProgressBar {
        public ToolkitProgressBar() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }
        protected override void OnPaint(PaintEventArgs e) {
            Rectangle track = ClientRectangle;
            if (track.Width < 3 || track.Height < 3) return;
            e.Graphics.Clear(BackColor);
            ControlPaint.DrawBorder(e.Graphics, track, SystemColors.ControlDark, ButtonBorderStyle.Solid);
            track.Inflate(-1, -1);
            int range = Maximum - Minimum;
            if (range <= 0) return;
            track.Width = (int)Math.Round(track.Width * ((Value - Minimum) / (double)range));
            using (var brush = new SolidBrush(ForeColor)) {
                e.Graphics.FillRectangle(brush, track);
            }
        }
    }
}
'@
}
$Global:pbar = New-Object ProfileMigrator.ToolkitProgressBar
$Global:pbar.ForeColor = $stepReadyColor
$Global:pbar.BackColor = [Drawing.Color]::FromArgb(231, 236, 241)
$Global:pbar.Location = New-Object System.Drawing.Point(15, 554)
$Global:pbar.Size = New-Object System.Drawing.Size(995, 18)
$Global:pbar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$Global:pbar.Visible = $true
$form.Controls.Add($Global:pbar)

# 3. Заголовок консоли лога
$lblLogTitle = New-Object System.Windows.Forms.Label
$lblLogTitle.Text = "Журнал операций и диагностика:"
$lblLogTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblLogTitle.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
$lblLogTitle.Location = New-Object System.Drawing.Point(15, 580)
$lblLogTitle.Size = New-Object System.Drawing.Size(350, 16)
$lblLogTitle.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($lblLogTitle)

# Лог-консоль
$txtLogs = New-Object System.Windows.Forms.TextBox
$txtLogs.Multiline = $true
$txtLogs.ScrollBars = "Vertical"
$txtLogs.ReadOnly = $true
$txtLogs.BackColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
$txtLogs.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
$txtLogs.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$txtLogs.Location = New-Object System.Drawing.Point(15, 600)
$txtLogs.Size = New-Object System.Drawing.Size(995, 98)
$txtLogs.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($txtLogs)

$btnExpandLog = New-Object System.Windows.Forms.Button
$btnExpandLog.Text = "Развернуть журнал"
$btnExpandLog.SetBounds(850, 574, 160, 24)
$btnExpandLog.Anchor = 'Bottom, Right'
Set-MainButtonStyle $btnExpandLog
$form.Controls.Add($btnExpandLog)
$script:logWindow = $null
$script:expandedLog = $null
$btnExpandLog.Add_Click({
    if ($script:logWindow) {
        if ($script:logWindow.WindowState -eq 'Minimized') { $script:logWindow.WindowState = 'Normal' }
        $script:logWindow.Activate()
        return
    }
    $script:logWindow = New-Object System.Windows.Forms.Form
    $script:logWindow.Text = "Журнал операций и диагностика"
    $script:logWindow.Size = New-Object Drawing.Size(1000, 550)
    $script:logWindow.MinimumSize = New-Object Drawing.Size(460, 220)
    $script:logWindow.StartPosition = 'CenterParent'
    if ($form.Icon) { $script:logWindow.Icon = $form.Icon }
    $script:expandedLog = New-Object System.Windows.Forms.TextBox
    $script:expandedLog.Multiline = $true
    $script:expandedLog.ReadOnly = $true
    $script:expandedLog.ScrollBars = 'Both'
    $script:expandedLog.WordWrap = $false
    $script:expandedLog.Dock = 'Fill'
    $script:expandedLog.Font = $txtLogs.Font
    $script:expandedLog.ForeColor = $txtLogs.ForeColor
    $script:expandedLog.BackColor = $txtLogs.BackColor
    $script:expandedLog.Text = $txtLogs.Text
    $script:logWindow.Controls.Add($script:expandedLog)
    $script:logWindow.Add_FormClosed({
        $script:expandedLog = $null
        $script:logWindow = $null
    })
    $script:logWindow.Show($form)
    $script:expandedLog.SelectionStart = $script:expandedLog.TextLength
    $script:expandedLog.ScrollToCaret()
})

function Write-GuiLog ($msg, $type = "INFO") {
    $time = (Get-Date).ToString("HH:mm:ss")
    $line = "[$time] [$type] $msg`r`n"
    $txtLogs.AppendText($line)
    if ($script:expandedLog) { $script:expandedLog.AppendText($line) }
    $txtLogs.SelectionStart = $txtLogs.Text.Length
    $txtLogs.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Окно красивой справки в виде вкладок и карточек
$btnHelp.Add_Click({
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.MinimizeBox = $false
    $helpForm.ShowInTaskbar = $false
    $helpForm.Text = "Справочный центр ProfileMigrator"
    $helpForm.Size = New-Object System.Drawing.Size(960, 690)
    $helpForm.MinimumSize = New-Object System.Drawing.Size(960, 690)
    $helpForm.StartPosition = "CenterParent"
    $helpForm.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $helpForm.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    if ($form.Icon -ne $null) { $helpForm.Icon = $form.Icon }

    # Верхняя панель заголовка
    $hHeader = New-Object System.Windows.Forms.Panel
    $hHeader.Dock = "Top"
    $hHeader.Height = 60
    $hHeader.BackColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $helpForm.Controls.Add($hHeader)

    $hTitle = New-Object System.Windows.Forms.Label
    $hTitle.Text = "Руководство администратора по миграции профилей"
    $hTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $hTitle.ForeColor = [System.Drawing.Color]::White
    $hTitle.Location = New-Object System.Drawing.Point(20, 10)
    $hTitle.AutoSize = $true
    $hHeader.Controls.Add($hTitle)

    $hSub = New-Object System.Windows.Forms.Label
    $hSub.Text = "Назначение всех шагов, параметры безопасности и рекомендации по переводу пользователей"
    $hSub.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $hSub.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
    $hSub.Location = New-Object System.Drawing.Point(22, 35)
    $hSub.AutoSize = $true
    $hHeader.Controls.Add($hSub)

    # Нижняя панель с кнопкой закрытия
    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Dock = "Bottom"
    $pnlBottom.Height = 45
    $pnlBottom.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
    $helpForm.Controls.Add($pnlBottom)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Закрыть справку"
    $btnClose.Size = New-Object System.Drawing.Size(140, 30)
    $btnClose.Location = New-Object System.Drawing.Point(790, 8)
    $btnClose.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(24, 43, 73)
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.FlatStyle = "Flat"
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnClose.Add_Click({ $helpForm.Close() })
    $pnlBottom.Controls.Add($btnClose)

    # Separate the tabs from the dark header with a visible 12 px gutter.
    $helpHeaderGap = New-Object System.Windows.Forms.Panel
    $helpHeaderGap.Dock = "Top"
    $helpHeaderGap.Height = 12
    $helpHeaderGap.BackColor = $helpForm.BackColor
    $helpForm.Controls.Add($helpHeaderGap)
    $helpHeaderGap.BringToFront()

    # Вкладки (TabControl)
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $tabControl.Padding = New-Object System.Drawing.Point(14, 8)
    $helpForm.Controls.Add($tabControl)
    # Dock Fill must use the space remaining between the header and footer.
    $tabControl.BringToFront()

    function Add-HelpTab ($Title, [string[]]$Blocks) {
        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $Title
        $tab.BackColor = [System.Drawing.Color]::White
        
        $pnlScroll = New-Object System.Windows.Forms.Panel
        $pnlScroll.Dock = "Fill"
        $pnlScroll.AutoScroll = $true
        $tab.Controls.Add($pnlScroll)

        $y = 15
        foreach ($block in $Blocks) {
            $parts = $block -split "\|"
            $hdr = $parts[0].Trim()
            $body = $parts[1].Trim()

            $card = New-Object System.Windows.Forms.Panel
            $card.Location = New-Object System.Drawing.Point(15, $y)
            $card.Width = 880
            $card.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
            $card.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
            $card.BorderStyle = "FixedSingle"

            # Цветная полоса слева
            $stripe = New-Object System.Windows.Forms.Panel
            $stripe.Dock = "Left"
            $stripe.Width = 6
            $stripe.BackColor = $stepReadyColor
            $card.Controls.Add($stripe)

            $lblH = New-Object System.Windows.Forms.Label
            $lblH.Text = $hdr
            $lblH.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold)
            $lblH.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
            $lblH.Location = New-Object System.Drawing.Point(18, 10)
            $lblH.AutoSize = $true
            $card.Controls.Add($lblH)

            $lblB = New-Object System.Windows.Forms.Label
            $lblB.Text = $body
            $lblB.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $lblB.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
            $lblB.Location = New-Object System.Drawing.Point(18, 36)
            $lblB.MaximumSize = New-Object System.Drawing.Size(810, 0)
            $lblB.AutoSize = $true
            $card.Controls.Add($lblB)

            $card.Height = $lblB.Bottom + 14
            $pnlScroll.Controls.Add($card)
            $y += $card.Height + 12
        }

        # Keep 24 px after the final card reachable at the end of scrolling.
        $pnlScroll.AutoScrollMinSize = New-Object System.Drawing.Size(0, ($y + 12))
        $tabControl.TabPages.Add($tab) | Out-Null
    }

    # Вкладка 1: Кнопки действий (Шаги 1-7)
    $actionsHelp = @(
        "1. Аудит OU (Поиск)|• Ищет учетные записи в указанном SearchBase и его вложенных подразделениях AD.`r`n• Запрашивает lastLogon у найденных контроллеров домена и берет наибольшее полученное значение. При недоступности DC результат может быть неполным.`r`n• Отмечает включенные учетные записи, у которых число дней после входа не превышает лимит. Если дата не получена, отображает «Никогда» и не ставит галочку.`r`n• Результат выводится в таблицу. CSV-отчет и ActiveUsersList.txt эта версия не создает.",
        "Оценка объема|• Сканирует папки TargetFolders и настройки 1С на сервере-источнике для отмеченных пользователей в режиме dry-run (Robocopy /L) без копирования данных.`r`n• Определяет суммарный объем данных и число файлов, отображая результат в столбце «Размер». Двойной клик на ячейке открывает детализацию по каждой папке.`r`n• Клик по заголовку столбца «Размер» выполняет численную сортировку таблицы, позволяя мгновенно выявить самые тяжелые профили.`r`n• Проверяет доступное место на целевом диске (обычно C:\) и предупреждает при риске переполнения диска (Disk Full).`r`n• В строке счетчиков отображает суммарный объем выбранных пользователей в реальном времени.",
        "2. Миграция выбранных|• После подтверждения обрабатывает только отмеченных пользователей по SID из AD. Зарегистрированный профиль использует по пути из ProfileList.`r`n• Если записи SID нет, но папка логина в системном каталоге профилей уже есть, регистрирует эту же папку без переименования и дубликатов. Подготовка добавляет только отсутствующие файлы из локального Default, сохраняет имеющиеся файлы и NTUSER.DAT, настраивает ACL.`r`n• Если нет ни записи, ни папки, создает профиль средствами Windows. Пользователь не должен входить на целевой сервер во время миграции.`r`n• Копирует папки TargetFolders с подпапками, включая пустые. По умолчанию: Desktop, Documents, Downloads, Outlook, Crypto Pro, Google, YandexBrowser. Переходы junction исключаются (/XJ).`r`n• Отдельно копирует ibases.v8i и 1CEStart.cfg из AppData\Roaming\1C\1CEStart, если они есть. Это настройки запуска, а не сами базы 1С.`r`n• Переносит данные, атрибуты и время файлов; затем применяет ACL на целевом профиле. При ошибках часть данных уже может быть скопирована.`r`n• Отличающиеся файлы назначения могут быть перезаписаны при последующем копировании с источника. Лишние файлы на цели и данные источника не удаляются.`r`n• При включении опции Outlook шаг 2 сохраняет существующие PST и извлекает почтовые настройки из временной копии исходного NTUSER.DAT. Загруженные, конфликтующие и некорректные целевые профили не обрабатываются.",
        "3. Докачка дельты (/XO)|• Выполняет тот же перенос для отмеченных пользователей, добавляя /XO: пропускает файлы источника, которые старее соответствующих файлов на цели.`r`n• Robocopy сравнивает текущее состояние файлов; дата первой миграции не хранится. Новые и отличающиеся файлы копируются с учетом /XO.`r`n• При отсутствии профиля и папки создает профиль Windows, затем применяет ACL. Если папка уже есть, но SID не зарегистрирован, требуется сначала шаг 2; дубликат папки не создается.`r`n• NTUSER.DAT сохраняется; лишние файлы на цели не удаляются. Пользователь не должен входить на целевой сервер во время копирования.`r`n• Время выполнения зависит от объема данных, сети и дисков. После докачки повторите валидацию.",
        "4. Изоляция прав ACL|• Работает с отмеченными, зарегистрированными и незагруженными профилями. Эта же настройка вызывается при миграции и докачке.`r`n• На корне профиля отключает наследование с сохранением унаследованных записей как явных (/inheritance:d).`r`n• Удаляет записи Everyone, Authenticated Users и Users в дереве профиля. Владельцем корня назначает пользователя по SID.`r`n• Выдает FullControl пользователю, SYSTEM и Administrators на корень, NTUSER.DAT и рекурсивно на существующие папки TargetFolders и настроек 1С.`r`n• Остальные явные разрешения не очищаются. Операция не является полной проверкой всех возможных доступов к профилю.",
        "5. Валидация профилей|• Проверяет только отмеченных пользователей и не копирует данные, не создает профили и не исправляет ACL.`r`n• Проверяет SID из AD, регистрацию и путь ProfileList, конфликты .bak, наличие непустого NTUSER.DAT и допустимое состояние целевого профиля.`r`n• Проверяет наличие всех папок TargetFolders; затем базовые FullControl для пользователя, SYSTEM и Administrators на корне, NTUSER.DAT и этих папках.`r`n• Отметка «1С» в примечании означает только наличие ibases.v8i. Его отсутствие само по себе не запрещает статус «Готов».`r`n• «Готов» — результат перечисленных проверок. Содержимое файлов, права каждого вложенного файла, вход в RDS и работу приложений проверяют отдельно.",
        "6. Проверка ярлыков|• Ищет файлы .lnk в Desktop и его подпапках только у отмеченных пользователей с зарегистрированным профилем. Поиск и подтвержденное удаление доступны при открытом сеансе пользователя.`r`n• В список попадают ярлыки с непустой целью, для которой не найден доступный файл или каталог. Для целей в AppData другого профиля предлагает существующий файл по тому же относительному пути у владельца ярлыка. Можно выбрать другую цель вручную; отметьте строки галочками и назначьте «Заменить» или «Удалить». Кнопки «Выбрать все» и «Снять выбор» меняют отметки. «Применить» выполняет только назначенные действия у отмеченных строк. Для ярлыков профильной установки Яндекс Браузера общий список показывает старый путь и найденный browser.exe в Program Files. После подтверждения изменяется цель с сохранением аргументов. Меню «Пуск», .url и другие папки не сканируются.`r`n• Перед удалением показывает пользователя, путь ярлыка и цель. Отмена поиска или закрытие списка сохраняет все файлы; удаление требует подтверждения.`r`n• Перед удалением повторно проверяет профиль и цель; изменившиеся ярлыки и ярлыки с доступной целью пропускает.`r`n• Недоступный сетевой ресурс тоже может попасть в список. Проверьте его доступность. Удаляются только перечисленные файлы .lnk.",
        "7. Безопасность C:\Users|• Меняет ACL каталога из поля «Целевая папка» (по умолчанию C:\Users), независимо от отметок пользователей.`r`n• Удаляет на этом каталоге записи Everyone и Users, затем выдает Users чтение/выполнение, создание файлов и подпапок: (CI)(RX,WD,AD). Права наследуются подпапками.`r`n• Не отключает наследование и не выполняет изоляцию каждого профиля. Защита чужих профилей зависит от их фактических ACL.`r`n• На этом шаге коды завершения icacls не проверяются. Сообщение о завершении не заменяет проверку итоговых прав.",
        "Остановить процесс|• Запрашивает остановку текущей операции; если она запустила Robocopy или icacls, завершает этот процесс.`r`n• Запрос AD или файловая проверка могут завершиться до обработки отмены: мгновенная остановка не гарантируется.`r`n• Уже созданные профили, скопированные файлы, примененные ACL и удаленные ярлыки не откатываются.`r`n• Кнопка доступна во время операции. После остановки проверьте журнал и состояние затронутых профилей."
    )
    Add-HelpTab "Основные шаги (1-5 и остановка)" $actionsHelp

    # Вкладка 2: Настройки и параметры
    $settingsHelp = @(
        "Сервер-источник|• Имя или IP компьютера, с которого читаются данные. Программу запускают на целевом сервере.`r`n• Для каждого логина AD источник всегда имеет вид \\сервер\C$\Users\логин. Путь ProfileList исходного сервера программа не определяет.`r`n• Требуются SMB-доступ (TCP 445) и права чтения C$. Каталог профиля на источнике должен совпадать с логином; другое размещение требует отдельной подготовки.",
        "Целевая папка|• Для новых профилей укажите системный каталог профилей Windows (обычно C:\Users). Программа не меняет системное размещение профилей.`r`n• Для существующего профиля используется фактический путь из ProfileList по SID, даже если он отличается от указанного каталога. Фактический путь каждого профиля выводится в журнале миграции.",
        "Целевое OU (SearchBase)|• Область поиска учетных записей в AD в формате Distinguished Name; пользователи никуда не перемещаются.`r`n• Если настройка отсутствует, программа пытается определить корень домена. При неудаче подставляет DC=domain,DC=local — замените его действительным значением.`r`n• Пример: OU=TerminalUsers,OU=Branch,DC=company,DC=local. Используются SID учетных записей текущего AD; сопоставление разных доменов не реализовано.",
        "Лимит неактивности (дней)|• Используется при аудите для автоматической расстановки галочек; по умолчанию 90 дней.`r`n• Галочку получают включенные учетные записи с полученной датой входа в пределах лимита включительно.`r`n• Изменение лимита не пересчитывает текущую таблицу: запустите аудит повторно или измените отметки вручную.",
        "Выбор папок и сохранение настроек|• В основном окне нажмите «Папки переноса» и отметьте нужные папки. Чтобы добавить свою, введите путь относительно профиля и нажмите «Добавить».`r`n• «Применить» включает выбранные папки для текущего запуска. «Отмена» возвращает прежний выбор. Звездочка на кнопке указывает на несохраненные изменения.`r`n• Нажмите «Сохранить параметры». Программа запишет список TargetFolders, сервер-источник, OU, целевую папку, лимит неактивности и остальные настройки в pm_config.json рядом с EXE. После сохранения звездочка исчезнет.`r`n• При первом запуске программа создаст файл, если он отсутствует. При последующих запусках она загрузит сохраненные настройки и отметки папок.",
        "Папки TargetFolders|• Общий для выбранных пользователей массив папок относительно их профиля. Если параметра нет, копируются Desktop, Documents и Downloads.`r`n• Явный список заменяет стандартный. Пустой массив, абсолютные пути, UNC, маски и выход через .. не допускаются.`r`n• Отсутствие заданной папки на источнике отмечается как ошибка миграции; отсутствие на цели выявляется при валидации.`r`n• Два файла настроек запуска 1С копируются отдельно. Примеры дополнительных папок и ограничения переноса приложений приведены на вкладке «Профили и ПО» этой справки.",
        "Журналы и повторный запуск|• LogsDir задает папку журналов Robocopy (по умолчанию C:\Migrate\Logs); она создается при копировании. Заранее создавать C:\Migrate не требуется.`r`n• «Сохранить список» записывает только отмеченных пользователей в AuditList.json в WorkDir (по умолчанию C:\Migrate). Повторное сохранение обновляет файл.`r`n• При следующем запуске нажмите «Загрузить список» с тем же источником и OU. Статусы и активность отражают сохраненные данные. Для обновления выполните аудит заново.`r`n• Уже созданные профили и файлы сохраняются. Повторная миграция или докачка использует их; автоматического продолжения остановленной операции нет."
    )
    Add-HelpTab "Параметры и настройки" $settingsHelp

    # Вкладка 3: Фильтрация и таблица
    $toolsHelp = @(
        "Выбор пользователей|• Для шагов 2-5 используются галочки в столбце «Выбор». Выделение строки цветом само по себе пользователя не выбирает.`r`n• «Выбрать все» и «Снять выбор» действуют на всю таблицу, включая скрытые фильтром строки.`r`n• Без отмеченных пользователей шаги 2-5 не запускаются. Аудит работает независимо от галочек.",
        "Только выбранные (Чекбокс)|• Скрывает неотмеченные строки, но не меняет галочки.`r`n• Сочетается с поиском по логину и ФИО. Отмеченные строки, скрытые поиском, тоже участвуют в операциях.",
        "Поиск по логину и ФИО|• Показывает строки, содержащие введенный текст в логине или ФИО, без учета регистра.`r`n• Поиск меняет только видимость строк. Чтобы исключить пользователя из операции, снимите его галочку.",
        "Счетчик (Всего / Выбрано)|• «Всего» — число строк, загруженных в текущую таблицу; после остановки аудита список может быть неполным.`r`n• «Выбрано» — число всех отмеченных строк, включая скрытые фильтрами.",
        "Готовность и примечание|• После аудита профиль имеет статус «Не проверен». «Скопирован», «Докачан (/XO)» и «ACL настроены» показывают результат соответствующей операции.`r`n• «Готов» выставляет валидация. Это не проверка входа в RDS или работоспособности приложений.`r`n• Валидация обновляет примечание: найденные проблемы, фактический путь и наличие ibases.v8i. Полный текст доступен при наведении на ячейку.",
        "Индикатор прогресса и журнал|• Процент отражает позицию в списке пользователей, ярлыков или этапов операции. Это не процент скопированных байтов и не оценка оставшегося времени.`r`n• Экранный журнал показывает события текущего запуска с временем и уровнем сообщения; отдельный файл этого журнала не ведется.`r`n• Для каждой папки TargetFolders Robocopy пишет файл в LogsDir. При повторном копировании одноименные журналы перезаписываются; отдельного файла Robocopy для настроек 1С нет."
    )
    Add-HelpTab "Таблица и фильтрация" $toolsHelp

    $profilesAndSoftwareHelp = @(
        'Подготовка и проверка профиля Windows|Миграция и докачка получают SID пользователя из AD и находят его профиль в ProfileList. Зарегистрированный профиль используется по указанному в реестре пути. При повторном шаге 2 отсутствующий параметр FullProfile добавляется как DWORD со значением 1; уже заданное значение не меняется. При отсутствии записи и папки профиль создает Windows в системном каталоге профилей; поле «Целевая папка» должно указывать этот каталог.',
        'Существующая папка без записи SID|Если папка логина уже есть, а записи SID нет, шаг 2 регистрирует эту же папку без переименования и создания дубликатов. Подготовка сохраняет существующие файлы, добавляет недостающие из локального шаблона Default и настраивает права. При отсутствии NTUSER.DAT он берется из Default целевого сервера; имеющийся NTUSER.DAT не заменяется. После регистрации выполняется обычное копирование с источника: отличающиеся файлы в TargetFolders могут быть перезаписаны. Для регистрации уже существующей папки используйте шаг 2, затем при необходимости шаг 3.',
        'NTUSER.DAT и ограничения|NTUSER.DAT с источника не переносится. Проверяются наличие и ненулевой размер файла, но не его внутреннее содержимое. Профили с .bak, незавершенной регистрацией, конфликтами путей, загрузкой в системе, признаками повреждения, временного, перемещаемого или обязательного профиля не обрабатываются. Не допускайте входа пользователя на целевой сервер до завершения операций.',
        'Валидация профиля|Валидация не создает профили и не исправляет ошибки. «Готов» означает, что прошли проверки регистрации, состояния профиля, NTUSER.DAT, наличия всех папок TargetFolders и базовых ACL на корне, NTUSER.DAT и этих папках. Содержимое файлов и права каждого вложенного объекта не сравниваются. Наличие ibases.v8i показывается отдельно и не является условием «Готов». Вход в RDS, запуск 1С и других приложений проверяют отдельно.',
        'Специализированное ПО|Программа копирует пользовательские файлы приложений из папок, указанных в TargetFolders. Сами приложения нужно установить на целевом сервере отдельно. Кроме отдельной опции Outlook, настройки из реестра не переносятся. Утилита не выполняет активацию лицензий, установку сертификатов или настройку приложений.',
        'Выбор папок приложений в программе|Настройка переноса: используется pm_config.json рядом с EXE. Если файла еще нет, он создается при первом запуске. Нажмите «Папки переноса» в основном окне, отметьте нужные папки или добавьте свои, затем нажмите «Применить» и «Сохранить параметры». Выбор используется при миграции и валидации, а после сохранения восстанавливается при следующем запуске. Вручную редактировать JSON не требуется. Пути задаются относительно папки профиля пользователя, например AppData\Roaming\Vendor\Product или WorkData. Абсолютные пути, UNC и выход через .. не поддерживаются; данные за пределами профиля переносятся отдельно.',
        'Стандартные папки и настройки 1С|Если TargetFolders отсутствует, используются Desktop, Documents и Downloads. Явно заданный массив заменяет этот список: сохраните в нем стандартные папки, если их тоже нужно перенести. Настройки запуска 1С (ibases.v8i и 1CEStart.cfg) копируются отдельно при наличии на источнике.',
        'Классический Outlook → Outlook 2021/2024 (экспериментально)|В окне подтверждения шага 2 включите перенос Outlook. Завершите сеансы пользователя на обоих серверах. Сбор идет с исходного сервера по SMB. На источнике проверяются профили Outlook 2010, Office 15.0 (2013) и Office 16.0 (2016 и новее). В пределах одной ветки используется профиль по умолчанию или единственный профиль. При неоднозначности администратор выбирает профиль. На цели используется ветка Office 16.0. Выберите папки с PST, которые еще нужно скопировать. Уже перенесенный PST по прежнему пути используется независимо от выбора папок. При отсутствии указанного PST или OST, а для IMAP также при отсутствии файлов .pst/.ost в AppData\Local\Microsoft\Outlook на цели настройки пропускаются со статусом «Внимание». Для автоматической привязки сохраняйте абсолютные пути. На цели существующие PST сохраняются, а при наличии почтового профиля импорт настроек пропускается. После первого входа проверьте авторизацию, доставку POP3 в нужный файл и синхронизацию папок IMAP.',
        'Пример выбора папок Thunderbird|Пример для Thunderbird: в окне «Папки переноса» отметьте Desktop, Documents и Downloads. В поле «Другая папка» введите AppData\Roaming\Thunderbird и нажмите «Добавить». Нажмите «Применить», затем «Сохранить параметры». Этот список будет записан в TargetFolders файла pm_config.json.',
        'Перенос Thunderbird|Thunderbird: при стандартной настройке папка AppData\Roaming\Thunderbird содержит почтовые профили, их настройки и сохраненные локально письма. Перед копированием закройте Thunderbird на обоих серверах; используйте одинаковые версии приложения. Нестандартные папки с почтой внутри профиля Windows добавьте в TargetFolders; папки вне профиля перенесите отдельно. При необходимости настройте новые пути в Thunderbird. После переноса проверьте открытие нужного профиля и доступ к почте.',
        'КриптоПро CSP и электронная подпись|КриптоПро CSP и электронная подпись: автоматическое восстановление сертификатов и ключевых контейнеров в этой программе не реализовано. Добавление папки в TargetFolders копирует только ее файлы. Перенос или подключение контейнера, установку личного сертификата и его привязку к закрытому ключу выполняйте отдельно штатными средствами КриптоПро с учетом типа хранилища или носителя. Универсальной папки, копирование которой восстанавливает электронную подпись, нет. После настройки проверьте подпись под учетной записью пользователя.',
        'Другие приложения|Другие приложения: добавляйте фактические папки с пользовательскими данными внутри профиля. Настройки, зависящие от реестра, имени компьютера или старых абсолютных путей, могут потребовать отдельного восстановления средствами приложения.',
        'Проверка данных приложений|Проверка: список TargetFolders общий для всех выбранных пользователей. Отсутствующая или недоступная папка на источнике отмечается как ошибка миграции; отсутствие папки на целевом сервере дает статус «Внимание» при валидации. Валидация проверяет наличие папок и базовые права доступа, но не сравнивает содержимое файлов и не проверяет работу приложений, почты или электронной подписи.'
    )
    Add-HelpTab "Профили и ПО" $profilesAndSoftwareHelp

    try { $helpForm.ShowDialog($form) | Out-Null }
    finally { $helpForm.Dispose() }
})

$btnAbout.Add_Click({
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.MinimizeBox = $false
    $aboutForm.Text = "О программе — ProfileMigrator"
    $aboutForm.ClientSize = New-Object System.Drawing.Size(900, 700)
    $aboutForm.MinimumSize = New-Object System.Drawing.Size(800, 600)
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $aboutForm.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $aboutForm.ShowInTaskbar = $false
    if ($form.Icon) { $aboutForm.Icon = $form.Icon }

    # Fixed header and footer surround a scrollable page.
    $aboutLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $aboutLayout.Dock = "Fill"
    $aboutLayout.ColumnCount = 1
    $aboutLayout.RowCount = 3
    [void]$aboutLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100)))
    [void]$aboutLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 72)))
    [void]$aboutLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Percent", 100)))
    [void]$aboutLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle("Absolute", 48)))
    $aboutForm.Controls.Add($aboutLayout)

    $aboutHeader = New-Object System.Windows.Forms.Panel
    $aboutHeader.Dock = "Fill"
    $aboutHeader.Margin = New-Object System.Windows.Forms.Padding(0)
    $aboutHeader.BackColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $aboutTitle = New-Object System.Windows.Forms.Label
    $aboutTitle.Text = "RDS Profile Migration Toolkit"
    $aboutTitle.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $aboutTitle.ForeColor = [System.Drawing.Color]::White
    $aboutTitle.AutoSize = $true
    $aboutTitle.Location = New-Object System.Drawing.Point(18, 10)
    $aboutHeader.Controls.Add($aboutTitle)
    $aboutSubtitle = New-Object System.Windows.Forms.Label
    $aboutSubtitle.Text = "Версия $script:AppVersion  •  Перенос пользовательских данных на сервер RDS"
    $aboutSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
    $aboutSubtitle.AutoSize = $true
    $aboutSubtitle.Location = New-Object System.Drawing.Point(19, 43)
    $aboutHeader.Controls.Add($aboutSubtitle)
    $aboutLayout.Controls.Add($aboutHeader, 0, 0)

    $aboutScroll = New-Object System.Windows.Forms.Panel
    $aboutScroll.Dock = "Fill"
    $aboutScroll.AutoScroll = $true
    $aboutScroll.Margin = New-Object System.Windows.Forms.Padding(0)
    $aboutLayout.Controls.Add($aboutScroll, 0, 1)
    $aboutContent = New-Object System.Windows.Forms.TableLayoutPanel
    $aboutContent.Location = New-Object System.Drawing.Point(0, 0)
    $aboutContent.AutoSize = $true
    $aboutContent.AutoSizeMode = "GrowAndShrink"
    $aboutContent.ColumnCount = 1
    $aboutContent.RowCount = 0
    $aboutContent.Padding = New-Object System.Windows.Forms.Padding(14, 12, 14, 14)
    [void]$aboutContent.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle("Percent", 100)))
    $aboutScroll.Controls.Add($aboutContent)

    function Add-AboutLabel ($Text, [bool]$Heading = $false) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Text
        $label.AutoSize = $true
        $label.Dock = "Fill"
        $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
        $label.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
        if ($Heading) {
            $label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
            $label.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
            $label.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 8)
        }
        $aboutContent.Controls.Add($label, 0, $aboutContent.RowCount)
        $aboutContent.RowCount++
    }
    $aboutAuthor = New-Object System.Windows.Forms.LinkLabel
    $aboutAuthor.Text = "Автор: altshift902  •  GitHub: https://github.com/altshift902"
    $aboutAuthor.AutoSize = $true
    $aboutAuthor.Dock = "Fill"
    $aboutAuthor.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
    $aboutAuthor.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $aboutAuthor.LinkColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
    $aboutAuthor.ActiveLinkColor = [System.Drawing.Color]::FromArgb(24, 43, 73)
    $aboutAuthor.VisitedLinkColor = $aboutAuthor.LinkColor
    $aboutAuthor.LinkBehavior = "HoverUnderline"
    $aboutAuthor.Links.Clear()
    [void]$aboutAuthor.Links.Add($aboutAuthor.Text.IndexOf('https://'), 'https://github.com/altshift902'.Length, 'https://github.com/altshift902')
    $aboutAuthor.Add_LinkClicked({
        param($sender, $e)
        try {
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo
            $startInfo.FileName = [string]$e.Link.LinkData
            $startInfo.UseShellExecute = $true
            [void][System.Diagnostics.Process]::Start($startInfo)
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show($sender.FindForm(), "Не удалось открыть ссылку: $($_.Exception.Message)", "GitHub", "OK", "Warning")
        }
    })
    $aboutContent.Controls.Add($aboutAuthor, 0, $aboutContent.RowCount)
    $aboutContent.RowCount++
    Add-AboutLabel "Утилита запускается на целевом сервере и переносит пользовательские данные из обычных папок профилей Windows по SMB. Выполняет поиск учетных записей в AD, подготовку целевого профиля, копирование папок TargetFolders и настроек запуска 1С, назначение ACL, проверку профилей и очистку ярлыков после подтверждения. Миграция, докачка, ACL, валидация и очистка обрабатывают только отмеченных пользователей. Настройка корневого каталога выполняется отдельно для пути из поля «Целевая папка»."
    Add-AboutLabel "Матрица совместимости" $true
    Add-AboutLabel "Сценарии файлового переноса (источник → цель); требуют проверки в вашей среде"

    $matrix = New-Object System.Windows.Forms.DataGridView
    $matrix.Dock = "Top"
    $matrix.ScrollBars = "None"
    $matrix.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    $matrix.ReadOnly = $true
    $matrix.AllowUserToAddRows = $false
    $matrix.AllowUserToDeleteRows = $false
    $matrix.AllowUserToResizeRows = $false
    $matrix.RowHeadersVisible = $false
    $matrix.MultiSelect = $false
    $matrix.SelectionMode = "FullRowSelect"
    $matrix.BackgroundColor = [System.Drawing.Color]::White
    $matrix.BorderStyle = "FixedSingle"
    $matrix.GridColor = [System.Drawing.Color]::FromArgb(226, 232, 240)
    $matrix.EnableHeadersVisualStyles = $false
    $matrix.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $matrix.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $matrix.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $matrix.ColumnHeadersHeight = 32
    $matrix.AutoSizeColumnsMode = "Fill"
    $matrix.AutoSizeRowsMode = "AllCells"
    $matrix.DefaultCellStyle.WrapMode = "True"
    $matrix.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(236, 253, 245)
    $matrix.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(3)
    [void]$matrix.Columns.Add("Source", "Источник (Source)")
    [void]$matrix.Columns.Add("Target", "Назначение (Target)")
    [void]$matrix.Columns.Add("Status", "Статус")
    [void]$matrix.Columns.Add("Notes", "Примечания")
    $matrix.Columns[0].FillWeight = 105
    $matrix.Columns[1].FillWeight = 85
    $matrix.Columns[2].FillWeight = 100
    $matrix.Columns[3].FillWeight = 170
    foreach ($entry in @(
        @("Server 2008 R2", "Server 2025", "Полная поддержка", "Проверьте SMB-доступ и политики аутентификации."),
        @("Server 2012 / 2012 R2", "Server 2025", "Полная поддержка", "Проверьте C$ и соответствие папки логину AD."),
        @("Server 2016", "Server 2025", "Полная поддержка", "Тот же перенос папок TargetFolders по SMB."),
        @("Server 2019 / 2022", "Server 2025", "Полная поддержка", "Скорость зависит от сети, дисков и объема файлов."),
        @("Windows 10 / 11", "Server 2025", "Полная поддержка", "Копирование C:\Users\логин; нужны учетные записи AD.")
    )) { [void]$matrix.Rows.Add([object[]]$entry) }
    # Fit the border to the header and actual rows, including wrapped text.
    $fitMatrixHeight = {
        $matrix.AutoResizeRows([System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells)
        $rowsHeight = $matrix.Rows.GetRowsHeight([System.Windows.Forms.DataGridViewElementStates]::Visible)
        $matrixHeight = $matrix.ColumnHeadersHeight + $rowsHeight + 2
        if ($matrix.Height -ne $matrixHeight) { $matrix.Height = $matrixHeight }
    }
    $matrix.Add_ColumnWidthChanged($fitMatrixHeight)
    & $fitMatrixHeight
    $aboutContent.Controls.Add($matrix, 0, $aboutContent.RowCount)
    $aboutContent.RowCount++
    Add-AboutLabel "Целевой сценарий — Windows Server 2025 (RDS). Для всех строк используется один алгоритм: чтение \\сервер\C$\Users\логин и работа с SID из AD на целевом сервере. Версии ОС и диалект SMB программа не проверяет и не выбирает. Матрица не подтверждает перенос приложений, сертификатов или полную совместимость среды; перед переводом пользователей проверьте доступ и пробную миграцию."
    Add-AboutLabel "Системные требования" $true
    Add-AboutLabel @"
• Запуск от имени администратора на целевом 64-разрядном сервере Windows с графической средой Desktop Experience.
• Windows PowerShell 5.1 (powershell.exe), .NET Framework и Windows Forms целевой ОС.
• Модуль ActiveDirectory (RSAT), доступ к контроллерам домена и чтению учетных записей AD.
• SMB (TCP 445) и права чтения \\Source\C$\Users\логин с целевого сервера; имя папки источника совпадает с логином AD.
• Robocopy, icacls, WScript.Shell и доступ к локальному Win32_UserProfile через CIM.
• Свободное место для профилей; запись рядом с EXE для pm_config.json, в LogsDir и во временный каталог. LogsDir по умолчанию — C:\Migrate\Logs.
• Рабочая область экрана от 1040 × 750 с учетом масштаба Windows.
"@
    Add-AboutLabel "Переносятся папки из конфигурации и два файла настроек запуска 1С (ibases.v8i и 1CEStart.cfg). Полное клонирование профиля, перенос контейнеров FSLogix/UPD и сопоставление разных учетных записей или доменов не реализованы. Роли, лицензии и права входа в RDS настраиваются отдельно. Перед копированием завершите сеансы пользователя и приложения на источнике и цели: загруженный профиль проверяется только на целевом сервере."

    $resizeAbout = {
        $contentWidth = [Math]::Max(400, $aboutScroll.ClientSize.Width - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth)
        $aboutContent.MaximumSize = New-Object System.Drawing.Size($contentWidth, 0)
        $aboutContent.MinimumSize = New-Object System.Drawing.Size($contentWidth, 0)
        foreach ($control in $aboutContent.Controls) {
            if ($control -is [System.Windows.Forms.Label]) {
                $control.MaximumSize = New-Object System.Drawing.Size(($contentWidth - 28), 0)
            }
        }
        $aboutContent.Width = $contentWidth
        $aboutContent.PerformLayout()
        & $fitMatrixHeight
    }
    $aboutScroll.Add_Resize($resizeAbout)
    & $resizeAbout

    $aboutFooter = New-Object System.Windows.Forms.FlowLayoutPanel
    $aboutFooter.Dock = "Fill"
    $aboutFooter.Margin = New-Object System.Windows.Forms.Padding(0)
    $aboutFooter.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
    $aboutFooter.Padding = New-Object System.Windows.Forms.Padding(8)
    $aboutFooter.FlowDirection = "RightToLeft"
    $aboutClose = New-Object System.Windows.Forms.Button
    $aboutClose.Text = "Закрыть"
    $aboutClose.Size = New-Object System.Drawing.Size(110, 30)
    $aboutClose.BackColor = [System.Drawing.Color]::FromArgb(24, 43, 73)
    $aboutClose.ForeColor = [System.Drawing.Color]::White
    $aboutClose.FlatStyle = "Flat"
    $aboutClose.FlatAppearance.BorderSize = 0
    $aboutClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $aboutClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $aboutFooter.Controls.Add($aboutClose)
    $aboutLayout.Controls.Add($aboutFooter, 0, 2)
    $aboutForm.CancelButton = $aboutClose
    try { [void]$aboutForm.ShowDialog($form) }
    finally { $aboutForm.Dispose() }
})
# --- Логика событий ---
$btnSaveCfg.Add_Click({
    $Global:Config.SourceServer = $txtSrc.Text.Trim()
    $Global:Config.TargetOU = $txtOU.Text.Trim()
    $Global:Config.ProfilesRoot = $txtDst.Text.Trim()
    $Global:Config.InactiveDaysLimit = [int]$numLimit.Value
    try { Save-Config }
    catch {
        Write-GuiLog $_.Exception.Message "ERROR"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Ошибка сохранения", "OK", "Error") | Out-Null
        return
    }
    $btnFolders.Text = "Папки переноса ($($Global:Config.TargetFolders.Count))"
    $tooltip.SetToolTip($btnFolders, ($Global:Config.TargetFolders -join '; '))
    Write-GuiLog "Параметры успешно сохранены в $($Global:ConfigFile)" "SUCCESS"
    [System.Windows.Forms.MessageBox]::Show("Параметры успешно сохранены!", "Уведомление", "OK", "Information")
})

$btnSelectAll.Add_Click({
    foreach ($row in $gridUsers.Rows) { $row.Cells[0].Value = $true }
    Update-CountersAndFilter
})

$btnDeselectAll.Add_Click({
    foreach ($row in $gridUsers.Rows) { $row.Cells[0].Value = $false }
    Update-CountersAndFilter
})

$chkOnlySelected.Add_CheckedChanged({
    Update-CountersAndFilter
})

$txtSearch.Add_TextChanged({
    Update-CountersAndFilter
})

$gridUsers.Add_CellContentClick({
    param($s, $e)
    if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
        $gridUsers.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        Update-CountersAndFilter
    }
})

# --- Неблокирующий запуск утилит, отмена и управление состоянием интерфейса ---
$Global:CancelRequested = $false
$Global:CurrentProc = $null

function Invoke-CliCommand ($ExePath, [string[]]$ArgsList) {
    if ($Global:CancelRequested) { return -1 }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ExePath
    $psi.Arguments = ($ArgsList | ForEach-Object {
        if ($_ -match '[\s"]') { '"{0}"' -f ($_ -replace '"', '\"') } else { $_ }
    }) -join ' '
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

    $proc = [System.Diagnostics.Process]::Start($psi)
    $Global:CurrentProc = $proc
    if ($proc) {
        # Пока процесс работает (копирование Robocopy, icacls), прокачиваем очередь сообщений Windows,
        # благодаря чему окно никогда не зависает и не переходит в статус "Не отвечает"
        while (-not $proc.WaitForExit(100)) {
            [System.Windows.Forms.Application]::DoEvents()
            if ($Global:CancelRequested) {
                try {
                    if (-not $proc.HasExited) {
                        $proc.Kill()
                        Write-GuiLog "Процесс $($proc.ProcessName) принудительно остановлен." "WARN"
                    }
                } catch {}
                break
            }
        }
        $code = if ($proc.HasExited) { $proc.ExitCode } else { -1 }
        $proc.Dispose()
        $Global:CurrentProc = $null
        return $code
    }
    return -1
}

function Measure-FolderDataSize ($FolderPath) {
    if ($Global:CancelRequested) { return @{ Bytes = [int64]0; Files = [int64]0; Cancelled = $true } }
    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        return @{ Bytes = [int64]0; Files = [int64]0; Cancelled = $false }
    }
    $dummy = Join-Path $env:TEMP ('PM_DryRun_' + [guid]::NewGuid().ToString('N'))
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "robocopy.exe"
    $psi.Arguments = "`"$FolderPath`" `"$dummy`" /L /E /BYTES /NJH /NFL /NDL /XJ /R:0 /W:0 /MT:16"
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.CreateNoWindow = $true
    try {
        $psi.StandardOutputEncoding = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
    } catch {}

    $proc = [System.Diagnostics.Process]::Start($psi)
    $Global:CurrentProc = $proc

    while (-not $proc.WaitForExit(100)) {
        [System.Windows.Forms.Application]::DoEvents()
        if ($Global:CancelRequested) {
            try { if (-not $proc.HasExited) { $proc.Kill() } } catch {}
            break
        }
    }
    $out = $proc.StandardOutput.ReadToEnd()
    $proc.Dispose()
    $Global:CurrentProc = $null

    if ($Global:CancelRequested) {
        return @{ Bytes = [int64]0; Files = [int64]0; Cancelled = $true }
    }

    $bytes = [int64]0
    $files = [int64]0
    $statLines = @()
    foreach ($line in ($out -split "\r?\n")) {
        if ($line -match '^\s*(?:Files|Файлов)\s*:\s*(\d+)') { $files = [int64]$matches[1] }
        elseif ($line -match '^\s*(?:Bytes|Байт)\s*:\s*(\d+)') { $bytes = [int64]$matches[1] }
        elseif ($line -match '^\s*[\p{L}\s]+\s*:\s*(\d+)') { $statLines += [int64]$matches[1] }
    }
    if ($bytes -eq 0 -and $statLines.Count -ge 3) {
        if ($files -eq 0) { $files = $statLines[1] }
        $bytes = $statLines[2]
    }
    return @{ Bytes = $bytes; Files = $files; Cancelled = $false }
}

function Get-UserDataTransferSize ([string]$Server, [string]$UserName, [string[]]$Folders) {
    if ($Global:CancelRequested) { return @{ Bytes = [int64]0; Files = [int64]0; Exists = $false; Cancelled = $true; Details = @{} } }
    $SrcBase = "\\$Server\C$\Users\$UserName"
    if (-not (Test-Path -LiteralPath $SrcBase -PathType Container)) {
        return @{ Bytes = [int64]0; Files = [int64]0; Exists = $false; Cancelled = $false; Details = @{} }
    }
    $totalBytes = [int64]0
    $totalFiles = [int64]0
    $details = [ordered]@{}

    foreach ($folder in $Folders) {
        if ($Global:CancelRequested) { return @{ Bytes = $totalBytes; Files = $totalFiles; Exists = $true; Cancelled = $true; Details = $details } }
        $s = Join-Path $SrcBase $folder
        if (Test-Path -LiteralPath $s -PathType Container) {
            $meas = Measure-FolderDataSize $s
            if ($meas.Cancelled) { return @{ Bytes = $totalBytes; Files = $totalFiles; Exists = $true; Cancelled = $true; Details = $details } }
            $totalBytes += $meas.Bytes
            $totalFiles += $meas.Files
            $details[$folder] = @{ Bytes = $meas.Bytes; Files = $meas.Files }
        }
    }

    # 1С:Предприятие (ibases.v8i, 1CEStart.cfg)
    $src1c = Join-Path $SrcBase "AppData\Roaming\1C\1CEStart"
    if (Test-Path -LiteralPath $src1c -PathType Container) {
        $files1c = 0
        $bytes1c = [int64]0
        foreach ($file in @('ibases.v8i', '1CEStart.cfg')) {
            $fp = Join-Path $src1c $file
            if (Test-Path -LiteralPath $fp -PathType Leaf) {
                try {
                    $fi = Get-Item -LiteralPath $fp -Force -ErrorAction SilentlyContinue
                    if ($fi) {
                        $bytes1c += $fi.Length
                        $files1c++
                    }
                } catch {}
            }
        }
        if ($files1c -gt 0) {
            $totalBytes += $bytes1c
            $totalFiles += $files1c
            $details['1C (Настройки)'] = @{ Bytes = $bytes1c; Files = $files1c }
        }
    }

    return @{ Bytes = $totalBytes; Files = $totalFiles; Exists = $true; Cancelled = $false; Details = $details }
}

function Update-Progress ($val, $status = "") {
    $v = [math]::Max(0, [math]::Min(100, [int]$val))
    $Global:pbar.Value = $v
    $Global:pbar.Invalidate()
    $Global:lblProgressPercent.Text = "$v%"
    if ($status) {
        $Global:lblProgressTitle.Text = "Прогресс операции: $status"
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# These buttons remain interactive while the form displays the wait cursor.
foreach ($button in @($btnHelp, $btnAbout, $btnExpandLog, $btnStop)) {
    $button.Cursor = [System.Windows.Forms.Cursors]::Arrow
}

function Set-UiBusy ($busy, $statusText = "") {
    $grpSettings.Enabled = -not $busy
    $gridUsers.Enabled = -not $busy
    foreach ($stepButton in @($btnAudit, $btnEstimate, $btnSync, $btnDelta, $btnValidate, $btnShortcuts)) {
        $stepButton.Enabled = -not $busy
        $stepButton.BackColor = if ($busy) { $stepBusyColor } else { $stepReadyColor }
    }
    $btnSaveCfg.Enabled = -not $busy
    $btnLoadList.Enabled = -not $busy
    $btnSaveList.Enabled = -not $busy -and $gridUsers.Rows.Count -gt 0
    $btnSelectAll.Enabled = -not $busy
    $btnDeselectAll.Enabled = -not $busy
    $txtSearch.Enabled = -not $busy
    $chkOnlySelected.Enabled = -not $busy

    if ($busy) {
        $Global:CancelRequested = $false
        $btnStop.Enabled = $true
        $btnStop.BackColor = [System.Drawing.Color]::FromArgb(231, 76, 60)
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        if ($statusText) {
            $Global:lblProgressTitle.Text = "Прогресс операции: $statusText"
        }
    } else {
        $btnStop.Enabled = $false
        $btnStop.BackColor = [System.Drawing.Color]::FromArgb(160, 170, 180)
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    [System.Windows.Forms.Application]::DoEvents()
}

$btnStop.Add_Click({
    if (-not $Global:CancelRequested) {
        $Global:CancelRequested = $true
        $btnStop.Enabled = $false
        $btnStop.BackColor = [System.Drawing.Color]::FromArgb(160, 170, 180)
        Write-GuiLog "ЗАПРОС ОСТАНОВКИ: Прерывание текущего процесса..." "WARN"
        Update-Progress $Global:pbar.Value "Остановка текущей операции..."
        if ($Global:CurrentProc) {
            try {
                if (-not $Global:CurrentProc.HasExited) {
                    $Global:CurrentProc.Kill()
                }
            } catch {}
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
})

function Ensure-ActiveDirectoryModule {
    if (Get-Module -ListAvailable ActiveDirectory) {
        return $true
    }

    $isServer = $false
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $isServer = ($os.ProductType -ne 1)
    } catch {
        $isServer = ($null -ne (Get-Command Install-WindowsFeature -ErrorAction SilentlyContinue))
    }

    $msg = "Модуль Active Directory (RSAT) не найден в системе.`r`n`r`n"
    if ($isServer) {
        $msg += "Выполнить автоматическую установку компонента Windows Server (RSAT-AD-PowerShell)?`r`n(Процесс займет 1-2 минуты)"
    } else {
        $msg += "Выполнить автоматическую установку RSAT Active Directory через Windows Capability?`r`n(Требуется подключение к интернету)"
    }

    $choice = [System.Windows.Forms.MessageBox]::Show(
        $msg,
        "Установка RSAT Active Directory",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        throw "Для работы необходим модуль ActiveDirectory (RSAT). Установка отменена пользователем."
    }

    Write-GuiLog "Запуск автоматической установки RSAT Active Directory..." "INFO"
    Update-Progress 15 "Установка модуля RSAT Active Directory..."

    try {
        if ($isServer) {
            $cmd = Get-Command Install-WindowsFeature -ErrorAction SilentlyContinue
            if (-not $cmd) {
                throw "Команда Install-WindowsFeature не найдена. Установите RSAT через Диспетчер серверов."
            }
            Write-GuiLog "Выполняется установка компонента Windows Server: RSAT-AD-PowerShell..." "INFO"
            $result = Install-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction Stop
            if (-not $result.Success) {
                throw "Не удалось установить RSAT-AD-PowerShell (ExitCode: $($result.ExitCode))."
            }
        } else {
            $cmd = Get-Command Add-WindowsCapability -ErrorAction SilentlyContinue
            if (-not $cmd) {
                throw "Команда Add-WindowsCapability не найдена. Установите компонент через Дополнительные компоненты Windows."
            }
            Write-GuiLog "Выполняется установка компонента Windows Capability: Rsat.ActiveDirectory.DS-LDS.Tools..." "INFO"
            Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop | Out-Null
        }

        Update-Progress 90 "Проверка доступности модуля..."
        if (-not (Get-Module -ListAvailable ActiveDirectory)) {
            throw "Компонент установлен, но модуль ActiveDirectory не обнаружен в PSModulePath."
        }

        Write-GuiLog "Модуль RSAT Active Directory успешно установлен!" "SUCCESS"
        Update-Progress 100 "RSAT Active Directory установлен"
        return $true
    } catch {
        Write-GuiLog "Ошибка установки RSAT Active Directory: $($_.Exception.Message)" "ERROR"
        throw
    }
}
# Обработчик 1: Аудит
$btnAudit.Add_Click({
    Set-UiBusy $true "Аудит OU в Active Directory..."
    Update-Progress 0 "Опрос контроллеров домена..."

    try {
        Write-GuiLog "Запуск аудита OU: $($txtOU.Text)..." "INFO"
        Ensure-ActiveDirectoryModule
        Import-Module ActiveDirectory -ErrorAction Stop
        $DomainControllers = Get-ADDomainController -Filter * -ErrorAction Stop
        Write-GuiLog "Обнаружено контроллеров домена: $($DomainControllers.Count)" "INFO"

        $Users = Get-ADUser -SearchBase $txtOU.Text -Filter * -Properties DisplayName, Enabled, WhenCreated, lastLogonTimestamp -ErrorAction Stop
        if (-not $Users) {
            Write-GuiLog "Пользователи в OU не найдены!" "WARN"
            Update-Progress 0 "Пользователи не найдены"
            return
        }

        $gridUsers.Rows.Clear()
        $script:auditListContext = @{ Server=$txtSrc.Text.Trim(); OU=$txtOU.Text.Trim(); AuditTime=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }
        $total = $Users.Count
        $cur = 0

        foreach ($User in $Users) {
            if ($Global:CancelRequested) {
                Write-GuiLog "Аудит прерван пользователем." "WARN"
                break
            }

            $cur++
            $pct = [math]::Min(100, [int](($cur / $total) * 100))
            Update-Progress $pct "Аудит пользователей: $cur из $total"
            if ($cur % 5 -eq 0 -or $cur -eq $total) {
                Write-GuiLog "Аудит пользователей: $cur/$total..." "INFO"
            }

            $SamAccountName = $User.SamAccountName
            $LatestLogonTime = 0
            foreach ($DC in $DomainControllers) {
                if ($Global:CancelRequested) { break }
                try {
                    $DCUser = Get-ADUser -Identity $SamAccountName -Server $DC.HostName -Properties lastLogon -ErrorAction Stop
                    if ($DCUser.lastLogon -and ($DCUser.lastLogon -gt $LatestLogonTime)) {
                        $LatestLogonTime = $DCUser.lastLogon
                    }
                } catch {}
                [System.Windows.Forms.Application]::DoEvents()
            }

            $LastLogonDate = if ($LatestLogonTime -gt 0) { [DateTime]::FromFileTime($LatestLogonTime) } else { $null }
            $DaysAgo = if ($LastLogonDate) { [int]((Get-Date) - $LastLogonDate).TotalDays } else { 999999 }
            $isActive = ($User.Enabled -and $DaysAgo -le $numLimit.Value)

            $dateStr = if ($LastLogonDate) { $LastLogonDate.ToString("yyyy-MM-dd HH:mm") } else { "Никогда" }
            $daysStr = if ($DaysAgo -eq 999999) { "-" } else { $DaysAgo.ToString() }

            $rowIdx = $gridUsers.Rows.Add($isActive, $SamAccountName, $User.DisplayName, $dateStr, $daysStr, "Не проверен", $(if ($isActive) { "Активен" } else { "Неактивен" }), "—")
            if (-not $isActive) {
                $gridUsers.Rows[$rowIdx].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        Update-CountersAndFilter
        if ($Global:CancelRequested) {
            Write-GuiLog "Аудит остановлен. Загружено пользователей: $($gridUsers.Rows.Count)" "WARN"
            Update-Progress $Global:pbar.Value "Аудит остановлен пользователем"
        } else {
            Write-GuiLog "Аудит завершен. Всего пользователей: $($gridUsers.Rows.Count)" "SUCCESS"
            Update-Progress 100 "Аудит завершен ($($gridUsers.Rows.Count) пользователей)"
        }
    } catch {
        Write-GuiLog "Ошибка аудита: $($_.Exception.Message)" "ERROR"
        Update-Progress 0 "Ошибка аудита"
    } finally {
        Set-UiBusy $false
    }
})

function Resolve-ProfileIdentity ([string]$UserName) {
    if (Get-Command -Name 'Ensure-ActiveDirectoryModule' -ErrorAction SilentlyContinue) { Ensure-ActiveDirectoryModule }
    if ([string]::IsNullOrWhiteSpace($UserName) -or $UserName -match '[\\/:*?"<>|\x00-\x1f]' -or
        $UserName -in @('.', '..') -or $UserName -match '[ .]$') { throw 'Недопустимый логин пользователя.' }
    Import-Module ActiveDirectory -ErrorAction Stop
    $account = Get-ADUser -Identity $UserName -Properties SID -ErrorAction Stop
    if ($account.SamAccountName -ne $UserName) { throw "Неоднозначная учетная запись: $UserName" }
    $sid = [System.Security.Principal.SecurityIdentifier]$account.SID.Value
    if (-not $sid.IsAccountSid()) { throw "Не удалось определить SID пользователя $UserName" }
    return [pscustomobject]@{ UserName = $account.SamAccountName; Sid = $sid.Value }
}

function ConvertTo-LocalProfilePath ([string]$Path) {
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if ($expanded -notmatch '^[A-Za-z]:\\' -or $expanded -match '[*?"<>|]') {
        throw "Ожидается абсолютный локальный путь профиля: $Path"
    }
    return [IO.Path]::GetFullPath($expanded).TrimEnd('\')
}

function Get-ProfileRegistrySnapshot {
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    $profileList = $null
    try {
        $profileList = $baseKey.OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList')
        if ($null -eq $profileList) { throw 'Раздел ProfileList отсутствует.' }
        $entries = @(foreach ($name in $profileList.GetSubKeyNames()) {
            # A staging key is not a registered profile, including after a process crash.
            if ($name -match '^ProfileMigrator\.Pending\.[0-9a-f]{32}$') { continue }
            $key = $profileList.OpenSubKey($name)
            if ($null -eq $key) { throw "Не удалось прочитать ProfileList\$name" }
            try { [pscustomobject]@{
                Sid = $name
                Path = [string]$key.GetValue('ProfileImagePath')
                PendingRegistration = $key.GetValue('ProfileMigratorRegistrationPending')
            } }
            finally { $key.Dispose() }
        })
        return [pscustomobject]@{
            Root = [string]$profileList.GetValue('ProfilesDirectory')
            DefaultPath = [string]$profileList.GetValue('Default')
            Entries = $entries
        }
    } finally {
        if ($profileList) { $profileList.Dispose() }
        $baseKey.Dispose()
    }
}

function Assert-ProfileLocation ($Identity, [string]$Path, $Snapshot) {
    $path = ConvertTo-LocalProfilePath $Path
    $root = ConvertTo-LocalProfilePath $Snapshot.Root
    $defaultPath = if ($Snapshot.DefaultPath) { ConvertTo-LocalProfilePath $Snapshot.DefaultPath } else { Join-Path $root 'Default' }
    if ($path -eq [IO.Path]::GetPathRoot($path).TrimEnd('\') -or $path -eq $root -or
        $path -eq $defaultPath -or $path.StartsWith($defaultPath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase) -or
        $path -eq (Join-Path $root 'Default') -or $path -eq (Join-Path $root 'Public')) {
        throw "Небезопасный путь профиля: $path"
    }
    foreach ($protected in @($env:windir, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData)) {
        if ($protected -and ($path -eq $protected -or $path.StartsWith($protected.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase))) {
            throw "Путь профиля указывает на системный каталог: $path"
        }
    }
    foreach ($entry in $Snapshot.Entries) {
        if ($entry.Sid -ne $Identity.Sid -and $entry.Path -and (ConvertTo-LocalProfilePath $entry.Path) -eq $path) {
            throw "Каталог $path также зарегистрирован для $($entry.Sid)."
        }
    }
    return $path
}

function Find-RegisteredProfile ($Identity, $Snapshot) {
    if (@($Snapshot.Entries | Where-Object { $_.Sid -eq "$($Identity.Sid).bak" }).Count) {
        throw "Для SID $($Identity.Sid) найдена запись .bak. Сначала восстановите профиль средствами Windows."
    }
    $matches = @($Snapshot.Entries | Where-Object { $_.Sid -eq $Identity.Sid })
    if ($matches.Count -eq 0) { return $null }
    if ($matches.Count -ne 1) { throw 'Обнаружены дублирующиеся записи SID.' }
    if ($matches[0].PendingRegistration) {
        throw "Для SID $($Identity.Sid) найдена незавершенная регистрация ProfileMigrator. Проверьте запись ProfileList перед повторной миграцией; данные папки сохранены."
    }
    $path = Assert-ProfileLocation $Identity $matches[0].Path $Snapshot
    return [pscustomobject]@{ UserName = $Identity.UserName; Sid = $Identity.Sid; Path = $path }
}

function Assert-ProfileOffline ($Profile, [switch]$AllowLoaded) {
    $records = @(Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($Profile.Sid)'" -ErrorAction Stop)
    foreach ($record in $records) {
        if ($record.Loaded -and -not $AllowLoaded) { throw 'Профиль загружен. Завершите все сеансы пользователя на целевом сервере.' }
        if ($record.Special -or ($record.Status -band 9)) { throw 'Профиль системный, временный или поврежденный.' }
        if ($record.RoamingConfigured -or ($record.Status -band 6)) { throw 'Перемещаемые и обязательные профили требуют отдельного сценария миграции.' }
        if ($record.LocalPath -and (ConvertTo-LocalProfilePath $record.LocalPath) -ne $Profile.Path) {
            throw 'Путь профиля в Win32_UserProfile не совпадает с ProfileList.'
        }
    }
}

function Assert-ProfileDirectory ([string]$Path) {
    $directory = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if (-not $directory.PSIsContainer) { throw 'Путь профиля не является каталогом.' }
    $ancestor = $directory
    while ($ancestor) {
        if ($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Путь профиля проходит через ссылку: $($ancestor.FullName)" }
        $ancestor = $ancestor.Parent
    }
}

function Assert-ProfileFiles ($Profile) {
    Assert-ProfileDirectory $Profile.Path
    $hive = Get-Item -LiteralPath (Join-Path $Profile.Path 'NTUSER.DAT') -Force -ErrorAction Stop
    if ($hive.PSIsContainer -or $hive.Length -eq 0 -or ($hive.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw 'NTUSER.DAT отсутствует, пуст или является ссылкой.'
    }
}

function Initialize-ProfileApi {
    if (-not ('ProfileMigrator.NativeProfile' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
namespace ProfileMigrator {
    public static class NativeProfile {
        [DllImport("userenv.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern int CreateProfile(string sid, string userName, StringBuilder profilePath, uint capacity);
    }
}
'@ -ErrorAction Stop
    }
}

function Invoke-CreateWindowsProfile ($Identity) {
    Initialize-ProfileApi
    # Use the conventional MAX_PATH output buffer for the profile-service API,
    # not the extended filesystem path limit (32768).
    $buffer = New-Object Text.StringBuilder 260
    $result = [ProfileMigrator.NativeProfile]::CreateProfile($Identity.Sid, $Identity.UserName, $buffer, [uint32]$buffer.Capacity)
    return [pscustomobject]@{ HResult = $result; Path = $buffer.ToString() }
}

function Initialize-ProfileRegistrationApi {
    if (-not ('ProfileMigrator.ProfileRegistrationSession' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
namespace ProfileMigrator {
    public sealed class ProfileRegistrationSession : IDisposable {
        RegistryKey root;
        public RegistryKey Parent { get; private set; }
        public RegistryKey Key { get; private set; }
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int RegCreateKeyExW(SafeRegistryHandle parent, string subkey, uint reserved,
            string keyClass, uint options, uint access, IntPtr security,
            out SafeRegistryHandle key, out uint disposition);

        // Only this adapter touches the real registry. Tests replace it with an in-memory store.
        // The disposition check prevents writing to a SID key created concurrently.
        public static ProfileRegistrationSession CreateNew(string sidText) {
            var sid = new SecurityIdentifier(sidText);
            if (!sid.IsAccountSid() || sid.Value != sidText) throw new ArgumentException("Invalid account SID");
            var session = new ProfileRegistrationSession();
            SafeRegistryHandle handle = null;
            try {
                session.root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                session.Parent = session.root.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", true);
                if (session.Parent == null) throw new IOException("ProfileList parent key is missing.");
                uint disposition;
                int error = RegCreateKeyExW(session.Parent.Handle, sidText, 0, null, 0,
                    0x2011F, IntPtr.Zero, out handle, out disposition);
                if (error != 0) throw new Win32Exception(error);
                if (disposition != 1) {
                    throw new IOException("SID already exists in ProfileList; nothing was overwritten.");
                }
                session.Key = RegistryKey.FromHandle(handle, RegistryView.Registry64);
                handle = null; // RegistryKey now owns the SafeRegistryHandle.
                return session;
            } catch {
                if (handle != null) handle.Dispose();
                session.Dispose();
                throw;
            }
        }
        public void Dispose() {
            if (Key != null) { Key.Dispose(); Key = null; }
            if (Parent != null) { Parent.Dispose(); Parent = null; }
            if (root != null) { root.Dispose(); root = null; }
        }
    }
}
'@ -ErrorAction Stop
    }
}

function New-ProfileRegistrationSession ([string]$Sid) {
    Initialize-ProfileRegistrationApi
    return [ProfileMigrator.ProfileRegistrationSession]::CreateNew($Sid)
}

function Write-NewProfileRegistration ($Profile) {
    $sid = New-Object Security.Principal.SecurityIdentifier($Profile.Sid)
    if (-not $sid.IsAccountSid() -or $sid.Value -cne $Profile.Sid) { throw 'Недопустимый SID профиля.' }
    $path = ConvertTo-LocalProfilePath $Profile.Path
    $sidBytes = New-Object byte[] $sid.BinaryLength
    $sid.GetBinaryForm($sidBytes, 0)
    if ($Global:CancelRequested) { throw 'Операция остановлена. Регистрация SID не выполнена.' }

    $session = New-ProfileRegistrationSession $sid.Value
    $token = [guid]::NewGuid().ToString('N')
    $marker = 'ProfileMigratorRegistrationPending'
    $completed = $false
    try {
        $key = $session.Key
        $key.SetValue($marker, $token, [Microsoft.Win32.RegistryValueKind]::String)
        $key.SetValue('Sid', $sidBytes, [Microsoft.Win32.RegistryValueKind]::Binary)
        foreach ($name in @('Flags', 'State', 'RefCount')) {
            $key.SetValue($name, 0, [Microsoft.Win32.RegistryValueKind]::DWord)
        }
        $key.SetValue('FullProfile', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        # Publish the path only after all required values exist. No UI events during writes.
        $key.SetValue('ProfileImagePath', $path, [Microsoft.Win32.RegistryValueKind]::ExpandString)
        $key.Flush()
        $key.DeleteValue($marker, $true)
        $completed = $true
    } catch {
        $writeError = $_.Exception.Message
        # Never delete an existing registration, an externally changed key, or any subtree.
        $current = $null
        try {
            $current = $session.Parent.OpenSubKey($sid.Value)
            if ($null -ne $current) {
                $names = @($current.GetValueNames())
                $unknown = @($names | Where-Object { $_ -notin @($marker, 'Sid', 'Flags', 'State', 'RefCount', 'FullProfile', 'ProfileImagePath') })
                $storedPath = [string]$current.GetValue('ProfileImagePath')
                $changedKnownValue = $false
                foreach ($name in @('Flags', 'State', 'RefCount')) {
                    if ($name -in $names -and $current.GetValue($name) -ne 0) { $changedKnownValue = $true }
                }
                if ('FullProfile' -in $names -and $current.GetValue('FullProfile') -ne 1) { $changedKnownValue = $true }
                if ('Sid' -in $names -and [Convert]::ToBase64String([byte[]]$current.GetValue('Sid')) -cne [Convert]::ToBase64String($sidBytes)) {
                    $changedKnownValue = $true
                }
                $owned = ($current.GetValue($marker) -ceq $token -and $session.Key.GetValue($marker) -ceq $token)
                $empty = ($names.Count -eq 0 -and $session.Key.ValueCount -eq 0)
                if (($owned -or $empty) -and -not $unknown.Count -and -not $changedKnownValue -and $current.SubKeyCount -eq 0 -and
                    (-not $storedPath -or $storedPath -eq $path)) {
                    $session.Parent.DeleteSubKey($sid.Value, $false)
                } else { throw 'Запись изменена другим процессом; автоматическое удаление запрещено.' }
            }
        } catch {
            throw "Ошибка регистрации: $writeError. Запись SID не удалена: $($_.Exception.Message) Проверьте ProfileList; файлы пользователя не удалялись."
        } finally { if ($current) { $current.Dispose() } }
        throw "Ошибка регистрации: $writeError. Новая запись SID удалена; файлы пользователя сохранены."
    } finally { $session.Dispose() }
    if (-not $completed) { throw 'Регистрация SID не завершена.' }
}

function Open-RegisteredProfileKey ([string]$Sid) {
    $identity = New-Object Security.Principal.SecurityIdentifier($Sid)
    if (-not $identity.IsAccountSid() -or $identity.Value -cne $Sid) { throw 'Недопустимый SID профиля.' }
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        # Open only: never recreate a key that disappeared during validation.
        return $baseKey.OpenSubKey(('SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' + $Sid), $true)
    } finally { $baseKey.Dispose() }
}

function Ensure-ProfileFullProfile ($Profile) {
    if ($Global:CancelRequested) { throw 'Операция остановлена.' }
    $registered = Find-RegisteredProfile $Profile (Get-ProfileRegistrySnapshot)
    if (-not $registered -or $registered.Path -ne $Profile.Path) { throw 'Регистрация профиля изменилась. Повторите проверку SID и пути.' }
    Assert-ProfileFiles $Profile
    Assert-ProfileOffline $Profile
    $key = Open-RegisteredProfileKey $Profile.Sid
    if ($null -eq $key) { throw 'Запись SID исчезла из ProfileList. Автоматическое пересоздание отменено.' }
    try {
        if ($key.GetValue('ProfileMigratorRegistrationPending')) { throw 'Обнаружена незавершенная регистрация профиля.' }
        if ((ConvertTo-LocalProfilePath ([string]$key.GetValue('ProfileImagePath'))) -ne $Profile.Path) {
            throw 'Путь ProfileImagePath изменился; параметры профиля не обновлялись.'
        }
        $storedSid = $key.GetValue('Sid')
        if ($null -ne $storedSid -and (New-Object Security.Principal.SecurityIdentifier($storedSid, 0)).Value -ne $Profile.Sid) {
            throw 'Значение Sid не совпадает с именем записи ProfileList.'
        }
        # Do not overwrite values assigned by Windows or another administrator.
        if ('FullProfile' -in $key.GetValueNames()) { return }
        if ($Global:CancelRequested) { throw 'Операция остановлена.' }
        $key.SetValue('FullProfile', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        Write-GuiLog "[$($Profile.UserName)] В существующую запись SID добавлен FullProfile = 1 (DWORD)." 'SUCCESS'
    } finally { $key.Dispose() }
}

function Assert-UnregisteredProfile ($Profile) {
    $snapshot = Get-ProfileRegistrySnapshot
    $registered = Find-RegisteredProfile $Profile $snapshot
    if ($registered) { throw "Регистрация SID уже появилась: $($registered.Path). Повторите операцию после проверки." }
    [void](Assert-ProfileLocation $Profile $Profile.Path $snapshot)
    Assert-ProfileDirectory $Profile.Path
    Assert-ProfileOffline $Profile
    # Also reject a directory reported by CIM for a different SID.
    foreach ($record in @(Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop)) {
        if ($record.LocalPath -and (ConvertTo-LocalProfilePath $record.LocalPath) -eq $Profile.Path -and $record.SID -ne $Profile.Sid) {
            throw "Каталог $($Profile.Path) используется другим профилем Windows ($($record.SID))."
        }
    }
}

function Copy-MissingProfileTemplate ($Profile, [string]$DefaultPath) {
    Assert-ProfileDirectory $DefaultPath
    Assert-ProfileFiles ([pscustomobject]@{ Path = $DefaultPath })
    # A partial/corrupt hive is never replaced with Default.
    $targetHive = Join-Path $Profile.Path 'NTUSER.DAT'
    if (Test-Path -LiteralPath $targetHive) { Assert-ProfileFiles $Profile }
    $queue = New-Object 'System.Collections.Generic.Queue[string]'
    $queue.Enqueue('')
    $preservedHives = @{}
    while ($queue.Count -gt 0) {
        if ($Global:CancelRequested) { throw 'Подготовка профиля остановлена. Регистрация SID не выполнена.' }
        $relativeDirectory = $queue.Dequeue()
        $sourceDirectory = if ($relativeDirectory) { Join-Path $DefaultPath $relativeDirectory } else { $DefaultPath }
        foreach ($item in @(Get-ChildItem -LiteralPath $sourceDirectory -Force -ErrorAction Stop)) {
            if ($Global:CancelRequested) { throw 'Подготовка профиля остановлена. Регистрация SID не выполнена.' }
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
            $relative = if ($relativeDirectory) { Join-Path $relativeDirectory $item.Name } else { $item.Name }
            Assert-ProfileTargetPath $Profile.Path $relative
            $destination = Join-Path $Profile.Path $relative
            # Never mix Default's transaction logs with an existing user's hive.
            $hiveName = if ($item.Name -like 'NTUSER*') { 'NTUSER.DAT' } elseif ($item.Name -like 'UsrClass.dat*') { 'UsrClass.dat' } else { $null }
            if ($hiveName) {
                $hiveRelative = if ($relativeDirectory) { Join-Path $relativeDirectory $hiveName } else { $hiveName }
                if (-not $preservedHives.ContainsKey($hiveRelative)) {
                    $preservedHives[$hiveRelative] = Test-Path -LiteralPath (Join-Path $Profile.Path $hiveRelative)
                    $targetDirectory = if ($relativeDirectory) { Join-Path $Profile.Path $relativeDirectory } else { $Profile.Path }
                    if (-not $preservedHives[$hiveRelative] -and [IO.Directory]::Exists($targetDirectory)) {
                        $orphanLogs = @(Get-ChildItem -LiteralPath $targetDirectory -Force -File -ErrorAction Stop |
                            Where-Object { $_.Name -like "$hiveName*" })
                        if ($orphanLogs.Count) { throw "Найдены служебные файлы $hiveName без самого файла реестра в $targetDirectory. Требуется проверка профиля." }
                    }
                }
                if ($preservedHives[$hiveRelative]) { continue }
            }
            if ($item.PSIsContainer) {
                if (Test-Path -LiteralPath $destination) { Assert-ProfileDirectory $destination }
                else { [void][IO.Directory]::CreateDirectory($destination) }
                $queue.Enqueue($relative)
            } elseif (-not (Test-Path -LiteralPath $destination)) {
                [IO.File]::Copy($item.FullName, $destination, $false)
            }
        }
        [Windows.Forms.Application]::DoEvents()
    }
}

function Register-ExistingProfileDirectory ($Identity, [string]$Path, $Snapshot) {
    $profile = [pscustomobject]@{ UserName = $Identity.UserName; Sid = $Identity.Sid; Path = $Path }
    Assert-UnregisteredProfile $profile
    if ($Global:CancelRequested) { throw 'Операция остановлена.' }
    $hivePath = Join-Path $Path 'NTUSER.DAT'
    if (Test-Path -LiteralPath $hivePath) { Assert-ProfileFiles $profile }
    Write-GuiLog "[$($Identity.UserName)] Подготовка существующей папки без перемещения данных: $Path" 'INFO'
    $defaultPath = if ($Snapshot.DefaultPath) { ConvertTo-LocalProfilePath $Snapshot.DefaultPath } else { Join-Path (ConvertTo-LocalProfilePath $Snapshot.Root) 'Default' }
    Copy-MissingProfileTemplate $profile $defaultPath
    Assert-ProfileFiles $profile
    Assert-UnregisteredProfile $profile
    Apply-SingleACL $profile
    Assert-UnregisteredProfile $profile
    if ($Global:CancelRequested) { throw 'Операция остановлена. Регистрация SID не выполнена.' }
    # Do not publish a ProfileList entry if NTUSER.DAT is locked/loaded.
    $hiveLock = [IO.File]::Open($hivePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    try { Write-NewProfileRegistration $profile }
    finally { $hiveLock.Dispose() }
    $registered = Find-RegisteredProfile $Identity (Get-ProfileRegistrySnapshot)
    if (-not $registered -or $registered.Path -ne $Path) { throw 'Не удалось подтвердить регистрацию существующей папки.' }
    Write-GuiLog "[$($Identity.UserName)] Существующая папка зарегистрирована: SID $($Identity.Sid), путь $Path. Данные подготовки не перезаписывались." 'SUCCESS'
    return $registered
}

function Get-UserProfileContext ([string]$UserName, [switch]$Create, [switch]$AdoptExisting) {
    $identity = Resolve-ProfileIdentity $UserName
    $snapshot = Get-ProfileRegistrySnapshot
    $profile = Find-RegisteredProfile $identity $snapshot
    if (-not $profile -and $Create) {
        $systemRoot = ConvertTo-LocalProfilePath $snapshot.Root
        if ((ConvertTo-LocalProfilePath $txtDst.Text.Trim()) -ne $systemRoot) {
            throw "Windows создает новые профили в $systemRoot. Укажите этот каталог в целевой папке. Существующие профили используются по зарегистрированному пути."
        }
        if ($Global:CancelRequested) { throw 'Операция остановлена.' }
        $candidatePath = Join-Path $systemRoot $identity.UserName
        $candidateItem = $null
        try { $candidateItem = Get-Item -LiteralPath $candidatePath -Force -ErrorAction Stop }
        catch [System.Management.Automation.ItemNotFoundException] { }
        catch { throw "Не удалось проверить существующую папку $candidatePath на $env:COMPUTERNAME. Создание нового профиля отменено: $($_.Exception.Message)" }
        Write-GuiLog "[$UserName] Сервер: $env:COMPUTERNAME; путь: $candidatePath; папка найдена: $($null -ne $candidateItem)." 'INFO'
        if ($null -ne $candidateItem) {
            if (-not $AdoptExisting) { throw "Папка $candidatePath уже есть, но SID не зарегистрирован. Выполните шаг 2 для регистрации в этой папке." }
            [void](Assert-ProfileLocation $identity $candidatePath $snapshot)
            $profile = Register-ExistingProfileDirectory $identity $candidatePath $snapshot
            Assert-ProfileFiles $profile
            Assert-ProfileOffline $profile
            return $profile
        }
        Write-GuiLog "[$UserName] Создание профиля Windows для SID $($identity.Sid)..." "INFO"
        if ($Global:CancelRequested) { throw 'Операция остановлена.' }
        $result = Invoke-CreateWindowsProfile $identity
        # Another process may have created the same profile between lookup and API call.
        if ($result.HResult -ne 0 -and $result.HResult -ne -2147024713) {
            $nativeCode = [int]$result.HResult -band 0xFFFF
            $nativeMessage = (New-Object ComponentModel.Win32Exception($nativeCode)).Message
            throw ('CreateProfile завершился с ошибкой 0x{0:X8}: {1}. Сервер: {2}; SID: {3}; ожидаемая папка: {4}.' -f [int]$result.HResult, $nativeMessage, $env:COMPUTERNAME, $identity.Sid, $candidatePath)
        }
        $profile = Find-RegisteredProfile $identity (Get-ProfileRegistrySnapshot)
        if (-not $profile) { throw 'После CreateProfile запись SID в ProfileList не найдена.' }
        if ($result.HResult -eq 0 -and (ConvertTo-LocalProfilePath $result.Path) -ne $profile.Path) {
            throw 'Путь, возвращенный CreateProfile, не совпадает с ProfileList.'
        }
    }
    if (-not $profile) { throw "Профиль для SID $($identity.Sid) не зарегистрирован. Сначала выполните миграцию." }
    Assert-ProfileFiles $profile
    if ($Create) { Assert-ProfileOffline $profile }
    if ($Create -and $AdoptExisting) { Ensure-ProfileFullProfile $profile }
    return $profile
}

function Assert-ProfileTargetPath ([string]$ProfilePath, [string]$RelativePath) {
    $current = $ProfilePath
    foreach ($part in $RelativePath.Split('\')) {
        $current = Join-Path $current $part
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Целевая папка проходит через ссылку: $current" }
        }
    }
}

function Test-ProfilePermissions ($Profile) {
    $paths = @($Profile.Path, (Join-Path $Profile.Path 'NTUSER.DAT'))
    $paths += @($Global:Config.TargetFolders | ForEach-Object { Join-Path $Profile.Path $_ })
    foreach ($path in $paths) {
        $acl = Get-Acl -LiteralPath $path -ErrorAction Stop
        $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
        foreach ($sid in @($Profile.Sid, 'S-1-5-18', 'S-1-5-32-544')) {
            $allow = 0
            foreach ($rule in $rules) {
                if ($rule.PropagationFlags -band [Security.AccessControl.PropagationFlags]::InheritOnly) { continue }
                $ruleSid = $rule.IdentityReference.Value
                if ($rule.AccessControlType -eq 'Deny' -and $ruleSid -in @($sid, 'S-1-1-0', 'S-1-5-11', 'S-1-5-32-545')) { return $false }
                if ($rule.AccessControlType -eq 'Allow' -and $ruleSid -eq $sid) { $allow = $allow -bor [int]$rule.FileSystemRights }
            }
            $full = [int][Security.AccessControl.FileSystemRights]::FullControl
            if (($allow -band $full) -ne $full) { return $false }
        }
    }
    return $true
}

function Get-ProfileReadiness ([string]$UserName) {
    $profile = Get-UserProfileContext $UserName
    Assert-ProfileOffline $profile
    foreach ($folder in $Global:Config.TargetFolders) { Assert-ProfileTargetPath $profile.Path $folder }
    $missing = @(Get-MissingTargetFolders $profile.Path)
    $issues = @()
    if ($missing.Count) { $issues += 'Нет папок: ' + ($missing -join ', ') }
    elseif (-not (Test-ProfilePermissions $profile)) { $issues += 'Проверьте ACL пользователя, SYSTEM и Administrators.' }
    return [pscustomobject]@{ Ready = ($issues.Count -eq 0); Profile = $profile; Issues = $issues }
}

function Copy-MissingOutlookPst ($Source, $Destination, $ProfileRoot) {
    $pending = New-Object 'System.Collections.Generic.Stack[string]'
    $pending.Push($Source)
    while ($pending.Count) {
        $directory = $pending.Pop()
        if ($Global:CancelRequested) { throw 'Копирование PST остановлено.' }
        $directoryInfo = Get-Item -LiteralPath $directory -Force -ErrorAction Stop
        if ($directoryInfo.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
        foreach ($entry in $directoryInfo.EnumerateFileSystemInfos()) {
            if ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
            if ($entry -is [IO.DirectoryInfo]) { $pending.Push($entry.FullName); continue }
            if ($entry.Extension -ne '.pst') { continue }
            $relative = $entry.FullName.Substring($Source.Length).TrimStart('\')
            $target = Join-Path $Destination $relative
            Assert-ProfileTargetPath $ProfileRoot $target.Substring($ProfileRoot.Length).TrimStart('\')
            if ([IO.File]::Exists($target)) { continue }
            $parent = [IO.Path]::GetDirectoryName($target)
            $staging = Join-Path $parent ([guid]::NewGuid().ToString('N') + '.pm-part')
            $reader = $entry.Open('Open','Read','Read')
            try {
                $writer = [IO.File]::Open($staging,'CreateNew','Write','None')
                try {
                    $buffer = New-Object byte[] (1024*1024)
                    while (($count = $reader.Read($buffer,0,$buffer.Length)) -gt 0) {
                        $writer.Write($buffer,0,$count)
                        [Windows.Forms.Application]::DoEvents()
                        if ($Global:CancelRequested) { throw 'Копирование PST остановлено.' }
                    }
                } finally { $writer.Dispose() }
                [IO.File]::SetLastWriteTimeUtc($staging,$entry.LastWriteTimeUtc)
                [IO.File]::Move($staging,$target)
            } finally {
                $reader.Dispose()
                if ([IO.File]::Exists($staging)) { [IO.File]::Delete($staging) }
            }
        }
    }
}

function Read-OutlookSections ($Key, [string]$RelativePath) {
    if ($Global:CancelRequested) { throw 'Сбор настроек Outlook остановлен.' }
    $values = @(foreach ($name in $Key.GetValueNames()) {
        if ($name -match '(?i)password|passwd|oauth|credential|token|secret') { continue }
        [pscustomobject]@{ Name = $name; Kind = $Key.GetValueKind($name)
            Data = $Key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) }
    })
    [pscustomobject]@{ Path = $RelativePath; Values = $values }
    foreach ($name in $Key.GetSubKeyNames()) {
        $child = $Key.OpenSubKey($name)
        try {
            $path = if ($RelativePath) { "$RelativePath\$name" } else { $name }
            Read-OutlookSections $child $path
        } finally { if ($child) { $child.Dispose() } }
    }
}

function Skip-OutlookMissingData ([string]$Reason) {
    $exception = New-Object System.IO.FileNotFoundException($Reason)
    $exception.Data['OutlookDataMissing'] = $true
    throw $exception
}

function Test-OutlookMigrationPlan ($Sections, $Profile, [string]$SourceUserRoot) {
    $accounts = 0
    $pop = 0
    $stores = 0
    $imapCaches = 0
    foreach ($section in $Sections) {
        if ($section.Path -match '^9375CFF0413111d3B88A00104B2A6676\\[^\\]+$') {
            $names = @($section.Values | ForEach-Object { $_.Name })
            if ($names -contains 'POP3 Server') { $accounts++; $pop++ }
            elseif ($names -contains 'IMAP Server') { $accounts++ }
            else {
                # The account manager also lists MAPI data stores and address books.
                $service = ($section.Values | Where-Object { $_.Name -eq 'Service Name' } | Select-Object -First 1).Data
                if ($service -is [byte[]]) {
                    $encoding = if ($service.Length -ge 2 -and $service[1] -eq 0) { [Text.Encoding]::Unicode } else { [Text.Encoding]::ASCII }
                    $service = $encoding.GetString($service)
                }
                $service = ([string]$service).TrimEnd([char]0)
                if ($service -notin @('MSPST MS','MSUPST MS','CONTAB','OMSAB')) {
                    $id = Split-Path $section.Path -Leaf
                    $fields = ($names | Sort-Object | Select-Object -First 30) -join ', '
                    $kind = if ($service -match '^[A-Za-z0-9 _-]{1,32}$') { $service } else { 'не определена' }
                    throw "Раздел Outlook $id не распознан как POP3, IMAP или служебный. Служба: $kind. Имена параметров: [$fields]."
                }
            }
        }
        foreach ($value in $section.Values) {
            if ($value.Name -notin @('001f6700','001e6700')) { continue }
            $path = if ($value.Data -is [string]) { $value.Data }
                elseif ($value.Data -is [byte[]] -and $value.Name -eq '001f6700') { [Text.Encoding]::Unicode.GetString($value.Data) }
                elseif ($value.Data -is [byte[]]) { [Text.Encoding]::Default.GetString($value.Data) }
                else { throw 'Неизвестный формат пути к файлу Outlook.' }
            $path = $path.TrimEnd([char]0)
            if ($path -notmatch '(?i)\.(pst|ost)$') { throw "Неизвестное хранилище Outlook: $path" }
            $path = ConvertTo-LocalProfilePath $path
            $prefix = $SourceUserRoot.TrimEnd('\') + '\'
            if (-not $path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw "PST вне исходного профиля: $path" }
            $relative = $path.Substring($prefix.Length)
            if ((Join-Path $Profile.Path $relative) -ne $path) { throw "На цели изменится абсолютный путь PST: $path. Требуется ручная привязка." }
            Assert-ProfileTargetPath $Profile.Path $relative
            if ([IO.Path]::GetExtension($path) -eq '.ost') {
                if (-not (Test-Path -LiteralPath $path -PathType Leaf -ErrorAction Stop)) {
                    Skip-OutlookMissingData "OST не найден на цели: $path"
                }
                $imapCaches++
                continue
            }
            if (-not (Test-Path -LiteralPath $path -PathType Leaf -ErrorAction Stop)) { Skip-OutlookMissingData "PST не найден на цели: $path. Проверьте файл и выберите его исходную папку для копирования при необходимости." }
            Write-GuiLog "[$($Profile.UserName)] Найден PST на цели: $path" 'INFO'
            $stores++
        }
    }
    if (-not $accounts) { throw 'В исходном профиле не найдены учетные записи POP3 или IMAP.' }
    if ($imapCaches -and $accounts -eq $pop) { throw 'В профиле найден OST без учетной записи IMAP.' }
    if ($pop -and -not $stores) { throw 'Не удалось определить PST профиля POP3.' }
    if ($accounts -gt $pop) {
        $relativeFolder = 'AppData\Local\Microsoft\Outlook'
        Assert-ProfileTargetPath $Profile.Path $relativeFolder
        $folder = Join-Path $Profile.Path $relativeFolder
        $mailFile = if (Test-Path -LiteralPath $folder -PathType Container -ErrorAction Stop) {
            Get-ChildItem -LiteralPath $folder -File -Force -ErrorAction Stop |
                Where-Object { $_.Extension -in '.pst','.ost' } | Select-Object -First 1
        }
        if (-not $mailFile) { Skip-OutlookMissingData "Нет файлов данных IMAP (.pst или .ost) на цели: $folder" }
    }
    return $accounts
}

function Invoke-OutlookHiveCommand ([string]$Verb, [string]$Mount, [string]$File) {
    if ($Mount -notmatch '^PM_Outlook_[a-f0-9]{32}$') { throw 'Недопустимое имя временного раздела.' }
    $arguments = @($Verb, "HKU\$Mount")
    if ($Verb -eq 'load') { $arguments += $File }
    $output = & "$env:SystemRoot\System32\reg.exe" @arguments 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Outlook: reg $Verb HKU\$Mount : $output" }
}

function Get-OutlookSourceProfiles ($Hive) {
    $branches = [ordered]@{
        'Office 16.0' = 'Software\Microsoft\Office\16.0\Outlook'
        'Office 15.0' = 'Software\Microsoft\Office\15.0\Outlook'
        'Outlook 2010' = 'Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem'
    }
    foreach ($label in $branches.Keys) {
        $parent = $Hive.OpenSubKey($branches[$label])
        if (-not $parent) { continue }
        try {
            $profiles = $parent.OpenSubKey('Profiles')
            if (-not $profiles) { continue }
            try {
                $default = if ($label -eq 'Outlook 2010') { $profiles.GetValue('DefaultProfile') } else { $parent.GetValue('DefaultProfile') }
                foreach ($name in $profiles.GetSubKeyNames()) {
                    [pscustomobject]@{ Branch=$label; Name=$name; IsDefault=($name -eq $default)
                        Path=($branches[$label] + '\Profiles\' + $name) }
                }
            } finally { $profiles.Dispose() }
        } finally { $parent.Dispose() }
    }
}

function Select-OutlookSourceProfile ($Candidates, $UserName) {
    if ($Candidates.Count -eq 1) { return $Candidates[0] }
    $branches = @($Candidates.Branch | Select-Object -Unique)
    $defaults = @($Candidates | Where-Object IsDefault)
    if ($branches.Count -eq 1 -and $defaults.Count -eq 1) { return $defaults[0] }

    $dialog = New-Object Windows.Forms.Form
    try {
        $dialog.Text = "Исходный профиль Outlook — $UserName"
        $dialog.ClientSize = New-Object Drawing.Size(540,130)
        $dialog.MinimizeBox = $false
        $dialog.FormBorderStyle = 'FixedDialog'
        $dialog.StartPosition = 'CenterParent'
        $dialog.MaximizeBox = $false
        $label = New-Object Windows.Forms.Label
        $label.SetBounds(12,12,510,24)
        $label.Text = 'Найдено несколько профилей. Выберите профиль для переноса:'
        $list = New-Object Windows.Forms.ComboBox
        $list.SetBounds(12,40,510,26)
        $list.DropDownStyle = 'DropDownList'
        foreach ($c in $Candidates) {
            $suffix = if ($c.IsDefault) { ' (по умолчанию)' } else { '' }
            [void]$list.Items.Add("$($c.Branch) — $($c.Name)$suffix")
        }
        $ok = New-Object Windows.Forms.Button
        $ok.Text = 'Выбрать'
        $ok.SetBounds(280,84,105,30)
        $ok.DialogResult = 'OK'
        $ok.Enabled = $false
        $list.Add_SelectedIndexChanged({ $ok.Enabled = $list.SelectedIndex -ge 0 })
        $cancel = New-Object Windows.Forms.Button
        $cancel.SetBounds(393,84,130,30)
        $cancel.Text = 'Пропустить'
        $cancel.DialogResult = 'Cancel'
        $dialog.Controls.AddRange(@($label,$list,$ok,$cancel))
        $dialog.CancelButton = $cancel
        if ($dialog.ShowDialog($form) -eq 'OK') { return $Candidates[$list.SelectedIndex] }
    } finally { $dialog.Dispose() }
}

function Import-ClassicOutlookProfile ($Profile, [string]$SourceDirectory) {
    Assert-ProfileOffline $Profile
    Assert-ProfileFiles $Profile
    $sourceMount = 'PM_Outlook_' + [guid]::NewGuid().ToString('N')
    $targetMount = 'PM_Outlook_' + [guid]::NewGuid().ToString('N')
    $temp = Join-Path ([IO.Path]::GetTempPath()) $sourceMount
    $acl = New-Object Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true,$false)
    foreach ($id in @('S-1-5-32-544','S-1-5-18')) {
        $principal = New-Object Security.Principal.SecurityIdentifier($id)
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($principal,'FullControl','ContainerInherit,ObjectInherit','None','Allow')))
    }
    [void][IO.Directory]::CreateDirectory($temp,$acl)
    $localHive = Join-Path $temp 'NTUSER.DAT'
    $users = [Microsoft.Win32.RegistryKey]::OpenBaseKey('Users','Registry64')
    $sourceMounted = $false
    $targetMounted = $false
    try {
        $sourceHive = Join-Path $SourceDirectory 'NTUSER.DAT'
        # Hold the source hive against writes while copying it and its recovery logs.
        $inputHive = [IO.File]::Open($sourceHive,'Open','Read','Read')
        try {
            $out = [IO.File]::Create($localHive)
            try { $inputHive.CopyTo($out) } finally { $out.Dispose() }
            foreach ($suffix in @('.LOG1','.LOG2','.LOG')) {
                if ([IO.File]::Exists($sourceHive + $suffix)) { [IO.File]::Copy($sourceHive + $suffix,$localHive + $suffix,$false) }
            }
        } finally { $inputHive.Dispose() }
        Invoke-OutlookHiveCommand load $sourceMount $localHive
        $sourceMounted = $true
        $hive = $users.OpenSubKey($sourceMount)
        try {
            $candidates = @(Get-OutlookSourceProfiles $hive)
            if (-not $candidates.Count) { Write-GuiLog "[$($Profile.UserName)] Исходные профили классического Outlook не найдены." 'INFO'; return }
            $selected = Select-OutlookSourceProfile $candidates $Profile.UserName
            if (-not $selected) { Write-GuiLog "[$($Profile.UserName)] Импорт Outlook пропущен по выбору администратора." 'INFO'; return }
            $name = $selected.Name
            Write-GuiLog "[$($Profile.UserName)] Исходный профиль Outlook: $($selected.Branch), $name." 'INFO'
            $key = $hive.OpenSubKey($selected.Path)
            try { $sections = @(Read-OutlookSections $key '') } finally { $key.Dispose() }
        } finally { $hive.Dispose() }
        Invoke-OutlookHiveCommand unload $sourceMount
        $sourceMounted = $false
        $accounts = Test-OutlookMigrationPlan $sections $Profile ("C:\Users\" + $Profile.UserName)
        if ($Global:CancelRequested) { throw 'Перенос Outlook остановлен.' }
        Assert-ProfileOffline $Profile
        Invoke-OutlookHiveCommand load $targetMount (Join-Path $Profile.Path 'NTUSER.DAT')
        $targetMounted = $true
        $office = $users.CreateSubKey("$targetMount\Software\Microsoft\Office\16.0\Outlook")
        try {
            $dest = $office.CreateSubKey('Profiles')
            try {
                if ($office.GetValue('DefaultProfile') -or $dest.SubKeyCount) {
                    Write-GuiLog "[$($Profile.UserName)] На цели уже есть профиль Outlook. Настройки пропущены." 'WARN'
                    return
                }
                $newKey = $dest.CreateSubKey($name)
                try {
                    try {
                        foreach ($section in $sections) {
                            if ($Global:CancelRequested) { throw 'Перенос Outlook остановлен.' }
                            $item = if ($section.Path) { $newKey.CreateSubKey($section.Path) } else { $newKey }
                            try { foreach ($v in $section.Values) { $item.SetValue($v.Name,$v.Data,$v.Kind) } }
                            finally { if ($section.Path) { $item.Dispose() } }
                        }
                        $newKey.Flush()
                    } finally { $newKey.Dispose() }
                    $office.SetValue('DefaultProfile',$name,[Microsoft.Win32.RegistryValueKind]::String)
                } catch { $dest.DeleteSubKeyTree($name,$false); throw }
            } finally { $dest.Dispose() }
        } finally { $office.Dispose() }
        Invoke-OutlookHiveCommand unload $targetMount
        $targetMounted = $false
        Write-GuiLog "[$($Profile.UserName)] Настройки профиля Outlook перенесены ($accounts уч. зап.). Проверьте авторизацию и папку доставки в классическом Outlook на цели." 'INFO'
    } finally {
        $users.Dispose()
        try { if ($targetMounted) { Invoke-OutlookHiveCommand unload $targetMount } }
        finally {
            if ($sourceMounted) { Invoke-OutlookHiveCommand unload $sourceMount }
            # Only the generated staging directory inside this process's temp root.
            $expected = [IO.Path]::Combine([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),$sourceMount)
            if ([IO.Path]::GetFullPath($temp) -ne $expected -or $sourceMount -notmatch '^PM_Outlook_[a-f0-9]{32}$') { throw 'Недопустимый временный путь.' }
            [IO.Directory]::Delete($temp,$true)
        }
    }
}

function Sync-UserFiles ($UserName, $isDelta, [switch]$MigrateOutlook) {
    try { return (Invoke-UserFileSync $UserName $isDelta -MigrateOutlook:$MigrateOutlook) }
    catch {
        Write-GuiLog "[$UserName] Миграция не выполнена: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Invoke-UserFileSync ($UserName, $isDelta, [switch]$MigrateOutlook) {
    if ($Global:CancelRequested) { return $false }

    $SrcBase = "\\$($txtSrc.Text.Trim())\C$\Users\$UserName"

    if (-not (Test-Path $SrcBase)) {
        Write-GuiLog "[$UserName] Профиль-источник не найден ($SrcBase)" "WARN"
        return $false
    }
    $profile = Get-UserProfileContext $UserName -Create -AdoptExisting:(-not $isDelta)
    $DstBase = $profile.Path
    Write-GuiLog "[$UserName] SID $($profile.Sid), целевой профиль: $DstBase" "INFO"
    if ($Global:CancelRequested) { return $false }
    if (-not (Test-Path -LiteralPath $Global:Config.LogsDir -PathType Container)) {
        New-Item -Path $Global:Config.LogsDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    $copyOk = $true
    $folderIndex = 0

    foreach ($folder in $Global:Config.TargetFolders) {
        if ($Global:CancelRequested) { return $false }
        $folderIndex++

        $s = Join-Path $SrcBase $folder
        $d = Join-Path $DstBase $folder
        Assert-ProfileTargetPath $DstBase $folder
        Assert-ProfileOffline $profile
        try {
            $sourceFolder = Get-Item -LiteralPath $s -Force -ErrorAction Stop
        } catch [System.Management.Automation.ItemNotFoundException] {
            # Confirm the share and source profile are still reachable.
            Get-Item -LiteralPath $SrcBase -Force -ErrorAction Stop | Out-Null
            Write-GuiLog "[$UserName] Папка отсутствует на источнике, пропущена: $folder" "INFO"
            continue
        }
        if (-not $sourceFolder.PSIsContainer) { throw "Вместо папки на источнике найден файл: $folder" }
        Write-GuiLog "[$UserName] Копирование: $folder..." "INFO"
        $logFolderName = $folder -replace '[\\/]', '_'
        $log = Join-Path $Global:Config.LogsDir "${UserName}_${folderIndex}_${logFolderName}.log"
        $args = @("$s", "$d", "/E", "/ZB", "/XJ", "/COPY:DAT", "/R:1", "/W:1", "/NP", "/LOG:$log")
        if ($isDelta) { $args += "/XO" }
        $protectPst = $MigrateOutlook -and -not $isDelta
        if ($protectPst) { $args += @('/XF','*.pst') }
        $resCode = Invoke-CliCommand "robocopy.exe" $args
        if ($protectPst -and $resCode -ge 0 -and $resCode -lt 8 -and -not $Global:CancelRequested) {
            Copy-MissingOutlookPst $s $d $DstBase
        }
        if ($Global:CancelRequested) { return $false }
        if ($resCode -lt 0 -or $resCode -ge 8) {
            $copyOk = $false
            Write-GuiLog "[$UserName] Ошибка копирования $folder (Robocopy: $resCode). Лог: $log" "ERROR"
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    if ($Global:CancelRequested) { return $false }

    # 1C:Предприятие
    $Src1C = Join-Path $SrcBase "AppData\Roaming\1C\1CEStart"
    $Dst1C = Join-Path $DstBase "AppData\Roaming\1C\1CEStart"
    Assert-ProfileTargetPath $DstBase 'AppData\Roaming\1C\1CEStart'
    if (Test-Path $Src1C) {
        if (-not (Test-Path $Dst1C)) { New-Item -Path $Dst1C -ItemType Directory -Force | Out-Null }
        Write-GuiLog "[$UserName] Синхронизация настроек 1C (список баз)..." "INFO"
        $args1c = @("$Src1C", "$Dst1C", "ibases.v8i", "1CEStart.cfg", "/COPY:DAT", "/R:1", "/W:1", "/NP")
        if ($isDelta) { $args1c += "/XO" }
        $res1c = Invoke-CliCommand "robocopy.exe" $args1c
        if ($Global:CancelRequested) { return $false }
        if ($res1c -lt 0 -or $res1c -ge 8) {
            $copyOk = $false
            Write-GuiLog "[$UserName] Ошибка копирования настроек 1С (Robocopy: $res1c)." "ERROR"
        } else {
            Write-GuiLog "[$UserName] Настройки 1С скопированы." "INFO"
        }
    }
    [System.Windows.Forms.Application]::DoEvents()

    if ($Global:CancelRequested) { return $false }

    # ACL
    Write-GuiLog "[$UserName] Применение прав доступа (ACL)..." "INFO"
    Apply-SingleACL $profile
    Write-GuiLog "[$UserName] Права доступа (ACL) настроены." "INFO"
    if ($copyOk -and $MigrateOutlook -and -not $isDelta -and -not $Global:CancelRequested) {
        try { Import-ClassicOutlookProfile $profile $SrcBase }
        catch {
            if ($_.Exception.Data['OutlookDataMissing']) {
                Write-GuiLog "[$UserName] Файлы скопированы. Настройки Outlook пропущены. $($_.Exception.Message)" 'WARN'
                return 'OutlookDataMissing'
            }
            Write-GuiLog "[$UserName] Файлы скопированы. Перенос настроек Outlook не завершен: $($_.Exception.Message)" 'ERROR'
            return $false
        }
    }
    return ($copyOk -and -not $Global:CancelRequested)
}

function Invoke-ProfileAclCommand ([string[]]$Arguments) {
    if ($Global:CancelRequested) { throw 'Операция остановлена.' }
    $code = Invoke-CliCommand 'icacls.exe' $Arguments
    if ($Global:CancelRequested -or $code -ne 0) { throw "Не удалось применить ACL (icacls: $code)." }
}

function Apply-SingleACL ($Profile) {
    Assert-ProfileOffline $Profile
    Assert-ProfileFiles $Profile
    $userSid = '*' + $Profile.Sid
    $path = $Profile.Path
    Invoke-ProfileAclCommand @($path, '/inheritance:d', '/L', '/C', '/Q')
    Invoke-ProfileAclCommand @($path, '/remove', '*S-1-1-0', '*S-1-5-11', '*S-1-5-32-545', '/T', '/L', '/C', '/Q')
    Invoke-ProfileAclCommand @($path, '/setowner', $userSid, '/L', '/C', '/Q')
    Invoke-ProfileAclCommand @($path, '/grant', "${userSid}:(OI)(CI)F", '*S-1-5-18:(OI)(CI)F', '*S-1-5-32-544:(OI)(CI)F', '/L', '/C', '/Q')
    Invoke-ProfileAclCommand @((Join-Path $path 'NTUSER.DAT'), '/grant', "${userSid}:F", '*S-1-5-18:F', '*S-1-5-32-544:F', '/L', '/C', '/Q')
    foreach ($folder in @($Global:Config.TargetFolders) + @('AppData\Roaming\1C\1CEStart')) {
        Assert-ProfileTargetPath $path $folder
        $target = Join-Path $path $folder
        if (Test-Path -LiteralPath $target -PathType Container) {
            Invoke-ProfileAclCommand @($target, '/grant', "${userSid}:(OI)(CI)F", '*S-1-5-18:(OI)(CI)F', '*S-1-5-32-544:(OI)(CI)F', '/T', '/L', '/C', '/Q')
        }
    }
}

# Обработчик: Оценка объема
$btnEstimate.Add_Click({
    [void]$gridUsers.EndEdit()
    $selectedRows = @($gridUsers.Rows | Where-Object { $_.Cells[0].Value -eq $true })
    if ($selectedRows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show($form, "Отметьте пользователей для оценки объема переносимых данных!", "Внимание", "OK", "Warning")
        return
    }

    $srcServer = $txtSrc.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($srcServer)) {
        [void][System.Windows.Forms.MessageBox]::Show($form, "Укажите имя сервера-источника!", "Внимание", "OK", "Warning")
        return
    }

    Set-UiBusy $true "Оценка объема переносимых данных..."
    Update-Progress 0 "Запуск оценки объема для $($selectedRows.Count) профилей..."
    try {
        $targetRoot = [System.IO.Path]::GetPathRoot($Global:Config.ProfilesRoot)
        $targetFree = Get-DriveFreeSpace $targetRoot
        $freeMsg = if ($targetFree -ne $null) { " Свободно на целевом диске $($targetRoot): $(Format-ByteSize $targetFree)." } else { "" }
        Write-GuiLog "Старт оценки объема данных $($selectedRows.Count) профилей с сервера \\$srcServer...$freeMsg" "INFO"

        $totalUsers = $selectedRows.Count
        $cur = 0
        $grandBytes = [int64]0
        $grandFiles = [int64]0
        $notFoundCount = 0

        foreach ($row in $selectedRows) {
            if ($Global:CancelRequested) {
                Write-GuiLog "Оценка объема прервана пользователем." "WARN"
                break
            }

            $cur++
            $user = [string]$row.Cells[1].Value
            $pct = [math]::Min(100, [int](($cur / $totalUsers) * 100))
            Update-Progress $pct "Оценка объема: $user ($cur из $totalUsers)"
            Write-GuiLog "[$cur/$totalUsers] Оценка профиля: $user..." "INFO"

            $res = Get-UserDataTransferSize $srcServer $user $Global:Config.TargetFolders

            if ($Global:CancelRequested -or $res.Cancelled) {
                Write-GuiLog "Оценка объема прервана пользователем." "WARN"
                break
            }

            if (-not $res.Exists) {
                $notFoundCount++
                $row.Cells[7].Value = "Не найден"
                $row.Cells[7].Style.ForeColor = [System.Drawing.Color]::DarkOrange
                $row.Cells[7].ToolTipText = "Профиль $user не найден на сервере-источнике \\$srcServer"
                $row.Tag = [pscustomobject]@{ EstimatedBytes = [int64]0; EstimatedFiles = [int64]0; ProfileExists = $false; Details = @{} }
                Write-GuiLog "[$user] Папка профиля на источнике не найдена." "WARN"
            } else {
                $grandBytes += $res.Bytes
                $grandFiles += $res.Files
                $fmtSize = Format-ByteSize $res.Bytes
                $row.Cells[7].Value = $fmtSize
                $row.Cells[7].Style.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
                $row.Cells[7].ToolTipText = "Файлов: $($res.Files) | Объем: $(Format-ByteSize $res.Bytes) ($($res.Bytes) байт). Двойной клик — детализация."
                $row.Tag = [pscustomobject]@{ EstimatedBytes = $res.Bytes; EstimatedFiles = $res.Files; ProfileExists = $true; Details = $res.Details }
                Write-GuiLog "[$user] Оценочный объем: $fmtSize (файлов: $($res.Files))." "INFO"
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        Update-CountersAndFilter

        if (-not $Global:CancelRequested) {
            $grandFmt = Format-ByteSize $grandBytes
            $summaryMsg = "Оценка объема завершена. Обработано профилей: $cur (общий объем: $grandFmt, файлов: $grandFiles)."
            if ($targetFree -ne $null) {
                $summaryMsg += " Свободно на целевом диске $($targetRoot): $(Format-ByteSize $targetFree)."
            }
            if ($targetFree -ne $null -and $grandBytes -gt $targetFree) {
                Write-GuiLog "$summaryMsg ВНИМАНИЕ: Оценочный объем превышает свободное место на диске $targetRoot!" "WARN"
                Update-Progress 100 "Оценка завершена: $grandFmt (ВНИМАНИЕ: МЕСТА МАЛО!)"
            } else {
                Write-GuiLog $summaryMsg "SUCCESS"
                Update-Progress 100 "Оценка объема завершена ($grandFmt)"
            }
        } else {
            Update-Progress $Global:pbar.Value "Оценка объема остановлена пользователем"
        }
    } catch {
        Write-GuiLog "Ошибка оценки объема данных: $($_.Exception.Message)" "ERROR"
        Update-Progress $Global:pbar.Value "Ошибка оценки объема"
    } finally {
        Set-UiBusy $false
    }
})

# Обработчик 2: Пакетная синхронизация
$btnSync.Add_Click({
    [void]$gridUsers.EndEdit()
    $selectedRows = @($gridUsers.Rows | Where-Object { $_.Cells[0].Value -eq $true })
    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Отметьте галочками пользователей для миграции!", "Внимание", "OK", "Warning")
        return
    }

    # Проверка свободного места перед стартом
    $targetRoot = [System.IO.Path]::GetPathRoot($Global:Config.ProfilesRoot)
    $targetFree = Get-DriveFreeSpace $targetRoot
    $estimatedBytes = [int64]0
    $hasEstimated = $false
    foreach ($r in $selectedRows) {
        if ($r.Tag -and $r.Tag.PSObject.Properties['EstimatedBytes'] -and $r.Tag.EstimatedBytes -ge 0) {
            $estimatedBytes += [int64]$r.Tag.EstimatedBytes
            $hasEstimated = $true
        }
    }

    if ($hasEstimated -and $targetFree -ne $null -and $estimatedBytes -gt $targetFree) {
        $warnText = "ВНИМАНИЕ: Оценочный объем данных выбранных пользователей ($(Format-ByteSize $estimatedBytes)) превышает свободное место на целевом диске $targetRoot ($(Format-ByteSize $targetFree))!`r`n`r`nПродолжение может привести к переполнению диска.`r`nВы уверены, что хотите продолжить миграцию?"
        if ([System.Windows.Forms.MessageBox]::Show($form, $warnText, "Предупреждение: Недостаточно места", "YesNo", "Warning") -ne 'Yes') {
            Write-GuiLog "Миграция отменена пользователем из-за нехватки места на целевом диске." "WARN"
            return
        }
    }

    $sizeNotice = ""
    if ($hasEstimated) {
        $sizeNotice = " Оценочный объем данных: $(Format-ByteSize $estimatedBytes)."
        if ($targetFree -ne $null) {
            $sizeNotice += " Свободно на диске $($targetRoot): $(Format-ByteSize $targetFree)."
        }
    }

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Шаг 2. Миграция выбранных пользователей"
    $dialog.ClientSize = New-Object Drawing.Size(560,200)
    $dialog.AutoScaleMode = 'Font'
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.StartPosition = 'CenterParent'
    $dialog.MinimizeBox = $false
    $dialog.MaximizeBox = $false
    $notice = New-Object System.Windows.Forms.Label
    $notice.SetBounds(14,12,532,78)
    $notice.Text = "Выбрано пользователей: $($selectedRows.Count).$sizeNotice Завершите их сеансы на обоих серверах до начала миграции. Отличающиеся файлы в выбранных папках могут быть обновлены."
    $outlook = New-Object System.Windows.Forms.CheckBox
    $outlook.SetBounds(14,90,532,44)
    $outlook.Text = 'Перенести настройки классического Outlook (POP3 и IMAP, экспериментально). Сохранить уже имеющиеся PST.'
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.SetBounds(328,154,108,30)
    $startButton.Text = 'Начать'
    $startButton.DialogResult = 'OK'
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.SetBounds(442,154,104,30)
    $cancelButton.Text = 'Отмена'
    $cancelButton.DialogResult = 'Cancel'
    $dialog.Controls.AddRange(@($notice,$outlook,$startButton,$cancelButton))
    $dialog.AcceptButton = $startButton
    $dialog.ActiveControl = $startButton
    $dialog.CancelButton = $cancelButton
    try {
        if ($dialog.ShowDialog($btnSync.FindForm()) -ne 'OK') { return }
        $includeOutlook = $outlook.Checked
    } finally { $dialog.Dispose() }

    Set-UiBusy $true "Пакетная миграция профилей..."
    Update-Progress 0 "Запуск миграции $($selectedRows.Count) профилей..."
    try {
        $mode = if ($includeOutlook) { 'с переносом настроек Outlook' } else { 'без переноса настроек Outlook' }
        Write-GuiLog "Старт миграции $($selectedRows.Count) профилей $mode." "INFO"
        $cur = 0
        $failed = 0
        $warnings = 0

        foreach ($row in $selectedRows) {
            if ($Global:CancelRequested) {
                Write-GuiLog "Пакетная миграция прервана пользователем." "WARN"
                break
            }

            $cur++
            $user = $row.Cells[1].Value
            $pct = [math]::Min(100, [int](($cur / $selectedRows.Count) * 100))
            Update-Progress $pct "Миграция: $user ($cur из $($selectedRows.Count))"
            Write-GuiLog "[$cur/$($selectedRows.Count)] Копирование профиля: $user..." "INFO"

            $ok = if ($includeOutlook) {
                Sync-UserFiles $user $false -MigrateOutlook
            } else {
                Sync-UserFiles $user $false
            }
            if ($Global:CancelRequested) {
                $row.Cells[5].Value = "Прервано"
                $row.Cells[5].Style.ForeColor = [System.Drawing.Color]::DarkOrange
                break
            }

            if ($ok -is [string] -and $ok -eq 'OutlookDataMissing') {
                $warnings++
                $row.Cells[5].Value = 'Внимание'
                $row.Cells[5].Style.ForeColor = [System.Drawing.Color]::DarkOrange
            } elseif ($ok) {
                $row.Cells[5].Value = "Скопирован"
                $row.Cells[5].Style.ForeColor = [System.Drawing.Color]::Green
            } else {
                $failed++
                $row.Cells[5].Value = "Ошибка"
                $row.Cells[5].Style.ForeColor = [System.Drawing.Color]::Red
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        if (-not $Global:CancelRequested) {
            Write-GuiLog "Миграция завершена. Успешно: $($cur - $failed - $warnings); с предупреждениями: $warnings; ошибок: $failed." $(if ($failed -or $warnings) { 'WARN' } else { 'SUCCESS' })
            Update-Progress 100 "Миграция завершена, с предупреждениями: $warnings, ошибок: $failed"
        } else {
            Update-Progress $Global:pbar.Value "Миграция остановлена пользователем"
        }
    } finally {
        Set-UiBusy $false
    }
})

# Обработчик 3: Докачка дельты
$btnDelta.Add_Click({
    $selectedRows = @($gridUsers.Rows | Where-Object { $_.Cells[0].Value -eq $true })
    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Отметьте хотя бы одного пользователя для докачки!", "Внимание", "OK", "Warning")
        return
    }

    Set-UiBusy $true "Точечная докачка дельты (/XO)..."
    Update-Progress 0 "Запуск докачки дельты (/XO)..."
    try {
        Write-GuiLog "Запуск точечной докачки дельты (/XO) для $($selectedRows.Count) профилей..." "INFO"
        $cur = 0
        $failed = 0

        foreach ($row in $selectedRows) {
            if ($Global:CancelRequested) {
                Write-GuiLog "Докачка дельты прервана пользователем." "WARN"
                break
            }

            $cur++
            $user = $row.Cells[1].Value
            $pct = [math]::Min(100, [int](($cur / $selectedRows.Count) * 100))
            Update-Progress $pct "Докачка дельты: $user ($cur из $($selectedRows.Count))"
            Write-GuiLog "[$cur/$($selectedRows.Count)] Докачка дельты: $user..." "INFO"

            $ok = Sync-UserFiles $user $true
            if ($Global:CancelRequested) {
                $row.Cells[5].Value = "Прервано"
                $row.Cells[5].Style.ForeColor = [System.Drawing.Color]::DarkOrange
                break
            }

            if ($ok) {
                $row.Cells[5].Value = "Докачан (/XO)"
                $row.Cells[5].Style.ForeColor = [System.Drawing.Color]::Green
            } else {
                $failed++
                $row.Cells[5].Value = "Ошибка"
                $row.Cells[5].Style.ForeColor = [System.Drawing.Color]::Red
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        if (-not $Global:CancelRequested) {
            Write-GuiLog "Докачка завершена. Успешно: $($cur - $failed); ошибок: $failed." $(if ($failed) { 'WARN' } else { 'SUCCESS' })
            Update-Progress 100 "Докачка завершена, ошибок: $failed"
        } else {
            Update-Progress $Global:pbar.Value "Докачка остановлена пользователем"
        }
    } finally {
        Set-UiBusy $false
    }
})

# Контекстное меню таблицы: ручная изоляция ACL для отмеченных профилей
$ctxGrid = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemAcl = $ctxGrid.Items.Add("Изолировать права ACL (для отмеченных)")
$menuItemAcl.Add_Click({
    $selectedRows = @($gridUsers.Rows | Where-Object { $_.Cells[0].Value -eq $true })
    if ($selectedRows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Отметьте пользователей для изоляции прав!", "Внимание", "OK", "Warning")
        return
    }

    Set-UiBusy $true "Изоляция прав доступа ACL..."
    Update-Progress 0 "Изоляция прав доступа ACL..."
    try {
        Write-GuiLog "Изоляция прав ACL для $($selectedRows.Count) профилей..." "INFO"
        $cur = 0
        $failed = 0

        foreach ($row in $selectedRows) {
            if ($Global:CancelRequested) {
                Write-GuiLog "Изоляция прав ACL прервана пользователем." "WARN"
                break
            }

            $cur++
            $pct = [math]::Min(100, [int](($cur / $selectedRows.Count) * 100))
            $user = $row.Cells[1].Value
            Update-Progress $pct "Настройка ACL: $user ($cur из $($selectedRows.Count))"
            Write-GuiLog "[$cur/$($selectedRows.Count)] Настройка прав ACL: $user..." "INFO"
            try {
                $profile = Get-UserProfileContext $user
                Apply-SingleACL $profile
                $row.Cells[5].Value = 'ACL настроены'
            } catch {
                $failed++
                $row.Cells[5].Value = 'Ошибка ACL'
                Write-GuiLog "[$user] $($_.Exception.Message)" 'ERROR'
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        if (-not $Global:CancelRequested) {
            Write-GuiLog "Настройка ACL завершена, ошибок: $failed." $(if ($failed) { 'WARN' } else { 'SUCCESS' })
            Update-Progress 100 "Настройка ACL завершена, ошибок: $failed"
        } else {
            Update-Progress $Global:pbar.Value "Изоляция прав остановлена пользователем"
        }
    } finally {
        Set-UiBusy $false
    }
})
$gridUsers.ContextMenuStrip = $ctxGrid

function Get-ShortcutTarget ($Shell, [string]$Path) {
    $shortcut = $null
    try {
        $shortcut = $Shell.CreateShortcut($Path)
        return [string]$shortcut.TargetPath
    } finally {
        if ($null -ne $shortcut -and [Runtime.InteropServices.Marshal]::IsComObject($shortcut)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut)
        }
    }
}

function Find-SystemYandexBrowser {
    foreach ($root in @($env:ProgramW6432,$env:ProgramFiles,${env:ProgramFiles(x86)}) | Select-Object -Unique) {
        if ($root) {
            $exe = Join-Path $root 'Yandex\YandexBrowser\Application\browser.exe'
            if ([IO.File]::Exists($exe)) { return $exe }
        }
    }
}

function Get-BrokenUserShortcuts ([string[]]$UserNames, $Stats) {
    $shell = $null
    try {
        $shell = New-Object -ComObject WScript.Shell
        $systemBrowser = Find-SystemYandexBrowser
        $cur = 0
        foreach ($user in $UserNames) {
            if ($Global:CancelRequested) { break }
            $cur++
            Update-Progress ([int](100 * $cur / $UserNames.Count)) "Проверка ярлыков: $user ($cur из $($UserNames.Count))"
            if ($Global:CancelRequested) { break }
            try {
                $profile = Get-UserProfileContext $user
                Assert-ProfileOffline $profile -AllowLoaded
                Assert-ProfileTargetPath $profile.Path 'Desktop'
                $desk = Join-Path $profile.Path 'Desktop'
                if (-not (Test-Path -LiteralPath $desk -PathType Container -ErrorAction Stop)) { throw "Папка рабочего стола не найдена: $desk" }
                $links = Get-ChildItem -LiteralPath $desk -Filter '*.lnk' -File -Force -Recurse -ErrorAction Stop
                foreach ($link in $links) {
                    if ($Global:CancelRequested) { break }
                    try {
                        $relativePath = $link.FullName.Substring($profile.Path.TrimEnd('\').Length + 1)
                        Assert-ProfileTargetPath $profile.Path $relativePath
                        $target = Get-ShortcutTarget $shell $link.FullName
                        $Stats.Checked++
                        if ([string]::IsNullOrWhiteSpace($target)) { $Stats.NoPath++ }
                        $newTarget = $null
                        # Resolve profile-local applications against the owner of the shortcut.
                        foreach ($root in @((Split-Path $profile.Path -Parent), (Join-Path $env:SystemDrive 'Users')) | Select-Object -Unique) {
                            $prefix = $root.TrimEnd('\') + '\'
                            if (-not $target -or -not $target.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)) { continue }
                            if ($target.Substring($prefix.Length) -notmatch '^[^\\]+\\(AppData\\(?:Local|Roaming)\\.+)$') { continue }
                            $suffix = $Matches[1]
                            if ($suffix.Split('\') -contains '..' -or $suffix.Contains(':')) { break }
                            Assert-ProfileTargetPath $profile.Path $suffix
                            $owned = Join-Path $profile.Path $suffix
                            if ($suffix -eq 'AppData\Local\Yandex\YandexBrowser\Application\browser.exe' -and $systemBrowser) {
                                $newTarget = $systemBrowser
                            } elseif ($owned -ne $target -and [IO.File]::Exists($owned)) {
                                $newTarget = $owned
                            }
                            break
                        }
                        if ($newTarget -or (-not [string]::IsNullOrWhiteSpace($target) -and
                            -not [IO.File]::Exists($target) -and -not [IO.Directory]::Exists($target))) {
                            [pscustomobject]@{
                                UserName = $user; Sid = $profile.Sid; ProfilePath = $profile.Path
                                ShortcutPath = $link.FullName; RelativePath = $relativePath; TargetPath = $target; NewTarget = $newTarget; Action = $(if ($newTarget) { 'Заменить' } else { 'Пропустить' })
                            }
                        }
                    } catch {
                        $Stats.Skipped++
                        Write-GuiLog "[$user] Не удалось проверить ярлык $($link.FullName): $($_.Exception.Message)" 'WARN'
                    }
                    [Windows.Forms.Application]::DoEvents()
                }
            } catch {
                $Stats.Skipped++
                Write-GuiLog "[$user] Проверка ярлыков пропущена: $($_.Exception.Message)" 'WARN'
            }
            [Windows.Forms.Application]::DoEvents()
        }
    } finally {
        if ($null -ne $shell) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
    }
}

function Show-ShortcutCleanupConfirmation ([object[]]$Shortcuts) {
    $dialog = New-Object Windows.Forms.Form
    try {
        $dialog.Text = 'Проверка ярлыков'
        $dialog.Size = New-Object Drawing.Size(1100, 620)
        $dialog.MinimumSize = New-Object Drawing.Size(860, 420)
        $dialog.StartPosition = 'CenterParent'
        $dialog.Font = New-Object Drawing.Font('Segoe UI', 9)
        $dialog.BackColor = [Drawing.Color]::FromArgb(248, 250, 252)
        $dialog.Cursor = [Windows.Forms.Cursors]::Default
        $dialog.ShowInTaskbar = $false
        $dialog.MinimizeBox = $false

        $layout = New-Object Windows.Forms.TableLayoutPanel
        $layout.Dock = 'Fill'
        $layout.ColumnCount = 1
        $layout.RowCount = 4
        [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
        [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 80)))
        [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 44)))
        [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100)))
        [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute', 60)))
        $dialog.Controls.Add($layout)

        $heading = New-Object Windows.Forms.Label
        $heading.Dock = 'Fill'
        $heading.Margin = New-Object Windows.Forms.Padding(0)
        $heading.Padding = New-Object Windows.Forms.Padding(16, 12, 16, 8)
        $heading.BackColor = [Drawing.Color]::FromArgb(15, 23, 42)
        $heading.ForeColor = [Drawing.Color]::White
        $heading.Text = "Отметьте ярлыки галочками и назначьте действие: «Заменить» или «Удалить».`r`nДля замены нужен путь в столбце «Новая цель». Вручную его можно задать кнопкой «Выбрать файл…» для текущей строки.`r`n«Применить» обработает только отмеченные строки с назначенным действием."
        $layout.Controls.Add($heading, 0, 0)

        $list = New-Object Windows.Forms.DataGridView
        $list.Dock = 'Fill'
        $list.Margin = New-Object Windows.Forms.Padding(12, 12, 12, 0)
        $list.ReadOnly = $true
        $list.AllowUserToAddRows = $false
        $list.AllowUserToDeleteRows = $false
        $list.RowHeadersVisible = $false
        $list.AutoSizeColumnsMode = 'Fill'
        $list.AutoSizeRowsMode = 'AllCells'
        $list.DefaultCellStyle.WrapMode = 'True'
        $list.DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(236, 253, 245)
        $list.DefaultCellStyle.Padding = New-Object Windows.Forms.Padding(6, 4, 6, 4)
        $list.BackgroundColor = $dialog.BackColor
        $list.GridColor = [Drawing.Color]::FromArgb(220, 226, 232)
        $list.SelectionMode = 'FullRowSelect'
        $list.MultiSelect = $false
        $list.EnableHeadersVisualStyles = $false
        $list.ColumnHeadersDefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(30, 41, 59)
        $list.ColumnHeadersDefaultCellStyle.ForeColor = [Drawing.Color]::White
        $list.ColumnHeadersHeight = 34
        [void]$list.Columns.Add('UserName', 'Пользователь')
        [void]$list.Columns.Add('ShortcutPath', 'Путь ярлыка (.lnk)')
        [void]$list.Columns.Add('TargetPath', 'Цель ярлыка')
        $list.Columns[0].FillWeight = 17
        $list.Columns[1].FillWeight = 43
        $list.Columns[2].FillWeight = 40
        $list.Columns[2].HeaderText = 'Текущая цель'
        [void]$list.Columns.Add('NewTarget','Новая цель')
        $list.ReadOnly = $false
        foreach ($column in $list.Columns) { $column.ReadOnly = $true }
        [void]$list.Columns.Add('Action','Действие')
        $list.Columns['Action'].ReadOnly = $true
        $list.Columns['Action'].FillWeight = 18
        $check = New-Object Windows.Forms.DataGridViewCheckBoxColumn
        $check.Name = 'Checked'
        $check.HeaderText = 'Выбор'
        $check.AutoSizeMode = 'None'
        $check.Width = 58
        $check.ReadOnly = $false
        $check.SortMode = 'NotSortable'
        $list.Columns.Insert(0,$check)
        foreach ($item in $Shortcuts) {
            $index = $list.Rows.Add($false,$item.UserName,$item.ShortcutPath,$item.TargetPath,$item.NewTarget,'—')
            $list.Rows[$index].Tag = $item
        }
        $layout.Controls.Add($list, 0, 2)
        $bar = New-Object Windows.Forms.FlowLayoutPanel
        $bar.Dock = 'Fill'
        $bar.Padding = New-Object Windows.Forms.Padding(9,3,0,0)
        $layout.Controls.Add($bar, 0, 1)
        function Add-ShortcutToolButton ($caption,$onClick) {
            $button = New-Object Windows.Forms.Button
            $button.Text = $caption
            $button.AutoSize = $true
            $button.Add_Click($onClick)
            $bar.Controls.Add($button)
        }
        $setRows = {
            param($change)
            [void]$list.EndEdit()
            $list.Tag = 'Batch'
            try {
                foreach ($r in $list.Rows) {
                    if ($change -eq 'Все') { $r.Cells['Checked'].Value = $true }
                    elseif ($change -eq 'Нет') { $r.Cells['Checked'].Value = $false }
                    elseif ($r.Cells['Checked'].Value) { $r.Cells['Action'].Value = $change }
                }
            } finally { $list.Tag = $null; & $refresh }
        }
        Add-ShortcutToolButton 'Выбрать все' { & $setRows 'Все' }
        Add-ShortcutToolButton 'Снять выбор' { & $setRows 'Нет' }
        Add-ShortcutToolButton 'Заменить' { & $setRows 'Заменить' }
        Add-ShortcutToolButton 'Удалить' { & $setRows 'Удалить' }

        $footer = New-Object Windows.Forms.FlowLayoutPanel
        $footer.Dock = 'Fill'
        $footer.FlowDirection = 'RightToLeft'
        $footer.Padding = New-Object Windows.Forms.Padding(9, 9, 9, 9)
        $footer.Margin = New-Object Windows.Forms.Padding(0)
        $footer.BackColor = [Drawing.Color]::FromArgb(241, 245, 249)
        $cancel = New-Object Windows.Forms.Button
        $cancel.Text = 'Отмена'
        $cancel.Size = New-Object Drawing.Size(110, 32)
        $cancel.BackColor = [Drawing.Color]::FromArgb(26, 46, 80)
        $cancel.ForeColor = [Drawing.Color]::White
        $cancel.FlatStyle = 'Flat'
        $cancel.FlatAppearance.BorderSize = 0
        $cancel.DialogResult = [Windows.Forms.DialogResult]::Cancel
        $delete = New-Object Windows.Forms.Button
        $delete.Size = New-Object Drawing.Size(340, 32)
        $delete.BackColor = $stepReadyColor
        $delete.ForeColor = [Drawing.Color]::White
        $delete.FlatStyle = 'Flat'
        $refresh = {
            $replaceCount = 0; $deleteCount = 0; $valid = $true
            foreach ($r in $list.Rows) {
                if (-not $r.Cells['Checked'].Value) { continue }
                switch ($r.Cells['Action'].Value) {
                    'Заменить' { $replaceCount++; if (-not $r.Tag.NewTarget) { $valid = $false } }
                    'Удалить' { $deleteCount++ }
                }
            }
            $delete.Text = "Применить: заменить $replaceCount, удалить $deleteCount"
            $delete.Enabled = $valid -and ($replaceCount + $deleteCount -gt 0)
        }
        $list.Add_CurrentCellDirtyStateChanged({ if ($list.IsCurrentCellDirty) { [void]$list.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })
        $list.Add_CellValueChanged({ if (-not $list.Tag) { & $refresh } })
        $delete.Add_Click({
            [void]$list.EndEdit()
            foreach ($r in $list.Rows) { $r.Tag.Action = if ($r.Cells['Checked'].Value -and $r.Cells['Action'].Value -ne '—') { $r.Cells['Action'].Value } else { 'Пропустить' } }
            $dialog.DialogResult = [Windows.Forms.DialogResult]::OK
        })
        & $refresh
        $footer.Controls.Add($cancel)
        $footer.Controls.Add($delete)
        $choose = New-Object Windows.Forms.Button
        $choose.Text = 'Выбрать файл…'
        $choose.Size = New-Object Drawing.Size(145,32)
        $bar.Controls.Add($choose)
        $choose.Add_Click({
            $row = $list.CurrentRow
            if ($null -eq $row) { return }
            $picker = New-Object Windows.Forms.OpenFileDialog
            try {
                $picker.Title = 'Новая цель для ' + $row.Tag.ShortcutPath
                if ($picker.ShowDialog($dialog) -ne 'OK') { return }
                if ([IO.Path]::GetExtension($picker.FileName) -eq '.lnk') { return }
                $row.Tag.NewTarget = $picker.FileName
                $row.Cells['NewTarget'].Value = $picker.FileName
                $row.Cells['Checked'].Value = $true
                $row.Cells['Action'].Value = 'Заменить'
                & $refresh
            } finally { $picker.Dispose() }
        })
        $layout.Controls.Add($footer, 0, 3)
        $dialog.CancelButton = $cancel
        $dialog.AcceptButton = $cancel
        $dialog.ActiveControl = $cancel
        return ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK)
    } finally { $dialog.Dispose() }
}

function Invoke-ConfirmedShortcutChanges ([object[]]$Shortcuts, [switch]$Repair) {
    $result = [pscustomobject]@{ Deleted = 0; Repaired = 0; Skipped = 0; Errors = 0 }
    $shell = $null
    try {
        $shell = New-Object -ComObject WScript.Shell
        $cur = 0
        foreach ($item in $Shortcuts) {
            if ($Global:CancelRequested) { break }
            $cur++
            Update-Progress ([int](100 * $cur / $Shortcuts.Count)) "Обработка ярлыков: $cur из $($Shortcuts.Count)"
            if ($Global:CancelRequested) { break }
            try {
                # Повторно проверяем профиль и именно тот ярлык, который показан в списке.
                $profile = Get-UserProfileContext $item.UserName
                if ($profile.Sid -ne $item.Sid -or $profile.Path -ne $item.ProfilePath) {
                    throw 'Регистрация профиля изменилась после поиска.'
                }
                Assert-ProfileOffline $profile -AllowLoaded
                Assert-ProfileTargetPath $profile.Path $item.RelativePath
                if (-not [IO.File]::Exists($item.ShortcutPath)) {
                    $result.Skipped++
                    continue
                }
                if ($Repair) {
                    $shortcut = $shell.CreateShortcut($item.ShortcutPath)
                    try {
                        if ($shortcut.TargetPath -ne $item.TargetPath -or -not [IO.File]::Exists($item.NewTarget)) {
                            $result.Skipped++
                            Write-GuiLog "Изменился ярлык или новая цель недоступна: $($item.ShortcutPath)" 'WARN'
                            continue
                        }
                        $shortcut.TargetPath = $item.NewTarget
                        $shortcut.Save()
                        $result.Repaired++
                        Write-GuiLog "Исправлен ярлык: $($item.ShortcutPath)" 'INFO'
                    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) }
                    continue
                }
                $target = Get-ShortcutTarget $shell $item.ShortcutPath
                if ($target -cne $item.TargetPath -or [IO.File]::Exists($target) -or [IO.Directory]::Exists($target)) {
                    $result.Skipped++
                    Write-GuiLog "Ярлык сохранен (изменился или цель стала доступна): $($item.ShortcutPath)" 'INFO'
                    continue
                }
                if ($Global:CancelRequested) { break }
                Remove-Item -LiteralPath $item.ShortcutPath -Force -ErrorAction Stop
                $result.Deleted++
                Write-GuiLog "Удален битый ярлык: [$($item.UserName)] $($item.ShortcutPath)" 'WARN'
            } catch {
                $result.Errors++
                Write-GuiLog "Ошибка обработки ярлыка: $($item.ShortcutPath). $($_.Exception.Message)" 'ERROR'
            }
        }
    } finally {
        if ($null -ne $shell) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
    }
    return $result
}

# Обработчик 4: Валидация
$btnValidate.Add_Click({
    [void]$gridUsers.EndEdit()
    $selectedRows = @($gridUsers.Rows | Where-Object { $_.Cells[0].Value -eq $true })
    if ($selectedRows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Отметьте пользователей для валидации!", "Внимание", "OK", "Warning")
        return
    }
    Set-UiBusy $true "Валидация готовности профилей..."
    Update-Progress 0 "Валидация готовности профилей..."
    try {
        Write-GuiLog "Запуск валидации отмеченных профилей: $($selectedRows.Count)..." "INFO"
        $total = $selectedRows.Count
        $cur = 0

        foreach ($row in $selectedRows) {
            if ($Global:CancelRequested) {
                Write-GuiLog "Валидация прервана пользователем." "WARN"
                break
            }

            $cur++
            $pct = if ($total -gt 0) { [math]::Min(100, [int](($cur / $total) * 100)) } else { 0 }
            $user = $row.Cells[1].Value
            Update-Progress $pct "Валидация: $user ($cur из $total)"

            try {
                $check = Get-ProfileReadiness $user
                $row.Cells[5].Value = if ($check.Ready) { 'Готов' } else { 'Внимание' }
                $row.Cells[5].Style.ForeColor = if ($check.Ready) { [Drawing.Color]::Green } else { [Drawing.Color]::DarkOrange }
                $has1c = Test-Path -LiteralPath (Join-Path $check.Profile.Path 'AppData\Roaming\1C\1CEStart\ibases.v8i') -PathType Leaf
                $detail = if ($check.Ready) { 'SID, NTUSER.DAT, папки и базовые ACL: OK' } else { $check.Issues -join '; ' }
                $row.Cells[6].Value = "$detail; путь: $($check.Profile.Path); 1C: $has1c"
                if (-not $check.Ready) { Write-GuiLog "[$user] $detail" 'WARN' }
            } catch {
                $row.Cells[5].Value = "Внимание"
                $row.Cells[5].Style.ForeColor = [System.Drawing.Color]::DarkOrange
                $row.Cells[6].Value = $_.Exception.Message
                Write-GuiLog "[$user] $($_.Exception.Message)" 'WARN'
            }
            $row.Cells[6].ToolTipText = $row.Cells[6].Value
            [System.Windows.Forms.Application]::DoEvents()
        }

        if (-not $Global:CancelRequested) {
            Write-GuiLog "Валидация профилей завершена." "SUCCESS"
            Update-Progress 100 "Валидация профилей успешно завершена"
        } else {
            Update-Progress $Global:pbar.Value "Валидация остановлена пользователем"
        }
    } finally {
        Set-UiBusy $false
    }
})

# Обработчик 6: Очистка ярлыков
$btnShortcuts.Add_Click({
    [void]$gridUsers.EndEdit()
    $selectedRows = @($gridUsers.Rows | Where-Object { $_.Cells[0].Value -eq $true })
    if ($selectedRows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Отметьте пользователей для очистки ярлыков!", "Внимание", "OK", "Warning")
        return
    }
    Set-UiBusy $true "Очистка неработающих ярлыков..."
    Update-Progress 0 "Поиск неработающих ярлыков..."
    try {
        Write-GuiLog "Поиск неработающих ярлыков у отмеченных пользователей: $($selectedRows.Count)..." 'INFO'
        $stats = @{ Checked=0; Skipped=0; NoPath=0 }
        $candidates = @(Get-BrokenUserShortcuts @($selectedRows | ForEach-Object { [string]$_.Cells[1].Value }) $stats)
        if ($Global:CancelRequested) {
            Write-GuiLog 'Поиск ярлыков остановлен. Ничего не удалено.' 'WARN'
            Update-Progress $Global:pbar.Value 'Поиск остановлен; ничего не удалено'
            return
        }
        Write-GuiLog "Проверено ярлыков: $($stats.Checked); без файловой цели: $($stats.NoPath); пропущено профилей или ярлыков: $($stats.Skipped)." 'INFO'

        if ($candidates.Count -eq 0 -and ($stats.Skipped -or $stats.NoPath)) {
            Update-Progress 100 'Поиск завершен с пропусками'
            Write-GuiLog 'Среди проверенных файловых целей битых ярлыков не обнаружено. Часть объектов проверить не удалось.' 'WARN'
            return
        }
        if ($candidates.Count -eq 0) {
            Write-GuiLog 'У отмеченных пользователей не найдено неработающих ярлыков.' 'INFO'
            Update-Progress 100 'Неработающие ярлыки не найдены'
            return
        }
        Update-Progress 100 'Проверка списка ярлыков'
        if ($Global:CancelRequested -or -not (Show-ShortcutCleanupConfirmation $candidates)) {
            Write-GuiLog 'Обработка ярлыков отменена. Изменений нет.' 'INFO'
            Update-Progress 100 'Обработка отменена'
            return
        }
        $repairs = @($candidates | Where-Object { $_.Action -eq 'Заменить' -and $_.NewTarget })
        $deletions = @($candidates | Where-Object { $_.Action -eq 'Удалить' })
        $fixed = Invoke-ConfirmedShortcutChanges $repairs -Repair
        $result = Invoke-ConfirmedShortcutChanges $deletions
        $result.Errors += $fixed.Errors
        $summary = "Исправлено: $($fixed.Repaired); удалено: $($result.Deleted); пропущено: $($candidates.Count - $repairs.Count - $deletions.Count + $result.Skipped + $fixed.Skipped); ошибок: $($result.Errors)."
        if ($Global:CancelRequested) {
            Write-GuiLog "Очистка остановлена. $summary" 'WARN'
            Update-Progress $Global:pbar.Value "Очистка остановлена. $summary"
        } else {
            Write-GuiLog "Очистка завершена. $summary" $(if ($result.Errors) { 'WARN' } else { 'SUCCESS' })
            Update-Progress 100 "Очистка завершена. $summary"
        }
    } catch {
        Write-GuiLog "Ошибка очистки ярлыков: $($_.Exception.Message)" 'ERROR'
        Update-Progress $Global:pbar.Value 'Ошибка очистки ярлыков'
    } finally {
        Set-UiBusy $false
    }
})

Write-GuiLog "Программа готова к работе. Выполните аудит OU или нажмите «Загрузить список»." "INFO"
Write-GuiLog "Конфигурация: $Global:ConfigFile. Папок в TargetFolders: $($Global:Config.TargetFolders.Count)." "INFO"
if (-not (Get-Module -ListAvailable ActiveDirectory)) { Write-GuiLog "Внимание: Модуль ActiveDirectory (RSAT) не найден. При запуске аудита будет предложена автоматическая установка." "WARN" }

[System.Windows.Forms.Application]::Run($form)
