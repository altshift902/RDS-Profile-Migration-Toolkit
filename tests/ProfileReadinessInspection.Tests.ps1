$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

$names = @('Get-LocalProfileInspection', 'Get-MissingTargetFolders', 'ConvertTo-LocalProfilePath', 'Save-AuditList', 'Read-AuditList')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $names) { . ([scriptblock]::Create($node.Extent.Text)) }
}

function Assert ($Condition, $Message) { if (-not $Condition) { throw $Message } }

$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('PM_Readiness_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $fixtureRoot -ItemType Directory)

try {
    $Global:Config = [ordered]@{
        ProfilesRoot = Join-Path $fixtureRoot 'Profiles'
        TargetFolders = @('Desktop', 'Documents', 'Downloads')
    }
    [void](New-Item -Path $Global:Config.ProfilesRoot -ItemType Directory)

    # 1. User not migrated at all
    $mockSnapshot = [pscustomobject]@{
        Root = $Global:Config.ProfilesRoot
        DefaultPath = Join-Path $Global:Config.ProfilesRoot 'Default'
        Entries = @()
    }

    $res1 = Get-LocalProfileInspection -UserName 'newuser' -UserSid 'S-1-5-21-999-1001' -RegistrySnapshot $mockSnapshot
    Assert ($res1.IsMigrated -eq $false) 'New user should not be marked as migrated'
    Assert ($res1.Status -eq 'Не проверен') 'New user status should be "Не проверен"'

    # 2. User folder exists on disk, but NOT in ProfileList
    $unregDir = Join-Path $Global:Config.ProfilesRoot 'orphanuser'
    [void](New-Item -Path $unregDir -ItemType Directory)
    $res2 = Get-LocalProfileInspection -UserName 'orphanuser' -UserSid 'S-1-5-21-999-1002' -RegistrySnapshot $mockSnapshot
    Assert ($res2.IsMigrated -eq $true) 'Orphan directory should be marked as migrated'
    Assert ($res2.Status -eq 'Внимание') 'Orphan directory status should be "Внимание"'
    Assert ($res2.Notes -match 'не зарегистрирована в ProfileList') 'Orphan directory note should mention ProfileList registration missing'

    # 3. User registered in ProfileList, but folder missing on disk
    $missingDirSnap = [pscustomobject]@{
        Root = $Global:Config.ProfilesRoot
        DefaultPath = Join-Path $Global:Config.ProfilesRoot 'Default'
        Entries = @(
            [pscustomobject]@{
                Sid = 'S-1-5-21-999-1003'
                Path = Join-Path $Global:Config.ProfilesRoot 'ghostuser'
                PendingRegistration = $null
            }
        )
    }
    $res3 = Get-LocalProfileInspection -UserName 'ghostuser' -UserSid 'S-1-5-21-999-1003' -RegistrySnapshot $missingDirSnap
    Assert ($res3.IsMigrated -eq $true) 'Ghost user should be marked as migrated'
    Assert ($res3.Status -eq 'Внимание') 'Ghost user status should be "Внимание"'
    Assert ($res3.Notes -match 'папка не найдена') 'Ghost user note should mention missing folder'

    # 4. User registered in ProfileList, folder exists, but NTUSER.DAT or TargetFolders missing
    $partialDir = Join-Path $Global:Config.ProfilesRoot 'partialuser'
    [void](New-Item -Path $partialDir -ItemType Directory)
    [void](New-Item -Path (Join-Path $partialDir 'Desktop') -ItemType Directory)
    $partialSnap = [pscustomobject]@{
        Root = $Global:Config.ProfilesRoot
        DefaultPath = Join-Path $Global:Config.ProfilesRoot 'Default'
        Entries = @(
            [pscustomobject]@{
                Sid = 'S-1-5-21-999-1004'
                Path = $partialDir
                PendingRegistration = $null
            }
        )
    }

    # Mock Assert-ProfileOffline and Test-ProfilePermissions
    function Assert-ProfileOffline ($p, $AllowLoaded) {}
    function Test-ProfilePermissions ($p) { return $true }

    $res4 = Get-LocalProfileInspection -UserName 'partialuser' -UserSid 'S-1-5-21-999-1004' -RegistrySnapshot $partialSnap
    Assert ($res4.IsMigrated -eq $true) 'Partial user should be marked as migrated'
    Assert ($res4.Status -eq 'Внимание') 'Partial user status should be "Внимание"'
    Assert ($res4.Notes -match 'Нет папок' -or $res4.Notes -match 'Отсутствует NTUSER.DAT') 'Partial user should report missing elements'

    # 5. User registered in ProfileList, fully migrated (all target folders, NTUSER.DAT, 1C)
    $readyDir = Join-Path $Global:Config.ProfilesRoot 'readyuser'
    [void](New-Item -Path $readyDir -ItemType Directory)
    [void](New-Item -Path (Join-Path $readyDir 'NTUSER.DAT') -ItemType File)
    foreach ($tf in $Global:Config.TargetFolders) {
        [void](New-Item -Path (Join-Path $readyDir $tf) -ItemType Directory)
    }
    $cDir = Join-Path $readyDir 'AppData\Roaming\1C\1CEStart'
    [void](New-Item -Path $cDir -ItemType Directory)
    [void](New-Item -Path (Join-Path $cDir 'ibases.v8i') -ItemType File)

    $readySnap = [pscustomobject]@{
        Root = $Global:Config.ProfilesRoot
        DefaultPath = Join-Path $Global:Config.ProfilesRoot 'Default'
        Entries = @(
            [pscustomobject]@{
                Sid = 'S-1-5-21-999-1005'
                Path = $readyDir
                PendingRegistration = $null
            }
        )
    }

    $res5 = Get-LocalProfileInspection -UserName 'readyuser' -UserSid 'S-1-5-21-999-1005' -RegistrySnapshot $readySnap
    Assert ($res5.IsMigrated -eq $true) 'Ready user should be marked as migrated'
    Assert ($res5.Status -eq 'Готов') 'Ready user status should be "Готов"'
    Assert ($res5.ForeColor.Name -eq 'Green') 'Ready user ForeColor should be Green'
    Assert ($res5.Notes -match '1C: True' -or $res5.Notes -match '1C: True') 'Ready user should detect 1C bases'

    Write-Output 'PASS: Get-LocalProfileInspection correctly identifies non-migrated, orphan, partial, and fully ready profiles with appropriate status and colors.'
} finally {
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
