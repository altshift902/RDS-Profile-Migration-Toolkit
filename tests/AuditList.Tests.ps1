$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

$names = @('Save-AuditList', 'Read-AuditList')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $names) { . ([scriptblock]::Create($node.Extent.Text)) }
}

function Assert ($Condition, $Message) { if (-not $Condition) { throw $Message } }

$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('PM_AuditList_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $fixtureRoot -ItemType Directory)

try {
    $grid = New-Object Windows.Forms.DataGridView
    $grid.AllowUserToAddRows = $false
    [void]$grid.Columns.Add((New-Object Windows.Forms.DataGridViewCheckBoxColumn))
    1..6 | ForEach-Object { [void]$grid.Columns.Add("c$_", "c$_") }

    # Add 3 users: user1 (checked), user2 (unchecked), user3 (checked)
    [void]$grid.Rows.Add($true,  'user1', 'Иванов Иван',   '01.11.2026', '5',  'Готов',     'Примечание 1')
    [void]$grid.Rows.Add($false, 'user2', 'Петров Петр',   '01.08.2026', '95', 'Внимание',  'Неактивен')
    [void]$grid.Rows.Add($true,  'user3', 'Сидоров Сидор', '05.11.2026', '1',  'Не проверен','')

    $context = @{ Server = 'SRV-TEST'; OU = 'OU=Users,DC=test,DC=local'; AuditTime = '07.11.2026 10:00:00' }
    $outPath = Join-Path $fixtureRoot 'AuditList.json'

    # 1. Save list - must save only user1 and user3
    Save-AuditList $grid $context $outPath
    Assert (Test-Path -LiteralPath $outPath) 'AuditList.json was not created'

    $jsonRaw = Get-Content -LiteralPath $outPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert ($jsonRaw.Version -eq 1) 'Unexpected Version'
    Assert ($jsonRaw.Rows.Count -eq 2) 'Did not save exactly 2 checked users'
    Assert ($jsonRaw.Rows[0].Login -eq 'user1' -and $jsonRaw.Rows[0].Selected -eq $true) 'user1 missing or not selected'
    Assert ($jsonRaw.Rows[1].Login -eq 'user3' -and $jsonRaw.Rows[1].Selected -eq $true) 'user3 missing or not selected'
    $logins = @($jsonRaw.Rows | ForEach-Object { $_.Login })
    Assert ($logins -notcontains 'user2') 'Unchecked user2 was incorrectly saved'

    # 2. Read-AuditList - verify parsing
    $loaded = Read-AuditList $outPath 'SRV-TEST' 'OU=Users,DC=test,DC=local'
    Assert ($loaded.Rows.Count -eq 2) 'Read-AuditList did not return 2 rows'

    # 3. Test saving with NO checked rows - must throw
    $grid.Rows[0].Cells[0].Value = $false
    $grid.Rows[2].Cells[0].Value = $false
    $threw = $false
    try {
        Save-AuditList $grid $context (Join-Path $fixtureRoot 'Empty.json')
    } catch {
        $threw = $true
    }
    Assert $threw 'Saving with no checked rows did not throw an error'

    Write-Output 'PASS: Save-AuditList saves only checked users, ignores unchecked, rejects empty selection, and Read-AuditList round-trips correctly.'
} finally {
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
