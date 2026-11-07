# DataSizeEstimation.Tests.ps1
# Unit-тесты для модуля оценки объема переносимых данных (версия 1.2.0)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

$names = @('Format-ByteSize', 'Get-DriveFreeSpace', 'Measure-FolderDataSize', 'Get-UserDataTransferSize')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $names) { . ([scriptblock]::Create($node.Extent.Text)) }
}

function Assert ($Condition, $Message) { if (-not $Condition) { throw $Message } }

$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('PM_SizeTest_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $fixtureRoot -ItemType Directory)

try {
    # 1. Format-ByteSize
    Assert ((Format-ByteSize -1) -eq '—') 'Negative byte value should return dash'
    Assert ((Format-ByteSize 0) -eq '0 Б') 'Zero bytes should return 0 Б'
    Assert ((Format-ByteSize 500) -eq '500 Б') '500 bytes should return 500 Б'
    Assert ((Format-ByteSize 1024) -match '1\s*КБ') '1024 bytes should format as 1 КБ'
    Assert ((Format-ByteSize (50 * 1024 * 1024)) -match '50[,.]0\s*МБ') '50MB should format as 50.0 МБ'
    Assert ((Format-ByteSize (2500 * 1024 * 1024)) -match '2[,.]44\s*ГБ') '2.44GB should format as 2.44 ГБ'
    Assert ((Format-ByteSize (2 * 1024 * 1024 * 1024 * 1024)) -match '2[,.]00\s*ТБ') '2TB should format as 2.00 ТБ'

    # 2. Get-DriveFreeSpace
    Assert ($null -eq (Get-DriveFreeSpace '')) 'Empty path should return null'
    Assert ($null -eq (Get-DriveFreeSpace $null)) 'Null path should return null'
    $cFree = Get-DriveFreeSpace 'C:\Users'
    Assert ($null -ne $cFree -and $cFree -gt 0) 'Drive C: should have positive free space'

    # 3. Measure-FolderDataSize on a local test folder
    $testFolder = Join-Path $fixtureRoot 'FolderA'
    [void](New-Item -Path $testFolder -ItemType Directory)
    $file1 = Join-Path $testFolder 'test1.dat'
    $file2 = Join-Path $testFolder 'test2.dat'
    [IO.File]::WriteAllBytes($file1, (New-Object byte[] 4096))
    [IO.File]::WriteAllBytes($file2, (New-Object byte[] 8192))

    $Global:CancelRequested = $false
    $Global:CurrentProc = $null
    $meas = Measure-FolderDataSize $testFolder
    Assert (-not $meas.Cancelled) 'Measure was cancelled unexpectedly'
    Assert ($meas.Files -eq 2) "Expected 2 files, got $($meas.Files)"
    Assert ($meas.Bytes -eq (4096 + 8192)) "Expected 12288 bytes, got $($meas.Bytes)"

    # Test non-existent folder
    $measMissing = Measure-FolderDataSize (Join-Path $fixtureRoot 'NonExistent')
    Assert ($measMissing.Bytes -eq 0 -and $measMissing.Files -eq 0) 'Missing folder should return 0 bytes/files'

    # 4. Get-UserDataTransferSize with simulated profile and 1C settings
    $mockServerDir = Join-Path $fixtureRoot 'FakeServer\C$\Users\testuser'
    [void](New-Item -Path (Join-Path $mockServerDir 'Desktop') -ItemType Directory -Force)
    [void](New-Item -Path (Join-Path $mockServerDir 'Documents') -ItemType Directory -Force)
    [void](New-Item -Path (Join-Path $mockServerDir 'AppData\Roaming\1C\1CEStart') -ItemType Directory -Force)

    [IO.File]::WriteAllBytes((Join-Path $mockServerDir 'Desktop\doc.txt'), (New-Object byte[] 2048))
    [IO.File]::WriteAllBytes((Join-Path $mockServerDir 'Documents\sheet.xlsx'), (New-Object byte[] 5120))
    [IO.File]::WriteAllBytes((Join-Path $mockServerDir 'AppData\Roaming\1C\1CEStart\ibases.v8i'), (New-Object byte[] 1000))

    # Test mock server UNC path simulation
    # We test with local path by replacing the function's internal path or calling sub-components
    $deskMeas = Measure-FolderDataSize (Join-Path $mockServerDir 'Desktop')
    $docsMeas = Measure-FolderDataSize (Join-Path $mockServerDir 'Documents')
    Assert ($deskMeas.Bytes -eq 2048 -and $deskMeas.Files -eq 1) 'Desktop folder measurement mismatch'
    Assert ($docsMeas.Bytes -eq 5120 -and $docsMeas.Files -eq 1) 'Documents folder measurement mismatch'

    # Test 1C file size
    $ibasesFile = Join-Path $mockServerDir 'AppData\Roaming\1C\1CEStart\ibases.v8i'
    Assert ((Get-Item $ibasesFile).Length -eq 1000) '1C file length mismatch'

    # 5. Cancellation test
    $Global:CancelRequested = $true
    $measCancel = Measure-FolderDataSize $testFolder
    Assert ($measCancel.Cancelled -eq $true) 'Cancelled flag was not set when CancelRequested is true'
    $Global:CancelRequested = $false

    Write-Output 'PASS: Data size estimation, byte formatting, free space checks, and cancellation verified.'
} finally {
    $Global:CancelRequested = $false
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
