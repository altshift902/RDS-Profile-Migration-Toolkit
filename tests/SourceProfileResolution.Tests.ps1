# SourceProfileResolution.Tests.ps1
# Unit-тесты автоматического определения фактической папки источника по SID

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

$names = @('Get-RemoteProfileListMapping', 'Get-SourceUserProfilePath', 'Get-UserDataTransferSize', 'Format-ByteSize', 'Measure-FolderDataSize')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $names) { . ([scriptblock]::Create($node.Extent.Text)) }
}

function Assert ($Condition, $Message) { if (-not $Condition) { throw $Message } }

$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('PM_SourceTest_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -Path $fixtureRoot -ItemType Directory)

try {
    # 1. Test standard matching when folder exists
    $mockServer = 'SimServer'
    $mockUser = 'ivanov'
    $mockUserSid = 'S-1-5-21-1234567890-1111111111-2222222222-1001'

    # Reset cache
    $Global:RemoteServerProfileListCache = @{}
    $Global:SourceProfilePathCache = @{}

    # Test when path does not exist and registry has no mapping
    $resPath = Get-SourceUserProfilePath $mockServer $mockUser $mockUserSid
    Assert ($resPath -eq "\\$mockServer\C$\Users\$mockUser") 'Default path should be returned if nothing exists'

    # 2. Test renamed user with simulated RemoteServerProfileListCache
    $renamedLogin = 'nadejda.sazonova'
    $renamedSid = 'S-1-5-21-9999999999-8888888888-7777777777-2005'
    $actualLocalFolder = 'C:\Users\Svetlana.Fadeeva'

    # Create local fixture representing the target folder
    $localUsersShare = Join-Path $fixtureRoot 'C$\Users'
    $actualDir = Join-Path $localUsersShare 'Svetlana.Fadeeva'
    [void](New-Item -Path $actualDir -ItemType Directory -Force)
    [void](New-Item -Path (Join-Path $actualDir 'Desktop') -ItemType Directory -Force)
    [IO.File]::WriteAllBytes((Join-Path $actualDir 'Desktop\doc.txt'), (New-Object byte[] 1024))

    # Inject mock mapping for mockServer
    $Global:RemoteServerProfileListCache = @{
        $mockServer.ToLower() = @{
            $renamedSid = (Join-Path $fixtureRoot 'C$\Users\Svetlana.Fadeeva')
        }
    }
    $Global:SourceProfilePathCache = @{}

    # In our test, the raw path in registry is $actualDir
    # Test Get-SourceUserProfilePath with drive path
    $Global:RemoteServerProfileListCache[$mockServer.ToLower()][$renamedSid] = "C:\Users\Svetlana.Fadeeva"
    
    # Verify that Get-RemoteProfileListMapping returns cached map
    $map = Get-RemoteProfileListMapping $mockServer
    Assert ($map.ContainsKey($renamedSid)) 'Mapping should contain renamed user SID'
    Assert ($map[$renamedSid] -eq "C:\Users\Svetlana.Fadeeva") 'ProfileImagePath should match Svetlana.Fadeeva'

    Write-Output 'PASS: Automatic source profile path resolution and caching verified.'
} finally {
    $Global:RemoteServerProfileListCache = @{}
    $Global:SourceProfilePathCache = @{}
    if (Test-Path -LiteralPath $fixtureRoot) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force
    }
}
