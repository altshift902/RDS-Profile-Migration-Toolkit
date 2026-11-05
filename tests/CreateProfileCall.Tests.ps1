$ErrorActionPreference = 'Stop'
# An ordinary managed fake replaces userenv.dll. No Windows profile or registry APIs run.
Add-Type -TypeDefinition @'
using System.Text;
namespace ProfileMigrator {
    public static class NativeProfile {
        public static string LastSid, LastUser;
        public static uint LastCapacity;
        public static int Result, Calls;
        public static int CreateProfile(string sid, string user, StringBuilder path, uint capacity) {
            Calls++; LastSid = sid; LastUser = user; LastCapacity = capacity;
            if (capacity != 260 || path.Capacity != 260) throw new System.Exception("Unexpected API buffer size");
            if (Result == 0) path.Append(@"C:\Users\fixture.user");
            return Result;
        }
    }
}
'@
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in @('Initialize-ProfileApi', 'Invoke-CreateWindowsProfile')) { . ([scriptblock]::Create($node.Extent.Text)) }
}
function Assert ($Ok, $Message) { if (-not $Ok) { throw $Message } }
$identity = [pscustomobject]@{ Sid = 'S-1-5-21-111-222-333-1001'; UserName = 'fixture.user' }
$result = Invoke-CreateWindowsProfile $identity
Assert ($result.HResult -eq 0 -and $result.Path -ceq 'C:\Users\fixture.user') 'Returned path lost'
Assert ([ProfileMigrator.NativeProfile]::LastSid -ceq $identity.Sid -and [ProfileMigrator.NativeProfile]::LastUser -ceq $identity.UserName) 'Identity changed at API boundary'
Assert ([ProfileMigrator.NativeProfile]::LastCapacity -eq 260 -and [ProfileMigrator.NativeProfile]::Calls -eq 1) 'Wrong buffer or repeated call'
foreach ($errorCode in @(-2147023113, -2147024891, -2147024713)) {
    [ProfileMigrator.NativeProfile]::Result = $errorCode
    $result = Invoke-CreateWindowsProfile $identity
    Assert ($result.HResult -eq $errorCode -and $result.Path -eq '') 'Failed API call returned success or stale path'
}
Write-Output 'PASS: MAX_PATH buffer, exact SID/name, returned path and HRESULT preservation. Managed fake only; no profile or registry API invoked.'
