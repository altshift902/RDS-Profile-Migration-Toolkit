$ErrorActionPreference = 'Stop'
# No registry access: execute the production writer against an in-memory adapter.
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'GUI\ProfileMigrator_GUI.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$names = @('ConvertTo-LocalProfilePath', 'Write-NewProfileRegistration', 'Initialize-ProfileRegistrationApi',
    'Assert-ProfileLocation', 'Find-RegisteredProfile', 'Ensure-ProfileFullProfile')
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($node.Name -in $names) { . ([scriptblock]::Create($node.Extent.Text)) }
}
function Assert ($Ok, $Message) { if (-not $Ok) { throw $Message } }
function Assert-Throws ([scriptblock]$Action, $Message) {
    $caught = $false
    try { & $Action | Out-Null } catch { $caught = $true }
    Assert $caught $Message
}
Assert ($ast.Extent.Text -notmatch 'RegRenameKey|NtRenameKey|RegCreateKeyTransacted|KtmW32') 'Forbidden registry operation remains in GUI'
# Compile declarations only. Never call the native adapter or any real registry method.
Initialize-ProfileRegistrationApi
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using Microsoft.Win32;
namespace ProfileWriterTests {
    public class Key {
        public Dictionary<string, object> Values = new Dictionary<string, object>();
        public Dictionary<string, RegistryValueKind> Kinds = new Dictionary<string, RegistryValueKind>();
        public List<string> Writes = new List<string>();
        public string FailurePoint;
        public string ExternalChange;
        public int SubKeyCount { get; set; }
        public int ValueCount { get { return Values.Count; } }
        public void Fail(string point) {
            if (FailurePoint != point) return;
            FailurePoint = null;
            if (ExternalChange == "marker") Values["ProfileMigratorRegistrationPending"] = "foreign-token";
            if (ExternalChange == "value") Values["ForeignValue"] = 1;
            if (ExternalChange == "path") Values["ProfileImagePath"] = @"C:\OtherProfile";
            if (ExternalChange == "state") Values["State"] = 42;
            if (ExternalChange == "full") Values["FullProfile"] = 0;
            if (ExternalChange == "sid") Values["Sid"] = new byte[] { 1, 2, 3 };
            if (ExternalChange == "child") SubKeyCount = 1;
            throw new InvalidOperationException("Injected failure: " + point);
        }
        public void SetValue(string name, object value, RegistryValueKind kind) {
            Fail(name);
            Values[name] = value; Kinds[name] = kind; Writes.Add(name);
        }
        public object GetValue(string name) { return Values.ContainsKey(name) ? Values[name] : null; }
        public string[] GetValueNames() { var names = new List<string>(Values.Keys); return names.ToArray(); }
        public void Flush() { Fail("Flush"); }
        public void DeleteValue(string name, bool missingIsError) { Fail("Commit"); Values.Remove(name); }
        public void Dispose() {}
    }
    public class Parent {
        public Key Key;
        public bool Exists = true;
        public bool DeleteFails;
        public int Deletes;
        public string ExpectedSid;
        public Key OpenSubKey(string sid) {
            if (sid != ExpectedSid) throw new Exception("Wrong SID");
            return Exists ? Key : null;
        }
        public void DeleteSubKey(string sid, bool missingIsError) {
            if (sid != ExpectedSid || Key.SubKeyCount != 0) throw new Exception("Unsafe deletion");
            if (DeleteFails) throw new Exception("Delete denied");
            Exists = false; Deletes++;
        }
    }
    public class Session {
        public Key Key = new Key();
        public Parent Parent = new Parent();
        public bool Disposed;
        public Session() { Parent.Key = Key; }
        public void Dispose() { Disposed = true; }
    }
}
'@
$profile = [pscustomobject]@{ Sid = 'S-1-5-21-111-222-333-1001'; Path = 'C:\Fixture\tester' }
$script:session = $null
$script:rejectCreate = $false
$script:opened = 0
function New-ProfileRegistrationSession ([string]$Sid) {
    Assert ($Sid -eq $profile.Sid) 'Wrong SID at create boundary'
    $script:opened++
    if ($script:rejectCreate) { throw 'SID appeared concurrently; no new key was created' }
    return $script:session
}
function Reset-Writer {
    $Global:CancelRequested = $false
    $script:rejectCreate = $false
    $script:opened = 0
    $script:session = New-Object ProfileWriterTests.Session
    $script:session.Parent.ExpectedSid = $profile.Sid
}

Reset-Writer
Write-NewProfileRegistration $profile
Assert ($session.Disposed -and $session.Parent.Exists -and $session.Parent.Deletes -eq 0) 'Registration not retained'
Assert ($session.Key.GetValue('ProfileImagePath') -eq $profile.Path) 'Wrong path'
Assert ($session.Key.Kinds['ProfileImagePath'] -eq 'ExpandString') 'Wrong path type'
Assert ((New-Object Security.Principal.SecurityIdentifier($session.Key.GetValue('Sid'), 0)).Value -eq $profile.Sid) 'Wrong binary SID'
foreach ($name in @('Flags', 'State', 'RefCount')) {
    Assert ($session.Key.GetValue($name) -eq 0 -and $session.Key.Kinds[$name] -eq 'DWord') "Wrong $name"
}
Assert ($session.Key.GetValue('FullProfile') -eq 1 -and $session.Key.Kinds['FullProfile'] -eq 'DWord') 'FullProfile must be DWORD 1'
Assert ($session.Key.Writes[$session.Key.Writes.Count - 1] -eq 'ProfileImagePath') 'Path published before required values'
Assert ($null -eq $session.Key.GetValue('ProfileMigratorRegistrationPending')) 'Successful registration still pending'

foreach ($point in @('ProfileMigratorRegistrationPending', 'Sid', 'Flags', 'State', 'RefCount', 'FullProfile', 'ProfileImagePath', 'Flush', 'Commit')) {
    Reset-Writer
    $session.Key.FailurePoint = $point
    Assert-Throws { Write-NewProfileRegistration $profile } "Failure not reported: $point"
    Assert (-not $session.Parent.Exists -and $session.Parent.Deletes -eq 1 -and $session.Disposed) "Partial key not cleaned up: $point"
    Reset-Writer
    Write-NewProfileRegistration $profile
    Assert ($session.Parent.Exists) "Retry failed after $point"
}

foreach ($change in @('marker', 'value', 'path', 'child', 'state', 'sid', 'full')) {
    Reset-Writer
    $session.Key.FailurePoint = 'Flush'
    $session.Key.ExternalChange = $change
    Assert-Throws { Write-NewProfileRegistration $profile } "Concurrent change ignored: $change"
    Assert ($session.Parent.Exists -and $session.Parent.Deletes -eq 0 -and $session.Disposed) "Foreign data deleted: $change"
}
Reset-Writer
$session.Key.FailurePoint = 'Flush'
$session.Parent.DeleteFails = $true
Assert-Throws { Write-NewProfileRegistration $profile } 'Rollback failure not reported'
Assert ($session.Parent.Exists -and $session.Key.GetValue('ProfileMigratorRegistrationPending')) 'Incomplete entry lacks marker'

Reset-Writer
$session.Key.Values['ProfileImagePath'] = 'C:\AlreadyRegistered'
$script:rejectCreate = $true
Assert-Throws { Write-NewProfileRegistration $profile } 'Existing key was accepted for writing'
Assert ($session.Key.GetValue('ProfileImagePath') -eq 'C:\AlreadyRegistered' -and $session.Key.Writes.Count -eq 0 -and $session.Parent.Deletes -eq 0) 'Existing registration changed'

Reset-Writer
$Global:CancelRequested = $true
Assert-Throws { Write-NewProfileRegistration $profile } 'Cancellation ignored'
Assert ($script:opened -eq 0) 'Cancelled operation opened the store'
Reset-Writer
Assert-Throws { Write-NewProfileRegistration ([pscustomobject]@{ Sid = $profile.Sid; Path = '\\server\share' }) } 'UNC path accepted'
Assert-Throws { Write-NewProfileRegistration ([pscustomobject]@{ Sid = 'S-1-5-18'; Path = $profile.Path }) } 'System SID accepted'
Assert ($script:opened -eq 0) 'Invalid input opened the store'

# Backfill the one missing value in an existing, validated registration; never recreate it.
function Get-ProfileRegistrySnapshot { return $script:snapshot }
function Assert-ProfileFiles ($Profile) { if ($script:filesFailure) { throw 'Hive missing or invalid' } }
function Assert-ProfileOffline ($Profile) { if ($script:loaded) { throw 'Profile loaded' } }
function Write-GuiLog ($Message, $Kind) {}
function Open-RegisteredProfileKey ($Sid) {
    Assert ($Sid -eq $profile.Sid) 'Wrong existing key requested'
    $script:openExistingCalls++
    if ($script:disappeared) { return $null }
    if ($script:cancelOnOpen) { $Global:CancelRequested = $true }
    return $script:session.Key
}
function Reset-Backfill {
    Reset-Writer
    $script:snapshot = [pscustomobject]@{ Root = 'C:\Fixture'; Entries = @([pscustomobject]@{ Sid = $profile.Sid; Path = $profile.Path }) }
    $script:session.Key.Values['ProfileImagePath'] = $profile.Path
    $script:session.Key.Values['KeepCustomValue'] = 'unchanged'
    $script:openExistingCalls = 0
    $script:filesFailure = $false; $script:loaded = $false; $script:disappeared = $false; $script:cancelOnOpen = $false
}
Reset-Backfill
Ensure-ProfileFullProfile $profile
Assert ($session.Key.GetValue('FullProfile') -eq 1 -and $session.Key.Kinds['FullProfile'] -eq 'DWord') 'Missing FullProfile was not added'
Assert ($session.Key.Writes.Count -eq 1 -and $session.Key.Writes[0] -eq 'FullProfile') 'Backfill modified other values'
Ensure-ProfileFullProfile $profile
Assert ($session.Key.Writes.Count -eq 1 -and $session.Parent.Deletes -eq 0 -and $script:opened -eq 0) 'Repeat rewrote or recreated an existing key'
Assert ($session.Key.GetValue('ProfileImagePath') -eq $profile.Path -and $session.Key.GetValue('KeepCustomValue') -ceq 'unchanged') 'Backfill damaged existing registration'
foreach ($existingValue in @(0, 1, 'leave-existing-type')) {
    Reset-Backfill
    $session.Key.Values['FullProfile'] = $existingValue
    Ensure-ProfileFullProfile $profile
    Assert ($session.Key.Writes.Count -eq 0 -and $session.Key.GetValue('FullProfile') -ceq $existingValue) 'Existing FullProfile was overwritten'
}
foreach ($scenario in @('missing', 'changedPath', 'pending', 'bak', 'loaded', 'files', 'disappeared', 'cancel', 'cancelOnOpen', 'wrongSid', 'writeFailure')) {
    Reset-Backfill
    switch ($scenario) {
        'missing' { $script:snapshot.Entries = @() }
        'changedPath' { $session.Key.Values['ProfileImagePath'] = 'C:\OtherProfile' }
        'pending' { $session.Key.Values['ProfileMigratorRegistrationPending'] = 'incomplete' }
        'bak' { $script:snapshot.Entries += [pscustomobject]@{ Sid = ($profile.Sid + '.bak'); Path = $profile.Path } }
        'loaded' { $script:loaded = $true }
        'files' { $script:filesFailure = $true }
        'disappeared' { $script:disappeared = $true }
        'cancel' { $Global:CancelRequested = $true }
        'cancelOnOpen' { $script:cancelOnOpen = $true }
        'wrongSid' {
            $other = New-Object Security.Principal.SecurityIdentifier('S-1-5-21-111-222-333-1002')
            $otherBytes = New-Object byte[] $other.BinaryLength
            $other.GetBinaryForm($otherBytes, 0)
            $session.Key.Values['Sid'] = $otherBytes
        }
        'writeFailure' { $session.Key.FailurePoint = 'FullProfile' }
    }
    Assert-Throws { Ensure-ProfileFullProfile $profile } "Unsafe backfill allowed: $scenario"
    Assert ($session.Key.Writes.Count -eq 0 -and $session.Parent.Deletes -eq 0 -and $script:opened -eq 0) "Existing profile was changed after failure: $scenario"
}
Write-Output 'PASS: in-memory registry writer, FullProfile DWORD 1, missing-value backfill, existing value preservation, path last, failures, rollback, races, cancellation. No real registry access.'
