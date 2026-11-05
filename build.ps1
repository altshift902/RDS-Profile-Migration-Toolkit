$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot 'GUI\ProfileMigrator_GUI.ps1'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$dist = Join-Path $PSScriptRoot 'dist'
[void][IO.Directory]::CreateDirectory($dist)
$match = [regex]::Match([IO.File]::ReadAllText($source), '(?m)^\$script:AppVersion = ''(\d+\.\d+\.\d+)(?:-dev\.(\d+))?''\s*$')
if (-not $match.Success) { throw 'Expected AppVersion: major.minor.patch or major.minor.patch-dev.N' }
$baseVersion = $match.Groups[1].Value
$buildNumber = 0
$productVersion = $baseVersion
if ($match.Groups[2].Success) {
    $buildNumber = [int]$match.Groups[2].Value
    $productVersion = "$baseVersion-dev.$buildNumber"
}
$attributes = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString() + '.cs')
$payload = [IO.Path]::GetTempFileName()
try {
    [IO.File]::WriteAllLines($attributes, @(
        "[assembly: System.Reflection.AssemblyVersion(`"$baseVersion.0`")]",
        "[assembly: System.Reflection.AssemblyFileVersion(`"$baseVersion.$buildNumber`")]",
        "[assembly: System.Reflection.AssemblyInformationalVersion(`"$productVersion`")]"
    ))
    $bytes = [IO.File]::ReadAllBytes($source)
    $memory = New-Object IO.MemoryStream
    try {
        $gzip = New-Object IO.Compression.GZipStream($memory,[IO.Compression.CompressionMode]::Compress,$true)
        try { $gzip.Write($bytes,0,$bytes.Length) } finally { $gzip.Dispose() }
        $packed = $memory.ToArray()
    } finally { $memory.Dispose() }
    for ($i=0; $i -lt $packed.Length; $i++) { $packed[$i] = $packed[$i] -bxor 0xA7 }
    [IO.File]::WriteAllBytes($payload,$packed)
    & $compiler /nologo /target:winexe /platform:anycpu "/out:$dist\ProfileMigrator_GUI.exe" "/win32icon:$PSScriptRoot\app.ico" "/win32manifest:$PSScriptRoot\app.manifest" /reference:System.Windows.Forms.dll "/resource:$payload,payload.dat" "$PSScriptRoot\GUI\ProgramGUIProtected.cs" $attributes
    if ($LASTEXITCODE) { throw 'Ошибка компиляции EXE.' }
    Write-Output "$dist\ProfileMigrator_GUI.exe"
} finally { [IO.File]::Delete($payload); [IO.File]::Delete($attributes) }
