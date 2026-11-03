#requires -Version 7
<#
.SYNOPSIS
Runs the smoke app JIT, trimmed and native AOT, and checks the three outputs agree.

.DESCRIPTION
A trimmed or AOT app can build clean and still fail at run time, so each published binary is run and
must exit 0. The saved workbook is then compared with the JIT run on the parts that should be
byte-stable (sheet XML and styles). Anything else in the package, such as the core properties
timestamps, is excluded by name, not by loosening the comparison.

.PARAMETER Engine
Pinata, SixLabors, SkiaExplicit or SkiaZeroConfig. SkiaZeroConfig registers nothing and relies on the
default font engine probe, which is the path trimming used to break.

.PARAMETER Report
Also exercise XLibur.Report. Expected to fail under AOT, so combine with -Modes.

.PARAMETER Modes
Subset of Jit, Trimmed, Aot.
#>
param(
    [string]$Engine = "SkiaZeroConfig",
    [switch]$Report,
    [string[]]$Modes = @("Jit", "Trimmed", "Aot")
)

$ErrorActionPreference = "Stop"
$Modes = @($Modes | ForEach-Object { $_ -split ',' })
$project = Join-Path $PSScriptRoot "XLibur.AotSmoke.csproj"
$work = Join-Path ([IO.Path]::GetTempPath()) ("xlibur-aot-smoke-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $work | Out-Null

$rid = if ($IsWindows) { "win-x64" } elseif ($IsLinux) { "linux-x64" } elseif ($IsMacOS) { "osx-arm64" } else { throw "Unsupported OS" }
$exeName = if ($IsWindows) { "XLibur.AotSmoke.exe" } else { "XLibur.AotSmoke" }
$props = @("-p:Engine=$Engine")
if ($Report) { $props += "-p:Report=true" }

function Invoke-Dotnet {
    & dotnet @args
    if ($LASTEXITCODE -ne 0) { throw "dotnet $($args -join ' ') exited with $LASTEXITCODE" }
}

$outputs = @{}

foreach ($mode in $Modes) {
    Write-Host "::group::$Engine / $mode"
    $file = Join-Path $work "$mode.xlsx"
    switch ($mode) {
        "Jit" {
            Invoke-Dotnet run --project $project -c Release @props -- $file
        }
        default {
            $publishDir = Join-Path $work "publish-$mode"
            $flag = if ($mode -eq "Aot") { "-p:PublishAot=true" } else { "-p:PublishTrimmed=true" }
            Invoke-Dotnet publish $project -c Release -r $rid @props $flag -o $publishDir
            & (Join-Path $publishDir $exeName) $file
            if ($LASTEXITCODE -ne 0) { throw "$mode build exited with $LASTEXITCODE" }
        }
    }
    $outputs[$mode] = $file
    Write-Host "::endgroup::"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
function Get-Part([string]$archivePath, [string]$name) {
    $zip = [IO.Compression.ZipFile]::OpenRead($archivePath)
    try {
        $entry = $zip.GetEntry($name)
        if ($null -eq $entry) { throw "$archivePath has no part $name" }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
}

if ($outputs.ContainsKey("Jit")) {
    $stable = @("xl/worksheets/sheet1.xml", "xl/styles.xml")
    foreach ($mode in $outputs.Keys | Where-Object { $_ -ne "Jit" }) {
        foreach ($part in $stable) {
            if ((Get-Part $outputs["Jit"] $part) -ne (Get-Part $outputs[$mode] $part)) {
                throw "$part differs between the JIT run and the $mode run"
            }
        }
    }
    Write-Host "Outputs match the JIT run on: $($stable -join ', ')"
}

Write-Host "Smoke passed for $Engine ($($Modes -join ', '))"
