[CmdletBinding()]
param(
    [ValidateSet('workbook', 'formula', 'address', 'all')]
    [string] $Target = 'workbook',

    [string] $LibFuzzer = (Join-Path $PSScriptRoot 'tools\libfuzzer-dotnet-windows.exe'),

    [string] $Corpus,

    [int] $Timeout = 10,

    [int] $MaxTotalTime = 0,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $LibFuzzerArgument
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($Target -eq 'all') {
    $failedTargets = [Collections.Generic.List[string]]::new()
    foreach ($targetName in @('workbook', 'formula', 'address')) {
        Write-Host "=== Fuzz target: $targetName ===" -ForegroundColor Cyan
        try {
            $childArguments = @{
                Target = $targetName
                LibFuzzer = $LibFuzzer
                Timeout = $Timeout
                MaxTotalTime = $MaxTotalTime
            }
            if ($Corpus) { $childArguments.Corpus = $Corpus }
            if ($LibFuzzerArgument) { $childArguments.LibFuzzerArgument = $LibFuzzerArgument }
            & $PSCommandPath @childArguments
            if ($LASTEXITCODE -ne 0) { throw "libFuzzer exited with code $LASTEXITCODE" }
        }
        catch {
            $failedTargets.Add($targetName)
            Write-Warning "Target '$targetName' stopped after a fuzzing failure: $($_.Exception.Message)"
        }
    }

    if ($failedTargets.Count -gt 0) {
        throw "Fuzzing completed with failures in: $($failedTargets -join ', ')"
    }
    return
}

function Resolve-ExistingPath([string] $Path, [string] $Description) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Description was not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Invoke-Native([string] $File, [string[]] $Arguments) {
    Write-Host "> $File $($Arguments -join ' ')" -ForegroundColor DarkGray
    & $File @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $File"
    }
}

$repoRoot = (Resolve-Path -LiteralPath $PSScriptRoot).Path
$libFuzzerPath = Resolve-ExistingPath $LibFuzzer 'libfuzzer-dotnet-windows.exe'
$project = Join-Path $repoRoot 'XLibur.Fuzz\XLibur.Fuzz.csproj'
$workRoot = Join-Path $repoRoot 'temp\fuzz'
$publishRoot = Join-Path $workRoot 'publish'
$toolsRoot = Join-Path $workRoot 'tools'
$corpusPath = if ([string]::IsNullOrWhiteSpace($Corpus)) {
    Join-Path $workRoot (Join-Path 'corpus' $Target)
}
else {
    [IO.Path]::GetFullPath($Corpus, $repoRoot)
}
$artifactRoot = Join-Path $workRoot 'artifacts'

New-Item -ItemType Directory -Force -Path $publishRoot, $toolsRoot, $corpusPath, $artifactRoot | Out-Null

# SharpFuzz rewrites the target assembly in place. Always publish into a fresh,
# script-owned directory so a rerun never attempts to instrument an already
# instrumented XLibur.dll.
if (Test-Path -LiteralPath $publishRoot) {
    Remove-Item -LiteralPath $publishRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $publishRoot | Out-Null

$tool = Join-Path $toolsRoot 'sharpfuzz.exe'
if (-not (Test-Path -LiteralPath $tool -PathType Leaf)) {
    Invoke-Native 'dotnet' @('tool', 'install', 'SharpFuzz.CommandLine', '--tool-path', $toolsRoot, '--version', '2.3.0')
}

Invoke-Native 'dotnet' @('publish', $project, '--configuration', 'Release', '--framework', 'net10.0', '--output', $publishRoot, '--no-self-contained', '--no-restore')

$targetAssembly = Join-Path $publishRoot 'XLibur.dll'
Invoke-Native $tool @($targetAssembly)

$seedFiles = @(Get-ChildItem -LiteralPath $corpusPath -File -ErrorAction SilentlyContinue)
if ($seedFiles.Count -eq 0) {
    if ($Target -eq 'workbook') {
        $seed = Join-Path $repoRoot 'XLibur.Tests\Resource\TryToLoad\LO\xlsx\empty.xlsx'
        Copy-Item -LiteralPath $seed -Destination (Join-Path $corpusPath 'empty.xlsx')
    }
    elseif ($Target -eq 'formula') {
        [IO.File]::WriteAllText((Join-Path $corpusPath 'formula.txt'), 'SUM(1,2)', [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText((Join-Path $corpusPath 'formula-2.txt'), 'IF(A1>0,"yes","no")', [Text.UTF8Encoding]::new($false))
    }
    else {
        [IO.File]::WriteAllText((Join-Path $corpusPath 'address.txt'), 'Sheet1!$A$1:$C$10', [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText((Join-Path $corpusPath 'address-2.txt'), 'R1C1', [Text.UTF8Encoding]::new($false))
    }
}

$env:XLIBUR_FUZZ_TARGET = $Target
$harness = Join-Path $publishRoot 'XLibur.Fuzz.exe'
$artifactPrefix = $artifactRoot + [IO.Path]::DirectorySeparatorChar
$arguments = @("--target_path=$harness", "-timeout=$Timeout", "-artifact_prefix=$artifactPrefix")
if ($MaxTotalTime -gt 0) { $arguments += "-max_total_time=$MaxTotalTime" }
if ($Target -eq 'workbook') { $arguments += '-max_len=1048576' } else { $arguments += '-max_len=4096' }
if ($LibFuzzerArgument) { $arguments += $LibFuzzerArgument }
$arguments += $corpusPath

Invoke-Native $libFuzzerPath $arguments
