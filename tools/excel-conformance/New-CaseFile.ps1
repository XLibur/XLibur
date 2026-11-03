<#
.SYNOPSIS
    Expands function signatures into formula conformance case files (spec 59, Task 3).

.DESCRIPTION
    Reads signatures.psd1 and writes one case file per family with the Excel answer columns empty.
    Run Record-Corpus.ps1 on the output to fill them from Excel. No Excel is needed here.

    For each function it writes: the typical call; each argument replaced in turn by an empty cell,
    text that looks like a number, text that does not, the empty string, a boolean, an error, a range,
    an array and a cell holding numeric text; the boundary values of the argument's kind; which error
    wins when two arguments are errors; and calls with too few or too many arguments.

    Case ids are positional, so an existing file is never overwritten unless -Force is given, and
    -Force discards its recorded answers. Add later cases to a new file instead.

.PARAMETER Family
    Families to write. Defaults to every family in the signature table.

.PARAMETER Force
    Overwrite files that exist.
#>
[CmdletBinding()]
param(
    [string[]] $Family,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
$repo = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$outDir = Join-Path $repo 'src\XLibur.Tests\Resource\Other\FormulaConformance'
$functions = (Import-PowerShellDataFile (Join-Path $PSScriptRoot 'signatures.psd1')).Functions

# Pilot functions already have hand-written cases; do not generate them twice.
$existing = @{}
foreach ($file in Get-ChildItem $outDir -Filter *.tsv) {
    if ($file.Name -eq 'known-diffs.tsv') { continue }
    foreach ($line in [System.IO.File]::ReadAllLines($file.FullName)) {
        if ($line.Length -gt 0 -and $line[0] -ne '#') { $existing[$line.Split("`t")[0]] = $file.Name }
    }
}

# Boundary literals per kind. 9E307 is the largest literal Excel accepts in a formula.
$boundaries = @{
    num  = @('0', '-1', '0.5', '1E-10', '1E15', '9E307')
    int  = @('0', '-1', '1', '2.5', '255', '1000000')
    text = @('" "', '"a b"', '"ABC"', '"😀"', '"1"', '"-1.5"')
    bool = @('FALSE')
    any  = @('0', '-1', '"0"', '0.5')
    ref  = @()
}

function New-Variants([string] $kind) {
    $rangeSetup = if ($kind -eq 'text') { 'A1="a";A2="b";A3="c"' } else { 'A1=1;A2=2;A3=3' }
    $v = @(
        @{ Arg = 'A1'; Setup = ''; Note = 'empty cell' }
        @{ Arg = '"5"'; Setup = ''; Note = 'numeric text' }
        @{ Arg = '"x"'; Setup = ''; Note = 'non-numeric text' }
        @{ Arg = '""'; Setup = ''; Note = 'empty string' }
        @{ Arg = 'TRUE'; Setup = ''; Note = 'boolean' }
        @{ Arg = '#N/A'; Setup = ''; Note = 'error' }
        @{ Arg = 'A1:A3'; Setup = $rangeSetup; Note = 'range' }
        @{ Arg = '{1,2,3}'; Setup = ''; Note = 'array constant' }
        @{ Arg = 'A1'; Setup = 'A1="5"'; Note = 'cell holding numeric text' }
    )
    foreach ($b in $boundaries[$kind]) { $v += @{ Arg = $b; Setup = ''; Note = "boundary $b" } }
    return $v
}

function Get-Call([string] $name, [string[]] $callArgs) { "=$name($($callArgs -join ','))" }

# The first $n items; a range slice such as 0..($n - 1) reverses when $n is 0.
function Select-Take($items, [int] $n) { if ($n -le 0) { @() } else { @($items)[0..($n - 1)] } }

$byFamily = [ordered]@{}
foreach ($sig in $functions) {
    if ($Family -and $sig.F -notin $Family) { continue }
    if ($existing.ContainsKey($sig.N)) { Write-Verbose "skip $($sig.N): already in $($existing[$sig.N])"; continue }

    $kinds = @($sig.A); $typical = @($sig.T)
    $min = if ($sig.ContainsKey('Min')) { $sig.Min } else { $kinds.Count }
    $rows = New-Object System.Collections.Generic.List[object]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    function Add-Row([string] $setup, [string] $formula, [string] $note) {
        if ($seen.Add("$setup|$formula")) { $rows.Add(@{ Setup = $setup; Formula = $formula; Note = $note }) }
    }

    Add-Row '' (Get-Call $sig.N (Select-Take $typical $min)) 'typical'
    if ($min -eq 0) { Add-Row '' (Get-Call $sig.N @()) 'typical' }
    if ($kinds.Count -gt $min) { Add-Row '' (Get-Call $sig.N $typical) 'typical, all arguments' }

    for ($i = 0; $i -lt $kinds.Count; $i++) {
        foreach ($variant in (New-Variants $kinds[$i])) {
            $callArgs = @($typical); $callArgs[$i] = $variant.Arg
            $callArgs = Select-Take $callArgs ([Math]::Max($min, $i + 1))
            Add-Row $variant.Setup (Get-Call $sig.N $callArgs) "arg $($i + 1) $($variant.Note)"
        }
    }

    if ($kinds.Count -ge 2) {
        $callArgs = @($typical); $callArgs[0] = '#N/A'; $callArgs[1] = '#DIV/0!'
        Add-Row '' (Get-Call $sig.N (Select-Take $callArgs ([Math]::Max($min, 2)))) 'two errors, first argument first'
        $callArgs[0] = '#DIV/0!'; $callArgs[1] = '#N/A'
        Add-Row '' (Get-Call $sig.N (Select-Take $callArgs ([Math]::Max($min, 2)))) 'two errors, second argument first'
    }

    if ($min -ge 1) {
        if ($min -gt 1) { Add-Row '' (Get-Call $sig.N (Select-Take $typical ($min - 1))) 'too few arguments' }
        Add-Row '' (Get-Call $sig.N @()) 'no arguments'
    }
    if ($kinds.Count -gt 0 -and $kinds.Count -lt 3) {
        Add-Row '' (Get-Call $sig.N (@($typical) + $typical[-1])) 'too many arguments'
    }
    elseif ($kinds.Count -eq 0) {
        Add-Row '' (Get-Call $sig.N @('1')) 'too many arguments'
    }

    if (-not $byFamily.Contains($sig.F)) { $byFamily[$sig.F] = New-Object System.Collections.Generic.List[string] }
    foreach ($r in $rows) {
        $byFamily[$sig.F].Add(($sig.N, $sig.F, $r.Setup, $r.Formula, '', '', $r.Note) -join "`t")
    }
}

foreach ($fam in $byFamily.Keys) {
    $path = Join-Path $outDir "$fam.tsv"
    if ((Test-Path $path) -and -not $Force) { throw "$path exists. Pass -Force to overwrite it and lose its recorded answers." }
    $header = "# Generated by tools/excel-conformance/New-CaseFile.ps1 from signatures.psd1. function, family, setup, formula, excel_value, excel_type, note."
    [System.IO.File]::WriteAllLines($path, @($header) + $byFamily[$fam], (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "$fam.tsv: $($byFamily[$fam].Count) cases"
}
