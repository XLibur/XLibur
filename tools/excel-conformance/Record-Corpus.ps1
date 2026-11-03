<#
.SYNOPSIS
    Records what Excel computes for each case in a formula conformance corpus file (spec 59).

.DESCRIPTION
    Windows only, needs desktop Excel. For each case row it builds the setup cells in a scratch
    workbook, enters the formula with Range.Formula (legacy semantics, no spilling), forces a full
    calculation and writes back the value and type Excel produced. The file is stamped with the Excel
    build so the CI runner can reject a corpus that did not come from Excel.

    It starts its own Excel instance, closes only the workbook it created and quits only that
    instance. Existing Excel windows are never touched.

.PARAMETER Path
    Corpus files to record. Defaults to every *.tsv beside the test resources.
#>
[CmdletBinding()]
param(
    [string[]] $Path
)

$ErrorActionPreference = 'Stop'
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$repo = Resolve-Path (Join-Path $PSScriptRoot '..\..')
if (-not $Path) {
    $dir = Join-Path $repo 'src\XLibur.Tests\Resource\Other\FormulaConformance'
    $Path = Get-ChildItem $dir -Filter *.tsv | Where-Object { $_.Name -ne 'known-diffs.tsv' } | ForEach-Object FullName
}

$errorNames = @{
    2000 = '#NULL!'; 2007 = '#DIV/0!'; 2015 = '#VALUE!'; 2023 = '#REF!'; 2029 = '#NAME?'
    2036 = '#NUM!'; 2042 = '#N/A'; 2043 = '#GETTING_DATA'; 2045 = '#SPILL!'; 2046 = '#CONNECT!'
    2047 = '#BLOCKED!'; 2048 = '#UNKNOWN!'; 2049 = '#FIELD!'; 2050 = '#CALC!'
}

function ConvertFrom-Setup([string] $setup) {
    # A1=5;A2="x";A3=#N/A;A4=TRUE;A5==SUM(1,2)  ->  list of @{ Ref; Kind; Value }
    $entries = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($setup)) { return , $entries }
    $i = 0
    while ($i -lt $setup.Length) {
        $eq = $setup.IndexOf('=', $i)
        $ref = $setup.Substring($i, $eq - $i).Trim()
        $i = $eq + 1
        if ($setup[$i] -eq '"') {
            $sb = New-Object System.Text.StringBuilder
            $i++
            while ($true) {
                if ($setup[$i] -eq '"') {
                    if ($i + 1 -lt $setup.Length -and $setup[$i + 1] -eq '"') { [void]$sb.Append('"'); $i += 2; continue }
                    $i++; break
                }
                [void]$sb.Append($setup[$i]); $i++
            }
            $entries.Add(@{ Ref = $ref; Kind = 'text'; Value = $sb.ToString() })
        }
        else {
            $end = $setup.IndexOf(';', $i)
            if ($end -lt 0) { $end = $setup.Length }
            $raw = $setup.Substring($i, $end - $i)
            $i = $end
            if ($raw.Length -eq 0) { }
            elseif ($raw -eq 'TRUE' -or $raw -eq 'FALSE') { $entries.Add(@{ Ref = $ref; Kind = 'formula'; Value = "=$raw" }) }
            elseif ($raw.StartsWith('=')) { $entries.Add(@{ Ref = $ref; Kind = 'formula'; Value = $raw }) }
            elseif ($raw.StartsWith('#')) { $entries.Add(@{ Ref = $ref; Kind = 'formula'; Value = "=$raw" }) }
            else { $entries.Add(@{ Ref = $ref; Kind = 'number'; Value = [double]::Parse($raw, $inv) }) }
        }
        if ($i -lt $setup.Length -and $setup[$i] -eq ';') { $i++ }
    }
    return , $entries
}

function ConvertTo-Cell($value) {
    if ($null -eq $value) { return @('', 'empty') }
    if ($value -is [double]) { return @($value.ToString('R', $inv), 'number') }
    if ($value -is [bool]) { return @($(if ($value) { 'TRUE' } else { 'FALSE' }), 'bool') }
    if ($value -is [int]) {
        $code = $value -band 0xFFFF
        if ($errorNames.ContainsKey($code)) { return @($errorNames[$code], 'error') }
        throw "Unknown Excel error code $value"
    }
    if ($value -is [string]) {
        $sb = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt $value.Length; $i++) {
            $ch = $value[$i]
            if ($ch -eq '\') { [void]$sb.Append('\\') }
            elseif ($ch -eq "`t") { [void]$sb.Append('\t') }
            elseif ($ch -eq "`n") { [void]$sb.Append('\n') }
            elseif ($ch -eq "`r") { [void]$sb.Append('\r') }
            elseif ([char]::IsHighSurrogate($ch) -and ($i + 1) -lt $value.Length -and [char]::IsLowSurrogate($value[$i + 1])) {
                [void]$sb.Append($ch).Append($value[$i + 1]); $i++
            }
            elseif ([char]::IsSurrogate($ch)) { [void]$sb.Append('\u' + ([int]$ch).ToString('X4')) }
            else { [void]$sb.Append($ch) }
        }
        return @($sb.ToString(), 'text')
    }
    throw "Unexpected Value2 type $($value.GetType().FullName)"
}

$xl = $null; $wb = $null; $ws = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    if ($xl.International(3) -ne '.' -or $xl.International(4) -ne ',') {
        throw "Excel decimal/thousands separators are not . and , - the corpus is en-US only."
    }
    $stamp = "$($xl.Version).$($xl.Build) ($($xl.Name) on $([System.Environment]::OSVersion.VersionString))"
    $today = (Get-Date).ToString('yyyy-MM-dd', $inv)

    foreach ($file in $Path) {
        $lines = [System.IO.File]::ReadAllLines($file)
        $out = New-Object System.Collections.Generic.List[string]
        $out.Add("# excel_build: $stamp")
        $out.Add("# recorded: $today")
        $count = 0
        foreach ($line in $lines) {
            if ($line.Length -eq 0) { $out.Add($line); continue }
            if ($line.StartsWith('#')) {
                if ($line -notmatch '^# (excel_build|recorded):') { $out.Add($line) }
                continue
            }
            $f = $line.Split("`t")
            # function family setup formula [excel_value excel_type note]
            $function = $f[0]; $family = $f[1]; $setup = $f[2]; $formula = $f[3]
            $note = if ($f.Length -gt 6) { $f[6] } else { '' }

            $ws.Cells.Clear() | Out-Null
            $ws.Cells.NumberFormat = 'General' | Out-Null
            foreach ($entry in (ConvertFrom-Setup $setup)) {
                $cell = $ws.Range($entry.Ref)
                switch ($entry.Kind) {
                    'number' { $arr = New-Object 'object[,]' 1, 1; $arr[0, 0] = [double]$entry.Value; $cell.Value2 = $arr }
                    'text' {
                        if ($entry.Value.Length -eq 0) { $cell.Formula = '=""' }
                        else { $arr = New-Object 'object[,]' 1, 1; $arr[0, 0] = "'" + $entry.Value; $cell.Value2 = $arr }
                    }
                    'formula' { $cell.Formula = $entry.Value }
                }
            }

            $target = $ws.Range('K1')
            $rejected = $false
            try { $target.Formula = $formula } catch { $rejected = $true }
            if ($rejected) {
                $value = ''; $type = 'rejected'
            }
            else {
                $xl.CalculateFull()
                $cellValue = $target.Value2
                if ($target.HasSpill) { throw "Case '$function $formula' spilled; arrays are not supported yet." }
                $value, $type = ConvertTo-Cell $cellValue
            }
            $out.Add(($function, $family, $setup, $formula, $value, $type, $note) -join "`t")
            $count++
        }
        [System.IO.File]::WriteAllLines($file, $out, (New-Object System.Text.UTF8Encoding($false)))
        Write-Host "Recorded $count cases in $file"
    }
}
finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    if ($xl) { $xl.Quit() }
    foreach ($o in @($ws, $wb, $xl)) { if ($o) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
