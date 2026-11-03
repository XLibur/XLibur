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

# The text of a quoted setup value, starting just after the opening quote. "" is a quote.
# Returns the text and the index after the closing quote.
function Read-QuotedText([string] $setup, [int] $start) {
    $sb = New-Object System.Text.StringBuilder
    $i = $start
    while ($true) {
        $isQuote = $setup[$i] -eq '"'
        if ($isQuote -and ($i + 1) -lt $setup.Length -and $setup[$i + 1] -eq '"') { [void]$sb.Append('"'); $i += 2 }
        elseif ($isQuote) { return @{ Text = $sb.ToString(); Next = $i + 1 } }
        else { [void]$sb.Append($setup[$i]); $i++ }
    }
}

# An unquoted setup value: a number, TRUE or FALSE, an error such as #N/A, or =formula. Empty is an empty cell.
function ConvertTo-SetupEntry([string] $ref, [string] $raw) {
    if ($raw.Length -eq 0) { return $null }
    if ($raw -eq 'TRUE' -or $raw -eq 'FALSE') { return @{ Ref = $ref; Kind = 'formula'; Value = "=$raw" } }
    if ($raw.StartsWith('=')) { return @{ Ref = $ref; Kind = 'formula'; Value = $raw } }
    if ($raw.StartsWith('#')) { return @{ Ref = $ref; Kind = 'formula'; Value = "=$raw" } }
    return @{ Ref = $ref; Kind = 'number'; Value = [double]::Parse($raw, $inv) }
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
            $quoted = Read-QuotedText $setup ($i + 1)
            $entries.Add(@{ Ref = $ref; Kind = 'text'; Value = $quoted.Text })
            $i = $quoted.Next
        }
        else {
            $end = $setup.IndexOf(';', $i)
            if ($end -lt 0) { $end = $setup.Length }
            $entry = ConvertTo-SetupEntry $ref $setup.Substring($i, $end - $i)
            if ($entry) { $entries.Add($entry) }
            $i = $end
        }
        if ($i -lt $setup.Length -and $setup[$i] -eq ';') { $i++ }
    }
    return , $entries
}

# Escapes \ tab LF CR, and a surrogate that is not half of a valid pair as \uXXXX.
function ConvertTo-EscapedText([string] $value) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $value.Length; $i++) {
        $ch = $value[$i]
        $isPair = [char]::IsHighSurrogate($ch) -and ($i + 1) -lt $value.Length -and [char]::IsLowSurrogate($value[$i + 1])
        if ($ch -eq '\') { [void]$sb.Append('\\') }
        elseif ($ch -eq "`t") { [void]$sb.Append('\t') }
        elseif ($ch -eq "`n") { [void]$sb.Append('\n') }
        elseif ($ch -eq "`r") { [void]$sb.Append('\r') }
        elseif ($isPair) { [void]$sb.Append($ch).Append($value[$i + 1]); $i++ }
        elseif ([char]::IsSurrogate($ch)) { [void]$sb.Append('\u' + ([int]$ch).ToString('X4')) }
        else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function ConvertTo-ErrorName([int] $value) {
    $code = $value -band 0xFFFF
    if (-not $errorNames.ContainsKey($code)) { throw "Unknown Excel error code $value" }
    return $errorNames[$code]
}

function ConvertTo-Cell($value) {
    if ($null -eq $value) { return @('', 'empty') }
    if ($value -is [double]) { return @($value.ToString('R', $inv), 'number') }
    if ($value -is [bool]) { return @($(if ($value) { 'TRUE' } else { 'FALSE' }), 'bool') }
    if ($value -is [int]) { return @((ConvertTo-ErrorName $value), 'error') }
    if ($value -is [string]) { return @((ConvertTo-EscapedText $value), 'text') }
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
