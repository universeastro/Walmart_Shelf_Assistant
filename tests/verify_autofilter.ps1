param(
    [int]$Rows = 5,
    # 1-based data row index the AutoFilter should leave visible.
    [int]$KeepIndex = 3,
    [string]$Target,
    [int]$DataStartRow = 7,
    [switch]$KeepOutput
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

# README tells the user to "save the filter or hidden state in Excel first,
# then import". AutoFilter is therefore the PRIMARY path this feature serves,
# not an exotic variant - but it reaches `.Hidden` through a different Excel
# mechanism than a hand-hidden row does. This test pins that down: a row
# removed by AutoFilter must be reported hidden and must be skipped.
#
# Confirmed here (2026-09-10): AutoFilter-hidden rows DO report Hidden=$true,
# so both paths converge on the same check in excel_mapper.ps1.
#
# UsedRange can extend past the last data row (formatting only). Those trailing
# empty rows are hidden by the filter too, but they must not inflate
# `rowsHiddenSkipped`: the counter is about hidden rows that contained mapped
# data, not formatting-only rows.

if (-not $Target) { $Target = Join-Path (Split-Path $PSScriptRoot -Parent) '文件\01\B模板01.xlsx' }

$repo = Split-Path $PSScriptRoot -Parent
$mapper = Join-Path $repo 'excel_mapper.ps1'
$generator = Join-Path $PSScriptRoot 'make_multirow_fixture.ps1'

$work = Join-Path $env:TEMP ('wsa_autofilter_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null
$fixture = Join-Path $work 'A_filtered.xlsx'
$output = Join-Path $work 'out.xlsx'

$failures = New-Object System.Collections.Generic.List[string]
$excel = $null; $wb = $null
try {
    Write-Output "== generate fixture: rows=$Rows keepIndex=$KeepIndex =="
    & $generator -OutputPath $fixture -Rows $Rows | ForEach-Object { Write-Output "   $_" }
    if (-not (Test-Path -LiteralPath $fixture)) { throw 'Fixture generation failed.' }

    Write-Output '== apply AutoFilter =='
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($fixture), 0, $false)
    $ws = $wb.Worksheets.Item(1)
    $used = $ws.UsedRange
    $firstCol = $used.Column
    $lastCol = $used.Column + $used.Columns.Count - 1
    $lastRow = $used.Row + $used.Rows.Count - 1

    # Find the Title column by its generated value, not by a header literal.
    $titleCol = 0
    for ($c = $firstCol; $c -le $lastCol; $c++) {
        if ([string]$ws.Cells.Item($used.Row + 1, $c).Value2 -like 'Sample product*') { $titleCol = $c; break }
    }
    if ($titleCol -eq 0) { throw 'Title column not found in fixture.' }

    $rng = $ws.Range($ws.Cells.Item($used.Row, $firstCol), $ws.Cells.Item($lastRow, $lastCol))
    [void]$rng.AutoFilter()
    [void]$rng.AutoFilter(($titleCol - $firstCol + 1), "Sample product $KeepIndex")

    # Assert the filter DID hide something - otherwise the test below would
    # pass trivially on an unfiltered fixture.
    $hiddenCount = 0
    for ($r = $used.Row + 1; $r -le $lastRow; $r++) {
        $rr = $ws.Rows.Item($r)
        try { $h = [bool]$rr.Hidden } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rr) }
        if ($h) { $hiddenCount++ }
    }
    Write-Output "   filter hid $hiddenCount row(s) of $($lastRow - $used.Row) data-range rows"
    if ($hiddenCount -eq 0) { $failures.Add('AutoFilter hid nothing - fixture or filter is wrong, test would be vacuous.') }

    $wb.SaveAs([System.IO.Path]::GetFullPath($fixture))
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($o in @($wb, $excel)) {
        if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Output '== run mapper (RowMode=Visible) =='
$raw = & powershell -NoProfile -ExecutionPolicy Bypass -File $mapper `
    -SourcePath $fixture -TargetPath $Target -OutputPath $output -RowMode Visible 2>&1
$exitCode = $LASTEXITCODE
$json = $null
try { $json = ($raw | Out-String).Trim() | ConvertFrom-Json } catch { }
if ($null -eq $json) {
    Write-Output ($raw | Out-String)
    throw "Mapper produced no parseable JSON (exit=$exitCode)."
}
Write-Output ("   exit=$exitCode success=$($json.success)")
Write-Output ("   rowsRead=$($json.rowsRead) rowsWritten=$($json.rowsWritten) rowsHiddenSkipped=$($json.rowsHiddenSkipped)")

# Exactly one row survives the filter, and the other data rows are counted as
# hidden. Formatting-only rows beyond the data range must not affect the count.
if ($json.rowsRead -ne 1) { $failures.Add("rowsRead: expected 1, got $($json.rowsRead)") }
if ($json.rowsWritten -ne 1) { $failures.Add("rowsWritten: expected 1, got $($json.rowsWritten)") }
if ($json.rowsHiddenSkipped -ne ($Rows - 1)) {
    $failures.Add("rowsHiddenSkipped: expected $($Rows - 1), got $($json.rowsHiddenSkipped)")
}

Write-Output '== verify the surviving row =='
$mapping = $json.mappings | Where-Object { $_.sourceColumn -eq 'N' } | Select-Object -First 1
if ($null -eq $mapping) { throw 'Mapper reported no mapping for source column N (Title).' }

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($json.output), 0, $true)
    $ws = $wb.Worksheets.Item($json.targetSheet)

    $n = 0
    foreach ($ch in $mapping.targetColumn.ToUpperInvariant().ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    $got = [string]$ws.Cells.Item($DataStartRow, $n).Value2
    $expected = "Sample product $KeepIndex"
    Write-Output "   $($mapping.targetColumn)$DataStartRow : expected [$expected] got [$got]"
    if ($got -ne $expected) { $failures.Add("$($mapping.targetColumn)$DataStartRow : expected [$expected] got [$got]") }

    $spill = [string]$ws.Cells.Item($DataStartRow + 1, $n).Value2
    if (-not [string]::IsNullOrEmpty($spill)) { $failures.Add("spill at $($mapping.targetColumn)$($DataStartRow+1) = [$spill]") }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($o in @($wb, $excel)) {
        if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    if ($KeepOutput) { Write-Output "   kept: $work" }
    elseif (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($failures.Count -gt 0) {
    Write-Output ''
    Write-Output "FAILED ($($failures.Count)):"
    foreach ($f in $failures) { Write-Output "  - $f" }
    exit 1
}
Write-Output ''
Write-Output 'PASS: AutoFilter-hidden rows are skipped and the single visible row lands correctly.'
exit 0
