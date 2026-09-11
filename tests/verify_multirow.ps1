param(
    [int]$Rows = 5,
    # 1-based data row to leave blank in the source. 0 = no blank row.
    [int]$BlankAt = 0,
    # Comma-separated 1-based data row indices to HIDE in the source, e.g. '3,5'.
    [string]$HideAt = '',
    [ValidateSet('Visible', 'All')][string]$RowMode = 'Visible',
    [string]$Target,
    [int]$DataStartRow = 7,
    [switch]$KeepOutput
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

# The expectation is keyed by the SOURCE column letter that the mapper itself
# reports in its JSON contract. Nothing is hardcoded on the target side, so
# changing the mapping table cannot silently turn this check into a no-op.
function Get-ExpectedValue([string]$SourceLetter, [int]$Index) {
    switch ($SourceLetter) {
        'D'  { return ('CUST-{0:000}' -f $Index) }
        'E'  { return ('GROUP-{0:000}' -f $Index) }
        'L'  { return "Red-$Index" }
        'N'  { return "Sample product $Index" }
        'O'  { return "Long description $Index" }
        'P'  { return "Feature one $Index" }
        'Q'  { return "Feature two $Index" }
        'R'  { return "Feature three $Index" }
        'S'  { return "Feature four $Index" }
        'T'  { return "Feature five $Index" }
        'AG' { return "https://img.example/one-$Index.jpg" }
        'AI' { return "https://img.example/main-$Index.jpg" }
        'AX' { return "https://img.example/swatch-$Index.jpg" }
        default { return $null }
    }
}

function ConvertTo-ColumnNumber([string]$Letter) {
    $n = 0
    foreach ($ch in $Letter.ToUpperInvariant().ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    return $n
}

if (-not $Target) { $Target = Join-Path (Split-Path $PSScriptRoot -Parent) '文件\B模板01.xlsx' }

$repo = Split-Path $PSScriptRoot -Parent
$mapper = Join-Path $repo 'excel_mapper.ps1'
$generator = Join-Path $PSScriptRoot 'make_multirow_fixture.ps1'

$work = Join-Path $env:TEMP ('wsa_multirow_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null
$fixture = Join-Path $work 'A_multirow.xlsx'
$output = Join-Path $work 'out.xlsx'

$failures = New-Object System.Collections.Generic.List[string]
$excel = $null; $wb = $null
try {
    Write-Output "== generate fixture: rows=$Rows blankAt=$BlankAt =="
    & $generator -OutputPath $fixture -Rows $Rows -BlankAt $BlankAt -HideAt $HideAt | ForEach-Object { Write-Output "   $_" }
    if (-not (Test-Path -LiteralPath $fixture)) { throw 'Fixture generation failed.' }

    Write-Output '== run mapper =='
    $raw = & powershell -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $fixture -TargetPath $Target -OutputPath $output -RowMode $RowMode 2>&1
    $exitCode = $LASTEXITCODE
    $json = $null
    try { $json = ($raw | Out-String).Trim() | ConvertFrom-Json } catch { }
    if ($null -eq $json) {
        Write-Output ($raw | Out-String)
        throw "Mapper produced no parseable JSON (exit=$exitCode)."
    }
    Write-Output ("   exit=$exitCode success=$($json.success)")
    Write-Output ("   rowsRead=$($json.rowsRead) rowsWritten=$($json.rowsWritten) sheet=$($json.targetSheet)")

    # Which source rows should reach the target: the blank one is dropped, and
    # in Visible mode so are the hidden ones. The survivors keep their ORIGINAL
    # 1-based index, which stays anchored to the value the generator wrote
    # (row index i always holds the '...$i' variant), so a row-order or
    # row-count bug cannot hide behind a shifted expectation.
    $hideSet = @()
    foreach ($token in ($HideAt -split ',')) {
        $n = 0
        if ([int]::TryParse($token.Trim(), [ref]$n)) { $hideSet += $n }
    }
    $included = @()
    for ($i = 1; $i -le $Rows; $i++) {
        if ($i -eq $BlankAt) { continue }
        if ($RowMode -eq 'Visible' -and $hideSet -contains $i) { continue }
        $included += $i
    }
    $expectedWritten = $included.Count
    Write-Output "   mode=$RowMode hidden=[$($hideSet -join ',')] expected rows=$expectedWritten"
    if ($json.rowsRead -ne $expectedWritten) { $failures.Add("rowsRead: expected $expectedWritten, got $($json.rowsRead)") }
    if ($json.rowsWritten -ne $expectedWritten) { $failures.Add("rowsWritten: expected $expectedWritten, got $($json.rowsWritten)") }

    Write-Output '== verify target cells =='
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($json.output), 0, $true)
    $ws = $wb.Worksheets.Item($json.targetSheet)

    $checked = 0
    $skippedSources = @($json.mappings | Where-Object { $null -eq (Get-ExpectedValue $_.sourceColumn 1) } | ForEach-Object { $_.sourceColumn })
    foreach ($m in $json.mappings) {
        if ($null -eq (Get-ExpectedValue $m.sourceColumn 1)) { continue }
        $col = ConvertTo-ColumnNumber $m.targetColumn
        for ($k = 0; $k -lt $included.Count; $k++) {
            # Verified behaviour: skipped source rows (blank, or hidden in
            # Visible mode) are DROPPED, not copied as empty rows, and the
            # survivors are placed contiguously starting at DataStartRow.
            # This is compaction, not index preservation - asserting otherwise
            # would encode a behaviour the mapper does not have.
            $index = $included[$k]
            $targetRow = $DataStartRow + $k
            $expected = Get-ExpectedValue $m.sourceColumn $index
            $got = [string]$ws.Cells.Item($targetRow, $col).Value2
            $checked++
            if ($got -ne $expected) {
                $failures.Add("$($m.targetColumn)$targetRow : expected [$expected] got [$got]  (source $($m.sourceColumn), row index $index)")
            }
        }
    }

    # Nothing may spill past the last written row.
    $spillRow = $DataStartRow + $expectedWritten
    foreach ($m in $json.mappings) {
        if ($null -eq (Get-ExpectedValue $m.sourceColumn 1)) { continue }
        $col = ConvertTo-ColumnNumber $m.targetColumn
        $spill = [string]$ws.Cells.Item($spillRow, $col).Value2
        if (-not [string]::IsNullOrEmpty($spill)) {
            $failures.Add("spill at $($m.targetColumn)$spillRow = [$spill] (should be empty)")
        }
    }

    Write-Output "   mappings checked: $($json.mappings.Count)   cells checked: $checked"
    if ($skippedSources.Count -gt 0) {
        Write-Output "   NOTE: no expectation defined for source column(s) $($skippedSources -join ', ') - not checked"
    }
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
Write-Output 'PASS: all multi-row assertions hold.'
exit 0
