param(
    [string]$Source,
    [string]$Target,
    [int]$DataStartRow = 7,
    [switch]$KeepOutput
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

# 需求第 15 条要求不破坏模板自带格式。2026-09-10 起 excel_mapper.ps1 的
# Set-CellValue 对新写入 Color 使用 -4131（Left），其他字段使用 5（Fill），
# 这是对写入单元格**主动修改水平对齐**。
#
# tests/compare_ooxml.py 只看 dataValidation / conditionalFormatting / 公式，
# **完全看不到对齐**——只跑它就说「格式保留通过」是一句空话。本脚本补这个盲区。
#
# 三条断言：
#   1. 新写入 Color 左对齐，其余映射字段填充对齐
#   2. 表头区（数据行以上）与模板**逐格一致**——防止对齐写入连带污染共享样式
#   3. 写入区之外（溢出行的映射列、非映射列）与模板逐格一致
#
# 第 2 条是本轮真正的风险点：Excel 里改一个单元格的对齐，通常会给它分配新的
# cellXfs 项而不动别人，但这是**要验的假设**，不是可以直接信的常识。

if (-not $Source) { $Source = Join-Path (Split-Path $PSScriptRoot -Parent) 'tests\fixtures\A_sample.xls' }
if (-not $Target) { $Target = Join-Path (Split-Path $PSScriptRoot -Parent) '文件\B模板01.xlsx' }

$repo = Split-Path $PSScriptRoot -Parent
$mapper = Join-Path $repo 'excel_mapper.ps1'

$work = Join-Path $env:TEMP ('wsa_align_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null
$output = Join-Path $work 'out.xlsx'

$failures = New-Object System.Collections.Generic.List[string]

function Get-Alignment($Worksheet, [int]$Row, [int]$Column) {
    $cell = $Worksheet.Cells.Item($Row, $Column)
    try {
        # Merged cells return $null here; that is a value to compare, not an error.
        return $cell.HorizontalAlignment
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
    }
}

function ConvertTo-ColumnNumber([string]$Letter) {
    $n = 0
    foreach ($ch in $Letter.ToUpperInvariant().ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    return $n
}

Write-Output '== run mapper =='
$raw = & powershell -NoProfile -ExecutionPolicy Bypass -File $mapper `
    -SourcePath $Source -TargetPath $Target -OutputPath $output 2>&1
$exitCode = $LASTEXITCODE
$json = $null
try { $json = ($raw | Out-String).Trim() | ConvertFrom-Json } catch { }
if ($null -eq $json) {
    Write-Output ($raw | Out-String)
    throw "Mapper produced no parseable JSON (exit=$exitCode)."
}
Write-Output "   exit=$exitCode success=$($json.success) rowsWritten=$($json.rowsWritten)"
$written = [int]$json.rowsWritten
if ($exitCode -ne 0 -or -not $json.success -or $written -eq 0) { throw 'No successful data write to validate.' }
if ($json.writeStartRow) { $DataStartRow = [int]$json.writeStartRow }

$mappedColumns = @($json.mappings | ForEach-Object { ConvertTo-ColumnNumber $_.targetColumn })
Write-Output "   mapped target columns: $(($json.mappings | ForEach-Object { $_.targetColumn }) -join ', ')"

$excel = $null; $tplWb = $null; $outWb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $tplWb = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($Target), 0, $true)
    $outWb = $excel.Workbooks.Open([System.IO.Path]::GetFullPath($json.output), 0, $true)
    $tpl = $tplWb.Worksheets.Item($json.targetSheet)
    $out = $outWb.Worksheets.Item($json.targetSheet)

    $used = $tpl.UsedRange
    $lastCol = $used.Column + $used.Columns.Count - 1
    $lastRow = $used.Row + $used.Rows.Count - 1
    Write-Output "   template used range: rows 1..$lastRow, cols 1..$lastCol"

    # --- 断言 1：按实际 Color 表头独立确认对齐 ---
    $checkedData = 0
    $colorColumns = @()
    $headerMatches = 0
    for ($r = 1; $r -le 10; $r++) {
        $labels = @{}
        for ($c = $used.Column; $c -le $lastCol; $c++) {
            $labels[$c] = ([string]$tpl.Cells.Item($r, $c).Text).Trim().ToLowerInvariant()
        }
        if ($labels.Values -contains 'sku' -and $labels.Values -contains 'product name') {
            $headerMatches++
            $colorColumns = @($labels.Keys | Where-Object { $labels[$_] -eq 'color' })
        }
    }
    if ($headerMatches -ne 1 -or $colorColumns.Count -ne 1 -or $mappedColumns -notcontains $colorColumns[0]) { throw 'Expected one mapped Color header on the field-name row.' }
    for ($k = 0; $k -lt $written; $k++) {
        $r = $DataStartRow + $k
        foreach ($c in $mappedColumns) {
            $a = Get-Alignment $out $r $c
            $checkedData++
            $expected = if ($colorColumns -contains $c) { -4131 } else { 5 }
            if ($null -eq $a -or [int]$a -ne $expected) {
                $failures.Add("data cell r${r}c${c}: HorizontalAlignment expected $expected, got [$a]")
            }
        }
    }
    Write-Output "== assertion 1: Color is Left-aligned; other mapped fields are Fill-aligned =="
    Write-Output "   checked $checkedData cell(s)"

    # --- 断言 2：表头区与模板逐格一致 ---
    $checkedHeader = 0
    for ($r = $used.Row; $r -lt $DataStartRow; $r++) {
        for ($c = $used.Column; $c -le $lastCol; $c++) {
            $ta = Get-Alignment $tpl $r $c
            $oa = Get-Alignment $out $r $c
            $checkedHeader++
            if ($ta -ne $oa) {
                $failures.Add("header cell r${r}c${c}: template alignment [$ta] but output [$oa] - template style was disturbed")
            }
        }
    }
    Write-Output '== assertion 2: header area unchanged vs template =='
    Write-Output "   checked $checkedHeader cell(s) over rows $($used.Row)..$($DataStartRow - 1)"

    # --- 断言 3：写入区之外与模板逐格一致 ---
    # 注意模板的 UsedRange 只到表头（rows 1..6），所以「溢出行」通常根本不在
    # UsedRange 内，只比对溢出行会得到一个几乎空转的断言。改为扫描数据区
    # 一个 10 行的块 × 全部列：唯一允许出现差异的位置是「实际写入的行 × 映射列」，
    # 而那已由断言 1 覆盖，此处必须全部相同。
    $checkedOutside = 0
    $blockEnd = $DataStartRow + [Math]::Max($written, 1) + 9
    for ($r = $DataStartRow; $r -le $blockEnd; $r++) {
        for ($c = $used.Column; $c -le $lastCol; $c++) {
            $isWrittenCell = ($r -lt ($DataStartRow + $written)) -and ($mappedColumns -contains $c)
            if ($isWrittenCell) { continue }
            $ta = Get-Alignment $tpl $r $c
            $oa = Get-Alignment $out $r $c
            $checkedOutside++
            if ($ta -ne $oa) {
                $failures.Add("untouched cell r${r}c${c}: template alignment [$ta] but output [$oa]")
            }
        }
    }
    Write-Output '== assertion 3: cells outside the written area unchanged =='
    Write-Output "   checked $checkedOutside cell(s) over rows $DataStartRow..$blockEnd x cols $($used.Column)..$lastCol"
}
finally {
    if ($tplWb) { try { $tplWb.Close($false) } catch {} }
    if ($outWb) { try { $outWb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($o in @($tplWb, $outWb, $excel)) {
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
Write-Output 'PASS: Color is Left-aligned, other mapped cells are Fill-aligned; other alignments are unchanged.'
exit 0
