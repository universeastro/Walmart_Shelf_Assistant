# Verify that excel_mapper.ps1 wrote each source field into the correct target column.
#
# ASCII-only on purpose: PowerShell 5.1 reads BOM-less .ps1 as GBK and mangles
# any non-ASCII literal. Keep this file free of non-ASCII characters.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File verify_mapping.ps1 `
#     -SourcePath <file A with data> -TargetPath <template B> [-OutputPath <result>]
#
# If -OutputPath is omitted the mapper is run first and the result verified.

param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$TargetPath,
    [string]$OutputPath,
    [string]$MapperPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $MapperPath) {
    # $PSScriptRoot is <repo>/.claude/skills/verify-mapping/scripts -> climb four levels.
    $repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
    $MapperPath = Join-Path $repoRoot 'excel_mapper.ps1'
}
if (-not $OutputPath) {
    $OutputPath = Join-Path $env:TEMP ('wsa_verify_' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsx')
}

# source column -> target column, keyed by the mapper's mapping table.
# Kept here so the check does not depend on parsing the mapper's JSON output.
$Pairs = @(
    @('D',  'D',  'CustomSKU -> sku'),
    @('N',  'H',  'Title -> product name'),
    @('O',  'Q',  'LongDescription -> site description'),
    @('P',  'R',  'Bullet1 -> key features (+)'),
    @('Q',  'S',  'Bullet2 -> key features 1 (+)'),
    @('R',  'T',  'Bullet3 -> key features 2 (+)'),
    @('AI', 'U',  'Link2 -> main image url'),
    @('L',  'Y',  'Color -> color'),
    @('AG', 'AD', 'Link1 -> additional image url (+)'),
    @('AI', 'AE', 'Link2b -> additional image url 1 (+)'),
    @('E',  'BM', 'ParentSKU -> variant group id'),
    @('AX', 'BQ', 'Thumb100 -> swatch image url')
)

function Get-ColNum([string]$Letter) {
    $n = 0
    foreach ($ch in $Letter.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    return $n
}

# Merged cells report empty .Text on every cell except the anchor, which makes
# a naive read look like a missing value. Always resolve through MergeArea.
function Get-CellText($Worksheet, [int]$Row, [int]$Col) {
    $cell = $Worksheet.Cells.Item($Row, $Col)
    if ($cell.MergeCells) { return [string]$cell.MergeArea.Cells.Item(1, 1).Text }
    return [string]$cell.Text
}

# Sheet file names are NOT stable across an Excel re-save: the mapper's output
# can renumber worksheets/sheetN.xml. Resolve sheets by workbook.xml name.
function Get-SheetNameToFile([string]$Root) {
    $map = @{}
    $wb = [xml](Get-Content -Encoding UTF8 (Join-Path $Root 'xl/workbook.xml'))
    $rels = [xml](Get-Content -Encoding UTF8 (Join-Path $Root 'xl/_rels/workbook.xml.rels'))
    $relMap = @{}
    foreach ($r in $rels.Relationships.Relationship) { $relMap[$r.Id] = ($r.Target -replace '^/?xl/', '') }
    foreach ($s in $wb.workbook.sheets.sheet) {
        $rid = $s.id
        if (-not $rid) { $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') }
        $map[$s.name] = $relMap[$rid]
    }
    return $map
}

if (-not (Test-Path $OutputPath)) {
    Write-Output "Running mapper: $SourcePath -> $TargetPath"
    & $MapperPath -SourcePath $SourcePath -TargetPath $TargetPath -OutputPath $OutputPath | Out-Null
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $src = $excel.Workbooks.Open((Resolve-Path $SourcePath).Path, 0, $true)
    $out = $excel.Workbooks.Open((Resolve-Path $OutputPath).Path, 0, $true)

    $sw = $src.Worksheets.Item(1)
    $targetSheetName = 'Product Content And Site Exp'
    $ow = $out.Worksheets.Item($targetSheetName)

    # Find the data row by locating a known source value, not by "first
    # non-empty cell" -- rows 2..5 of the template are hierarchical headers.
    $probe = [string]$sw.Cells.Item(2, (Get-ColNum 'D')).Text
    if ($probe -eq '') { throw "Source row 2 column D is empty; nothing to verify." }

    $u = $ow.UsedRange
    $lastR = $u.Row + $u.Rows.Count - 1
    $lastC = $u.Column + $u.Columns.Count - 1
    $dataRow = -1
    for ($r = $u.Row; $r -le $lastR; $r++) {
        for ($c = $u.Column; $c -le $lastC; $c++) {
            if ((Get-CellText $ow $r $c) -eq $probe) { $dataRow = $r; break }
        }
        if ($dataRow -gt 0) { break }
    }

    Write-Output "Probe value '$probe' found at target row: $dataRow"
    Write-Output ''
    if ($dataRow -lt 0) {
        Write-Output 'FAIL: source data was not written into the output at all.'
        exit 2
    }

    Write-Output ('{0,-4} {1,-4} {2,-28} {3,-28} {4}' -f 'SRC', 'TGT', 'SOURCE', 'TARGET', 'VERDICT')
    Write-Output ('-' * 96)
    $pass = 0; $fail = 0; $empty = 0
    foreach ($p in $Pairs) {
        $sv = [string]$sw.Cells.Item(2, (Get-ColNum $p[0])).Text
        $tv = Get-CellText $ow $dataRow (Get-ColNum $p[1])
        if ($sv -eq $tv) {
            if ($sv -eq '') { $res = 'both-empty'; $empty++ } else { $res = 'OK'; $pass++ }
        } else { $res = '** MISMATCH **'; $fail++ }
        $a = $sv; $b = $tv
        if ($a.Length -gt 26) { $a = $a.Substring(0, 26) + '..' }
        if ($b.Length -gt 26) { $b = $b.Substring(0, 26) + '..' }
        Write-Output ('{0,-4} {1,-4} {2,-28} {3,-28} {4}  {5}' -f $p[0], $p[1], $a, $b, $res, $p[2])
    }
    Write-Output ('-' * 96)
    Write-Output "match=$pass  mismatch=$fail  both-empty=$empty"

    $src.Close($false)
    $out.Close($false)
    if ($fail -gt 0) { exit 1 }
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
