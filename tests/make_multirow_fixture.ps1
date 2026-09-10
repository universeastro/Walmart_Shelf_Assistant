param(
    [string]$SourceFixture,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [int]$Rows = 5,
    # 1-based data row index to leave completely blank in every mapped column.
    # 0 = no blank row. Probes how the mapper treats gaps in the source.
    [int]$BlankAt = 0,
    # Comma-separated 1-based data row indices to HIDE (as if the user had
    # filtered them out in Excel). e.g. '3,5'.
    [string]$HideAt = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

if (-not $SourceFixture) { $SourceFixture = Join-Path $PSScriptRoot 'fixtures\A_sample.xls' }
if ($Rows -lt 1) { throw 'Rows must be >= 1.' }
if ($BlankAt -lt 0 -or $BlankAt -gt $Rows) { throw 'BlankAt must be 0..Rows.' }

# Source columns the mapper reads. Header names are authoritative (project
# rule: the documented column letters are stale). Fallback letters are kept
# so the fixture can still be rebuilt if the sample is ever re-exported.
#
# 父SKU is deliberately written to E, NOT to its own header on B: the mapper
# maps that field via fallback-column because B is a helper column. Writing
# the fixture to B would make the test pass while the real source stays empty.
$COLUMNS = @(
    @{ Role = 'Sku';      Header = '自定义SKU';           Fallback = 'D';  ByName = $true  },
    @{ Role = 'Parent';   Header = '父SKU';               Fallback = 'E';  ByName = $false },
    @{ Role = 'Color';    Header = '颜色';                Fallback = 'L';  ByName = $true  },
    @{ Role = 'Title';    Header = '标题';                Fallback = 'N';  ByName = $true  },
    @{ Role = 'LongDesc'; Header = '长描述';              Fallback = 'O';  ByName = $true  },
    @{ Role = 'F1';       Header = '五点1';               Fallback = 'P';  ByName = $true  },
    @{ Role = 'F2';       Header = '五点2';               Fallback = 'Q';  ByName = $true  },
    @{ Role = 'F3';       Header = '五点3';               Fallback = 'R';  ByName = $true  },
    @{ Role = 'F4';       Header = '五点4';               Fallback = 'S';  ByName = $true  },
    @{ Role = 'F5';       Header = '五点5';               Fallback = 'T';  ByName = $true  },
    @{ Role = 'Link1';    Header = '小平台专用链接 1';     Fallback = 'AG'; ByName = $true  },
    @{ Role = 'Link2';    Header = '小平台专用链接 2';     Fallback = 'AI'; ByName = $true  },
    @{ Role = 'Swatch';   Header = '代理链接100*100缩率图'; Fallback = 'AX'; ByName = $true  }
)

function ConvertTo-ColumnNumber([string]$Letter) {
    $n = 0
    foreach ($ch in $Letter.ToUpperInvariant().ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    return $n
}

function ConvertTo-ColumnLetter([int]$Number) {
    $r = ''
    while ($Number -gt 0) { $Number--; $r = [char](65 + ($Number % 26)) + $r; $Number = [math]::Floor($Number / 26) }
    return $r
}

# Read headers with Value2, not Text: .Text is the DISPLAYED string and is
# truncated to the column width, so a narrow column can hide the real header.
function Find-HeaderColumn($Worksheet, [string]$Header) {
    $used = $Worksheet.UsedRange
    $last = $used.Column + $used.Columns.Count - 1
    for ($c = $used.Column; $c -le $last; $c++) {
        if ([string]$Worksheet.Cells.Item($used.Row, $c).Value2 -eq $Header) { return $c }
    }
    return 0
}

$tempPath = $null
$excel = $null; $wb = $null; $ws = $null
try {
    $source = [System.IO.Path]::GetFullPath($SourceFixture)
    if (-not (Test-Path -LiteralPath $source)) { throw "Source fixture not found: $source" }
    $out = [System.IO.Path]::GetFullPath($OutputPath)
    if (-not [System.IO.Path]::GetExtension($out)) { $out += '.xlsx' }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Edit a throwaway copy so the committed fixture is never touched.
    $tempPath = Join-Path ([System.IO.Path]::GetDirectoryName($out)) ('.wsa_fixture_' + [guid]::NewGuid().ToString('N') + '.xlsx')
    Copy-Item -LiteralPath $source -Destination $tempPath -Force
    $wb = $excel.Workbooks.Open($tempPath, 0, $false)
    $ws = $wb.Worksheets.Item(1)

    # Resolve every column before writing anything, so a missing header fails
    # loudly instead of silently producing a half-built fixture.
    $resolved = New-Object System.Collections.Generic.List[object]
    foreach ($col in $COLUMNS) {
        $number = 0
        if ($col.ByName) { $number = Find-HeaderColumn $ws $col.Header }
        if ($number -eq 0) { $number = ConvertTo-ColumnNumber $col.Fallback }
        if ($number -eq 0) { throw "Cannot resolve column for $($col.Role)" }
        $resolved.Add([pscustomobject]@{ Role = $col.Role; Column = $number; Letter = ConvertTo-ColumnLetter $number })
    }

    $headerRow = $ws.UsedRange.Row

    # Clear the data area first: the sample already holds one row, and a
    # leftover value would make a "blank" row look populated.
    for ($r = $headerRow + 1; $r -le $headerRow + $Rows; $r++) {
        # ClearContents, NOT `.Value2 = $null`. Assigning $null to that
        # property makes every SUBSEQUENT write in this script silently fail
        # (3/3 runs, all 13 columns, readback empty) while raising no error.
        # `.ClearContents()` and `.Value2 = ''` both behave. The mechanism was
        # not isolated - a minimal repro of the same loop works fine - so this
        # is recorded as an observed trap, not an explained one.
        # [void] is load-bearing: an unsuppressed ClearContents() emits one
        # empty value per call, so a 5x13 clear would prepend 65 blank lines
        # to this script's output and corrupt anything parsing it.
        foreach ($m in $resolved) { [void]$ws.Cells.Item($r, $m.Column).ClearContents() }
    }

    for ($i = 1; $i -le $Rows; $i++) {
        if ($i -eq $BlankAt) { continue }
        $row = $headerRow + $i
        $values = @{
            Sku      = 'CUST-{0:000}' -f $i
            Parent   = 'GROUP-{0:000}' -f $i
            Color    = "Red-$i"
            Title    = "Sample product $i"
            LongDesc = "Long description $i"
            F1       = "Feature one $i"
            F2       = "Feature two $i"
            F3       = "Feature three $i"
            F4       = "Feature four $i"
            F5       = "Feature five $i"
            Link1    = "https://img.example/one-$i.jpg"
            Link2    = "https://img.example/main-$i.jpg"
            Swatch   = "https://img.example/swatch-$i.jpg"
        }
        foreach ($m in $resolved) { $ws.Cells.Item($row, $m.Column).Value2 = $values[$m.Role] }
    }

    # Hide AFTER writing: a hidden row still holds its values, which is the
    # whole point - the mapper must skip it by visibility, not by emptiness.
    $hidden = New-Object System.Collections.Generic.List[int]
    foreach ($token in ($HideAt -split ',')) {
        $n = 0
        if (-not [int]::TryParse($token.Trim(), [ref]$n)) { continue }
        if ($n -lt 1 -or $n -gt $Rows) { throw "HideAt index $n out of range 1..$Rows" }
        $rowRange = $ws.Rows.Item($headerRow + $n)
        try { $rowRange.Hidden = $true } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rowRange) }
        $hidden.Add($n)
    }

    # A hidden row cannot be read back through the same path as a visible one,
    # so assert visibility separately from value content.
    foreach ($n in $hidden) {
        $rowRange = $ws.Rows.Item($headerRow + $n)
        try { $isHidden = [bool]$rowRange.Hidden } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rowRange) }
        if (-not $isHidden) { throw "Row $($headerRow + $n) was not actually hidden." }
    }

    # Assert the values actually stuck. A protected sheet, a merged-cell
    # anchor, or a silently-ignored write would otherwise yield an all-blank
    # fixture that every downstream assertion passes against.
    for ($i = 1; $i -le $Rows; $i++) {
        if ($i -eq $BlankAt) { continue }
        $row = $headerRow + $i
        foreach ($m in $resolved) {
            $got = [string]$ws.Cells.Item($row, $m.Column).Value2
            if ([string]::IsNullOrEmpty($got)) {
                throw "Write did not stick: row $row column $($m.Letter) ($($m.Role)) reads back empty."
            }
        }
    }

    $wb.SaveAs($out)

    # Re-open the saved file and re-assert: catches a write that survives in
    # memory but is lost on save.
    $verifyWb = $excel.Workbooks.Open($out, 0, $true)
    try {
        $verifyWs = $verifyWb.Worksheets.Item(1)
        for ($i = 1; $i -le $Rows; $i++) {
            if ($i -eq $BlankAt) { continue }
            $row = $headerRow + $i
            foreach ($m in $resolved) {
                $got = [string]$verifyWs.Cells.Item($row, $m.Column).Value2
                if ([string]::IsNullOrEmpty($got)) {
                    throw "Saved fixture lost data: row $row column $($m.Letter) ($($m.Role)) is empty after SaveAs."
                }
            }
        }
    }
    finally {
        if ($verifyWb) { try { $verifyWb.Close($false) } catch {} }
    }

    Write-Output ('fixture=' + $out)
    Write-Output ('source=' + $source + '  headerRow=' + $headerRow + '  rows=' + $Rows + '  blankAt=' + $BlankAt + '  hidden=[' + ($hidden -join ',') + ']')
    foreach ($m in $resolved) { Write-Output ('  ' + $m.Role.PadRight(9) + $m.Letter.PadRight(4) + 'col' + $m.Column) }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($o in @($ws, $wb, $excel)) {
        if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    if ($tempPath -and (Test-Path -LiteralPath $tempPath)) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
