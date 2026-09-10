param(
    [Parameter(Mandatory=$true)][string]$TemplatePath,
    [Parameter(Mandatory=$true)][string]$FixturePath
)

# Builds a formula-bearing variant of the B template, used to test requirement 15:
# "template's own formats (including formulas) must not be destroyed".
#
# Deliberately ASCII-only: PowerShell 5.1 decodes a BOM-less .ps1 as GBK, so any
# non-ASCII literal here would be mangled. Pass Chinese paths in as parameters.
#
# Formulas go far to the right of the columns excel_mapper.ps1 writes to
# (D,H,Q,R,S,T,U,Y,AD,AE,BM,BQ), so they survive the mapping run and can be
# re-read from the output.
#
# Cells are PROBED, not hardcoded. Writing to a non-anchor cell of a merged range
# silently does nothing and reads back as an empty formula - the template has
# merged cells scattered across these columns, so guessing addresses produces a
# fixture that looks fine and contains no formulas at all.

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

$TARGET = 'Product Content And Site Exp'
$OTHER  = 'Data Definitions'
$COLS = 79..140      # CA..EL, far right of every mapped column
$ROWS = 1..4

function Get-FreeAddresses($ws, $n) {
    $out = @()
    foreach ($c in $COLS) {
        foreach ($r in $ROWS) {
            $cell = $ws.Cells.Item($r, $c)
            if ($cell.MergeCells) { continue }
            if ($cell.HasFormula) { continue }
            if (-not [string]::IsNullOrEmpty([string]$cell.Value2)) { continue }
            # A Text-formatted (@) cell stores an assigned formula as a literal
            # string and reads back looking like one. The sheet is protected and
            # forbids changing NumberFormat, so skip those cells rather than
            # trying to reformat them.
            if ([string]$cell.NumberFormat -match '@') { continue }
            $out += $cell.Address(0, 0)
            if ($out.Count -ge $n) { return $out }
        }
    }
    return $out
}

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Copy-Item -LiteralPath (Resolve-Path $TemplatePath).Path -Destination $FixturePath -Force
    $wb = $excel.Workbooks.Open((Resolve-Path $FixturePath).Path, 0, $false)

    $ws = $wb.Worksheets.Item($TARGET)
    $dd = $wb.Worksheets.Item($OTHER)

    # @() is load-bearing: PowerShell unrolls a single-element array into a scalar
    # on return, so a one-element result would make $otherCells[0] index into the
    # STRING ('CA1'[0] -> 'C') and Range('C') blows up with 0x800A03EC.
    $targetCells = @(Get-FreeAddresses $ws 4)
    if ($targetCells.Count -lt 4) { throw "only found $($targetCells.Count) free cells on $TARGET" }
    $otherCells = @(Get-FreeAddresses $dd 1)
    if ($otherCells.Count -lt 1) { throw "no free cell on $OTHER" }

    Write-Output ('probed target cells: ' + ($targetCells -join ', '))
    Write-Output ('probed other cell  : ' + ($otherCells -join ', '))
    Write-Output ''

    # Each assignment is isolated so a failure names the exact cell and formula.
    $steps = @(
        @{ Addr = $targetCells[0]; Kind = 'normal';    F = '=COUNTA(D7:D500)' },
        @{ Addr = $targetCells[1]; Kind = 'crosssheet'; F = "='$OTHER'!A1" },
        # LEN over a range needs CSE, so this stays a genuine array formula, and
        # unlike =SUM(D7:D9*1) it evaluates cleanly when D7 holds TEXT (a text
        # value times 1 is #VALUE!, which looks like mapper damage but is not).
        @{ Addr = $targetCells[2]; Kind = 'array';     F = '=SUM(LEN(D7:D9))' },
        @{ Addr = $targetCells[3]; Kind = 'refcell';   F = '=D7&"-suffix"' },
        @{ Addr = $otherCells[0];  Kind = 'crossback'; F = "='$TARGET'!D7" }
    )

    foreach ($s in $steps) {
        $sheet = if ($s.Kind -eq 'crossback') { $dd } else { $ws }
        $cell = $sheet.Range($s.Addr)

        try {
            if ($s.Kind -eq 'array') {
                $cell.FormulaArray = $s.F
            } else {
                $cell.Formula = $s.F
            }
            # Assert it really landed as a formula, not as text.
            if (-not $cell.HasFormula) {
                throw "cell $($s.Addr) reports HasFormula=False after the write (stored as text?)"
            }
            if ($s.Kind -eq 'array' -and -not $cell.HasArray) {
                throw "cell $($s.Addr) reports HasArray=False after FormulaArray (not an array formula)"
            }
            Write-Output ('  wrote ' + $s.Kind.PadRight(11) + $s.Addr.PadRight(8) + ' -> ' + $s.F)
        }
        catch {
            Write-Output ('  FAILED ' + $s.Kind.PadRight(11) + $s.Addr.PadRight(8) + ' -> ' + $s.F)
            Write-Output ('         merged=' + $cell.MergeCells + ' inTable=' + ($null -ne $cell.ListObject))
            throw
        }
    }

    $wb.Save()

    Write-Output ('target used range = ' + $ws.UsedRange.Address(0, 0))
    Write-Output ''
    Write-Output ('{0,-8} {1,-12} {2}' -f 'CELL', 'HASARRAY', 'FORMULA')
    Write-Output ('-' * 70)
    foreach ($a in $targetCells) {
        Write-Output ('{0,-8} {1,-12} {2}' -f $a, $ws.Range($a).HasArray, $ws.Range($a).Formula)
    }
    Write-Output ('{0,-8} {1,-12} {2}' -f $otherCells[0], $dd.Range($otherCells[0]).HasArray, $dd.Range($otherCells[0]).Formula)
    Write-Output ''
    Write-Output 'FIXTURE_OK'
}
catch {
    Write-Output ('FIXTURE_FAIL at script line ' + $_.InvocationInfo.ScriptLineNumber + ': ' + $_.Exception.Message)
    Write-Output ('  statement: ' + ([string]$_.InvocationInfo.Line).Trim())
    exit 1
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($o in @($wb, $excel)) {
        if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
