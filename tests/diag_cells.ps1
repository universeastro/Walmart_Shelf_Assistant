param(
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)][string]$SheetName,
    # Comma-separated string, not [string[]]: an array parameter cannot be bound
    # from an external shell reliably (each element is read as a positional arg).
    [string]$Cells = 'CA1,CA3,CB1,CB3,CC1,CC3'
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Output ('Application.ReferenceStyle = ' + $excel.ReferenceStyle + '  (1 = xlA1, -1 = xlR1C1)')
    $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path, 0, $true)
    $ws = $wb.Worksheets.Item($SheetName)

    Write-Output ('sheet=' + $ws.Name + '  used=' + $ws.UsedRange.Address(0,0))
    Write-Output ''
    foreach ($addr in ($Cells -split ',')) {
        $c = $ws.Range($addr)
        Write-Output ('--- ' + $addr + ' ---')
        Write-Output ('  MergeCells   = ' + $c.MergeCells)
        Write-Output ('  HasFormula   = ' + $c.HasFormula)
        Write-Output ('  HasArray     = ' + $c.HasArray)
        Write-Output ('  ListObject   = ' + $(if ($null -eq $c.ListObject) { '(none)' } else { $c.ListObject.Name }))
        Write-Output ('  Formula      = ' + [string]$c.Formula)
        Write-Output ('  FormulaR1C1  = ' + [string]$c.FormulaR1C1)
        Write-Output ('  Value2       = ' + [string]$c.Value2)
        Write-Output ('  NumberFormat = ' + [string]$c.NumberFormat)
    }
    Write-Output ''
    Write-Output '== all formulas on this sheet =='
    foreach ($c in $ws.UsedRange.SpecialCells(-4123)) {   # xlCellTypeFormulas
        Write-Output ('  ' + $c.Address(0,0) + '  array=' + $c.HasArray + '  ' + [string]$c.Formula)
    }
    Write-Output '== end =='
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($o in @($wb, $excel)) {
        if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
