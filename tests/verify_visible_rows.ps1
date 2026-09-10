param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [ValidateSet('Visible', 'All')][string]$RowMode = 'Visible'
)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$src = $null
$out = $null
try {
    $src = $excel.Workbooks.Open((Resolve-Path $SourcePath).Path, 0, $true)
    $out = $excel.Workbooks.Open((Resolve-Path $OutputPath).Path, 0, $true)
    $s = $src.Worksheets.Item(1)
    $t = $out.Worksheets.Item('Product Content And Site Exp')
    $x = $s.UsedRange.Value2
    $y = $t.UsedRange.Value2
    $pairs = @(@(4,4),@(14,8),@(15,17),@(16,18),@(17,19),@(18,20),@(35,21),@(12,25),@(33,30),@(35,31),@(5,65),@(50,69))
    $dest = 7
    $hidden = 0
    $checked = 0
    for ($r=2; $r -le $x.GetLength(0); $r++) {
        if ($s.Rows.Item($r).Hidden -and $RowMode -eq 'Visible') { $hidden++; continue }
        $populated = $false
        foreach ($pair in $pairs) {
            if ($null -ne $x[$r,$pair[0]] -and [string]$x[$r,$pair[0]] -ne '') { $populated = $true }
        }
        if (-not $populated) { continue }
        foreach ($pair in $pairs) {
            if ([string]$x[$r,$pair[0]] -cne [string]$y[$dest,$pair[1]]) {
                throw "Mismatch source row $r column $($pair[0]); output row $dest column $($pair[1])"
            }
            $checked++
        }
        $dest++
    }
    if ($y.GetLength(0) -ne $dest-1) { throw 'Unexpected output row count' }
    Write-Output "PASS: mode=$RowMode; data rows=$($dest-7); skipped hidden=$hidden; compared cells=$checked; contiguous output starts at row 7"
} finally {
    if ($src) { try { $src.Close($false) } catch {} }
    if ($out) { try { $out.Close($false) } catch {} }
    try { $excel.Quit() } catch {}
}
