$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$work = Join-Path $env:TEMP ('wsa_append_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $work)
$template = Join-Path $repo '文件\B模板01.xlsx'
$populated = Join-Path $work 'populated.xlsx'
$conflict = Join-Path $work 'conflict.xlsx'
$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($template, 0, $true)
    $ws = $wb.Worksheets.Item('Product Content And Site Exp')
    $ws.Range('D7').Value2 = 'OLD-SKU'
    $ws.Range('H7').Value2 = 'OLD-TITLE'
    $ws.Range('G16').Value2 = 0
    $ws.Rows.Item(16).Hidden = $true
    $ws.Range('J50').NumberFormat = '0.00'
    $ws.Range('J50').Formula = '=1+2'
    if (-not $ws.Range('J50').HasFormula) { throw 'Fixture formula J50 was not created' }
    $ws.Range('D100').Interior.Color = 65535
    $wb.SaveAs($populated, 51)
    $ws.Range('H20').NumberFormat = '0.00'
    $ws.Range('H20').Formula = '=1+1'
    if (-not $ws.Range('H20').HasFormula) { throw 'Fixture formula H20 was not created' }
    $wb.SaveAs($conflict, 51)
    $wb.Close($false)
    $wb = $null
    $excel.Quit()
    $excel = $null

    foreach ($mode in @('Visible', 'All')) {
        & (Join-Path $PSScriptRoot 'verify_multirow.ps1') -Rows 5 -HideAt '3,5' -RowMode $mode -Target $populated -DataStartRow 20
        if ($LASTEXITCODE -ne 0) { throw "Multirow append failed: $mode" }
    }

    $output = Join-Path $work 'output.xlsx'
    $before = (Get-FileHash $populated).Hash
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'excel_mapper.ps1') -SourcePath (Join-Path $PSScriptRoot 'fixtures\A_sample.xls') -TargetPath $populated -OutputPath $output
    if ($LASTEXITCODE -ne 0) { throw ($raw | Out-String) }
    $result = $raw | ConvertFrom-Json
    if ($result.existingLastRow -ne 16 -or $result.writeStartRow -ne 20) { throw 'Wrong append location' }
    if ((Get-FileHash $populated).Hash -ne $before) { throw 'Original template changed' }
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($output, 0, $true)
    $ws = $wb.Worksheets.Item('Product Content And Site Exp')
    if ($ws.Range('D7').Value2 -cne 'OLD-SKU' -or $ws.Range('H7').Value2 -cne 'OLD-TITLE') { throw 'Existing data changed' }
    if ($ws.Range('G16').Value2 -ne 0) { throw 'Existing non-mapped data changed' }
    if ($ws.Range('J50').Formula -ne '=1+2') { throw 'Formula changed' }
    if ($excel.WorksheetFunction.CountA($ws.Range('A17:CP19')) -ne 0) { throw 'Gap is not three empty rows' }
    if ($ws.Range('D20').Value2 -cne 'CUST-001') { throw 'New row missing' }
    $wb.Close($false)
    $wb = $null
    $excel.Quit()
    $excel = $null

    $blocked = Join-Path $work 'blocked.xlsx'
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'excel_mapper.ps1') -SourcePath (Join-Path $PSScriptRoot 'fixtures\A_sample.xls') -TargetPath $conflict -OutputPath $blocked
    if ($LASTEXITCODE -eq 0 -or (Test-Path $blocked)) { throw 'Formula collision was not blocked' }
    $failure = $raw | ConvertFrom-Json
    if ($failure.success -or $failure.stage -ne '检查追加位置' -or $failure.message -notmatch 'H20:H20') {
        throw 'Failure was not the expected formula collision'
    }
    Write-Output 'PASS: append location, both row modes, old data, gap, template hash, formula protection.'
} finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    Write-Output "Audit artifacts: $work"
}
