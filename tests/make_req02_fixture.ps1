$ErrorActionPreference = 'Stop'
$fixtureDir = Join-Path $PSScriptRoot 'fixtures'
$sourcePath = Join-Path $fixtureDir 'A_req02_sample.xlsx'
$targetPath = Join-Path $fixtureDir 'B_req02_sample.xlsx'
$excel = $null
$workbook = $null
try {
    [void](New-Item -ItemType Directory -Path $fixtureDir -Force)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Source'
    $sheet.Range('A1').Value2 = 'SKU'
    $sheet.Range('D1').Value2 = '自定义'
    $sheet.Range('A2').Value2 = 'SKU-001'
    $sheet.Range('D2').Value2 = 'PLATFORM-001'
    $sheet.Range('A3').Value2 = 'SKU-HIDDEN'
    $sheet.Range('D3').Value2 = 'PLATFORM-HIDDEN'
    $sheet.Rows.Item(3).Hidden = $true
    $sheet.Range('A5').Value2 = 'SKU-002'
    $sheet.Range('D5').Value2 = 'PLATFORM-002'
    $workbook.SaveAs($sourcePath, 51)
    $workbook.Close($false)
    $workbook = $null

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Sheet1'
    $sheet.Range('A1').Value2 = 'SKU'
    $sheet.Range('B1').Value2 = '平台SKU'
    $sheet.Range('A1:B1').Font.Bold = $true
    $sheet.Range('A1:B1').Interior.Color = 65535
    for ($row = 2; $row -le 8; $row++) {
        $sheet.Cells.Item($row, 1).Value2 = "OLD-$row"
        $sheet.Cells.Item($row, 2).Value2 = "OLD-PLATFORM-$row"
    }
    $sheet.Range('C2').Formula = '=1+1'
    $workbook.SaveAs($targetPath, 51)
    $workbook.Close($false)
    $workbook = $null
    Write-Output "Created: $sourcePath"
    Write-Output "Created: $targetPath"
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
}
