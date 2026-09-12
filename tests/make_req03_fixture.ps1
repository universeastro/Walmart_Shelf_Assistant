$ErrorActionPreference = 'Stop'
$fixtureDir = Join-Path $PSScriptRoot 'fixtures'
$sourcePath = Join-Path $fixtureDir 'A_req03_sample.xls'
$targetPath = Join-Path $fixtureDir 'B_req03_sample.xls'
$excel = $null
$workbook = $null
try {
    [void](New-Item -ItemType Directory -Path $fixtureDir -Force)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) { $workbook.Worksheets.Item(2).Delete() }
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Source'
    foreach ($item in @(
        @(4, '自定义'), @(27, 'SKU价(￥)'), @(31, '重量(g)'),
        @(51, '长'), @(52, '宽'), @(53, '高')
    )) {
        $sheet.Cells.Item(1, $item[0]).Value2 = $item[1]
        $sheet.Columns.Item($item[0]).NumberFormat = '@'
    }
    $rows = @(
        @{ Row = 2; Values = @('SKU-TEXT-001', '47', '21.5', '50', '30', '12') },
        @{ Row = 3; Values = @('SKU-HIDDEN', '18.25', '7', '1', '2', '3') },
        @{ Row = 5; Values = @('SKU-TEXT-002', 'not-a-number', '', '10.5', '20', '30') }
    )
    $columns = @(4, 27, 31, 51, 52, 53)
    foreach ($row in $rows) {
        for ($index = 0; $index -lt $columns.Count; $index++) {
            $sheet.Cells.Item($row.Row, $columns[$index]).Value2 = $row.Values[$index]
        }
    }
    $sheet.Rows.Item(3).Hidden = $true
    $workbook.SaveAs($sourcePath, -4143)
    $workbook.Close($false)
    $workbook = $null

    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) { $workbook.Worksheets.Item(2).Delete() }
    $price = $workbook.Worksheets.Item(1)
    $price.Name = '价格'
    $target = $workbook.Worksheets.Add()
    $target.Name = '导入 单位转换'
    $shipping = $workbook.Worksheets.Add()
    $shipping.Name = '运费表（公式数据，不动）'
    $headers = @(
        @(1, 'SKU(直接从sheet 1导入）'), @(2, 'SKU价(￥)'), @(3, '重量(g)'),
        @(8, '长'), @(9, '宽'), @(10, '高')
    )
    foreach ($sheet in @($price, $target)) {
        foreach ($item in $headers) { $sheet.Cells.Item(1, $item[0]).Value2 = $item[1] }
    }
    $target.Range('A1:R1').Font.Bold = $true
    $target.Range('A1:R1').Interior.Color = 15773696
    $target.Range('A2:A20').NumberFormat = '@'
    $target.Range('Q1:Q20').NumberFormat = '@'
    $target.Range('Q2:Q20').Value2 = '@100'
    $target.Range('D1').Value2 = '重量(lb)'
    $target.Range('E1').Value2 = '重量(lb)两位小数'
    $target.Range('K1').Value2 = '长in'
    $target.Range('L1').Value2 = '宽in'
    $target.Range('M1').Value2 = '高in'
    $target.Range('N1').Value2 = '体积'
    $target.Range('R1').Value2 = '结果'
    $target.Range('D2').Formula = '=IF(C2="","",C2/453.592)'
    $target.Range('E2').Formula = '=IF(D2="","",ROUND(D2,2))'
    $target.Range('K2').Formula = '=IF(H2="","",H2/2.54)'
    $target.Range('L2').Formula = '=IF(I2="","",I2/2.54)'
    $target.Range('M2').Formula = '=IF(J2="","",J2/2.54)'
    $target.Range('N2').Formula = '=IF(COUNT(H2:J2)=3,H2*I2*J2,"")'
    $target.Range('R2').Formula = '=IF(A2="","",A2&"-OK")'
    foreach ($letter in @('D', 'E', 'K', 'L', 'M', 'N', 'R')) {
        [void]$target.Range("${letter}2:${letter}20").FillDown()
    }
    $shipping.Range('A1').Value2 = 'KEEP-SHIPPING'
    $shipping.Range('A2').Formula = '=1+1'
    $workbook.SaveAs($targetPath, -4143)
    $workbook.Close($false)
    $workbook = $null
    Write-Output "Created: $sourcePath"
    Write-Output "Created: $targetPath"
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
}
