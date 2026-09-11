$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$source = Join-Path $PSScriptRoot 'fixtures\A_req02_sample.xlsx'
$target = Join-Path $PSScriptRoot 'fixtures\B_req02_sample.xlsx'
$work = Join-Path $env:TEMP ('wsa_req02_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $work)
$excel = $null
$sourceWb = $null
$targetWb = $null
try {
    if (-not (Test-Path -LiteralPath $source) -or -not (Test-Path -LiteralPath $target)) {
        throw '缺少 Req02 合成夹具；请先运行 tests/make_req02_fixture.ps1。'
    }
    $before = (Get-FileHash -LiteralPath $target).Hash
    $emptySource = Join-Path $work 'empty-source.xlsx'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $sourceWb = $excel.Workbooks.Add()
    $emptySheet = $sourceWb.Worksheets.Item(1)
    $emptySheet.Range('A1').Value2 = 'SKU'
    $emptySheet.Range('D1').Value2 = '自定义'
    $sourceWb.SaveAs($emptySource, 51)
    $sourceWb.Close($false)
    $sourceWb = $null
    $excel.Quit()
    $excel = $null

    $cases = @(
        @{ Source = $source; RowMode = 'Visible'; WriteMode = 'Append'; ExpectedRows = 2; ExpectedStart = 12 },
        @{ Source = $source; RowMode = 'All'; WriteMode = 'Append'; ExpectedRows = 3; ExpectedStart = 12 },
        @{ Source = $source; RowMode = 'Visible'; WriteMode = 'Replace'; ExpectedRows = 2; ExpectedStart = 2 },
        @{ Source = $emptySource; RowMode = 'Visible'; WriteMode = 'Replace'; ExpectedRows = 0; ExpectedStart = 2 }
    )
    foreach ($case in $cases) {
        $caseSource = $case.Source
        $mode = $case.RowMode
        $writeMode = $case.WriteMode
        $output = Join-Path $work "output-$mode-$writeMode-$($case.ExpectedRows).xlsx"
        $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'excel_mapper.ps1') `
            -SourcePath $caseSource -TargetPath $target -OutputPath $output -Profile Req02 `
            -WriteMode $writeMode -RowMode $mode
        if ($LASTEXITCODE -ne 0) { throw ($raw | Out-String) }
        $result = $raw | ConvertFrom-Json
        if (-not $result.success -or $result.profile -ne 'Req02') { throw "Req02 映射未成功: $mode" }
        if ($result.writeMode -ne $writeMode) { throw "支线写入模式摘要错误: $mode/$writeMode" }
        if ($result.rowsRead -ne $case.ExpectedRows -or $result.rowsWritten -ne $case.ExpectedRows) {
            throw "支线读写行数不符合夹具契约: $mode/$writeMode"
        }
        if ($result.existingLastRow -ne 8 -or $result.writeStartRow -ne $case.ExpectedStart -or $result.mappings.Count -ne 2) {
            throw "支线既有末行、写入起点或映射数量错误: $mode/$writeMode"
        }
        if ((Get-FileHash -LiteralPath $target).Hash -ne $before) { throw '原始支线 B 文件被修改。' }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $sourceWb = $excel.Workbooks.Open($caseSource, 0, $true)
        $targetWb = $excel.Workbooks.Open($output, 0, $true)
        $s = $sourceWb.Worksheets.Item(1)
        $t = $targetWb.Worksheets.Item('Sheet1')
        $lastSourceRow = $s.UsedRange.Row + $s.UsedRange.Rows.Count - 1
        if ($writeMode -eq 'Append' -or $case.ExpectedRows -eq 0) {
            for ($row = 2; $row -le 8; $row++) {
                if ([string]$t.Cells.Item($row, 1).Value2 -cne "OLD-$row" -or
                    [string]$t.Cells.Item($row, 2).Value2 -cne "OLD-PLATFORM-$row") {
                    throw "追加模式破坏既有目标数据: row $row ($mode)"
                }
            }
            for ($row = 9; $row -le 11; $row++) {
                if ($null -ne $t.Cells.Item($row, 1).Value2 -or $null -ne $t.Cells.Item($row, 2).Value2) {
                    throw "追加间隔行被意外写入: row $row ($mode)"
                }
            }
        }
        $dest = $case.ExpectedStart
        $checked = 0
        for ($row = 2; $row -le $lastSourceRow; $row++) {
            if ($mode -eq 'Visible' -and [bool]$s.Rows.Item($row).Hidden) { continue }
            $sku = $s.Cells.Item($row, 1).Value2
            $platformSku = $s.Cells.Item($row, 4).Value2
            if (($null -eq $sku -or [string]$sku -eq '') -and ($null -eq $platformSku -or [string]$platformSku -eq '')) { continue }
            if ([string]$sku -cne [string]$t.Cells.Item($dest, 1).Value2) { throw "SKU 不匹配: A$row -> B$dest ($mode)" }
            if ([string]$platformSku -cne [string]$t.Cells.Item($dest, 2).Value2) { throw "平台SKU 不匹配: A$row -> B$dest ($mode)" }
            $checked += 2
            $dest++
        }
        if ($dest -ne ($result.writeStartRow + $result.rowsWritten)) { throw "支线目标行数不连续: $mode/$writeMode" }
        if ($writeMode -eq 'Replace' -and $case.ExpectedRows -gt 0) {
            for ($row = $dest; $row -le 8; $row++) {
                if ($null -ne $t.Cells.Item($row, 1).Value2 -or $null -ne $t.Cells.Item($row, 2).Value2) {
                    throw "替换模式未清理旧数据: row $row ($mode)"
                }
            }
        }
        if ($t.Range('C2').Formula -ne '=1+1' -or $t.Range('C2').Value2 -ne 2) { throw "目标公式未保留或未正确计算: $mode" }
        if ($t.Range('A1').Value2 -cne 'SKU' -or $t.Range('B1').Value2 -cne '平台SKU') { throw "目标表头被修改: $mode" }
        $sourceWb.Close($false)
        $targetWb.Close($false)
        $sourceWb = $null
        $targetWb = $null
        $excel.Quit()
        $excel = $null
        Write-Output "PASS: Req02 rowMode=$mode writeMode=$writeMode start=$($result.writeStartRow); rows=$($result.rowsWritten); hidden=$($result.rowsHiddenSkipped); cells=$checked; formula-preserved=true"
    }
    Write-Output 'PASS: Req02 original-hash-unchanged=true'
} finally {
    if ($sourceWb) { $sourceWb.Close($false) }
    if ($targetWb) { $targetWb.Close($false) }
    if ($excel) { $excel.Quit() }
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
