$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
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
    $aliasSource = Join-Path $work 'alias-source.xlsx'
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
    Copy-Item -LiteralPath $source -Destination $aliasSource
    $sourceWb = $excel.Workbooks.Open($aliasSource, 0, $false)
    $sourceWb.Worksheets.Item(1).Range('D1').Value2 = '自定义SKU'
    $sourceWb.Save()
    $sourceWb.Close($false)
    $sourceWb = $null
    $excel.Quit()
    $excel = $null

    $cases = @(
        @{ Source = $source; RowMode = 'Visible'; WriteMode = 'Append'; ExpectedRows = 2; ExpectedStart = 12 },
        @{ Source = $source; RowMode = 'All'; WriteMode = 'Append'; ExpectedRows = 3; ExpectedStart = 12 },
        @{ Source = $aliasSource; RowMode = 'Visible'; WriteMode = 'Append'; ExpectedRows = 2; ExpectedStart = 12 },
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

    $existingOutput = Join-Path $work 'existing-output.xlsx'
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'excel_mapper.ps1') `
        -SourcePath $source -TargetPath $target -OutputPath $existingOutput -Profile Req02 `
        -OutputMode Replace -RowMode Visible
    if ($LASTEXITCODE -ne 0) { throw ($raw | Out-String) }
    $first = $raw | ConvertFrom-Json
    if ($first.outputMode -ne 'Replace' -or $first.writeStartRow -ne 12 -or $first.rowsWritten -ne 2) {
        throw '首次生成现有输出夹具失败。'
    }

    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'excel_mapper.ps1') `
        -SourcePath $source -TargetPath $target -OutputPath $existingOutput -Profile Req02 `
        -OutputMode AppendExisting -RowMode Visible
    if ($LASTEXITCODE -ne 0) { throw ($raw | Out-String) }
    $appended = $raw | ConvertFrom-Json
    if ($appended.outputMode -ne 'AppendExisting' -or $appended.existingLastRow -ne 13 -or
        $appended.writeStartRow -ne 17 -or $appended.rowsWritten -ne 2) {
        throw '继续写入没有按现有输出末行 + 4 追加。'
    }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $targetWb = $excel.Workbooks.Open($existingOutput, 0, $true)
    $t = $targetWb.Worksheets.Item('Sheet1')
    foreach ($pair in @(@(12, 17), @(13, 18))) {
        foreach ($column in 1..2) {
            if ([string]$t.Cells.Item($pair[0], $column).Value2 -cne [string]$t.Cells.Item($pair[1], $column).Value2) {
                throw "继续写入未保留旧批次或新批次不一致: $($pair[0]) -> $($pair[1]), column $column"
            }
        }
    }
    foreach ($row in 14..16) {
        if ($null -ne $t.Cells.Item($row, 1).Value2 -or $null -ne $t.Cells.Item($row, 2).Value2) {
            throw "继续写入未保留 3 个空白间隔行: row $row"
        }
    }
    $targetWb.Close($false)
    $targetWb = $null
    $excel.Quit()
    $excel = $null

    $blockedExisting = Join-Path $work 'blocked-existing.xlsx'
    Copy-Item -LiteralPath $existingOutput -Destination $blockedExisting
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $targetWb = $excel.Workbooks.Open($blockedExisting, 0, $false)
    $targetWb.Worksheets.Item('Sheet1').Range('A22').Formula = '=1+1'
    $targetWb.Save()
    $targetWb.Close($false)
    $targetWb = $null
    $excel.Quit()
    $excel = $null
    $blockedBefore = (Get-FileHash -LiteralPath $blockedExisting).Hash
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'excel_mapper.ps1') `
        -SourcePath $source -TargetPath $target -OutputPath $blockedExisting -Profile Req02 `
        -OutputMode AppendExisting -RowMode Visible
    if ($LASTEXITCODE -eq 0) { throw '继续写入没有阻止公式冲突。' }
    $failure = $raw | ConvertFrom-Json
    if ($failure.success -or $failure.stage -ne '检查追加位置' -or $failure.message -notmatch 'A22:A23') {
        throw '继续写入失败原因不是预期的公式冲突。'
    }
    if ((Get-FileHash -LiteralPath $blockedExisting).Hash -ne $blockedBefore) {
        throw '继续写入失败时修改了现有输出文件。'
    }

    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'excel_mapper.ps1') `
        -SourcePath $source -TargetPath $target -OutputPath $existingOutput -Profile Req02 `
        -OutputMode Replace -RowMode Visible
    if ($LASTEXITCODE -ne 0) { throw ($raw | Out-String) }
    $replaced = $raw | ConvertFrom-Json
    if ($replaced.outputMode -ne 'Replace' -or $replaced.existingLastRow -ne 8 -or
        $replaced.writeStartRow -ne 12 -or $replaced.rowsWritten -ne 2) {
        throw '替换输出没有重新以文件 B 为基底生成。'
    }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $targetWb = $excel.Workbooks.Open($existingOutput, 0, $true)
    $t = $targetWb.Worksheets.Item('Sheet1')
    if ($null -ne $t.Cells.Item(17, 1).Value2 -or $null -ne $t.Cells.Item(17, 2).Value2) {
        throw '替换输出错误保留了上一次继续写入的数据。'
    }
    $targetWb.Close($false)
    $targetWb = $null
    $excel.Quit()
    $excel = $null
    if ((Get-FileHash -LiteralPath $target).Hash -ne $before) { throw '输出模式验证修改了原始支线 B 文件。' }
    Write-Output 'PASS: existing output AppendExisting and Replace semantics.'
    Write-Output 'PASS: Req02 original-hash-unchanged=true'
} finally {
    if ($sourceWb) { $sourceWb.Close($false) }
    if ($targetWb) { $targetWb.Close($false) }
    if ($excel) { $excel.Quit() }
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
