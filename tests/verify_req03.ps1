$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$repo = Split-Path $PSScriptRoot -Parent
$mapper = Join-Path $repo 'excel_mapper.ps1'
$source = Join-Path $PSScriptRoot 'fixtures\A_req03_sample.xls'
$target = Join-Path $PSScriptRoot 'fixtures\B_req03_sample.xls'
$work = Join-Path $env:TEMP ('wsa_req03_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $work)
$excel = $null
$sourceWb = $null
$targetWb = $null
$templateWb = $null

function Invoke-Mapper([string]$Source, [string]$Target, [string]$Output, [string]$RowMode = 'Visible', [string]$OutputMode = 'Replace') {
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $Source -TargetPath $Target -OutputPath $Output -Profile Req03 `
        -RowMode $RowMode -OutputMode $OutputMode
    if ($LASTEXITCODE -ne 0) { throw ($raw | Out-String) }
    return $raw | ConvertFrom-Json
}

function Assert-Workbook(
    [string]$Output,
    [string]$RowMode,
    [int]$ExpectedRows,
    [int]$ExpectedStart,
    [string]$SourcePath = $source,
    [string]$TemplatePath = $target,
    [int[]]$SourceColumns = @(4, 27, 31, 51, 52, 53),
    [int[]]$TargetColumns = @(1, 2, 3, 8, 9, 10)
) {
    $script:excel = New-Object -ComObject Excel.Application
    $script:excel.Visible = $false
    $script:excel.DisplayAlerts = $false
    $script:sourceWb = $script:excel.Workbooks.Open($SourcePath, 0, $true)
    $script:templateWb = $script:excel.Workbooks.Open($TemplatePath, 0, $true)
    $script:targetWb = $script:excel.Workbooks.Open($Output, 0, $true)
    try {
        $s = $script:sourceWb.Worksheets.Item(1)
        $templateSheet = $script:templateWb.Worksheets.Item('导入 单位转换')
        $outputSheet = $script:targetWb.Worksheets.Item('导入 单位转换')
        $dest = $ExpectedStart
        for ($row = 2; $row -le 5; $row++) {
            if ($RowMode -eq 'Visible' -and [bool]$s.Rows.Item($row).Hidden) { continue }
            $values = @($SourceColumns | ForEach-Object { $s.Cells.Item($row, $_).Value2 })
            if (-not ($values | Where-Object { $null -ne $_ -and [string]$_ -ne '' })) { continue }
            for ($index = 0; $index -lt $SourceColumns.Count; $index++) {
                $actual = $outputSheet.Cells.Item($dest, $TargetColumns[$index]).Value2
                if ($index -eq 0 -or $values[$index] -eq 'not-a-number' -or [string]$values[$index] -eq '') {
                    if ([string]$actual -cne [string]$values[$index]) {
                        throw "Req03 文本值不匹配: source R$row C$($SourceColumns[$index]) -> target R$dest C$($TargetColumns[$index])"
                    }
                } else {
                    if ($actual -isnot [double] -or [double]$actual -ne [double]$values[$index]) {
                        throw "Req03 数值类型或值不匹配: source R$row C$($SourceColumns[$index]) -> target R$dest C$($TargetColumns[$index])"
                    }
                }
            }
            $dest++
        }
        if ($dest -ne ($ExpectedStart + $ExpectedRows)) { throw 'Req03 写入行数不连续。' }

        foreach ($letter in @('D', 'E', 'K', 'L', 'M', 'N', 'R')) {
            for ($row = 2; $row -le 20; $row++) {
                if ($outputSheet.Range("$letter$row").Formula -cne $templateSheet.Range("$letter$row").Formula) {
                    throw "Req03 修改了公式: $letter$row"
                }
            }
        }
        for ($row = 2; $row -le 20; $row++) {
            if ([string]$outputSheet.Range("Q$row").Value2 -cne [string]$templateSheet.Range("Q$row").Value2) {
                throw "Req03 修改了非目标常量: Q$row"
            }
        }
        foreach ($column in $TargetColumns) {
            for ($row = 1; $row -le 20; $row++) {
                $left = $templateSheet.Cells.Item($row, $column)
                $right = $outputSheet.Cells.Item($row, $column)
                if ([string]$left.NumberFormat -cne [string]$right.NumberFormat -or
                    [string]$left.HorizontalAlignment -cne [string]$right.HorizontalAlignment -or
                    [string]$left.Interior.Color -cne [string]$right.Interior.Color -or
                    [string]$left.Font.Name -cne [string]$right.Font.Name -or
                    [string]$left.Font.Size -cne [string]$right.Font.Size -or
                    [string]$left.Font.Bold -cne [string]$right.Font.Bold) {
                    throw "Req03 修改了目标单元格格式: R$row C$column"
                }
            }
        }
        $price = $script:targetWb.Worksheets.Item('价格')
        if ([string]$price.Range('A1').Value2 -cne 'SKU(直接从sheet 1导入）' -or $null -ne $price.Range('A2').Value2) {
            throw 'Req03 错误写入价格工作表。'
        }
        $shipping = $script:targetWb.Worksheets.Item('运费表（公式数据，不动）')
        if ([string]$shipping.Range('A1').Value2 -cne 'KEEP-SHIPPING' -or $shipping.Range('A2').Formula -cne '=1+1') {
            throw 'Req03 修改了运费工作表。'
        }
    } finally {
        $script:sourceWb.Close($false)
        $script:templateWb.Close($false)
        $script:targetWb.Close($false)
        $script:sourceWb = $null
        $script:templateWb = $null
        $script:targetWb = $null
        $script:excel.Quit()
        $script:excel = $null
    }
}

try {
    if (-not (Test-Path -LiteralPath $source) -or -not (Test-Path -LiteralPath $target)) {
        throw '缺少 Req03 合成夹具；请先运行 tests/make_req03_fixture.ps1。'
    }
    $before = (Get-FileHash -Algorithm MD5 -LiteralPath $target).Hash
    foreach ($case in @(
        @{ Mode = 'Visible'; Rows = 2; Hidden = 1 },
        @{ Mode = 'All'; Rows = 3; Hidden = 0 }
    )) {
        $output = Join-Path $work ("output-" + $case.Mode + '.xls')
        $result = Invoke-Mapper $source $target $output $case.Mode
        if (-not $result.success -or $result.profile -ne 'Req03' -or $result.targetSheet -cne '导入 单位转换') {
            throw "Req03 映射摘要错误: $($case.Mode)"
        }
        if ($result.rowsRead -ne $case.Rows -or $result.rowsWritten -ne $case.Rows -or
            $result.rowsHiddenSkipped -ne $case.Hidden -or $result.existingLastRow -ne 0 -or
            $result.writeStartRow -ne 2 -or $result.mappings.Count -ne 6) {
            throw "Req03 行数、起点或映射数量错误: $($case.Mode)"
        }
        if (@($result.mappings | Where-Object { $_.sourceMethod -ne 'header' -or $_.targetMethod -ne 'header' }).Count -ne 0) {
            throw "Req03 未优先使用表头: $($case.Mode)"
        }
        Assert-Workbook $output $case.Mode $case.Rows 2
        if ((Get-FileHash -Algorithm MD5 -LiteralPath $target).Hash -ne $before) { throw 'Req03 修改了原始 B 夹具。' }
        $signature = [IO.File]::ReadAllBytes($output)[0..7]
        if (($signature | ForEach-Object { $_.ToString('X2') }) -join ' ' -ne 'D0 CF 11 E0 A1 B1 1A E1') {
            throw 'Req03 输出不是有效的二进制 .xls。'
        }
        Write-Output "PASS: Req03 rowMode=$($case.Mode) rows=$($case.Rows) start=2 mappings=6 format-and-formulas-preserved=true"
    }

    $existing = Join-Path $work 'append-existing.xls'
    $first = Invoke-Mapper $source $target $existing Visible
    $second = Invoke-Mapper $source $target $existing Visible AppendExisting
    if ($first.writeStartRow -ne 2 -or $second.existingLastRow -ne 3 -or $second.writeStartRow -ne 7) {
        throw 'Req03 继续写入未按末行 + 4 追加。'
    }
    Assert-Workbook $existing Visible 2 7
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $targetWb = $excel.Workbooks.Open($existing, 0, $true)
    $sheet = $targetWb.Worksheets.Item('导入 单位转换')
    foreach ($row in 4..6) {
        foreach ($column in @(1, 2, 3, 8, 9, 10)) {
            if ($null -ne $sheet.Cells.Item($row, $column).Value2) { throw "Req03 三行间隔被写入: R$row C$column" }
        }
    }
    $targetWb.Close($false)
    $targetWb = $null
    $excel.Quit()
    $excel = $null
    Write-Output 'PASS: Req03 AppendExisting start=7 and rows 4-6 remain blank.'

    $fallbackSource = Join-Path $work 'fallback-source.xls'
    $fallbackTarget = Join-Path $work 'fallback-target.xls'
    Copy-Item -LiteralPath $source -Destination $fallbackSource
    Copy-Item -LiteralPath $target -Destination $fallbackTarget
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $sourceWb = $excel.Workbooks.Open($fallbackSource, 0, $false)
    foreach ($column in @(4, 27, 31, 51, 52, 53)) { $sourceWb.Worksheets.Item(1).Cells.Item(1, $column).Value2 = 'unknown' }
    $sourceWb.Save()
    $sourceWb.Close($false)
    $sourceWb = $null
    $targetWb = $excel.Workbooks.Open($fallbackTarget, 0, $false)
    foreach ($column in @(1, 2, 3, 8, 9, 10)) { $targetWb.Worksheets.Item('导入 单位转换').Cells.Item(1, $column).Value2 = 'unknown' }
    $targetWb.Save()
    $targetWb.Close($false)
    $targetWb = $null
    $excel.Quit()
    $excel = $null
    $fallbackOutput = Join-Path $work 'fallback-output.xls'
    $fallback = Invoke-Mapper $fallbackSource $fallbackTarget $fallbackOutput Visible
    if (@($fallback.mappings | Where-Object { $_.sourceMethod -ne 'fallback-column' -or $_.targetMethod -ne 'fallback-column' }).Count -ne 0) {
        throw 'Req03 表头缺失时未完整走列字母兜底。'
    }
    Assert-Workbook $fallbackOutput Visible 2 2 $fallbackSource $fallbackTarget
    Write-Output 'PASS: Req03 source and target column-letter fallback with cell-by-cell values.'

    $movedSource = Join-Path $work 'moved-source.xls'
    $movedTarget = Join-Path $work 'moved-target.xls'
    Copy-Item -LiteralPath $source -Destination $movedSource
    Copy-Item -LiteralPath $target -Destination $movedTarget
    $originalSourceLetters = @('D', 'AA', 'AE', 'AY', 'AZ', 'BA')
    $movedSourceLetters = @('E', 'AB', 'AF', 'BB', 'BC', 'BD')
    $movedSourceColumns = @(5, 28, 32, 54, 55, 56)
    $movedTargetColumns = @(21, 22, 23, 24, 25, 26)
    $targetHeaders = @('SKU(直接从sheet 1导入）', 'SKU价(￥)', '重量(g)', '长', '宽', '高')
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $sourceWb = $excel.Workbooks.Open($movedSource, 0, $false)
    $sourceSheet = $sourceWb.Worksheets.Item(1)
    for ($index = 0; $index -lt $movedSourceLetters.Count; $index++) {
        $fromRange = $sourceSheet.Range("$($originalSourceLetters[$index])1:$($originalSourceLetters[$index])5")
        $toRange = $sourceSheet.Range("$($movedSourceLetters[$index])1:$($movedSourceLetters[$index])5")
        try { $toRange.Value2 = $fromRange.Value2 }
        finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($fromRange)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($toRange)
        }
        $sourceSheet.Range("$($originalSourceLetters[$index])1:$($originalSourceLetters[$index])5").ClearContents()
    }
    $sourceWb.Save()
    $sourceWb.Close($false)
    $sourceWb = $null
    $targetWb = $excel.Workbooks.Open($movedTarget, 0, $false)
    $targetSheet = $targetWb.Worksheets.Item('导入 单位转换')
    foreach ($letter in @('A', 'B', 'C', 'H', 'I', 'J')) { $targetSheet.Range("${letter}1").ClearContents() }
    for ($index = 0; $index -lt $movedTargetColumns.Count; $index++) {
        $targetSheet.Cells.Item(1, $movedTargetColumns[$index]).Value2 = $targetHeaders[$index]
    }
    $targetWb.Save()
    $targetWb.Close($false)
    $targetWb = $null
    $excel.Quit()
    $excel = $null
    $movedOutput = Join-Path $work 'moved-output.xls'
    $moved = Invoke-Mapper $movedSource $movedTarget $movedOutput Visible
    if (@($moved.mappings | Where-Object { $_.sourceMethod -ne 'header' -or $_.targetMethod -ne 'header' }).Count -ne 0) {
        throw 'Req03 非兜底列的表头未优先命中。'
    }
    Assert-Workbook $movedOutput Visible 2 2 $movedSource $movedTarget $movedSourceColumns $movedTargetColumns
    Write-Output 'PASS: Req03 header-first mapping uses moved non-fallback columns with cell-by-cell values.'

    $ambiguousTarget = Join-Path $work 'ambiguous-target.xls'
    Copy-Item -LiteralPath $target -Destination $ambiguousTarget
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $targetWb = $excel.Workbooks.Open($ambiguousTarget, 0, $false)
    $targetWb.Worksheets.Item('导入 单位转换').Range('G1').Value2 = '重量(g)'
    $targetWb.Save()
    $targetWb.Close($false)
    $targetWb = $null
    $excel.Quit()
    $excel = $null
    $ambiguousHash = (Get-FileHash -Algorithm MD5 -LiteralPath $ambiguousTarget).Hash
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $source -TargetPath $ambiguousTarget -OutputPath (Join-Path $work 'ambiguous-output.xls') `
        -Profile Req03
    $failure = $raw | ConvertFrom-Json
    if ($LASTEXITCODE -eq 0 -or $failure.success -or $failure.message -notmatch '多个候选列') {
        throw 'Req03 未阻止歧义目标表头。'
    }
    if ((Get-FileHash -Algorithm MD5 -LiteralPath $ambiguousTarget).Hash -ne $ambiguousHash) {
        throw 'Req03 歧义失败时修改了原始目标。'
    }
    Write-Output 'PASS: Req03 rejects ambiguous target headers.'

    $protectedTarget = Join-Path $work 'protected-target.xls'
    Copy-Item -LiteralPath $target -Destination $protectedTarget
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $targetWb = $excel.Workbooks.Open($protectedTarget, 0, $false)
    $targetWb.Worksheets.Item('导入 单位转换').Range('A2:A3').Merge()
    $targetWb.Save()
    $targetWb.Close($false)
    $targetWb = $null
    $excel.Quit()
    $excel = $null
    $protectedHash = (Get-FileHash -Algorithm MD5 -LiteralPath $protectedTarget).Hash
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $source -TargetPath $protectedTarget -OutputPath (Join-Path $work 'protected-output.xls') `
        -Profile Req03
    $failure = $raw | ConvertFrom-Json
    if ($LASTEXITCODE -eq 0 -or $failure.success -or $failure.message -notmatch '包含公式或合并单元格') {
        throw 'Req03 未阻止写入区域合并单元格冲突。'
    }
    if ((Get-FileHash -Algorithm MD5 -LiteralPath $protectedTarget).Hash -ne $protectedHash) {
        throw 'Req03 合并单元格冲突失败时修改了原始目标。'
    }
    Write-Output 'PASS: Req03 rejects merged cells in the write region.'

    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $source -TargetPath $target -OutputPath (Join-Path $work 'replace.xls') `
        -Profile Req03 -WriteMode Replace
    if ($LASTEXITCODE -eq 0 -or ($raw | ConvertFrom-Json).message -notmatch '仅适用于支线 Req02') {
        throw 'Req03 未拒绝 WriteMode Replace。'
    }
    Write-Output 'PASS: Req03 rejects WriteMode Replace.'

    $badOutput = Join-Path $work 'wrong-extension.xlsx'
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $source -TargetPath $target -OutputPath $badOutput -Profile Req03
    $failure = $raw | ConvertFrom-Json
    if ($LASTEXITCODE -eq 0 -or $failure.success -or $failure.message -notmatch '\.xls' -or
        (Test-Path -LiteralPath $badOutput)) {
        throw 'Req03 未拒绝非 .xls 输出路径。'
    }
    Write-Output 'PASS: Req03 rejects non-.xls output before creating a file.'
    Write-Output 'PASS: Req03 original-hash-unchanged=true'
} finally {
    if ($sourceWb) { $sourceWb.Close($false) }
    if ($targetWb) { $targetWb.Close($false) }
    if ($templateWb) { $templateWb.Close($false) }
    if ($excel) { $excel.Quit() }
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
