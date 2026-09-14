param(
    [string]$SourcePath = (Join-Path $PSScriptRoot '..\文件\04\A模板.xls'),
    [string]$TargetPath = (Join-Path $PSScriptRoot '..\文件\04\B模板.xlsx')
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$mapper = Join-Path $root 'excel_mapper.ps1'
$ooxmlComparer = Join-Path $root 'tests\compare_ooxml.py'
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('req04-verify-' + [guid]::NewGuid().ToString('N'))
$firstOutput = Join-Path $tempRoot 'first.xlsx'
$secondOutput = Join-Path $tempRoot 'second.xlsx'
$edgeSource = Join-Path $tempRoot 'edge-source.xls'
$duplicateSource = Join-Path $tempRoot 'duplicate-source.xls'
$unmatchedSource = Join-Path $tempRoot 'unmatched-source.xls'
$missingSkuSource = Join-Path $tempRoot 'missing-sku-source.xls'
$reorderedSource = Join-Path $tempRoot 'reordered-source.xls'
$conflictTarget = Join-Path $tempRoot 'conflict-target.xlsx'
$edgeOutput = Join-Path $tempRoot 'edge-output.xlsx'
$reorderedOutput = Join-Path $tempRoot 'reordered-output.xlsx'
$excel = $null
$sourceWb = $null
$firstWb = $null
$secondWb = $null
$templateWb = $null
$editor = $null
$book = $null
$edgeWb = $null
$reorderedSourceWb = $null
$reorderedOutputWb = $null
$testStage = '初始化'

function Normalize([object]$value) {
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim().ToLowerInvariant()
}

function Assert-True([bool]$condition, [string]$message) {
    if (-not $condition) { throw $message }
}

function Assert-Number([object]$actual, [object]$expected, [string]$label) {
    Assert-True ($null -ne $actual -and $null -ne $expected) "$label 缺少值。"
    $delta = [math]::Abs(([double]$actual) - ([double]$expected))
    Assert-True ($delta -lt 0.0000001) "$label 不一致：实际 $actual，期望 $expected。"
}

function Set-Value($worksheet, [string]$address, [object]$value) {
    $cell = $worksheet.Range($address)
    try { $cell.Value2 = $value }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
}

function Get-Value($worksheet, [string]$address) {
    $cell = $worksheet.Range($address)
    try { return $cell.Value2 }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
}

function Get-FormatSignature($worksheet, [string]$address) {
    $cell = $worksheet.Range($address)
    $font = $null
    $interior = $null
    try {
        $font = $cell.Font
        $interior = $cell.Interior
        return (@(
            [string]$cell.NumberFormat,
            [string]$font.Name,
            [string]$font.Size,
            [string]$font.Bold,
            [string]$font.Italic,
            [string]$font.Color,
            [string]$interior.Color,
            [string]$cell.VerticalAlignment,
            [string]$cell.WrapText
        ) -join '|')
    } finally {
        foreach ($item in @($font, $interior, $cell)) {
            if ($item) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
        }
    }
}

function Clear-Value($worksheet, [string]$address) {
    $cell = $worksheet.Range($address)
    try { $cell.ClearContents() }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
}

function Set-RowHidden($worksheet, [int]$row, [bool]$hidden) {
    $rowRange = $worksheet.Rows.Item($row)
    try { $rowRange.Hidden = $hidden }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($rowRange) }
}

try {
    $testStage = '准备目录与基线哈希'
    New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
    $originalSourceHash = (Get-FileHash -Algorithm MD5 -LiteralPath $SourcePath).Hash
    $originalHash = (Get-FileHash -Algorithm MD5 -LiteralPath $TargetPath).Hash

    $testStage = '首次与追加映射'
    $firstJson = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $SourcePath -TargetPath $TargetPath -OutputPath $firstOutput `
        -Profile Req04 -RowMode Visible -OutputMode Replace | Select-Object -Last 1
    Assert-True ($LASTEXITCODE -eq 0) "Req04 首次运行失败：$firstJson"
    $first = $firstJson | ConvertFrom-Json
    Assert-True ($first.success -eq $true) '首次运行未返回 success=true。'
    Assert-True ($first.rowsRead -eq 306 -and $first.rowsWritten -eq 306) "首次行数不正确：$($first.rowsRead)/$($first.rowsWritten)。"
    Assert-True ($first.writeStartRow -eq 13) "首次写入起点不正确：$($first.writeStartRow)。"
    Assert-True ($first.rowsHiddenSkipped -eq 0) "真实模板隐藏行计数不正确：$($first.rowsHiddenSkipped)。"
    Assert-True ($first.warnings.Count -eq 0) '真实模板不应产生数值转换警告。'
    Assert-True (($first.mappings.targetColumn -join ',') -eq 'K,AN,AP,AT,J,CR,AR') "目标映射列不正确：$($first.mappings.targetColumn -join ',')。"

    Copy-Item -LiteralPath $firstOutput -Destination $secondOutput -Force
    $secondJson = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $SourcePath -TargetPath $TargetPath -OutputPath $secondOutput `
        -Profile Req04 -RowMode Visible -OutputMode AppendExisting | Select-Object -Last 1
    Assert-True ($LASTEXITCODE -eq 0) "Req04 追加运行失败：$secondJson"
    $second = $secondJson | ConvertFrom-Json
    Assert-True ($second.success -eq $true) '追加运行未返回 success=true。'
    Assert-True ($second.existingLastRow -eq 318 -and $second.writeStartRow -eq 322) "追加起点不正确：末行 $($second.existingLastRow)，起点 $($second.writeStartRow)。"
    Assert-True ($second.rowsWritten -eq 306) "追加写入行数不正确：$($second.rowsWritten)。"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $sourceWb = $excel.Workbooks.Open((Resolve-Path $SourcePath).Path, 0, $true)
    $firstWb = $excel.Workbooks.Open($firstOutput, 0, $true)
    $secondWb = $excel.Workbooks.Open($secondOutput, 0, $true)
    $templateWb = $excel.Workbooks.Open((Resolve-Path $TargetPath).Path, 0, $true)
    $units = $sourceWb.Worksheets.Item('导入 单位转换')
    $prices = $sourceWb.Worksheets.Item('价格')
    $firstWs = $firstWb.Worksheets.Item('Product Content And Site Exp')
    $secondWs = $secondWb.Worksheets.Item('Product Content And Site Exp')
    $templateWs = $templateWb.Worksheets.Item('Product Content And Site Exp')
    for ($row = 7; $row -le 9; $row++) {
        Assert-True ([string]$firstWs.Cells.Item($row, 4).Value2 -eq [string]$templateWs.Cells.Item($row, 4).Value2) "B 自带数据行 $row 被覆盖。"
    }

    $expected = @{}
    for ($row = 2; $row -le 307; $row++) {
        $sku = Normalize $units.Cells.Item($row, 1).Value2
        if (-not $sku) { continue }
        $expected[$sku] = [pscustomobject]@{
            Weight = $units.Cells.Item($row, 5).Value2
            Length = $units.Cells.Item($row, 11).Value2
            Width = $units.Cells.Item($row, 12).Value2
            Height = $units.Cells.Item($row, 13).Value2
            SellingPrice = $prices.Cells.Item($row, 9).Value2
            MSRP = $prices.Cells.Item($row, 17).Value2
        }
    }
    Assert-True ($expected.Count -eq 306) "源 SKU 数量不正确：$($expected.Count)。"

    for ($index = 0; $index -lt 306; $index++) {
        $sourceRow = $index + 2
        $targetRow = $index + 13
        $sku = Normalize $units.Cells.Item($sourceRow, 1).Value2
        $item = $expected[$sku]
        Assert-Number $firstWs.Cells.Item($targetRow, 11).Value2 $item.Weight "首次 K$targetRow"
        Assert-Number $firstWs.Cells.Item($targetRow, 40).Value2 $item.Length "首次 AN$targetRow"
        Assert-Number $firstWs.Cells.Item($targetRow, 42).Value2 $item.Height "首次 AP$targetRow"
        Assert-Number $firstWs.Cells.Item($targetRow, 44).Value2 $item.Weight "首次 AR$targetRow"
        Assert-Number $firstWs.Cells.Item($targetRow, 46).Value2 $item.Width "首次 AT$targetRow"
        Assert-Number $firstWs.Cells.Item($targetRow, 10).Value2 $item.SellingPrice "首次 J$targetRow"
        Assert-Number $firstWs.Cells.Item($targetRow, 96).Value2 $item.MSRP "首次 CR$targetRow"
        Assert-True ([string]$firstWs.Cells.Item($targetRow, 4).Value2 -eq [string]$templateWs.Cells.Item($targetRow, 4).Value2) "首次 D$targetRow 被修改，Req04 不应写入 SKU。"
        Assert-Number $secondWs.Cells.Item($index + 322, 11).Value2 $item.Weight "追加 K$($index + 322)"
    }
    foreach ($row in @(13, 318)) {
        foreach ($column in @('J', 'K', 'AN', 'AP', 'AR', 'AT', 'CR')) {
            Assert-True ((Get-FormatSignature $firstWs "${column}${row}") -eq (Get-FormatSignature $templateWs "${column}${row}")) "首次 ${column}${row} 的非对齐格式发生变化。"
            $cell = $firstWs.Range("${column}${row}")
            try { Assert-True ([int]$cell.HorizontalAlignment -eq 5) "首次 ${column}${row} 未使用填充对齐。" }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
        }
    }
    foreach ($address in @('D13', 'BS13', 'K4')) {
        Assert-True ((Get-FormatSignature $firstWs $address) -eq (Get-FormatSignature $templateWs $address)) "未写入单元格 $address 的格式发生变化。"
    }
    for ($row = 319; $row -le 321; $row++) {
        foreach ($column in @(10, 11, 40, 42, 44, 46, 96)) {
            Assert-True ([string]$secondWs.Cells.Item($row, $column).Value2 -eq '') "追加间隔行 $row/$column 不为空。"
        }
    }
    $sourceWb.Close($false)
    $firstWb.Close($false)
    $secondWb.Close($false)
    $templateWb.Close($false)
    $excel.Quit()
    foreach ($item in @($units, $prices, $firstWs, $secondWs, $templateWs, $sourceWb, $firstWb, $secondWb, $templateWb, $excel)) {
        if ($item) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
    }
    $sourceWb = $null
    $firstWb = $null
    $secondWb = $null
    $templateWb = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $comparison = & py $ooxmlComparer $TargetPath $firstOutput 2>&1
    Assert-True ($LASTEXITCODE -eq 0) "Req04 输出的公式/条件格式/数据验证未保留：$($comparison -join [Environment]::NewLine)"

    $testStage = '制作边界夹具'
    foreach ($path in @($edgeSource, $duplicateSource, $unmatchedSource, $missingSkuSource, $reorderedSource)) {
        Copy-Item -LiteralPath $SourcePath -Destination $path -Force
    }
    Copy-Item -LiteralPath $TargetPath -Destination $conflictTarget -Force

    $editor = New-Object -ComObject Excel.Application
    $editor.Visible = $false
    $editor.DisplayAlerts = $false
    foreach ($case in @(
        @{ Path = $edgeSource; Kind = 'Edge' },
        @{ Path = $duplicateSource; Kind = 'Duplicate' },
        @{ Path = $unmatchedSource; Kind = 'Unmatched' },
        @{ Path = $missingSkuSource; Kind = 'Missing' },
        @{ Path = $reorderedSource; Kind = 'Reordered' }
    )) {
        $book = $editor.Workbooks.Open($case.Path, 0, $false)
        $unitSheet = $book.Worksheets.Item('导入 单位转换')
        $priceSheet = $book.Worksheets.Item('价格')
        if ($case.Kind -eq 'Edge') {
            Set-RowHidden $unitSheet 2 $true
            Set-RowHidden $priceSheet 2 $true
            Clear-Value $unitSheet 'A2'
            Set-Value $unitSheet 'E2' '1'
            Set-Value $priceSheet 'A2' 'REQ04-HIDDEN-PRICE-SKU'
            Set-Value $unitSheet 'E3' 'bad-number'
            Clear-Value $unitSheet 'L3'
            Clear-Value $priceSheet 'I3'
            Clear-Value $priceSheet 'Q3'
        } elseif ($case.Kind -eq 'Duplicate') {
            Set-Value $unitSheet 'A3' (Get-Value $unitSheet 'A2')
        } elseif ($case.Kind -eq 'Unmatched') {
            Set-Value $priceSheet 'A2' 'REQ04-UNMATCHED-SKU'
        } elseif ($case.Kind -eq 'Missing') {
            Clear-Value $unitSheet 'A2'
            Set-Value $unitSheet 'E2' '1'
        } else {
            $sku2 = Get-Value $priceSheet 'A2'
            $sku3 = Get-Value $priceSheet 'A3'
            Set-Value $priceSheet 'A2' $sku3
            Set-Value $priceSheet 'A3' $sku2
        }
        $book.Save()
        $book.Close($false)
        foreach ($item in @($unitSheet, $priceSheet, $book)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)
        }
    }
    $book = $editor.Workbooks.Open($conflictTarget, 0, $false)
    $conflictSheet = $book.Worksheets.Item('Product Content And Site Exp')
    $mergeRange = $conflictSheet.Range('K13:K14')
    try { $mergeRange.Merge() }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($mergeRange) }
    $book.Save()
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($conflictSheet)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    $editor.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($editor)
    $editor = $null
    $book = $null

    $testStage = '边界值与隐藏行映射'
    $edgeJson = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $edgeSource -TargetPath $TargetPath -OutputPath $edgeOutput `
        -Profile Req04 -RowMode Visible -OutputMode Replace | Select-Object -Last 1
    Assert-True ($LASTEXITCODE -eq 0) "Req04 边界运行失败：$edgeJson"
    $edge = $edgeJson | ConvertFrom-Json
    Assert-True ($edge.rowsRead -eq 305 -and $edge.rowsWritten -eq 305) "边界行数不正确：$($edge.rowsRead)/$($edge.rowsWritten)。"
    Assert-True ($edge.rowsHiddenSkipped -eq 2) "隐藏行计数应为 2，实际为 $($edge.rowsHiddenSkipped)。"
    Assert-True ($edge.warnings.Count -eq 2) "重量双写应产生 2 条文本回退警告，实际为 $($edge.warnings.Count)。"
    foreach ($warning in $edge.warnings) {
        Assert-True ($warning -like '*SKU *' -and $warning -like '*重量(lb)两位小数(E)*' -and $warning -like '*bad-number*') "警告缺少 SKU、源字段或原值：$warning"
        Assert-True ($warning -like '*目标 K.*' -or $warning -like '*目标 AR.*') "警告缺少目标字段：$warning"
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $edgeWb = $excel.Workbooks.Open($edgeOutput, 0, $true)
    $edgeWs = $edgeWb.Worksheets.Item('Product Content And Site Exp')
    Assert-True ([string](Get-Value $edgeWs 'K13') -eq 'bad-number') '数值转换失败后 K13 未按文本写入。'
    Assert-True ([string](Get-Value $edgeWs 'AR13') -eq 'bad-number') '数值转换失败后 AR13 未按文本写入。'
    foreach ($address in @('AT13', 'J13', 'CR13')) {
        Assert-True ([string](Get-Value $edgeWs $address) -eq '') "$address 应保持空值。"
    }
    Assert-True ([string](Get-Value $edgeWs 'AN13') -ne '' -and [string](Get-Value $edgeWs 'AP13') -ne '') '同一 SKU 的其他字段不应受空字段影响。'
    $edgeWb.Close($false)
    $excel.Quit()
    foreach ($item in @($edgeWs, $edgeWb, $excel)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)
    }
    $edgeWb = $null
    $excel = $null

    $reorderedJson = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $reorderedSource -TargetPath $TargetPath -OutputPath $reorderedOutput `
        -Profile Req04 -RowMode Visible -OutputMode Replace | Select-Object -Last 1
    Assert-True ($LASTEXITCODE -eq 0) "价格表乱序映射失败：$reorderedJson"
    $reordered = $reorderedJson | ConvertFrom-Json
    Assert-True ($reordered.rowsWritten -eq 306) "价格表乱序后的行数不正确：$($reordered.rowsWritten)。"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $reorderedSourceWb = $excel.Workbooks.Open($reorderedSource, 0, $true)
    $reorderedOutputWb = $excel.Workbooks.Open($reorderedOutput, 0, $true)
    $reorderedPrices = $reorderedSourceWb.Worksheets.Item('价格')
    $reorderedWs = $reorderedOutputWb.Worksheets.Item('Product Content And Site Exp')
    Assert-Number (Get-Value $reorderedWs 'J13') (Get-Value $reorderedPrices 'I3') '价格表乱序 J13'
    Assert-Number (Get-Value $reorderedWs 'CR13') (Get-Value $reorderedPrices 'Q3') '价格表乱序 CR13'
    Assert-Number (Get-Value $reorderedWs 'J14') (Get-Value $reorderedPrices 'I2') '价格表乱序 J14'
    Assert-Number (Get-Value $reorderedWs 'CR14') (Get-Value $reorderedPrices 'Q2') '价格表乱序 CR14'
    $reorderedSourceWb.Close($false)
    $reorderedOutputWb.Close($false)
    $excel.Quit()
    foreach ($item in @($reorderedPrices, $reorderedWs, $reorderedSourceWb, $reorderedOutputWb, $excel)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item)
    }
    $reorderedSourceWb = $null
    $reorderedOutputWb = $null
    $excel = $null

    $testStage = '拒绝型边界验证'
    foreach ($negative in @(
        @{ Label = '重复 SKU'; Source = $duplicateSource; Target = $TargetPath; Output = (Join-Path $tempRoot 'duplicate.xlsx'); RowMode = 'Visible'; Pattern = '重复 SKU' },
        @{ Label = '无法匹配 SKU'; Source = $unmatchedSource; Target = $TargetPath; Output = (Join-Path $tempRoot 'unmatched.xlsx'); RowMode = 'Visible'; Pattern = '未找到匹配项' },
        @{ Label = '缺失 SKU'; Source = $missingSkuSource; Target = $TargetPath; Output = (Join-Path $tempRoot 'missing.xlsx'); RowMode = 'Visible'; Pattern = 'SKU 为空' },
        @{ Label = '合并区域冲突'; Source = $SourcePath; Target = $conflictTarget; Output = (Join-Path $tempRoot 'conflict.xlsx'); RowMode = 'Visible'; Pattern = '包含公式或合并单元格' },
        @{ Label = 'All 模式'; Source = $SourcePath; Target = $TargetPath; Output = (Join-Path $tempRoot 'all.xlsx'); RowMode = 'All'; Pattern = '只支持“仅可见行”' }
    )) {
        $failureJson = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
            -SourcePath $negative.Source -TargetPath $negative.Target -OutputPath $negative.Output `
            -Profile Req04 -RowMode $negative.RowMode -OutputMode Replace | Select-Object -Last 1
        $failureCode = $LASTEXITCODE
        $failure = $failureJson | ConvertFrom-Json
        Assert-True ($failureCode -ne 0 -and $failure.success -eq $false) "$($negative.Label) 未返回失败。"
        Assert-True ($failure.message -like "*$($negative.Pattern)*") "$($negative.Label) 错误信息不正确：$($failure.message)"
        Assert-True (-not (Test-Path -LiteralPath $negative.Output)) "$($negative.Label) 失败后不应生成输出文件。"
    }

    $badExtension = Join-Path $tempRoot 'bad-extension.xls'
    $extensionJson = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mapper `
        -SourcePath $SourcePath -TargetPath $TargetPath -OutputPath $badExtension `
        -Profile Req04 -RowMode Visible -OutputMode Replace | Select-Object -Last 1
    $extensionCode = $LASTEXITCODE
    $extensionFailure = $extensionJson | ConvertFrom-Json
    Assert-True ($extensionCode -ne 0 -and $extensionFailure.message -like '*输出扩展名必须跟随文件 B*') '不匹配的输出扩展名未被拒绝。'
    Assert-True (-not (Test-Path -LiteralPath $badExtension)) '扩展名失败后不应生成输出文件。'

    $testStage = '最终哈希核对'
    $afterSourceHash = (Get-FileHash -Algorithm MD5 -LiteralPath $SourcePath).Hash
    $afterHash = (Get-FileHash -Algorithm MD5 -LiteralPath $TargetPath).Hash
    Assert-True ($originalSourceHash -eq $afterSourceHash) '原始 A 文件哈希发生变化。'
    Assert-True ($originalHash -eq $afterHash) '原始 B 模板哈希发生变化。'
    [pscustomobject]@{ success = $true; rows = 306; firstStart = $first.writeStartRow; appendStart = $second.writeStartRow; hiddenEdge = $edge.rowsHiddenSkipped; warningsEdge = $edge.warnings.Count; reorderedJoin = $true; negativeCases = 6; outputExtension = [IO.Path]::GetExtension($first.output) } | ConvertTo-Json -Compress
    exit 0
} catch {
    foreach ($openBook in @($sourceWb, $firstWb, $secondWb, $templateWb, $book, $edgeWb, $reorderedSourceWb, $reorderedOutputWb)) {
        if ($openBook) { try { $openBook.Close($false) } catch {} }
    }
    if ($editor) { try { $editor.Quit() } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    Write-Error ("{0}：{1}" -f $testStage, $_.Exception.Message)
    exit 1
} finally {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($tempRoot -and (Test-Path -LiteralPath $tempRoot)) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
