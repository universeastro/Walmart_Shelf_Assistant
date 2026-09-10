param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Normalize-Text([object]$Value) {
    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    $text = $text.Replace([char]0xA0, ' ')
    $text = $text -replace '\s+', ' '
    return $text.Trim().ToLowerInvariant()
}

function Get-ColumnLetter([int]$Number) {
    $result = ''
    while ($Number -gt 0) {
        $Number--
        $result = [char](65 + ($Number % 26)) + $result
        $Number = [math]::Floor($Number / 26)
    }
    return $result
}

function Get-CellText($Worksheet, [int]$Row, [int]$Column) {
    $cell = $Worksheet.Cells.Item($Row, $Column)
    try {
        if ($cell.MergeCells) {
            return [string]$cell.MergeArea.Cells.Item(1, 1).Text
        }
        return [string]$cell.Text
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
    }
}

function Get-CellValue($Worksheet, [int]$Row, [int]$Column) {
    return $Worksheet.Cells.Item($Row, $Column).Value2
}

function Find-SourceColumn($Worksheet, [string[]]$HeaderNames, [string]$FallbackLetter) {
    $used = $Worksheet.UsedRange
    $headerRow = $used.Row
    $maxColumn = $used.Column + $used.Columns.Count - 1
    $wanted = @($HeaderNames | ForEach-Object { Normalize-Text $_ } | Where-Object { $_ })
    for ($column = $used.Column; $column -le $maxColumn; $column++) {
        $text = Normalize-Text (Get-CellText $Worksheet $headerRow $column)
        if ($text -and $wanted -contains $text) {
            return [pscustomobject]@{ Column = $column; Header = Get-ColumnLetter $column; Method = 'header' }
        }
    }
    if ($FallbackLetter) {
        $number = 0
        foreach ($char in $FallbackLetter.ToUpperInvariant().ToCharArray()) {
            $number = $number * 26 + ([int][char]$char - 64)
        }
        return [pscustomobject]@{ Column = $number; Header = $FallbackLetter.ToUpperInvariant(); Method = 'fallback-column' }
    }
    return $null
}

function Find-TargetSheet($Workbook) {
    $best = $null
    $bestScore = -1
    foreach ($worksheet in $Workbook.Worksheets) {
        $used = $worksheet.UsedRange
        $score = 0
        $scanRows = [math]::Min(12, $used.Row + $used.Rows.Count - 1)
        $scanCols = [math]::Min(140, $used.Column + $used.Columns.Count - 1)
        for ($row = $used.Row; $row -le $scanRows; $row++) {
            for ($column = $used.Column; $column -le $scanCols; $column++) {
                $text = Normalize-Text (Get-CellText $worksheet $row $column)
                if ($text -eq 'sku') { $score += 3 }
                if ($text -eq 'product name') { $score += 3 }
                if ($text -eq 'site description') { $score += 2 }
                if ($text -eq 'main image url') { $score += 2 }
            }
        }
        if ($score -gt $bestScore) {
            $best = $worksheet
            $bestScore = $score
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($used)
    }
    if ($null -eq $best -or $bestScore -lt 5) {
        throw '未找到包含 SKU、Product Name 等字段的目标工作表。'
    }
    return $best
}

function Find-TargetColumns($Worksheet) {
    $used = $Worksheet.UsedRange
    $maxRow = $used.Row + $used.Rows.Count - 1
    $maxColumn = $used.Column + $used.Columns.Count - 1
    $headerRows = New-Object System.Collections.Generic.List[int]
    $descriptionRow = $null

    # The supplied Walmart template uses rows 2-5 as hierarchical headers.
    # Detect the last contiguous header row by looking for the required leaf labels.
    for ($row = $used.Row; $row -le [math]::Min($maxRow, 8); $row++) {
        $nonEmpty = 0
        for ($column = $used.Column; $column -le $maxColumn; $column++) {
            if ((Normalize-Text (Get-CellText $Worksheet $row $column))) { $nonEmpty++ }
        }
        if ($nonEmpty -gt 0) { $headerRows.Add($row) }
    }
    if ($headerRows.Count -eq 0) { throw '目标工作表没有可识别的表头。' }

    # Row 1 is the version marker in the template, not a field header.
    $headerRows = @($headerRows | Where-Object { $_ -gt $used.Row })
    if ($headerRows.Count -eq 0) { $headerRows = @($used.Row) }
    $headerLastRow = ($headerRows | Measure-Object -Maximum).Maximum

    for ($row = $headerLastRow + 1; $row -le [math]::Min($maxRow, $headerLastRow + 2); $row++) {
        $sample = Normalize-Text (Get-CellText $Worksheet $row ($used.Column + 3))
        if ($sample -match 'alphanumeric|decimal|closed list|url|date|number|boolean') {
            $descriptionRow = $row
            break
        }
    }
    $dataStart = if ($descriptionRow) { $descriptionRow + 1 } else { $headerLastRow + 1 }
    $columns = @{}
    for ($column = $used.Column; $column -le $maxColumn; $column++) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($row in $headerRows) {
            $text = Normalize-Text (Get-CellText $Worksheet $row $column)
            if ($text -and ($parts.Count -eq 0 -or $parts[$parts.Count - 1] -ne $text)) {
                $parts.Add($text)
            }
        }
        if ($parts.Count -gt 0) {
            $columns[$column] = [pscustomobject]@{
                Column = $column
                Letter = Get-ColumnLetter $column
                Path = ($parts -join ' > ')
                Leaf = $parts[$parts.Count - 1]
            }
        }
    }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($used)
    return [pscustomobject]@{ Columns = $columns; HeaderRows = $headerRows; DataStart = $dataStart }
}

function Find-TargetColumn($TargetColumns, [string[]]$LeafNames) {
    $wanted = @($LeafNames | ForEach-Object { Normalize-Text $_ } | Where-Object { $_ })
    foreach ($entry in $TargetColumns.GetEnumerator()) {
        if ($wanted -contains $entry.Value.Leaf) { return $entry.Value }
    }
    return $null
}

$excel = $null
$sourceWb = $null
$targetWb = $null
$sourceWs = $null
$targetWs = $null
$sourceUsed = $null
$result = [ordered]@{
    success = $false
    output = $OutputPath
    source = $SourcePath
    target = $TargetPath
    targetSheet = ''
    rowsRead = 0
    rowsWritten = 0
    mappings = @()
    skipped = @()
    message = ''
}

try {
    $resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
    $outputDirectory = [System.IO.Path]::GetDirectoryName($resolvedOutput)
    if (-not (Test-Path $outputDirectory)) { New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null }
    if ([System.IO.Path]::GetFullPath($TargetPath) -eq $resolvedOutput) { throw '输出文件不能覆盖原始 B 模板，请选择新的输出文件名。' }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $sourceWb = $excel.Workbooks.Open((Resolve-Path $SourcePath).Path, 0, $true)
    $targetWb = $excel.Workbooks.Open((Resolve-Path $TargetPath).Path, 0, $false)
    $sourceWs = $sourceWb.Worksheets.Item(1)
    $targetWs = Find-TargetSheet $targetWb
    $result.targetSheet = $targetWs.Name
    $targetInfo = Find-TargetColumns $targetWs
    $sourceUsed = $sourceWs.UsedRange
    $sourceHeaderRow = $sourceUsed.Row
    $sourceLastRow = $sourceUsed.Row + $sourceUsed.Rows.Count - 1

    $sourceDefinitions = @(
        @{ Key = '自定义SKU'; Names = @('自定义SKU'); Fallback = 'D'; Targets = @('SKU') },
        @{ Key = '标题'; Names = @('标题'); Fallback = 'N'; Targets = @('Product Name') },
        @{ Key = '长描述'; Names = @('长描述'); Fallback = 'O'; Targets = @('Site Description') },
        @{ Key = '五点1'; Names = @('五点1'); Fallback = 'P'; Targets = @('Key Features (+)') },
        @{ Key = '五点2'; Names = @('五点2'); Fallback = 'Q'; Targets = @('Key Features 1 (+)') },
        @{ Key = '五点3'; Names = @('五点3'); Fallback = 'R'; Targets = @('Key Features 2 (+)') },
        @{ Key = '五点4'; Names = @('五点4'); Fallback = 'S'; Targets = @('Key Features 3 (+)') },
        @{ Key = '五点5'; Names = @('五点5'); Fallback = 'T'; Targets = @('Key Features 4 (+)') },
        @{ Key = '小平台专用链接2'; Names = @('小平台专用链接 2', '小平台专用链接2'); Fallback = 'AI'; Targets = @('Main Image URL') },
        @{ Key = '颜色'; Names = @('颜色'); Fallback = 'L'; Targets = @('Color') },
        @{ Key = '小平台专用链接1'; Names = @('小平台专用链接 1', '小平台专用链接1'); Fallback = 'AG'; Targets = @('Additional Image URL (+)') },
        @{ Key = '小平台专用链接2附加'; Names = @('小平台专用链接 2', '小平台专用链接2'); Fallback = 'AI'; Targets = @('Additional Image URL 1 (+)') },
        @{ Key = '小平台专用链接3'; Names = @('小平台专用链接 3', '小平台专用链接3'); Fallback = 'AK'; Targets = @('Additional Image URL 2 (+)') },
        @{ Key = '小平台专用链接4'; Names = @('小平台专用链接 4', '小平台专用链接4'); Fallback = 'AM'; Targets = @('Additional Image URL 3 (+)') },
        @{ Key = '小平台专用链接5'; Names = @('小平台专用链接 5', '小平台专用链接5'); Fallback = 'AO'; Targets = @('Additional Image URL 4 (+)') },
        @{ Key = '小平台专用链接6'; Names = @('小平台专用链接 6', '小平台专用链接6'); Fallback = 'AQ'; Targets = @('Additional Image URL 5 (+)') },
        @{ Key = '小平台专用链接7'; Names = @('小平台专用链接 7', '小平台专用链接7'); Fallback = 'AS'; Targets = @('Additional Image URL 6 (+)') },
        @{ Key = '小平台专用链接8'; Names = @('小平台专用链接 8', '小平台专用链接8'); Fallback = 'AU'; Targets = @('Additional Image URL 7 (+)') },
        @{ Key = '父SKU'; Names = @('父SKU'); Fallback = 'E'; Targets = @('Variant Group ID') },
        @{ Key = '代理链接100*100缩率图'; Names = @('代理链接100*100缩率图'); Fallback = 'AX'; Targets = @('Swatch Image URL') }
    )

    $resolvedMappings = New-Object System.Collections.Generic.List[object]
    foreach ($definition in $sourceDefinitions) {
        $sourceColumn = Find-SourceColumn $sourceWs $definition.Names $definition.Fallback
        $targetColumn = Find-TargetColumn $targetInfo.Columns $definition.Targets
        if ($null -eq $sourceColumn) {
            $result.skipped += "源字段未找到: $($definition.Key)"
            continue
        }
        if ($null -eq $targetColumn) {
            $result.skipped += "目标字段未找到: $($definition.Targets -join ', ')"
            continue
        }
        $resolvedMappings.Add([pscustomobject]@{ Definition = $definition; Source = $sourceColumn; Target = $targetColumn })
        $result.mappings += [ordered]@{
            source = $definition.Key
            sourceColumn = $sourceColumn.Header
            sourceMethod = $sourceColumn.Method
            targetColumn = $targetColumn.Letter
            targetField = $targetColumn.Leaf
            targetPath = $targetColumn.Path
        }
    }

    $sourceDataStart = $sourceHeaderRow + 1
    $dataRows = New-Object System.Collections.Generic.List[int]
    for ($row = $sourceDataStart; $row -le $sourceLastRow; $row++) {
        $hasValue = $false
        foreach ($mapping in $resolvedMappings) {
            $value = Get-CellValue $sourceWs $row $mapping.Source.Column
            if ($null -ne $value -and [string]$value -ne '') { $hasValue = $true; break }
        }
        if ($hasValue) { $dataRows.Add($row) }
    }
    $result.rowsRead = $dataRows.Count

    for ($index = 0; $index -lt $dataRows.Count; $index++) {
        $sourceRow = $dataRows[$index]
        $targetRow = $targetInfo.DataStart + $index
        foreach ($mapping in $resolvedMappings) {
            $value = Get-CellValue $sourceWs $sourceRow $mapping.Source.Column
            if ($null -eq $value) { $value = '' }
            $targetWs.Cells.Item($targetRow, $mapping.Target.Column).Value2 = $value
        }
        $result.rowsWritten++
    }

    $targetWb.SaveCopyAs($resolvedOutput)
    $result.success = $true
    $result.message = if ($dataRows.Count -eq 0) { '未发现源数据行；已生成未填充的输出副本。' } else { '映射完成。' }
}
catch {
    $result.message = $_.Exception.Message
}
finally {
    if ($sourceWb) { $sourceWb.Close($false) }
    if ($targetWb) { $targetWb.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($object in @($sourceUsed, $sourceWs, $targetWs, $sourceWb, $targetWb, $excel)) {
        if ($object) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object) } catch {} }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$result | ConvertTo-Json -Depth 8 -Compress
if (-not $result.success) { exit 1 }
