param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,
    [ValidateSet('Mainline', 'Req02', 'Req03')]
    [string]$Profile = 'Mainline',
    [ValidateSet('Append', 'Replace')]
    [string]$WriteMode = 'Append',
    [ValidateSet('Visible', 'All')]
    [string]$RowMode = 'Visible',
    [ValidateSet('Replace', 'AppendExisting')]
    [string]$OutputMode = 'Replace'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

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

function Get-ColumnNumber([string]$Letter) {
    $number = 0
    foreach ($char in $Letter.ToUpperInvariant().ToCharArray()) {
        $number = $number * 26 + ([int][char]$char - 64)
    }
    return $number
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
    $cell = $Worksheet.Cells.Item($Row, $Column)
    try {
        return $cell.Value2
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
    }
}

function Set-CellValue($Worksheet, [int]$Row, [int]$Column, [object]$Value, [int]$HorizontalAlignment = 5) {
    $cell = $Worksheet.Cells.Item($Row, $Column)
    try {
        $cell.Value2 = $Value
        $cell.HorizontalAlignment = $HorizontalAlignment
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell)
    }
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
        $number = Get-ColumnNumber $FallbackLetter
        return [pscustomobject]@{ Column = $number; Header = $FallbackLetter.ToUpperInvariant(); Method = 'fallback-column' }
    }
    return $null
}

function Find-TargetSheet($Workbook, [string]$MappingProfile = 'Mainline') {
    if ($MappingProfile -eq 'Req03') {
        $named = $null
        try { $named = $Workbook.Worksheets.Item('导入 单位转换') } catch {}
        if ($null -ne $named) {
            if ($named.Visible -ne -1) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($named)
                throw '目标工作表“导入 单位转换”不是可见工作表，为保护模板已停止写入。'
            }
            return $named
        }

        $required = @(
            'sku(直接从sheet 1导入）', 'sku价(￥)', '重量(g)', '长', '宽', '高'
        )
        $matches = New-Object System.Collections.Generic.List[object]
        foreach ($worksheet in $Workbook.Worksheets) {
            if ($worksheet.Visible -ne -1) { continue }
            $used = $worksheet.UsedRange
            try {
                $maxRow = [math]::Min(10, $used.Row + $used.Rows.Count - 1)
                $maxColumn = [math]::Min(255, $used.Column + $used.Columns.Count - 1)
                for ($row = $used.Row; $row -le $maxRow; $row++) {
                    $found = @{}
                    for ($column = $used.Column; $column -le $maxColumn; $column++) {
                        $text = Normalize-Text (Get-CellText $worksheet $row $column)
                        if ($required -contains $text) { $found[$text] = $true }
                    }
                    if ($found.Count -eq $required.Count) {
                        $matches.Add($worksheet)
                        break
                    }
                }
            } finally {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($used)
            }
        }
        if ($matches.Count -eq 0) { throw '未找到工作表“导入 单位转换”，也未找到包含六个 Req03 目标字段的工作表。' }
        if ($matches.Count -gt 1) {
            $names = ($matches | ForEach-Object { $_.Name }) -join ', '
            throw "多个工作表同时包含六个 Req03 目标字段，无法安全选择：$names"
        }
        return $matches[0]
    }

    if ($MappingProfile -eq 'Req02') {
        $matches = New-Object System.Collections.Generic.List[object]
        foreach ($worksheet in $Workbook.Worksheets) {
            if ($worksheet.Visible -ne -1) { continue }
            $used = $worksheet.UsedRange
            try {
                $maxRow = [math]::Min(10, $used.Row + $used.Rows.Count - 1)
                $maxColumn = [math]::Min(50, $used.Column + $used.Columns.Count - 1)
                for ($row = $used.Row; $row -le $maxRow; $row++) {
                    $skuFound = $false
                    $platformSkuFound = $false
                    for ($column = $used.Column; $column -le $maxColumn; $column++) {
                        $text = Normalize-Text (Get-CellText $worksheet $row $column)
                        if ($text -eq 'sku') { $skuFound = $true }
                        if ($text -eq '平台sku') { $platformSkuFound = $true }
                    }
                    if ($skuFound -and $platformSkuFound) {
                        $matches.Add($worksheet)
                        break
                    }
                }
            } finally {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($used)
            }
        }
        if ($matches.Count -eq 0) { throw '未找到同时包含 SKU 和平台SKU 的支线目标工作表。' }
        if ($matches.Count -gt 1) {
            $names = ($matches | ForEach-Object { $_.Name }) -join ', '
            throw "多个工作表同时包含 SKU 和平台SKU，无法安全选择：$names"
        }
        return $matches[0]
    }

    $best = $null
    $bestScore = -1
    foreach ($worksheet in $Workbook.Worksheets) {
        # Hidden metadata sheets can contain the same labels as the visible
        # entry sheet but are never valid destinations for user data.
        if ($worksheet.Visible -ne -1) { continue }
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

function Find-TargetColumns($Worksheet, [string]$MappingProfile = 'Mainline') {
    $used = $Worksheet.UsedRange
    $maxRow = $used.Row + $used.Rows.Count - 1
    $maxColumn = $used.Column + $used.Columns.Count - 1
    $headerRows = New-Object System.Collections.Generic.List[int]
    $descriptionRow = $null
    $leafHeaderRow = $null
    $bestReq03HeaderRow = $null
    $bestReq03HeaderCount = -1

    # Find the row containing the leaf field names for the selected profile.
    # Rows above it are the hierarchical group headers. Rows below it contain
    # XML names and descriptions and must not participate in label matching.
    for ($row = $used.Row; $row -le [math]::Min($maxRow, 10); $row++) {
        $skuFound = $false
        $productNameFound = $false
        $platformSkuFound = $false
        $req03Headers = @{}
        for ($column = $used.Column; $column -le $maxColumn; $column++) {
            $text = Normalize-Text (Get-CellText $Worksheet $row $column)
            if ($text -eq 'sku') { $skuFound = $true }
            if ($text -eq 'product name') { $productNameFound = $true }
            if ($text -eq '平台sku') { $platformSkuFound = $true }
            if (@('sku(直接从sheet 1导入）', 'sku价(￥)', '重量(g)', '长', '宽', '高') -contains $text) {
                $req03Headers[$text] = $true
            }
        }
        if ($MappingProfile -eq 'Req03' -and $req03Headers.Count -gt $bestReq03HeaderCount) {
            $bestReq03HeaderRow = $row
            $bestReq03HeaderCount = $req03Headers.Count
        }
        $rowMatches = if ($MappingProfile -eq 'Req02') {
            $skuFound -and $platformSkuFound
        } elseif ($MappingProfile -eq 'Req03') {
            $req03Headers.Count -eq 6
        } else {
            $skuFound -and $productNameFound
        }
        if ($rowMatches) {
            $leafHeaderRow = $row
            break
        }
    }
    if ($MappingProfile -eq 'Req03' -and $null -eq $leafHeaderRow) {
        $leafHeaderRow = if ($null -ne $bestReq03HeaderRow) { $bestReq03HeaderRow } else { $used.Row }
    }
    if ($null -eq $leafHeaderRow) { throw '目标工作表没有可识别的字段表头。' }
    for ($row = $used.Row; $row -le $leafHeaderRow; $row++) {
        if ($row -gt $used.Row) { $headerRows.Add($row) }
    }
    if ($headerRows.Count -eq 0) { $headerRows.Add($leafHeaderRow) }
    $headerLastRow = $leafHeaderRow

    for ($row = $headerLastRow + 1; $MappingProfile -eq 'Mainline' -and $row -le [math]::Min($maxRow, $headerLastRow + 3); $row++) {
        $keywordHits = 0
        for ($column = $used.Column; $column -le $maxColumn; $column++) {
            $sample = Normalize-Text (Get-CellText $Worksheet $row $column)
            if ($sample -match 'alphanumeric,|decimal,|closed list -|dateonly,|number,|boolean,') {
                $keywordHits++
            }
        }
        # XML-name rows can contain isolated tokens such as "url" or
        # "date". A description row repeats type markers across the whole
        # schema, so require several independent hits.
        if ($keywordHits -ge 5) {
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
                Method = 'header'
            }
        }
    }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($used)
    return [pscustomobject]@{ Columns = $columns; HeaderRows = $headerRows; DataStart = $dataStart }
}

function Find-TargetColumn($TargetColumns, [string[]]$LeafNames, [string[]]$PathHints = @(), [string]$FallbackLetter = '') {
    $wanted = @($LeafNames | ForEach-Object { Normalize-Text $_ } | Where-Object { $_ })
    $candidates = @($TargetColumns.GetEnumerator() | Where-Object { $wanted -contains $_.Value.Leaf })
    if ($candidates.Count -eq 1) { return $candidates[0].Value }
    if ($candidates.Count -gt 1 -and $PathHints.Count -gt 0) {
        $hints = @($PathHints | ForEach-Object { Normalize-Text $_ })
        $pathMatches = @($candidates | Where-Object {
            foreach ($hint in $hints) {
                if ($_.Value.Path -like "*$hint*") { return $true }
            }
            return $false
        })
        if ($pathMatches.Count -eq 1) { return $pathMatches[0].Value }
    }
    if ($candidates.Count -gt 1) {
        $locations = ($candidates | ForEach-Object { $_.Value.Letter + ':' + $_.Value.Path }) -join '; '
        throw "目标字段名称存在多个候选列，无法安全匹配 [$($LeafNames -join ', ')]: $locations"
    }
    if ($FallbackLetter) {
        $number = Get-ColumnNumber $FallbackLetter
        return [pscustomobject]@{
            Column = $number
            Letter = $FallbackLetter.ToUpperInvariant()
            Path = $LeafNames[0]
            Leaf = $LeafNames[0]
            Method = 'fallback-column'
        }
    }
    return $null
}

function Get-LastRecordRow($Worksheet, [int]$DataStart) {
    $used = $Worksheet.UsedRange
    $constants = $null
    $last = 0
    try {
        # Constants distinguish records from preformatted rows and formula scaffolding.
        # Include all used columns, not only SKU or the mapped columns.
        try { $constants = $used.SpecialCells(2) } # xlCellTypeConstants
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2146827284) { throw }
        }
        if ($null -ne $constants) {
            foreach ($area in $constants.Areas) {
                try {
                    $bottom = $area.Row + $area.Rows.Count - 1
                    if ($bottom -ge $DataStart) { $last = [math]::Max($last, $bottom) }
                } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($area) }
            }
        }
        return $last
    } finally {
        foreach ($item in @($constants, $used)) {
            if ($null -ne $item) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
        }
    }
}

function Get-LastRecordRowForMappings($Worksheet, $Mappings, [int]$DataStart) {
    $used = $Worksheet.UsedRange
    try {
        $usedLastRow = $used.Row + $used.Rows.Count - 1
        if ($usedLastRow -lt $DataStart) { return 0 }
        $last = 0
        $seen = @{}
        foreach ($mapping in $Mappings) {
            $column = [int]$mapping.Target.Column
            if ($seen.ContainsKey($column)) { continue }
            $seen[$column] = $true
            $letter = Get-ColumnLetter $column
            $range = $Worksheet.Range("${letter}${DataStart}:${letter}${usedLastRow}")
            try {
                foreach ($cellType in @(2, -4123)) { # constants, then formulas
                    $cells = $null
                    try {
                        try { $cells = $range.SpecialCells($cellType) }
                        catch [System.Runtime.InteropServices.COMException] {
                            if ($_.Exception.HResult -ne -2146827284) { throw }
                        }
                        if ($null -ne $cells) {
                            foreach ($area in $cells.Areas) {
                                try {
                                    $bottom = $area.Row + $area.Rows.Count - 1
                                    $last = [math]::Max($last, $bottom)
                                } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($area) }
                            }
                        }
                    } finally {
                        if ($null -ne $cells) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cells) }
                    }
                }
            } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
        }
        return $last
    } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($used) }
}

function Convert-Req03Value([object]$Value, [string]$ValueType) {
    if ($null -eq $Value) { return '' }
    if ($ValueType -eq 'Text') { return [string]$Value }
    if ($Value -isnot [string]) { return $Value }
    $text = $Value.Trim()
    if (-not $text) { return '' }
    $number = [double]0
    if ([double]::TryParse(
        $text,
        [Globalization.NumberStyles]::Float,
        [Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )) {
        return $number
    }
    return $Value
}

function Assert-WriteRegion($Worksheet, $Mappings, [int]$Start, [int]$Count) {
    if ($Count -eq 0) { return }
    $end = $Start + $Count - 1
    if ($end -gt $Worksheet.Rows.Count) { throw '写入数据超过工作表最大行数。' }
    foreach ($mapping in $Mappings) {
        $letter = $mapping.Target.Letter
        $range = $Worksheet.Range("${letter}${Start}:${letter}${end}")
        try {
            # Mixed ranges return null, so only an explicit false is safe.
            if ($range.HasFormula -ne $false -or $range.MergeCells -ne $false) {
                throw "写入区域 ${letter}${Start}:${letter}${end} 包含公式或合并单元格，为保护模板已停止写入。"
            }
        } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
    }
}

$excel = $null
$sourceWb = $null
$targetWb = $null
$sourceWs = $null
$targetWs = $null
$sourceUsed = $null
$workingTargetPath = $null
$applicationOptimized = $false
$stage = '检查输出路径'
$result = [ordered]@{
    success = $false
    output = $OutputPath
    source = $SourcePath
    target = $TargetPath
    targetSheet = ''
    rowsRead = 0
    rowsWritten = 0
    rowsHiddenSkipped = 0
    existingLastRow = 0
    writeStartRow = 0
    rowMode = $RowMode
    profile = $Profile
    writeMode = $WriteMode
    outputMode = $OutputMode
    mappings = @()
    skipped = @()
    message = ''
}

try {
    if ($Profile -ne 'Req02' -and $WriteMode -ne 'Append') {
        throw 'WriteMode Replace 仅适用于支线 Req02。'
    }
    if ($OutputMode -eq 'AppendExisting' -and $WriteMode -ne 'Append') {
        throw '继续写入现有输出不能与 WriteMode Replace 同时使用。'
    }
    $resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
    if (Test-Path -LiteralPath $resolvedOutput -PathType Container) { throw '输出路径是文件夹，请指定完整的 Excel 文件名。' }
    if (-not [System.IO.Path]::GetExtension($resolvedOutput)) { $resolvedOutput += [System.IO.Path]::GetExtension($TargetPath) }
    if ($Profile -eq 'Req03' -and [System.IO.Path]::GetExtension($resolvedOutput) -ine '.xls') {
        throw '支线 Req03 的输出文件必须使用 .xls 格式。'
    }
    $result.output = $resolvedOutput
    $outputDirectory = [System.IO.Path]::GetDirectoryName($resolvedOutput)
    if (-not (Test-Path $outputDirectory)) { New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null }
    if ([System.IO.Path]::GetFullPath($TargetPath) -eq $resolvedOutput) { throw '输出文件不能覆盖原始 B 模板，请选择新的输出文件名。' }
    if ([System.IO.Path]::GetFullPath($SourcePath) -eq $resolvedOutput) { throw '输出文件不能覆盖文件 A。' }
    if ($OutputMode -eq 'AppendExisting' -and -not (Test-Path -LiteralPath $resolvedOutput -PathType Leaf)) {
        throw '选择继续写入时，输出文件必须已经存在。'
    }
    if (Test-Path -LiteralPath $resolvedOutput) {
        try {
            $probe = [IO.File]::Open($resolvedOutput, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            $probe.Dispose()
        } catch { throw "输出文件被占用或不可写：$resolvedOutput。请关闭该文件或选择新的输出文件名。" }
    }

    $stage = '启动 Excel'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $applicationOptimized = $true
    $stage = '读取文件 A'
    $sourceWb = $excel.Workbooks.Open((Resolve-Path $SourcePath).Path, 0, $true)
    # Work only on a private copy. Replace starts from B; AppendExisting starts
    # from the selected output so its current rows and workbook features survive.
    $baseTargetPath = if ($OutputMode -eq 'AppendExisting') { $resolvedOutput } else { (Resolve-Path $TargetPath).Path }
    $workingTargetPath = Join-Path $outputDirectory ('.walmart_shelf_assistant_' + [guid]::NewGuid().ToString('N') + [System.IO.Path]::GetExtension($baseTargetPath))
    Copy-Item -LiteralPath $baseTargetPath -Destination $workingTargetPath -Force
    $stage = if ($OutputMode -eq 'AppendExisting') { '读取现有输出副本' } else { '读取模板副本' }
    $targetWb = $excel.Workbooks.Open($workingTargetPath, 0, $false)
    $stage = '匹配并写入数据'
    $sourceWs = $sourceWb.Worksheets.Item(1)
    $targetWs = Find-TargetSheet $targetWb $Profile
    $result.targetSheet = $targetWs.Name
    $targetInfo = Find-TargetColumns $targetWs $Profile
    $sourceUsed = $sourceWs.UsedRange
    $sourceHeaderRow = $sourceUsed.Row
    $sourceLastRow = $sourceUsed.Row + $sourceUsed.Rows.Count - 1

    $sourceDefinitions = if ($Profile -eq 'Req02') { @(
        @{ Key = 'SKU'; Names = @('SKU'); Fallback = ''; Targets = @('SKU') },
        @{ Key = '自定义'; Names = @('自定义', '自定义SKU'); Fallback = ''; Targets = @('平台SKU') }
    ) } elseif ($Profile -eq 'Req03') { @(
        @{ Key = '自定义SKU'; Names = @('自定义SKU', '自定义'); Fallback = 'D'; Targets = @('SKU(直接从sheet 1导入）'); TargetFallback = 'A'; ValueType = 'Text' },
        @{ Key = 'SKU价(￥)'; Names = @('SKU价(￥)'); Fallback = 'AA'; Targets = @('SKU价(￥)'); TargetFallback = 'B'; ValueType = 'Number' },
        @{ Key = '重量(g)'; Names = @('重量(g)'); Fallback = 'AE'; Targets = @('重量(g)'); TargetFallback = 'C'; ValueType = 'Number' },
        @{ Key = '长'; Names = @('长'); Fallback = 'AY'; Targets = @('长'); TargetFallback = 'H'; ValueType = 'Number' },
        @{ Key = '宽'; Names = @('宽'); Fallback = 'AZ'; Targets = @('宽'); TargetFallback = 'I'; ValueType = 'Number' },
        @{ Key = '高'; Names = @('高'); Fallback = 'BA'; Targets = @('高'); TargetFallback = 'J'; ValueType = 'Number' }
    ) } else { @(
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
        # The supplied A template has an empty E header; preserve the
        # documented E-column mapping instead of treating the nearby 父SKU
        # helper column B as the source for Variant Group ID.
        @{ Key = '父SKU'; Names = @(); Fallback = 'E'; Targets = @('Variant Group ID') },
        @{ Key = '代理链接100*100缩率图'; Names = @('代理链接100*100缩率图'); Fallback = 'AX'; Targets = @('Swatch Image URL') }
    ) }

    $resolvedMappings = New-Object System.Collections.Generic.List[object]
    foreach ($definition in $sourceDefinitions) {
        $sourceColumn = Find-SourceColumn $sourceWs $definition.Names $definition.Fallback
        $targetColumn = Find-TargetColumn $targetInfo.Columns $definition.Targets @() $definition.TargetFallback
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
            targetMethod = if ($targetColumn.Method) { $targetColumn.Method } else { 'header' }
        }
    }

    if ($Profile -eq 'Req02' -and $resolvedMappings.Count -ne 2) {
        throw "支线映射不完整：必须同时找到源字段 SKU、自定义（兼容自定义SKU）及目标字段 SKU、平台SKU。"
    }

    $sourceDataStart = $sourceHeaderRow + 1
    $dataRows = New-Object System.Collections.Generic.List[int]
    for ($row = $sourceDataStart; $row -le $sourceLastRow; $row++) {
        $sourceRowRange = $sourceWs.Rows.Item($row)
        try { $isHidden = [bool]$sourceRowRange.Hidden }
        finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sourceRowRange) }
        if ($isHidden -and $RowMode -eq 'Visible') {
            $result.rowsHiddenSkipped++
            continue
        }
        $hasValue = $false
        foreach ($mapping in $resolvedMappings) {
            $value = Get-CellValue $sourceWs $row $mapping.Source.Column
            if ($null -ne $value -and [string]$value -ne '') { $hasValue = $true; break }
        }
        if ($hasValue) { $dataRows.Add($row) }
    }
    $result.rowsRead = $dataRows.Count

    $stage = if ($Profile -eq 'Req02' -and $WriteMode -eq 'Replace') { '检查替换区域' } else { '检查追加位置' }
    $result.existingLastRow = if ($Profile -eq 'Req03') {
        Get-LastRecordRowForMappings $targetWs $resolvedMappings $targetInfo.DataStart
    } else {
        Get-LastRecordRow $targetWs $targetInfo.DataStart
    }
    $writeStart = if ($Profile -eq 'Req02' -and $WriteMode -eq 'Replace') {
        $targetInfo.DataStart
    } elseif ($result.existingLastRow -gt 0) {
        $result.existingLastRow + 4
    } else {
        $targetInfo.DataStart
    }
    $result.writeStartRow = $writeStart
    if ($Profile -eq 'Req02' -and $WriteMode -eq 'Replace' -and $dataRows.Count -gt 0) {
        $clearEnd = [math]::Max($result.existingLastRow, $writeStart + $dataRows.Count - 1)
        if ($clearEnd -ge $writeStart) {
            Assert-WriteRegion $targetWs $resolvedMappings $writeStart ($clearEnd - $writeStart + 1)
            $stage = '清理支线旧数据'
            foreach ($mapping in $resolvedMappings) {
                $letter = $mapping.Target.Letter
                $range = $targetWs.Range("${letter}${writeStart}:${letter}${clearEnd}")
                try { $range.ClearContents() }
                finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
            }
        }
    } else {
        Assert-WriteRegion $targetWs $resolvedMappings $writeStart $dataRows.Count
    }
    $stage = if ($Profile -eq 'Req02') { '写入支线数据' } elseif ($Profile -eq 'Req03') { '写入支线 03 数据' } else { '写入追加数据' }

    if ($Profile -eq 'Req03' -and $dataRows.Count -gt 0) {
        foreach ($mapping in $resolvedMappings) {
            $block = New-Object 'object[,]' $dataRows.Count, 1
            for ($index = 0; $index -lt $dataRows.Count; $index++) {
                $value = Get-CellValue $sourceWs $dataRows[$index] $mapping.Source.Column
                $block[($index), 0] = Convert-Req03Value $value $mapping.Definition.ValueType
            }
            $letter = $mapping.Target.Letter
            $endRow = $writeStart + $dataRows.Count - 1
            $range = $targetWs.Range("${letter}${writeStart}:${letter}${endRow}")
            try { $range.Value2 = $block }
            finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
        }
        $result.rowsWritten = $dataRows.Count
    } else {
        for ($index = 0; $index -lt $dataRows.Count; $index++) {
            $sourceRow = $dataRows[$index]
            $targetRow = $writeStart + $index
            foreach ($mapping in $resolvedMappings) {
                $value = Get-CellValue $sourceWs $sourceRow $mapping.Source.Column
                if ($null -eq $value) { $value = '' }
                if ($Profile -eq 'Req02') {
                    $cell = $targetWs.Cells.Item($targetRow, $mapping.Target.Column)
                    try { $cell.Value2 = $value }
                    finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
                } else {
                    $alignment = if ((Normalize-Text $mapping.Target.Leaf) -eq 'color') { -4131 } else { 5 }
                    Set-CellValue $targetWs $targetRow $mapping.Target.Column $value $alignment
                }
            }
            $result.rowsWritten++
        }
    }

    $stage = '保存输出文件'
    $excel.EnableEvents = $true
    $excel.ScreenUpdating = $true
    $applicationOptimized = $false
    $targetWb.SaveAs($resolvedOutput)
    $result.output = $targetWb.FullName
    $result.success = $true
    $result.message = if ($dataRows.Count -eq 0) { '未发现源数据行；已生成未填充的输出副本。' } else { '映射完成。' }
}
catch {
    $result.message = $_.Exception.Message
    $result.stage = $stage
    if ($stage -eq '保存输出文件') {
        $result.message = "无法保存到 $resolvedOutput。请确认输出文件未在 Excel/WPS 中打开，或选择新的文件名。原始错误：" + $result.message
    }
    $result.errorLine = $_.InvocationInfo.ScriptLineNumber
    $result.errorType = $_.Exception.GetType().FullName
    if ($result.message -match '80010108|RPC_E_DISCONNECTED') {
        $result.message = 'Excel 连接已断开。请重新运行任务；如果再次失败，请提供本次使用的 A/B 文件路径。原始错误：' + $result.message
    }
}
finally {
    # A disconnected Excel instance must not suppress the original result.
    if ($excel -and $applicationOptimized) {
        try { $excel.EnableEvents = $true } catch {}
        try { $excel.ScreenUpdating = $true } catch {}
    }
    if ($sourceWb) { try { $sourceWb.Close($false) } catch {} }
    if ($targetWb) { try { $targetWb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    foreach ($object in @($sourceUsed, $sourceWs, $targetWs, $sourceWb, $targetWb, $excel)) {
        if ($object) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($object) } catch {} }
    }
    if ($workingTargetPath -and (Test-Path -LiteralPath $workingTargetPath)) {
        Remove-Item -LiteralPath $workingTargetPath -Force -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$result | ConvertTo-Json -Depth 8 -Compress
if (-not $result.success) { exit 1 }
