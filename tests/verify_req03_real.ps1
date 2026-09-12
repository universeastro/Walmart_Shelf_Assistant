$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$repo = Split-Path $PSScriptRoot -Parent
$source = Join-Path $repo '文件\03\A模板02.xls'
$target = Join-Path $repo '文件\03\XPY沃尔玛价格计算.xls'
$output = Join-Path $env:TEMP ('wsa_req03_real_' + [guid]::NewGuid().ToString('N') + '.xls')
$excel = $null
$sourceWb = $null
$templateWb = $null
$outputWb = $null

Add-Type -TypeDefinition @'
using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

public static class Req03ArrayHash {
    private static void Add(HashAlgorithm hash, string value) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        hash.TransformBlock(bytes, 0, bytes.Length, bytes, 0);
    }

    public static string Compute(object value) {
        using (SHA256 hash = SHA256.Create()) {
            Array array = value as Array;
            if (array == null) {
                Add(hash, Scalar(value));
            } else {
                Add(hash, array.Rank.ToString(CultureInfo.InvariantCulture));
                for (int dimension = 0; dimension < array.Rank; dimension++) {
                    Add(hash, ":" + array.GetLength(dimension).ToString(CultureInfo.InvariantCulture));
                }
                foreach (object item in array) Add(hash, "|" + Scalar(item));
            }
            hash.TransformFinalBlock(new byte[0], 0, 0);
            return BitConverter.ToString(hash.Hash).Replace("-", "");
        }
    }

    private static string Scalar(object value) {
        if (value == null) return "<null>";
        return value.GetType().FullName + ":" + Convert.ToString(value, CultureInfo.InvariantCulture);
    }
}
'@

function Get-FormulaHash($Worksheet, [string]$Address = '') {
    $range = $null
    if ($Address) { $range = $Worksheet.Range($Address) }
    else { $range = $Worksheet.UsedRange }
    try { return [Req03ArrayHash]::Compute($range.Formula) }
    finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range) }
}

try {
    if (-not (Test-Path -LiteralPath $source) -or -not (Test-Path -LiteralPath $target)) {
        throw '缺少 文件/03 的 Req03 真实 A/B 文件。'
    }
    $targetHash = (Get-FileHash -Algorithm MD5 -LiteralPath $target).Hash
    $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'excel_mapper.ps1') `
        -SourcePath $source -TargetPath $target -OutputPath $output -Profile Req03 -RowMode Visible
    if ($LASTEXITCODE -ne 0) { throw ($raw | Out-String) }
    $result = $raw | ConvertFrom-Json
    if (-not $result.success -or $result.rowsRead -ne 306 -or $result.rowsWritten -ne 306 -or
        $result.rowsHiddenSkipped -ne 66 -or $result.existingLastRow -ne 0 -or
        $result.writeStartRow -ne 2 -or $result.mappings.Count -ne 6 -or
        $result.targetSheet -cne '导入 单位转换') {
        throw '真实 Req03 映射摘要与冻结契约不一致。'
    }
    if ((Get-FileHash -Algorithm MD5 -LiteralPath $target).Hash -ne $targetHash) {
        throw '真实 Req03 原始 B 文件哈希发生变化。'
    }
    $signature = [IO.File]::ReadAllBytes($output)[0..7]
    if ((($signature | ForEach-Object { $_.ToString('X2') }) -join ' ') -ne 'D0 CF 11 E0 A1 B1 1A E1') {
        throw '真实 Req03 输出不是二进制 .xls。'
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $sourceWb = $excel.Workbooks.Open($source, 0, $true)
    $templateWb = $excel.Workbooks.Open($target, 0, $true)
    $outputWb = $excel.Workbooks.Open($output, 0, $true)
    $sourceSheet = $sourceWb.Worksheets.Item(1)
    $templateSheet = $templateWb.Worksheets.Item('导入 单位转换')
    $outputSheet = $outputWb.Worksheets.Item('导入 单位转换')
    $sourceColumns = @(4, 27, 31, 51, 52, 53)
    $targetColumns = @(1, 2, 3, 8, 9, 10)
    $destinationRow = 2
    $checked = 0
    for ($sourceRow = 2; $sourceRow -le 373; $sourceRow++) {
        if ([bool]$sourceSheet.Rows.Item($sourceRow).Hidden) { continue }
        $hasValue = $false
        foreach ($column in $sourceColumns) {
            $value = $sourceSheet.Cells.Item($sourceRow, $column).Value2
            if ($null -ne $value -and [string]$value -ne '') { $hasValue = $true; break }
        }
        if (-not $hasValue) { continue }
        for ($index = 0; $index -lt $sourceColumns.Count; $index++) {
            $expected = $sourceSheet.Cells.Item($sourceRow, $sourceColumns[$index]).Value2
            $actual = $outputSheet.Cells.Item($destinationRow, $targetColumns[$index]).Value2
            if ($index -eq 0) {
                if ($actual -isnot [string] -or [string]$actual -cne [string]$expected) {
                    throw "真实 Req03 SKU 文本不匹配: A R$sourceRow -> B R$destinationRow"
                }
            } elseif ($null -eq $expected -or [string]$expected -eq '') {
                if ($null -ne $actual -and [string]$actual -ne '') {
                    throw "真实 Req03 空数值未保留为空: A R$sourceRow -> B R$destinationRow C$($targetColumns[$index])"
                }
            } else {
                $number = [double]0
                $parsed = [double]::TryParse(
                    ([string]$expected).Trim(),
                    [Globalization.NumberStyles]::Float,
                    [Globalization.CultureInfo]::InvariantCulture,
                    [ref]$number
                )
                if ($parsed) {
                    if ($actual -isnot [double] -or [double]$actual -ne $number) {
                        throw "真实 Req03 数值类型或值不匹配: A R$sourceRow -> B R$destinationRow C$($targetColumns[$index])"
                    }
                } elseif ([string]$actual -cne [string]$expected) {
                    throw "真实 Req03 文本回退值不匹配: A R$sourceRow -> B R$destinationRow C$($targetColumns[$index])"
                }
            }
            $checked++
        }
        $destinationRow++
    }
    if ($destinationRow -ne 308 -or $checked -ne 1836) { throw '真实 Req03 逐格验证数量错误。' }

    foreach ($address in @('D1:E16601', 'K1:N16601', 'Q1:R5220')) {
        if ((Get-FormulaHash $templateSheet $address) -cne (Get-FormulaHash $outputSheet $address)) {
            throw "真实 Req03 修改了导入表保护区域: $address"
        }
    }
    foreach ($sheetName in @('价格', '运费表（公式数据，不动）')) {
        $left = $templateWb.Worksheets.Item($sheetName)
        $right = $outputWb.Worksheets.Item($sheetName)
        if ((Get-FormulaHash $left) -cne (Get-FormulaHash $right)) {
            throw "真实 Req03 修改了非目标工作表: $sheetName"
        }
    }
    Write-Output "PASS: Req03 real template rows=306 cells=$checked numeric-types=true protected-content=true original-hash=true output=.xls"
} finally {
    if ($sourceWb) { $sourceWb.Close($false) }
    if ($templateWb) { $templateWb.Close($false) }
    if ($outputWb) { $outputWb.Close($false) }
    if ($excel) { $excel.Quit() }
    if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue }
}
