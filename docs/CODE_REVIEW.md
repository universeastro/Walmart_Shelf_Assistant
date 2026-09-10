# 代码审查：excel_mapper.ps1

| | |
|---|---|
| 日期 | 2026-09-10 |
| 审查者 | Claude |
| 被审查版本 | `8cfafc5`（全文 295 行通读） |
| 审查方式 | 逐行阅读 + 对模板实测取证 |

**结论：无当前故障。** 映射功能正确（已验证 12/12）。以下 7 项均为
**潜伏问题与健壮性缺陷**——现在不发作，但会在换模板、加字段或上真实数据时发作。

按严重程度排序。第 1、2 项有实测证据，且**修复成本很低**，建议优先处理。

---

## 修复状态（2026-09-10 更新）

Codex 在 `39a28bb` 中处理了第 1、2、3、6、7 项。**已复验通过**：

```powershell
verify_mapping.ps1 -SourcePath _docx_qa/A_sample.xls -TargetPath 文件/B模板.xlsx
# -> 数据行仍为第 7 行，12/12 一致，退出码 0
```

| 项 | 状态 | 实现方式 |
|---|---|---|
| 1. 歧义匹配 | ✅ 已修 | 收集全部候选；有 `PathHints` 时用路径消歧；**仍不唯一则抛错**而非任选 |
| 2. 隐藏表 | ✅ 已修 | `if ($worksheet.Visible -ne -1) { continue }`（第 87 行） |
| 3. 描述行探测 | ✅ 已修 | 改为扫描全部列，要求 ≥5 列命中且正则收紧为 `alphanumeric,` 等带分隔符形式 |
| 4. 源表表头行数 | ⬜ 未处理 | 仍需先确认业务实际 |
| 5. 逐格 COM 写入 | ⬜ 未处理 | 真实数据前应做批量写入 |
| 6. COM 泄漏 | ✅ 已修 | 新增 `Set-CellValue`，`Get-CellValue` / `Set-CellValue` 均释放中间 RCW |
| 7. 只读打开 | ✅ 已修 | 改为 `Open(..., $true)`；**已实测确认只读下 `SaveCopyAs` 正常**，输出文件正常生成 |

### 复验补充说明

- **第 5 项（逐格 COM 写入）现在更值得做**：第 6 项已把每次调用的泄漏堵上，
  但数千次跨进程调用的开销本身仍在。合成样本只有 1 行，完全掩盖了这点。
- **第 1 项的抛错行为需要回归用例**：`Find-TargetColumn` 现在遇到歧义会 `throw`。
  当前映射表恰好没有指向 `unit` / `measure` 的字段，所以不触发。**但若将来新增
  这类字段而未提供 `PathHints`，映射会直接失败**（这是期望行为——宁可失败也不要写错列，
  但必须有测试覆盖，否则会在生产时才暴露）。
- **第 2 项为代码验证，非行为验证**：未构造出隐藏表得分反超的用例，
  仅确认代码路径存在且主流程仍正确选中可见表。

### 本次更新的过程失误（如实记录）

`39a28bb` 提交时，`git add -A` **把 Codex 同时写入的 `excel_mapper.ps1` 一并提交，
未经任何验证**。事后补验为通过，但这属于流程失控而非流程正确。

原因：编辑仓库文件与 `git add -A` 之间，Codex 在 13:55 写入了该文件，
而我未执行 `verify-and-archive` 第 0 步的 `git status` 检查。

**教训**：`git add -A` 在有并发写入者的仓库里是危险操作，应改为显式指定文件。

---

## 1. 目标列匹配只看 `Leaf`，丢弃了已算好的 `Path` ⚠️

**位置**：`Find-TargetColumn`，第 157-163 行

```powershell
foreach ($entry in $TargetColumns.GetEnumerator()) {
    if ($wanted -contains $entry.Value.Leaf) { return $entry.Value }
}
```

### 实测证据

模板中存在**重名的叶子表头**：

```
Leaf = 'unit'    出现 5 次：
    AM  assembled product width  > unit
    AK  assembled product weight > unit
    AI  assembled product height > unit
    AG  assembled product depth  > unit
    AA  net content              > unit

Leaf = 'measure' 出现 5 次：
    AL  assembled product width  > measure
    AJ  assembled product weight > measure
    AH  assembled product height > measure
    AF  assembled product depth  > measure
    AB  net content              > measure
```

### 影响

`Find-TargetColumns` 已经算出了完整的 `Path`
（如 `required for the item to be visible on walmart website > net content (netcontent) > unit`），
但 `Find-TargetColumn` 只比对最后一段 `Leaf`，**把用于消歧的信息丢掉了**。

更糟的是 `$TargetColumns` 是 `@{}`（**无序** `System.Collections.Hashtable`）。
当多个列共享同一个 `Leaf` 时，返回哪一列**取决于哈希枚举顺序**，不受代码控制。

**当前未触发**，因为 `$sourceDefinitions` 里没有映射到 `unit` / `measure` 的字段。
但一旦要填净含量、尺寸、重量（这是沃尔玛上架的高频字段），就会**静默写错列**。

### 建议

在叶子名有歧义时，用 `Path` 消歧。`Targets` 数组可以扩展为支持完整路径或父级限定，
例如先按 `Leaf` 收集候选，候选多于一个时再用 `Path` 匹配：

```powershell
$candidates = @($TargetColumns.GetEnumerator() | Where-Object { $wanted -contains $_.Value.Leaf })
if ($candidates.Count -eq 1) { return $candidates[0].Value }
# 多个候选 -> 用 Path 匹配，仍无法唯一确定就抛错而不是随便选一个
```

**关键原则：歧义时应当报错，而不是任选一列。** 写错列到沃尔玛后台是批量事故。

> ⚠️ 附带一条：**不要**把 `Find-TargetColumns` 第 135 行的 `$columns = @{}` 改成
> `[ordered]@{}`。`OrderedDictionary` 的索引器会把 `int` 参数当成**位置索引**
> 而非键，`$columns[$column] = ...` 会立刻抛 `ArgumentOutOfRangeException`
> （参数名 `index`）。本报告作者实测踩到过。

---

## 2. `Find-TargetSheet` 不排除隐藏工作表 ⚠️

**位置**：第 67-94 行

### 实测证据

按代码中的打分逻辑对四个工作表打分：

| 工作表 | 得分 | 可见性 | 命中 |
|---|---|---|---|
| Instructions and Examples | 0 | 可见 | — |
| Data Definitions | 0 | 可见 | — |
| **Hidden_product_content_and_sit** | **6** | **隐藏** | `sku`×2 |
| Product Content And Site Exp | 19 | 可见 | `sku`×4 等 |

接受阈值是 `$bestScore -lt 5`。

### 影响

隐藏表拿到 6 分，**越过了阈值**。当前因为正确表拿 19 分而安全——但这个余量
（19 vs 6）完全取决于本模板的巧合。

换一个字段较少、或 `sku` 字样出现次数较少的模板变体，隐藏表就可能胜出。
**届时数据会被写进一个隐藏工作表**，用户在界面上看不到任何异常，
打开输出文件也看不出问题，直到上架失败。

### 建议

跳过非可见工作表：

```powershell
if ($worksheet.Visible -ne -1) { continue }   # xlSheetVisible = -1
```

隐藏表本来就不该是写入目标。这一行改动成本极低，建议直接加上。

---

## 3. 描述行探测硬编码取第 4 列

**位置**：第 127-133 行

```powershell
$sample = Normalize-Text (Get-CellText $Worksheet $row ($used.Column + 3))
if ($sample -match 'alphanumeric|decimal|closed list|url|date|number|boolean') {
```

只取 `UsedRange` 的**第 4 列**判断这一行是不是字段类型描述行。
本模板该列恰好有类型关键字，所以能工作。

如果模板变体里第 4 列是空列或内容不同，`$descriptionRow` 保持 `$null`，
`$dataStart` 退化为 `$headerLastRow + 1`，**数据会整体上移一行**——
覆盖掉表头下方的行，且不报任何错。

### 建议

扫描表头行下方的多列（而非单列），按「命中类型关键字的列数占比」判断，
或要求至少 N 列命中才认定是描述行。

---

## 4. 假设源文件只有一个表头行

**位置**：第 200 行与第 252 行

```powershell
$sourceHeaderRow = $sourceUsed.Row          # 第 200 行
$sourceDataStart = $sourceHeaderRow + 1     # 第 252 行
```

`Find-SourceColumn` 同样只用 `$used.Row` 作为表头行。

这与目标表的多级表头处理**不对称**：目标表会智能定位叶子表头行，
源表却假定数据从第一行往下第二行开始。

如果某份选品表有两行表头（合并表头），**第二行表头会被当成数据**写进模板。

### 建议

需求文档未明确 A 表是否可能有多行表头。**先确认业务实际**，
若可能，应复用与目标表相同的「定位叶子表头行」逻辑。

---

## 5. 性能：逐格 COM 写入

**位置**：第 264-273 行

```powershell
$targetWs.Cells.Item($targetRow, $mapping.Target.Column).Value2 = $value
```

每个字段、每一行都是一次独立 COM 跨进程调用。

估算：真实选品表约 372 行 × 12 个字段 ≈ **4500 次 COM 调用**，
此外第 257 行扫描数据行时还有一轮同量级的读取。

**建议**：把待写区域组装成二维数组，一次性赋值：

```powershell
$block = New-Object 'object[,]' $rowCount, $colCount
# ... 填充 ...
$targetWs.Range($startCell, $endCell).Value2 = $block
```

这会带来数量级的速度提升。合成样本只有 1 行，完全掩盖了这个问题，
**上真实数据前值得先做**。

---

## 6. COM 对象泄漏

**位置**：`Get-CellText`（第 30-40 行）、`Get-CellValue`（第 42-44 行）

```powershell
return $Worksheet.Cells.Item($Row, $Column).Value2      # 中间 RCW 未释放
```

`Get-CellText` 里 `$cell.MergeArea` 与 `.Cells.Item(1,1)` 各自产生 RCW，
同样只释放了外层的 `$cell`。

`Get-CellValue` 在数据行扫描与写入两轮循环中被调用数千次，
泄漏的 RCW 会持续累积。脚本末尾的 `[GC]::Collect()` 能兜住最终回收，
但运行期间可能导致 Excel 进程内存增长、退出缓慢，甚至残留 `EXCEL.EXE` 进程。

**建议**：把中间对象也纳入 `finally` 释放，或在数据量增大后改为批量读写（见第 5 项，
批量方案会顺带消除绝大部分逐格调用）。

---

## 7. 目标工作簿以可写方式打开（次要）

**位置**：第 194 行

```powershell
$targetWb = $excel.Workbooks.Open((Resolve-Path $TargetPath).Path, 0, $false)
```

第三个参数 `$false` 表示**可写**。写入的是内存副本，且用 `SaveCopyAs` 落盘
（第 275 行），所以原始模板**不会被修改**——这一点是对的。

但以可写方式打开会在模板目录留下 Excel 锁文件（`~$B模板.xlsx`）。
若脚本异常终止（如强杀进程），锁文件会残留，下次打开模板时 Excel 提示
「文件已被占用」。

**建议**：改为以只读打开（`$true`）。只读工作簿仍可修改内存中的单元格并
`SaveCopyAs`，同时避免锁文件问题。需实测确认 `SaveCopyAs` 在只读模式下行为一致。

---

## 附：本次审查中的自查更正

前一份 `docs/VERIFICATION_REPORT.md` 中写「表头 2-5 行 + 字段描述行 6」有误。
实测 `headerRows = 2,3,4`，叶子表头行是**第 4 行**，第 6 行是描述行，
数据从第 7 行开始。已在验证报告中更正。

---

## 建议处理顺序

| 优先级 | 项 | 理由 |
|---|---|---|
| 高 | 2. 隐藏表 | 一行修复，消除一个无声故障模式 |
| 高 | 1. 歧义匹配 | 加字段时必然踩到，且写错列后果严重 |
| 中 | 3. 描述行探测 | 换模板即触发，会整体错行 |
| 中 | 5. 批量写入 | 上真实数据前必须做 |
| 中 | 4. 源表表头行数 | 需先确认业务实际 |
| 低 | 6. COM 泄漏 | 第 5 项做完后大部分自动消失 |
| 低 | 7. 只读打开 | 需实测验证 |
