# 代码审查：excel_mapper.ps1

| | |
|---|---|
| 日期 | 2026-09-10 |
| 审查者 | Claude |
| 被审查版本 | `8cfafc5`（全文 295 行通读） |
| 审查方式 | 逐行阅读 + 对模板实测取证 |

**结论：无当前故障。** 映射功能正确（已验证 12/12）。以下 10 项均为
**潜伏问题与健壮性缺陷**——现在不发作，但会在换模板、加字段或上真实数据时发作。

按严重程度排序。第 1、2 项有实测证据，且**修复成本很低**，建议优先处理。

---

## 修复状态（2026-09-10 更新）

Codex 在 `39a28bb` 中处理了第 1、2、3、6、7 项。**已复验通过**：

```powershell
verify_mapping.ps1 -SourcePath tests/fixtures/A_sample.xls -TargetPath 文件/B模板.xlsx
# -> 数据行仍为第 7 行，12/12 一致，退出码 0
```

| 项 | 状态 | 实现方式 |
|---|---|---|
| 1. 歧义匹配 | ✅ 已修 | 收集全部候选；有 `PathHints` 时用路径消歧；**仍不唯一则抛错**而非任选 |
| 2. 隐藏表 | ✅ 已修 | `if ($worksheet.Visible -ne -1) { continue }`（第 87 行） |
| 3. 描述行探测 | ✅ 已修 | 改为扫描全部列，要求 ≥5 列命中且正则收紧为 `alphanumeric,` 等带分隔符形式 |
| 4. 源表表头行数 | ⬜ 未处理 | 仍需先确认业务实际 |
| 5. 逐格 COM 写入 | ⬜ 未处理，**已实测量化** | 372 行 20 秒，不阻塞；见下方实测 |
| 6. COM 泄漏 | ✅ 已修 | 新增 `Set-CellValue`，`Get-CellValue` / `Set-CellValue` 均释放中间 RCW |
| 7. 只读打开 | ✅ 已修（**方案后被更换**） | `e7b1fcc` 改为「复制到输出目录的 GUID 临时文件 → 可写打开 → `SaveAs` → `finally` 删除」。见下方说明 |
| 8. `output` 字段与实际落盘路径不符 | ✅ 已修（`a854441`） | 无扩展名时按模板扩展名补齐；`SaveAs` 后从 `$targetWb.FullName` 回读真实路径。已实测确认输出无扩展名时 `output` 报告 `.xlsx` 且文件确实生成在该处 |

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

### 第 7 项的方案在 `e7b1fcc` 中被更换（不是回归）

原修复是「只读打开模板 `Open(..., $true)`」，我实测确认只读下 `SaveCopyAs` 正常。
`e7b1fcc` 换成了另一条路：**把模板复制到输出目录下的 GUID 临时文件，以可写方式打开那个副本，
`SaveAs` 到目标路径，`finally` 里删掉临时文件。**

新方案更彻底，理由：只读方案虽然避免了模板被改写，但**锁文件的问题只是转移**——
Excel 仍会在模板所在目录留下 `~$B模板.xlsx`。新方案下临时文件在输出目录，
模板目录**自始至终不会出现锁文件**。

实测确认（模板 md5 前后同为 `4ec60063`）：

| 检查 | 结果 |
|---|---|
| 原模板是否被修改 | 未修改（md5 逐字节一致） |
| 输出目录残留 | 跑入空目录后仅有 `out.xlsx`，无临时文件、无锁文件 |
| 失败后是否清理 | 构造 `SaveAs` 之后的失败路径，仍无残留 |

> ⚠️ **切换 `SaveCopyAs → SaveAs` 时我怀疑会丢格式，实测证伪了。**
> 理由是 `SaveAs` 会按扩展名推断文件格式（`.csv` 会退化成 CSV，丢掉全部格式和多表），
> 而 `SaveCopyAs` 不会。实测把输出命名为 `out.csv`，落盘文件头是 `PK..`（ZIP）、
> 268KB 完整工作簿——**格式没有被转换**。该风险未成立，记录在此以备后续换用
> 带 `FileFormat` 参数的重载时对照。

---

## 8. `output` 字段可能与实际落盘路径不符（低）

**位置**：`excel_mapper.ps1` 第 215 行 `output = $OutputPath`、第 320 行 `SaveAs($resolvedOutput)`

`SaveAs` 在路径缺少 Excel 可识别的扩展名时会**自行补上扩展名**。实测把输出指定为
`_docx_qa/_straycheck`（一个已存在的目录名），Excel 实际写出的是
`_docx_qa/_straycheck.xlsx`，而返回的 JSON 里 `output` 报的仍是
`_docx_qa/_straycheck`——**一个并不存在文件的路径**。GUI 会显示「处理完成」，
用户按提示去找文件却找不到。

**当前触发面很窄**：GUI 的保存对话框设了 `defaultextension=".xlsx"`，
正常走对话框不会踩到；只有用户选「所有文件」并手输无扩展名的名字时才会。

**建议**：`SaveAs` 后从 `$targetWb.FullName` 回读真实路径再填进 `result.output`，
而不是回填调用方传入的 `$OutputPath`。

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

### 实测（2026-09-10，多行夹具补齐后）

| 操作 | 规模 | 耗时 |
|---|---|---|
| `excel_mapper.ps1` 完整映射 | 372 行 × 12 字段（约 4464 次写入 + 同量级读取） | **20 秒** |
| 夹具生成脚本纯写入（对照） | 372 行 × 13 列 = 4836 次写入 | 27 秒（约 180 次/秒） |

**结论：20 秒尚未构成阻塞。** 本项从「上真实数据前必须做」降级为
**「值得做但不紧急」**——它换来的是数量级提速，但当前的绝对耗时用户可以接受。
若将来行数或字段数翻几倍，应重新评估。

> 注意：这里测的是**映射**耗时。整条链路还要加上 Excel 启动、
> 模板复制与 `SaveAs` 重新序列化整个包的开销，用户感知的总时长会更长。

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

## 9. `rowsHiddenSkipped` 把「隐藏的空行」也计入（低，仅影响日志数字）

**位置**：`excel_mapper.ps1` 的 `-RowMode Visible` 隐藏行分支

```powershell
if ($isHidden -and $RowMode -eq 'Visible') {
    $result.rowsHiddenSkipped++
    continue
}
```

判据只有 `Hidden`，**不看该行有没有数据**。而源表的 `UsedRange` 常比数据区
多出若干空行（仅有格式），这些行同样会被隐藏（手工隐藏时未必，但 **AutoFilter
筛选会连它们一起藏**），于是被计入跳过数。

### 实测证据（2026-09-10）

5 行夹具、标题列筛选 `= "Sample product 3"`，源表 `UsedRange = $A$1:$BC$8`，
第 7、8 行无值仅有格式：

```
rowsRead=1 rowsWritten=1 rowsHiddenSkipped=6
```

只有 4 行是用户筛掉的数据行，另 2 行是被筛选连带隐藏的空行。
**输出正确**（只写 1 行，目标 H7 = `Sample product 3`），
**只是 GUI 日志会显示「跳过 6 行」**。

### 影响

用户可能据此以为自己筛掉了 6 行而回头核对——是困惑，不是数据事故。
`tests/verify_autofilter.ps1` 已把该数字**排除在断言之外**。

### 建议

若要精确，在 `rowsHiddenSkipped++` 前加一次「该行是否有数据」的判定
（复用第 257 行扫描数据行时已有的判空逻辑即可）。**不紧急**，
但注意别把它和「空行压缩」的逻辑搅在一起——两者是独立的跳过原因，
将来若要分别统计，应拆成两个计数器。

---

## 10. 空值单元格也被强制设为「水平对齐：填充」（低，洁净度）

**位置**：`excel_mapper.ps1` 写入循环

```powershell
$value = Get-CellValue $sourceWs $sourceRow $mapping.Source.Column
if ($null -eq $value) { $value = '' }
Set-CellValue $targetWs $targetRow $mapping.Target.Column $value   # 内部会设 HorizontalAlignment
```

`Set-CellValue` 对**每个映射字段**都调用，源值为空时也一样。于是源表某字段为空时，
目标单元格被清空、但水平对齐仍被改成 5。

### 影响

**看不出差别**——单元格本来就没有内容。属于「改动范围比必要范围大」，
不是可见缺陷。**输出值完全正确**（`verify_mapping.ps1` 12/12）。

### 建议

若要「空值不碰格式」，在调用前判空即可。但**先确认这是否是期望行为**：
「填充」对齐在空单元格上无副作用，为一个看不见的差异增加一条分支，
未必划算。**不紧急，不建议现在动。**

特别提示：**不要**为了让空值跳过而对 `Set-CellValue` 的调用加条件判断，
除非同时确认数据行扫描用的 `$hasValue` 逻辑不受影响——那两处判空是独立的，
混在一起会重演第 9 项那类「两个跳过原因共用一个计数器」的问题。

---

## 11. 关闭窗口时正在处理：进度条动画与工作线程都不收尾（低）

`ShelfAssistant.destroy()` 只做两件事：`_hide_path_tip()` 和 `_save_paths()`，
然后 `super().destroy()`。**它不停进度条，也不理会工作线程。**

于是「处理中直接关窗口」有两条后果：

**一、进度条动画的 `after` 脚本留在事件队列里**，窗口销毁后照样触发，
控制台刷出：

```
cannot invoke "winfo" command: application has been destroyed
    while executing "winfo exists $pb"
    (procedure "ttk::progressbar::Autoincrement" line 4)
```

**二、工作线程跑完子进程后调 `self.after(0, self._finish, ...)`**，
此时 widget 已销毁 → 抛 `TclError`。该异常落进 `_run_worker` 的
`except Exception` 分支，而那个分支**又调了一次 `self.after`**：

```python
except Exception as exc:
    self.after(0, self._finish, {"success": False, "message": str(exc)}, 1)
```

第二次调用同样抛错，这次**无人接住**，工作线程带 traceback 死掉。

### 实测

已确定性复现 `after` 抛出的一类情形：主线程不在 `mainloop()` 里时，
跨线程调 `self.after` 抛 `RuntimeError: main thread is not in main loop`。
正常使用（用户跑 `mainloop()`）不会触发——**这条路径只在关窗口时可达**。

### 影响

**低。** 程序反正要关了，用户看不到。危害是：`except` 分支本意是
「出错也要把界面恢复可操作」，但它自己会抛错，等于**这个兜底是假的**——
一旦 `after` 因任何原因失败，`self.running` 永远是 `True`，控件永久锁死、
进度条永远转、没有弹窗也没有报错，只能杀进程。

### 建议

两处小改，都不影响正常流程：

1. `destroy()` 里加 `self.progress.stop()`。
2. `except` 分支用 `try: self.after(...) except tk.TclError: pass` 包一层，
   或者先判 `self.winfo_exists()`。关键是**兜底代码自己不能抛**。

**未改动**——留给 Codex 判断是否值得为「关窗口」这一条路径加防御。

---

## 12. 每次处理都弹出 PowerShell 黑框（中，用户可见）

**用户报告**：每次使用都会弹出一个 PowerShell 窗口。作为一个面向用户的 app，
不可接受。

### 成因

`_run_worker()` 启动映射脚本时**没有传 `creationflags`**：

```python
completed = subprocess.run(command, capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
```

程序以 `--windowed` 打包，**自身没有控制台**。而 `powershell.exe` 是控制台程序——
无控制台的进程启动控制台程序时，Windows 会**给子进程新分配一个控制台窗口**。
映射要跑约 20 秒，那个框就停留约 20 秒。

### 实测证据（2026-09-10）

用一个**同样 `--windowed` 打包**的探针 exe（自身无控制台，与真实 app 一致）
启动 PowerShell 子进程，由**子进程自己**通过 `GetConsoleWindow()` +
`IsWindowVisible()` 报告结果——绕开 stdout 重定向与代码页的干扰：

| 调用方式 | 子进程自报 |
|---|---|
| 无 `creationflags`（当前实现） | `hwnd=2364656 visible=True` |
| `creationflags=CREATE_NO_WINDOW` | `hwnd=0 visible=False` |

### 修法

```python
completed = subprocess.run(
    command, capture_output=True, text=True, encoding="utf-8", errors="replace",
    creationflags=subprocess.CREATE_NO_WINDOW,
)
```

**只有这一处**，其余逻辑一律不动。

`app.py` 靠 stdout 拿映射脚本输出的 JSON，所以必须确认加了标志后 stdout
仍能正常捕获。已实测：

```
returncode    : 0
stdout lines  : 1
json parsed   : True
rowsRead      : 1 rowsWritten: 1
output exists : True 268920
```

### 注意

`subprocess.CREATE_NO_WINDOW` 是 Windows 专有常量。本项目本来就是 Windows 专用
（用了 `ctypes.windll`），直接写即可；若要更保守可写
`getattr(subprocess, "CREATE_NO_WINDOW", 0)`。

### 改完必须重新打包并重装

用户桌面的快捷方式指向 `%LOCALAPPDATA%\Programs\WalmartShelfAssistant\` 里
**已安装的副本**，不会随源码改动自动更新：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
```

### 验证方式

改后重装并启动，点一次「开始填充」：**处理期间不应出现任何黑框**，
摘要与输出文件应与改动前一致。这一条只能用眼睛确认，没有自动化断言能替代。

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
| 低 | 9. `rowsHiddenSkipped` 计数 | 只影响日志数字，输出正确 |
| 低 | 10. 空值单元格也设对齐 | 看不出差别，不建议现在动 |
| 低 | 11. 关窗口时不收尾 | 仅影响关闭路径；但兜底代码自己是假的，值得顺手加固 |
| 中 | 12. 弹出 PowerShell 黑框 | 一行修复、用户每次都看得见，投入产出比最高 |
