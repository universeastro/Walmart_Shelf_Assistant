# AGENTS.md

本文件供在本仓库工作的 AI agent（Codex 等）自动读取。修改代码前请先读完。

## 项目

把文件 A（选品表）的商品字段，按多级表头名称填入文件 B（沃尔玛上架模板）的副本，
原始模板不被修改。

- `app.py` —— tkinter GUI，选 A/B/输出，子进程调用映射脚本
- `excel_mapper.ps1` —— 实际映射逻辑，PowerShell + **Excel COM**
- `文件/` —— A 模板、B 模板、需求文档
- `docs/VERIFICATION_REPORT.md` —— **独立验证报告，动手前先看**
- `docs/CODE_REVIEW.md` —— **代码审查，10 项潜伏问题与健壮性缺陷（含修复状态表）**
- `tests/fixtures/A_sample.xls` —— 测试夹具（**合成数据**），校验的唯一客观依据
- `.claude/skills/` —— Claude 侧的校验流程与脚本。**这里不会自动加载到你**，
  但下面「改完必须验证」一条已把该跑的脚本写全，可直接执行。

## 硬性约束

**1. 不要用 openpyxl / ExcelJS 保存 B 模板。**
需求第 15 条要求不破坏模板自带格式。openpyxl 保存会丢弃图表、条件格式等。
必须走 Excel COM 打开、写入、另存副本。

**2. 新增 `.ps1` 必须带 UTF-8 BOM。**
PowerShell 5.1 读无 BOM 的 `.ps1` 按 GBK 解码，中文字面量全乱、中文路径报错。
在 Bash 里用 heredoc 生成 `.ps1` 时最容易踩。脚本内容尽量保持纯 ASCII。

**3. 列字母不可信，表头名才是依据。**
需求文档里写的 B 表列字母与实际模板**系统性不符**（见验证报告 3.2）。
按列字母兜底（`fallback-column`）的映射需格外小心，改动前先确认。

**4. 测试夹具必须入库。**
放在 `tests/fixtures/`（`.gitignore` 已保留该路径）。不要放进被忽略的目录。
夹具是判断映射对错的唯一客观依据。

**5. 不要把 `Find-TargetColumns` 里的 `$columns = @{}` 改成 `[ordered]@{}`。**
`OrderedDictionary` 的索引器把 `int` 当**位置索引**而非键，
`$columns[$column] = ...` 会立刻抛 `ArgumentOutOfRangeException`（参数名 `index`）。
已实测踩到。

## 改完必须验证

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .claude/skills/verify-mapping/scripts/verify_mapping.ps1 `
  -SourcePath tests/fixtures/A_sample.xls `
  -TargetPath 文件/B模板.xlsx
```

退出码 `0` 全对 / `1` 有字段不一致 / `2` 源数据压根没写入。

**改动 `excel_mapper.ps1` 的映射表后，必须同步更新脚本里的 `$Pairs`**，
否则会拿旧契约校验新实现，得出假阳性。

## 两个已知的误报陷阱

**不要把条件格式规则数变少当成 bug。**
Excel 会合并条件相同的相邻列规则（48 条 → 43 条），覆盖范围完全等价。
判断依据是 **sqref 覆盖范围**是否相等，不是规则数。验证报告 2.2 节有实测数据。

**不要把 `rowsRead: 0` 当成缺陷。**
`文件/A模板.xls` 是纯表头空模板，返回 0 行是正确行为。要用带数据的文件验证。

## 待确认（需人工决策）

- A 表 **AI 列**同时映射到 Main Image URL 和 Additional Image URL 1 (+)，
  需求文档第 8 条与第 10 条自相矛盾。确认前不要擅自改动。
- A 表 **E 列**（空表头）强制走列字母兜底映射到 Variant Group ID，无名称校验。

## 当前测试缺口

- ~~公式保留完全未测~~ → **已补齐**：`tests/fixtures/B_formula.xlsx`
  （由 `tests/make_formula_fixture.ps1` 生成）含 5 条公式，实测 **5/5 存活且
  重算正确**，含 CSE 数组公式与跨表引用。见验证报告第 8 节。
  仍未覆盖：外部工作簿引用、定义名称、易失函数。
- ~~多行输入未测~~ → **已补齐**：`tests/make_multirow_fixture.ps1` +
  `tests/verify_multirow.ps1`。实测 5 行、含空行的 5 行、**372 行（真实规模，
  4464 格）全部通过**，行序与行数对应正确且无溢出。见验证报告第 9 节。
- ~~隐藏行跳过未测~~ → **已补齐**：`-HideAt` + `-RowMode`。
  `Visible`/`All` 两种模式、含空行组合、**全部行隐藏**的边界均已通过，
  372 行无性能回归。见验证报告第 10 节。
- ~~AutoFilter 筛选是否走同一条隐藏路径未测~~ → **已补齐**：实测确认
  AutoFilter 筛掉的行同样报告 `Hidden = $true`，与手工隐藏殊途同归。
  见 `tests/verify_autofilter.ps1` 与验证报告 10.9。
- 仍未覆盖：源表多行表头、源行数超过模板容量（实测上限 372 行）。
- 全部验证跑在合成样本上，未经真实业务文件验证。

## 两个必须知道的实测结论

**1. 源表中间的空行会被丢弃（压缩），不是复制成空行。**
5 行里第 3 行留空 → 读入 4 行，源第 4 行落到目标第 9 行。**源行号与目标行号会错位。**
若业务要求「A 表第 N 行 = B 表第 N 行」，当前实现不满足。

**2. 模板只有 3 个 Key Features 列，「五点4」「五点5」会被丢弃。**
实测确认 `文件/B模板.xlsx` 只有 `Key Features (+)` / `1 (+)` / `2 (+)`，
以及 `Additional Image URL (+)` / `1 (+)`。映射器会把这 8 个字段报进 `skipped`，
GUI 写进日志面板。**这是业务侧要拍板的事——不要擅自改映射去凑这些字段。**

## 一个容易被漏掉的盲区

**`tests/compare_ooxml.py` 不看对齐与样式。** 它只比 dataValidation、
conditionalFormatting 的 sqref 覆盖和公式。改动涉及**单元格格式**时，
只跑它然后说「格式保留通过」是没有依据的——必须补 `tests/verify_alignment.ps1`。

**写入的单元格会被主动设为「水平对齐：填充」**（`Set-CellValue` 里的
`HorizontalAlignment = 5`），这是 2026-09-10 起有意为之、README 已说明的行为。
实测确认它**不会**污染表头或写入区之外的单元格（564 + 1022 格逐格比对一致），
也**不影响** `Value2`，故不影响上架数据。

## 一个必须避开的坑

**不要用 `$ws.Cells.Item($r,$c).Value2 = $null` 清空单元格。**
实测该赋值会让**同一脚本之后的所有写入静默失效**（不报错、读回为空），
3/3 次确定性复现。用 `.ClearContents()` 或 `.Value2 = ''` 均正常。
原因未定位（最小复现脚本重现不出来），但现象稳定。
`excel_mapper.ps1` 第 329 行把 `$null` 转成 `''` 正是因此**必须保留**。

## 测试工具

- `tests/compare_ooxml.py` —— 比较模板与输出的 OOXML 部件、数据验证、
  条件格式覆盖范围与公式。用法 `py tests/compare_ooxml.py <模板> <输出>`。
- `tests/make_formula_fixture.ps1` —— 生成含公式的模板夹具。
- `tests/make_multirow_fixture.ps1` —— 由 `A_sample.xls` 展开成 N 行夹具，
  `-BlankAt n` 制造空行，`-HideAt '3,5'` 隐藏指定行（**先写后藏**，藏起来的行
  仍持有值，所以映射器只能靠可见性跳过它）。列按**表头名**定位。
- `tests/verify_multirow.ps1` —— 生成夹具 → 跑映射器 → 逐格比对行序与行数。
  `-Rows 372` 为真实规模；`-RowMode All|Visible` 与 `-HideAt` 测隐藏行语义。
- `tests/verify_visible_rows.ps1` —— Codex 侧的独立实现，整块
  `UsedRange.Value2` 读入二维数组比对，与上一个工具路径不同，互为交叉验证。
- `tests/verify_autofilter.ps1` —— 对夹具施加**真实 AutoFilter**（README 让用户
  做的第一步），断言筛掉的行被跳过、唯一可见行落到目标。无参数直接跑。
- `tests/verify_alignment.ps1` —— 断言写入单元格为水平对齐「填充」，
  且表头区与写入区之外与模板**逐格一致**。改 `Set-CellValue` 后必跑。
- `tests/smoke_completion_dialog.py` —— `CompletionDialog` 冒烟测试，
  用 `py` 跑。需可见桌面会话。
- `tests/diag_cells.ps1` —— 读取指定单元格的公式与计算值，排查用。
