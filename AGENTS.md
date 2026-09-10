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
- `docs/STRESS_TEST.md` —— **3000 行规模与进程稳定性压测**。含逐阶段耗时实测、
  一个已验证的免费提速（关 `ScreenUpdating`/`EnableEvents`，−25%，产物逐格一致）、
  以及耗时随机器负载波动 3 倍的数据
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
- ~~界面从没真跑过一次映射~~ → **已补齐**：`tests/verify_ui_integration.py`
  真起 Excel 走完整流程，12.5 秒跑通，并用「把 MAPPER 指向不存在的脚本」
  证伪过（确实会失败，不是空转）。见验证报告第 13 节。
- 仍未覆盖：源表多行表头、源行数超过模板容量（实测上限 372 行）。
- 仍未覆盖：高 DPI 缩放（125%/150%）下的界面布局。
- 全部验证跑在合成样本上，未经真实业务文件验证。

## 两个必须知道的实测结论

**1. 源表中间的空行会被丢弃（压缩），不是复制成空行。**
5 行里第 3 行留空 → 读入 4 行，源第 4 行落到目标第 9 行。**源行号与目标行号会错位。**
若业务要求「A 表第 N 行 = B 表第 N 行」，当前实现不满足。

**2. 模板只有 3 个 Key Features 列，「五点4」「五点5」会被丢弃。**
实测确认 `文件/B模板.xlsx` 只有 `Key Features (+)` / `1 (+)` / `2 (+)`，
以及 `Additional Image URL (+)` / `1 (+)`。映射器会把这 8 个字段报进 `skipped`，
GUI 写进日志面板。**这是业务侧要拍板的事——不要擅自改映射去凑这些字段。**

## 用户配置（2026-09-10 起）

`app.py` 会把三个路径存到
`%LOCALAPPDATA%\WalmartShelfAssistant\settings.json`，`destroy()` 与
`run_mapping()` 开头各保存一次，**打开又关闭但不改动任何路径时不写盘**。

两条给改动者：

1. **`ShelfAssistant(settings_path=...)` 是可注入的，测试必须用它。**
   无参构造会读写**用户真实的配置**——`tests/smoke_completion_dialog.py`
   已经因为无参构造踩过一次，记得保持注入临时路径。
2. **`_load_paths` 里的 `if isinstance(value, str):` 是承重的，不要删。**
   实测：绕过它之后，配置里一个非字符串值会让程序在 `__init__` 期间抛
   `AttributeError`，**窗口根本不出现**。该 `try` 只包住读文件与 `json.loads`，
   逐项赋值不在其中，所以类型检查必须留在原处。

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

## 另一个必须避开的坑：逗号比减号结合得更紧

**二维数组下标里的减法必须加括号。**

写批量赋值时很自然会写成：

```powershell
$block[$r - 1, $c - 1] = $value      # ❌
```

PowerShell 里**逗号运算符的优先级高于 `-`**，上式被解析成
`$r - (1, $c) - 1`，于是拿一个整数去减 `Object[]`，报出

```
方法调用失败，因为 [System.Object[]] 不包含名为"op_Subtraction"的方法。
```

这个报错**完全指不到真正的出错行**，极易误判。正确写法：

```powershell
$block[($r - 1), ($c - 1)] = $value  # ✅
```

实测代价：在压测脚手架里花了两轮排查，先误判成「`.ps1` 少了 UTF-8 BOM」
（本文件上面那条约束），改完纯 ASCII 后**报错一字不变**，才定位到优先级。

## 打包（已实测，2026-09-10）

入口是 `build.ps1`，产出 `dist\WalmartShelfAssistant\`（onedir，约 26 MB）。
**`.spec` 不入库**（已加进 `.gitignore`），参数写在 `build.ps1` 里以免漂移。

三条实测结论，改打包参数前先看：

1. **`app.py` 不需要任何冻结适配。** 实测冻结后
   `Path(__file__).resolve().parent` == `_internal`，而
   `--add-data "excel_mapper.ps1;."` 正好把脚本放进 `_internal\`，
   `MAPPER` 自然解析正确。**不要**去加 `sys._MEIPASS` 判断，那是多余的。
2. **目标电脑不需要装 Python。** 实测运行中的 exe 加载的是
   `_internal\python314.dll` / `tcl90.dll`，与系统 Python 无关。
3. **目标电脑必须装 Microsoft Excel。** 硬门槛，`excel_mapper.ps1` 走
   `New-Object -ComObject Excel.Application`，打包绕不过去。见 README。

`build.ps1` 末尾会断言 `_internal\excel_mapper.ps1` 存在——少了它程序能启动、
点「开始填充」才报错，属于最难排查的一类失败。

`install.ps1` 把 `dist\` **复制**到 `%LOCALAPPDATA%\Programs\WalmartShelfAssistant\`
并在桌面建快捷方式（免管理员权限）。**刻意不用链接指向 `dist\`**：
`build.ps1` 每次删除重建 `dist\`，链接会让快捷方式在下次打包后失效。
`dist\` 不存在时 `install.ps1` 会自动先跑 `build.ps1`。

## 写 Tk 多线程测试必须用 mainloop()，不能用 update() 忙等

**只有当主线程阻塞在 `mainloop()` 里**（`_tkinter` 的 `dispatching` 已置位）、
跨线程的 `self.after(0, cb)` 才会被排队投递。主线程若只是循环调
`window.update()` 泵事件，工作线程里的 `self.after(...)` 会直接抛
`RuntimeError: main thread is not in main loop`，**回调永远送不到**。

实测代价：按 `update()` 忙等写的端到端测试，**超时 180 秒**，看起来像
「映射器挂死」；换成真实 `mainloop()` 后 12.5 秒跑完。
`app.py` 本身没问题——**是测试脚手架的问题**。

正确写法见 `tests/verify_ui_integration.py` 的 `_mainloop_until()`：
`after(0, poll)` 轮询 + 到点或到超时 `quit()`，在 `mainloop()` 里等。

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
- `tests/verify_ui_integration.py` —— **GUI 与映射器的真实集成验证**（Claude 侧）。
  `tests/test_ui_layout.py` 把 `threading.Thread` patch 成 Mock，于是拼命令行、
  起子进程、解析 JSON、回填摘要那段**一行都没跑过**；这里真起 Excel 跑一遍，
  断言输出文件生成、摘要数字与映射器返回一致、控件恢复。另含「校验失败不留
  转圈」「重复点击被忽略」「状态栏随路径刷新」三例。约 16 秒，需 Excel。
- `tests/test_path_settings.py` —— 路径持久化的单元测试（Codex 侧）。
- `tests/verify_path_settings.py` —— 同一功能的**独立**边界验证（Claude 侧，
  与上一个角度不同）。两者都用
  `py -m unittest tests.<模块名> -v` 跑。

> 本仓库**没有装 pytest**，用 `unittest`。
- `tests/diag_cells.ps1` —— 读取指定单元格的公式与计算值，排查用。
