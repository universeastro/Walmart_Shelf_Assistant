# AGENTS.md

本文件供在本仓库工作的 AI agent（Codex 等）自动读取。修改代码前请先读完。

## 项目

把文件 A（选品表）的商品字段，按多级表头名称填入文件 B（沃尔玛上架模板）的副本，
原始模板不被修改。

- `app.py` —— tkinter GUI，选 A/B/输出，子进程调用映射脚本
- `excel_mapper.ps1` —— 实际映射逻辑，PowerShell + **Excel COM**
- `文件/` —— A 模板、B 模板、需求文档
- `docs/VERIFICATION_REPORT.md` —— **独立验证报告，动手前先看**
- `docs/CODE_REVIEW.md` —— **代码审查，8 项潜伏问题与健壮性缺陷（含修复状态表）**
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

- **公式保留完全未测**：现有模板公式数为 0。需要一个含公式的夹具
  （普通公式、跨表引用、数组公式各一）。这是需求 15 里风险最高的部分。
- 全部验证跑在合成样本上，未经真实业务文件验证。
