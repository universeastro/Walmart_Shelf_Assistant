# 映射验证报告

| | |
|---|---|
| 日期 | 2026-09-10 |
| 验证者 | Claude |
| 被验证版本 | `8cfafc5`（第一轮）、`e7b1fcc`（第二轮，见第 7 节） |
| 验证方式 | 独立复跑，**未采信 Codex 自产的 `B_filled.xlsx`** |
| 复跑命令 | 见文末「如何复跑」 |

---

## 结论摘要

| 项 | 结果 |
|---|---|
| 映射正确性 | ✅ 12/12 字段落在正确列 |
| 数据验证保留 | ✅ 30/30 |
| 条件格式保留 | ✅ 覆盖等价（48 条合并为 43 条，非丢失） |
| 工作表 / table / drawing | ✅ 无部件丢失 |
| **公式保留** | ❌ **未测到——模板公式数为 0** |
| 真实业务文件 | ❌ 未测，仅跑了合成样本 |

**无阻塞性缺陷。** 待办集中在「没测到」而非「测出错」：见第 4 节。

---

## 1. 映射正确性

输入 `tests/fixtures/A_sample.xls` → 输出由 `文件/B模板.xlsx` 生成。

数据写入**第 7 行**（表头 2-4 行，第 6 行是字段描述行，数据从第 7 行开始），位置正确。逐字段核对：

| 源列 | 目标列 | 源值 | 目标值 | 判定 |
|---|---|---|---|---|
| D | D | CUST-001 | CUST-001 | ✅ |
| N | H | Sample product | Sample product | ✅ |
| O | Q | Long description | Long description | ✅ |
| P | R | Feature one | Feature one | ✅ |
| Q | S | Feature two | Feature two | ✅ |
| R | T | Feature three | Feature three | ✅ |
| AI | U | https://img.example/main... | 同左 | ✅ |
| L | Y | Red | Red | ✅ |
| AG | AD | https://img.example/one... | 同左 | ✅ |
| AI | AE | https://img.example/main... | 同左 | ✅ |
| E | BM | GROUP-001 | GROUP-001 | ✅ |
| AX | BQ | https://img.example/swatch... | 同左 | ✅ |

`excel_mapper.ps1` 输出的 `rowsRead: 1, rowsWritten: 1`。

> **注意**：对空模板 `文件/A模板.xls` 跑映射会返回 `rowsRead: 0`，这是**正确行为**
> （该文件是纯表头模板，没有数据行），不是缺陷。验证时必须用带数据的文件。

---

## 2. 格式保留（需求第 15 条）

### 2.1 无部件丢失

对比模板与输出的 OOXML 包内部结构：33 个 table、4 个工作表、drawing、theme 全部存在，
无部件丢失。

但**几乎所有部件的内容都变了**（`styles.xml`、`theme1.xml`、全部 table、所有 worksheet、
`docProps` 等）。这说明输出是让 **Excel 把整个包重新序列化**了一遍，而非定向打补丁。
结构无损，但字节级全变。

### 2.2 ⚠️ 一个必须说明的误报

初测发现条件格式「48 条 → 43 条」，一度判定为格式丢失。**深挖后确认是误报**：

Excel 把**条件相同的相邻列规则合并**了：

| 模板（分离） | 输出（合并后） |
|---|---|
| `AA7:AA10000` + `AB7:AB10000` | `AA7:AB10000` |
| `AF7:AF10000` + `AG7:AG10000` | `AF7:AG10000` |
| `AH7:AH10000` + `AI7:AI10000` | `AH7:AI10000` |
| `AJ7:AJ10000` + `AK7:AK10000` | `AJ7:AK10000` |
| `AL7:AL10000` + `AM7:AM10000` | `AL7:AM10000` |

覆盖范围完全等价（`AA7:AA10000 + AB7:AB10000` ≡ `AA7:AB10000`），**语义无变化**。

**请勿把规则数变少当成 bug 去「修复」。** 判断依据应该是 sqref 覆盖范围是否相等，
不是块数或规则数是否相等。详见第 5 节的坑位说明。

### 2.3 数据验证

30 条 `dataValidation` 与 30 条 `formula1` 全部保留，覆盖范围一致。

---

## 3. 待确认问题（需要人工决策，不是代码缺陷）

### 3.1 AI 列的归属自相矛盾

需求文档内部冲突：

- 第 8 条：A 的 AI 列（小平台专用链接2）→ B 的 **Main Image URL**
- 第 10 条：A 的 AI 列（小平台专用链接2）→ B 的 **Additional Image URL 1 (+)**

当前实现是**两处都写**（`AI → U` 和 `AI → AE`）。

**需要确认**：AI 列到底应该只进其中一个，还是确实两处都要。同一张图片同时作为主图和附加图
是否是有意为之，请业务侧确认。

### 3.2 需求文档的列字母已失效

文档第 8~12 条写明的 B 表目标列，与 `文件/B模板.xlsx` 实际布局**系统性不符**：

| 文档写的 | 实测所在列 | 偏差 |
|---|---|---|
| Main Image URL → W | U | −2 |
| Color → AA | Y | −2 |
| Additional Image URL (+) → AF | AD | −2 |
| Additional Image URL 1 (+) → AG | AE | −2 |
| Variant Group ID → BU | BM | −8 |
| Swatch Image URL → BZ | BQ | −9 |

文档对应的是一个**更宽的模板版本**。映射器按表头名匹配，结果是对的——文档第 13 条
本来就规定表头名优先，所以这不影响当前正确性。

**但它推翻了列字母的权威性。** 凡是走 `fallback-column` 的映射，其正确性都建立在
同一个已失效的坐标系上。当前有一处：

```
父SKU => variant group id   sourceColumn=E   sourceMethod=fallback-column
```

这处在 `8cfafc5` 中把名称匹配去掉了（`Names = @()`），改为**强制走 E 列兜底**。
理由是 A 模板 E 列表头为空，而 B 列是 `父SKU` 辅助列，按名称匹配会取错。

这个决定本身合理，但**它意味着该映射不再有任何名称校验**：一旦源表列序变动，
会静默写错列且不报错。建议至少加一条断言（例如校验 E 列在数据行有值时才写）。

---

## 4. 未覆盖的测试（当前最大的风险敞口）

### 4.1 公式保留——需求第 15 条的核心部分

`文件/B模板.xlsx` 和 `_docx_qa/B_sample.xlsx` 的公式数**都是 0**，因此
「公式等自带格式不能被破坏」这条**完全没有被验证**。

这是需求第 15 条里风险最高的部分：Excel 重新序列化会丢弃它不理解的内容，
而公式（尤其跨表引用、外部链接、数组公式）正是常见 casualty。

**建议**：造一个含公式的模板夹具（至少覆盖普通公式、跨表引用、数组公式各一），
纳入回归测试。这是目前最值得补的一个测试。

### 4.2 真实业务文件

现有全部验证跑在合成样本（`tests/fixtures/A_sample.xls`，一行数据）上。
真实选品表动辄数百行，且列的实际情况可能更复杂。

**建议**：上线前用一份真实 A/B 跑一遍完整验证。

---

## 5. 工程问题

### 5.1 测试夹具未纳入版本控制 —— ✅ 已修复（2026-09-10）

**原问题**：`_docx_qa/` 被 `.gitignore` 整体忽略，而 `A_sample.xls`（唯一的测试夹具）
就在里面。夹具是判断代码对错的唯一客观依据，不入库则**校验无法复现**——
`AGENTS.md` 里那条「改完必须验证」的命令，引用的是一个 clone 下来根本不存在的路径。

**已修复**：夹具移到 `tests/fixtures/A_sample.xls` 并入库，全部 7 处引用同步更新。

夹具内容经确认是**纯合成数据**（`docProps` 创建者为 `Apache POI`，值全是
`CUST-001` / `GROUP-001` / `Sample product` / `img.example` 占位符），
表头是 A 表的字段名（属 schema，非业务数据），**不含真实商品信息**。

> ⚠️ 该文件扩展名是 `.xls`，**实际内容是 xlsx**（ZIP 包）。Excel COM 能正常读取，
> 但按扩展名判断类型的工具会误判。写脚本处理它时注意这一点。

### 5.2 `.ps1` 必须带 UTF-8 BOM

PowerShell 5.1 读无 BOM 的 `.ps1` 会按 GBK 解码，中文字面量全部乱码，中文路径直接
报「路径不存在」。`8cfafc5` 已给 `excel_mapper.ps1` 加 BOM，**新增任何 `.ps1` 都要注意**。

在 Bash 里用 heredoc 生成 `.ps1` 时尤其容易踩到（本报告作者踩过两次）。

---

## 7. 第二轮验证（`e7b1fcc`，2026-09-10）

Codex 把 B 模板的写入方式改为「复制到输出目录的 GUID 临时文件 → 可写打开 →
`SaveAs` → `finally` 删除」，并补充了错误路径信息与 stdout 的 UTF-8 编码。
基线哈希在校验前后**逐字节一致**，确认校验期间 Codex 未再写入。

| 检查项 | 结果 |
|---|---|
| 映射正确性 | 12/12，数据行仍为第 7 行，退出码 0 |
| 原模板完整性 | md5 前后同为 `4ec60063`，**逐字节未变** |
| 临时文件清理 | 跑入空目录后其中仅有 `out.xlsx`，全仓库无 `~$` 与 `.walmart_shelf_assistant_*` |
| 失败后清理 | 构造 `SaveAs` 之后的失败路径，仍无残留 |
| 错误路径 JSON | 源/目标不存在时均输出合法 JSON（`errorLine` 指向 235/238 行、`errorType=ItemNotFoundException`），退出码 1 |
| 输出目录自动创建 | 多级不存在的路径可自动创建 |
| 格式保留 | 57/57 部件；数据验证 30/30；条件格式 48→43 为已知的相邻列合并误报，覆盖等价 |
| 子进程编码 | `app.py` 以 `utf-8` 解码，与 mapper 新设的 stdout UTF-8 配套 |

**一个被实测证伪的怀疑（记录下来以免后人重走）**：切换 `SaveCopyAs → SaveAs` 时
怀疑会丢格式——因为 `SaveAs` 会按扩展名推断文件格式（`.csv` 会退化成 CSV，
丢掉全部格式和多张工作表），而 `SaveCopyAs` 不会。实测把输出命名为 `out.csv`，
落盘文件头是 `PK..`（ZIP）、268KB 完整工作簿，**格式并未被转换**。
该风险不成立。但若将来换用带 `FileFormat` 参数的重载，需要重新评估。

**本轮新增发现**：`output` 字段可能指向一个不存在的路径，见
`CODE_REVIEW.md` 第 8 项。

**未覆盖**：公式保留（模板公式数仍为 0）、真实业务文件、多行输入——
三项缺口与第一轮完全相同，**本轮没有缩小任何一项**。

---

## 6. 如何复跑验证

已固化为项目 skill，可直接调用：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .claude/skills/verify-mapping/scripts/verify_mapping.ps1 `
  -SourcePath tests/fixtures/A_sample.xls `
  -TargetPath 文件/B模板.xlsx
```

退出码：`0` 全对 / `1` 有字段不一致 / `2` 源数据压根没写入。

**改动映射表后，必须同步更新该脚本里的 `$Pairs`**，否则会拿旧契约校验新实现，
得出假阳性。

详细踩坑记录见 `.claude/skills/verify-mapping/SKILL.md`。
