---
name: verify-mapping
description: 验证 excel_mapper.ps1 是否把文件 A 的每个字段写进了文件 B 的正确列和正确行。当需要测试、调试、审查映射逻辑，或改动 excel_mapper.ps1 后确认没有回归时使用。
---

# 验证 A→B 字段映射

`excel_mapper.ps1` 的映射正确性**不能靠读代码判断**。表头匹配、合并单元格展开、
叶子里表头定位这几处，逻辑看着合理但结果可能是错的。「看着对」和「落对列」是两回事——
串一列，到沃尔玛后台就是批量事故。

## 用法

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .claude/skills/verify-mapping/scripts/verify_mapping.ps1 `
  -SourcePath <带数据的文件A> -TargetPath <模板B>
```

脚本会先跑一遍映射，再逐字段核对。省略 `-OutputPath` 时输出到临时目录。

- 退出码 `0` = 全部一致
- 退出码 `1` = 有字段不一致
- 退出码 `2` = 源数据压根没写进输出

## 判定依据

脚本内置了 12 组映射对（D→D、N→H、O→Q、P→R、Q→S、R→T、AI→U、L→Y、
AG→AD、AI→AE、E→BM、AX→BQ）。**映射表改动后必须同步更新这里的 `$Pairs`**，
否则脚本会拿旧契约去校验新实现，得出假阳性。

## 已知的坑（每条都实际踩过）

**1. `.ps1` 必须带 UTF-8 BOM。**
PowerShell 5.1 读无 BOM 的 `.ps1` 会按 GBK 解码，中文字面量全乱，
中文路径直接报「路径不存在」。这与 `excel_mapper.ps1` 加 BOM 是同一个原因。
本 skill 的脚本刻意写成纯 ASCII 来规避，改动时不要引入中文。

**2. 工作表文件编号会变，必须按名称对齐。**
Excel 重新保存后会重排 `xl/worksheets/sheetN.xml`（实测 `sheet2.xml.rels`
变成 `sheet3.xml.rels`）。直接对比两个包的 `sheet3.xml` 是在比两个不同的工作表，
会得出「条件格式丢了」的错误结论。要经过 `xl/workbook.xml` + `workbook.xml.rels`
解析出「工作表名 → 文件」再比对。

**3. 条件格式数量变少通常不是丢失。**
Excel 会把**条件相同的相邻列规则合并**。实测模板 48 条 → 输出 43 条，
但 `AA7:AA10000 + AB7:AB10000` 被合并成 `AA7:AB10000`，覆盖范围完全等价。
报「格式丢失」前，先比对 sqref 覆盖范围是否相等。

**4. 数据行不能靠「第一个非空单元格」找。**
模板 2-4 行是分级表头（**第 4 行是叶子表头**），第 6 行是字段描述行，
数据从**第 7 行**开始。脚本改为**搜索已知探针值**（源表 D2 的值）所在的格，
再以该行核对。

**5. 合并单元格的 `.Text` 只有锚点有值。**
其余格返回空字符串，会被误判成「字段没写进去」。必须走 `MergeArea`。

**6. `rowsRead: 0` 往往不是 bug。**
`文件/A模板.xls` 是纯表头空模板，没有数据行，返回 0 行是正确行为。
验证前先用带数据的文件（即 `tests/fixtures/A_sample.xls`）。

## 未覆盖的部分

- **公式保留**：当前模板公式数为 0，这条需求没有被测到。要验证第 15 条
  的公式部分，需要一个**含公式的模板**做夹具。
- **真实业务文件**：现有验证跑在合成样本上。上线前应拿一份真实的 A/B 跑一遍。

## 映射器输出字段

`excel_mapper.ps1` 向 stdout 输出单行 JSON：
`success` / `rowsRead` / `rowsWritten` / `targetSheet` /
`mappings[]`（含 `sourceColumn`、`targetColumn`、`sourceMethod`、`targetField`）/
`skipped[]`（未匹配到的目标字段）。`sourceMethod` 为 `header` 或 `fallback-column`，
**出现 `fallback-column` 的地方值得人工复核**——它意味着该字段没有按名称匹配上，
而是按硬编码列字母写入，源表结构一变就会静默写错。
