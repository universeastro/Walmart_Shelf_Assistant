# 追加写入变更记录

## 行为

- B 商品数据区最后一个含固定值的行号为 L 时，新批次从 L+4 开始；空模板沿用正常数据起始行。
- 检查所有已使用列，包含隐藏行、零值及未映射列；不回填已有记录间空隙。
- 仅有格式或公式的行不认作商品记录。无法自动区分公式生成的业务记录与预置公式；若商品记录完全由公式构成，应先人工确认，不将此类文件视为已验证适用。
- 写入前检查整个新批次的映射区域。存在公式、合并单元格或超过工作表行数时停止，不保存输出。
- 不插入整行、不清空间隔区域，不改变固定列映射和源行筛选顺序。间隔区域原有公式保留，因此“三行空行”指不写入本批商品数据，不保证原公式单元格视觉为空。
- JSON 新增 existingLastRow 和 writeStartRow，APP 处理记录显示追加位置。

## 验证状态

- PowerShell 语法检查通过，两个脚本均为 UTF-8 BOM。
- git diff --check 通过。
- 新增 tests/verify_append.ps1：创建临时 B 副本，覆盖旧记录间空隙、隐藏旧记录、末行未映射列为零、远处格式和公式、两种导出范围、三行间隔、原 B 哈希及公式冲突阻止写入。
- Excel 专项测试与原映射回归启动后未返回结果，已中止；沙箱外重试被审批服务 503 阻止。上述运行测试均未判定通过。
- 未修改用户业务工作簿，未更新桌面安装版。须先完成 Excel 验证，再打包安装。

## 待执行

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/verify_append.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/verify-mapping/scripts/verify_mapping.ps1 -SourcePath tests/fixtures/A_sample.xls -TargetPath 文件/B模板.xlsx
```
