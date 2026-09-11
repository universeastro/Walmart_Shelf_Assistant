# 任务清单

| 编号 | 负责人 | 状态 | 允许修改模块 | 依赖 |
|---|---|---|---|---|
| T01 | Codex | 已完成 | `PROJECT.md`、`TASKS.md`、`HANDOFF.md` | 无 |
| T02 | Claude | 已完成 | 需求/验收审查；未经协调不改实现文件 | 先读 T01 文档 |
| T03 | Codex | 已完成 | `excel_mapper.ps1`、映射配置/参数 | C-01 已按默认追加生效 |
| T04 | Codex | 已完成 | `app.py`、`README.md`、`build.ps1`、`requirements.txt`、`requirements-build.txt` | T03 接口稳定 |
| T05 | Codex | 已完成 | `tests/`（新增支线验证与回归路径） | T03/T04 |
| T06 | Claude | 已完成 | 审查结果与 `HANDOFF.md` | T03-T05；最终审查通过 |
| T07 | Codex | 已完成 | 整合、回归验证、提交 | T06 最终审查通过 |
| T08 | Claude | 已完成 | `文件/` 素材基线入库、`.gitignore` | 用户重组已完成 |
| T09 | Claude | 已完成 | `docs/WORKFLOW.md`（流程补齐） | 无 |
| T10 | Codex | 已完成 | `AGENTS.md`、`README.md` 文档同步 | T07 |
| T11 | Codex | 已完成 | `app.py`、`tests/test_path_settings.py`、`tests/test_ui_layout.py`、`tests/verify_path_settings.py`、`tests/verify_ui_integration.py`、`README.md` | 双页面与分方案路径记忆 |
| T12 | Claude | 已完成 | 需求、边界、回归与文档审查；审查结论追加 `HANDOFF.md` | T11 |
| T13 | Codex | 已完成 | 最终整合、验证与提交 | T12 通过 |
| T14 | Codex | 已完成 | `excel_mapper.ps1`、`tests/verify_req02.ps1`、`PROJECT.md`、`README.md`、`HANDOFF.md` | 支线兼容源表头“自定义SKU” |

## 接口约定

- 映射器新增 profile/方案选择时必须显式传入或可确定地从模板识别；不得覆盖 B01 默认契约。
- 映射摘要沿用现有 JSON 输出，新增字段需向后兼容。
- 目标表头缺失或歧义时，必须返回可读错误，禁止静默写入猜测列。
- 支线 profile 固定按表头定位 `SKU`、`平台SKU`；主线 profile 行为不变。
- **写入语义：默认 `Append`，显式 `-WriteMode Replace` 才覆盖。** 空目标从第 2 行写入；
  有既有记录时从末行 + 4 开始，真实 `文件/02/B模板.xls` 对应 R700。

## 治理规则

- **实际改动的文件必须落在某个任务的「允许修改模块」内。** 发现越界改动时，
  要么补进对应任务，要么补建任务——不得留白。2026-09-11 曾出现
  `AGENTS.md`、`README.md`、`build.ps1`、`requirements*.txt`、`文件/` 重组
  全部无任务覆盖的情况，本表已回填（见 T04/T08/T10）。
- 新增任务时才可新增编号；不得复用已关闭的编号。
- 状态仅使用：待开始、进行中、审查中、已完成、阻塞。
- 每次状态或接口变化同步 `HANDOFF.md`。
- 状态置 `阻塞` 时，必须在本表或 `HANDOFF.md` 写明**解除条件**。
