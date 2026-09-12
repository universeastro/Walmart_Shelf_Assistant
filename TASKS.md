# 任务清单

> 当前阶段：**03 并行功能**（分支 `feature/req03-parallel`，基线 `14ac950`）。
> T01–T22 为已关闭的 01/02 阶段记录，保留作回归依据，编号不复用。
> 03 阶段需求与验收标准见 `PROJECT.md`；C-01/C-02 已由用户裁决，C-03–C-05 已由视频直接验证，
> 仅 `.xls` 默认输出仍待用户确认（见 `HANDOFF.md`）。

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
| T15 | Codex | 已完成 | `app.py`、`excel_mapper.ps1`、`tests/test_ui_layout.py`、`tests/verify_ui_integration.py`、`tests/verify_req02.ps1`、`tests/verify_append.ps1`、`README.md`、`PROJECT.md`、`HANDOFF.md` | 已存在输出可选替换或继续写入 |
| T16 | Codex | 已完成 | `app.py`、`tests/test_ui_layout.py`、`HANDOFF.md` | T15；显示替换/继续写入结果及实际输出路径 |
| T17 | Codex | 已完成 | `app.py`、`tests/test_ui_layout.py`、`tests/verify_req02.ps1`、`HANDOFF.md` | Claude 清单：首次新建状态显示、CP936 验证兼容 |
| T18 | Codex | 已完成 | `AGENTS.md`、`HANDOFF.md`、`docs/WORKFLOW.md` | 用户确认本分支阶段完成；`文件/03/` 留待后续阶段，不纳入本分支 |
| T19 | Codex | 已完成 | `tests/verify_append.ps1`、`HANDOFF.md` | 回归脚本不得因嵌套脚本 `exit` 提前结束 |
| T20 | Claude | 已完成 | `docs/REVIEW_REQ02.md`、`HANDOFF.md` | 支线实现前历史审查及最终结论交接 |
| T21 | Codex | 已完成 | `app.py`、`tests/test_path_settings.py`、`tests/test_ui_layout.py`、`tests/verify_path_settings.py`、`PROJECT.md`、`README.md`、`TASKS.md`、`HANDOFF.md` | T11；主线/支线分别记忆导出范围 |
| T22 | Codex | 已完成 | `TASKS.md`、`HANDOFF.md`、Git 分支与远端引用 | T18/T21；阶段收口并快进合并到 `main` |
| T23 | Codex | 已完成 | `PROJECT.md`、`TASKS.md`、`HANDOFF.md`（03 阶段目标/约束/验收与分工） | 无；等待用户确认正式开始 |
| T24 | Claude | 已完成 | 只读需求、素材、视频、验收边界与接口冲突审查；结论写入 `HANDOFF.md` | T23；C-03–C-05 已有视频证据 |
| T25 | Codex | 待开始 | `excel_mapper.ps1`（Req03 profile、六字段映射、目标表/末行识别） | T24；用户确认正式开始 |
| T26 | Codex | 待开始 | `app.py`、`README.md`（第三方案页、路径与导出范围记忆、默认输出扩展名） | T25 接口稳定 |
| T27 | Codex | 待开始 | `tests/`（Req03 逐格验证与 01/02 回归路径） | T25/T26 |
| T28 | Claude | 待开始 | 独立需求/边界/回归/文档审查；结论写入 `HANDOFF.md` | T25–T27 |
| T29 | Codex | 待开始 | 整合、完整回归、`docs/WORKFLOW.md` 素材登记、提交 | T28 通过 |

## 接口约定

- 映射器新增 profile/方案选择时必须显式传入或可确定地从模板识别；不得覆盖 B01 默认契约。
- 映射摘要沿用现有 JSON 输出，新增字段需向后兼容。
- 目标表头缺失或歧义时，必须返回可读错误，禁止静默写入猜测列。
- 支线 profile 固定按表头定位 `SKU`、`平台SKU`；主线 profile 行为不变。
- **写入语义：默认 `Append`，显式 `-WriteMode Replace` 才覆盖。** 空目标从第 2 行写入；
  有既有记录时从末行 + 4 开始，真实 `文件/02/B模板.xls` 对应 R700。
  该口径对 03 同样适用（C-04：沿用 01/02 的「隔三行」定义）。

### 03 阶段新增约定

- `-Profile` 的 `ValidateSet` 扩为 `Mainline`、`Req02`、`Req03`；Req03 仅允许 `-WriteMode Append`。
- **目标工作表定位顺序**：先按名称精确匹配 `导入 单位转换`；未命中时退回表头集合判定
  （须同时含 `SKU(直接从sheet 1导入）`、`SKU价(￥)`、`重量(g)`、`长`、`宽`、`高`）；
  多命中时必须报错，禁止任选其一。目的是排除 `价格` 表——其 A1/B1 与目标表逐字符相同。
- **末行检测**：Req03 只扫描被映射的目标列（A、B、C、H、I、J），不得复用扫全 UsedRange 的
  `Get-LastRecordRow`。原因见 `PROJECT.md` 技术约束第 2 条（`Q1:Q5220` 常量会导致起点变成 5224）。
- **表头匹配必须精确**：`Normalize-Text` 后全等比较；不得做全角/半角括号转换或去符号模糊匹配。
- **字段级跳过**：单字段源或目标匹配不到时计入 `skipped` 并继续；B 的
  `SKU(直接从sheet 1导入）` 列匹配不到时整体失败。
- **值类型**：SKU 保持文本；B/C/H/I/J 五个目标列将完整可解析的数字文本写为数值，解析失败则
  保留文本。视频中的 306 个单列数值和 918 个三列尺寸均执行了“转为数字”。
- 映射摘要沿用现有 JSON 输出，新增字段需向后兼容。
- GUI 侧 `PROFILE_LABELS` 增 `"支线 03": "Req03"`；方案页、A/B/输出路径、导出范围按方案独立记忆；
  03 默认导出范围「仅可见行」（C-02）。
- 03 的输出默认扩展名沿用 B 的扩展名 `.xls`（**待用户确认**，见 `PROJECT.md` 当前状态）。

## 治理规则

- **实际改动的文件必须落在某个任务的「允许修改模块」内。** 发现越界改动时，
  要么补进对应任务，要么补建任务——不得留白。2026-09-11 曾出现
  `AGENTS.md`、`README.md`、`build.ps1`、`requirements*.txt`、`文件/` 重组
  全部无任务覆盖的情况，本表已回填（见 T04/T08/T10）。
- 新增任务时才可新增编号；不得复用已关闭的编号。
- 状态仅使用：待开始、进行中、审查中、已完成、阻塞。
- 每次状态或接口变化同步 `HANDOFF.md`。
- 状态置 `阻塞` 时，必须在本表或 `HANDOFF.md` 写明**解除条件**。
