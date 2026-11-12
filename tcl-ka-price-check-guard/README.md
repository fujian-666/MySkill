# tcl-ka-price-check-skill

TCL 华南大区 KA 价差核对的**规则固化层 + 交付门禁**：把多轮返工踩过的坑做成可执行断言，交付前跑不过就不许出。

> 定调：**这就是简单的匹配价格的事，有就有，没有就没有，事实胜于雄辩。**
> 本仓库不含任何业务数据、价格表原件、OA 号、审批人姓名、客户名称或本机绝对路径。
> 基线数值与本机路径放在 `config/profile.json`（gitignore）。

## 它解决什么

| 曾经反复出错 | 现在的挡法 |
|---|---|
| 型号 `85Z11L-A` / `65Z9H-JN` 精确匹配漏配，假报「无此型号」 | `model_key.py` 统一键 + 分级去后缀 + 等级机判定，自带自测 |
| 只查外部表就下「没有」结论，内嵌战区表其实有价 | 四路匹配链写死，每版全量回归 |
| 不同口径的数硬相减，凭空造出几百条差异 | 表头批注必读 + 口径不串 |
| 三个价格用一个简称抹平 | 第 10 项「三列不串词」+ 门禁 G5 |
| 校验只查「两边一致」，没查「标一致的行是否真一致」 | 一致性 10 项（含差额 =0 / ≠0 / 零基准豁免） |
| 沿用前一天版本号 | 门禁 G1 + `--next` 自动取当天下一版名 |
| 内部取数出处写进业务视图 | 业务视图只留待办标注，结论说明单独发负责人 |
| 为让校验变绿而放宽判据 | 先坐实 I=0/S=0 才豁免，第 8 项闭环 |

完整复盘见 [`references/POSTMORTEM.md`](references/POSTMORTEM.md)。

## 目录

```
.
├── SKILL.md                         # WorkBuddy Skill 入口（触发词、铁律、每版动作、脚本速查）
├── README.md
├── config/
│   └── profile.example.json         # 四元组配置模板（基线为 0，防呆）
├── scripts/
│   ├── profile.py                   # 配置加载（KA_PROFILE > profile.json > example）
│   ├── model_key.py                 # 型号归一化 + selftest
│   ├── identify_table.py            # 表型识别 + 月份交叉校验（只读）
│   ├── check_tz_cs_consistency.py   # 制表↔情况表 10 项一致性（只读）
│   ├── delivery_gate.py             # 21 项交付门禁 + --next（只读）
│   ├── backfill_reason_to_tz.py     # 业务原因回填（COM 写入，另存新版本）
│   └── style_check_sheet.py         # 核对情况表标准样式（COM，幂等）
└── references/
    ├── RULES.md                     # 现行口径全文
    ├── WORKFLOW.md                  # 五步执行层 / 情况表 / 金额链 / 回填
    ├── POSTMORTEM.md                # 版本链根因复盘
    └── EXCEL_COM_PITFALLS.md        # Excel/COM/openpyxl 硬坑
```

## 安装

**作为 WorkBuddy Skill（推荐）**

把整个目录复制到用户级技能目录：

```
~/.workbuddy/skills/tcl-ka-price-check-guard/
```

Windows 示例：

```
C:\Users\<你>\.workbuddy\skills\tcl-ka-price-check-guard\
```

**作为独立工具集**

```bash
pip install openpyxl
# 仅 backfill / style 需要 Excel COM：
pip install pywin32     # Windows + 已装 Excel
```

## 配置（换表必做）

```bash
cp config/profile.example.json config/profile.json
```

编辑 `config/profile.json`：

- `quad` / `owner` —— 四元组标签与交付命名里的负责人
- `sheets` / `header_row` / `first_row` / `last_row` / `cols` —— 按表头名核对后的真实列位，**不写死、不套别的表**
- `price_root` + `source_paths` —— 外部权威源表（大区表 / 总部表），按月登记
- `baseline` —— 当前交付件实测的 sheet 数 / 图片数 / 错误值数 / 真公式数 / 行数 / 金额链
- `word_counts` / `items` / `oa_evidence_rows` —— 三列专属词行数、情况表条目数、OA 证据行数

模板里基线全为 `0`，**不填就跑门禁必然 FAIL** —— 这是有意的防呆，逼你按当前四元组实测。

优先顺序：`load(path)` > 环境变量 `KA_PROFILE` > `config/profile.json` > `config/profile.example.json`。
同时维护多张表时，一个表一个 profile，用 `KA_PROFILE` 切换。

## 快速开始

```bash
# 0. 归一化自测（PASS 才可参与匹配）
python scripts/model_key.py

# 1. 表型识别，冲突即停
python scripts/identify_table.py "<xlsx>"

# 2. 每版交付前：命名建议
python scripts/delivery_gate.py --next "<任一现有xlsx>"

# 3. 十项一致性
python scripts/check_tz_cs_consistency.py "<xlsx>"

# 4. 二十一项门禁（不给 --prev 则跳过全簿 diff 段）
python scripts/delivery_gate.py "<xlsx>" \
    --prev "<上一版xlsx>" --expect-diff 86 --plan "<plan.json>"

# 5. 业务填完情况表 G/H 后回填（默认预览，--apply 才落盘）
python scripts/backfill_reason_to_tz.py --file "<源>" --out "<新版本>" --apply
```

退出码：`0` = 全过允许交付，`非0` = 有 FAIL，**禁止交付**。

## 每版固定动作

```
identify_table → openpyxl 只读建 plan（断言）→ COM 逐格写+逐格读回
→ 另存新文件（不覆盖）→ 独立 verify 全簿 diff → delivery_gate 全过
→ 记忆/日志 → 交付
```

## 设计原则

1. 规则不只是文字，必须落成脚本里的断言。
2. 只读工具不依赖 Excel，任何机器都能跑门禁。
3. 四元组隔离：所有表相关常量外置到 profile，脚本逻辑通用。
4. 只改该改的格：数值、公式、金额链零改动由全簿 diff 证明。
5. 业务视图最小化：内部出处不进表，结论单独发负责人。

## 注意

- 本仓库记录的是**方法论与校验逻辑**，不含可复现的业务数据；`config/profile.json`、xlsx、日志一律不入库。
- 涉及企业内部口径，默认**私有仓库**；如需公开，请先自行复核脱敏。
- 执行层（插列/刷值/抓明细）不在本仓库，见 `references/WORKFLOW.md` 的流程说明。
