# bookkeeping · 账目统计

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![Codex Skill](https://img.shields.io/badge/Codex-Skill-16A34A.svg)](SKILL.md)
[![Pro](https://img.shields.io/badge/Pro-%245%2Fmonth-ff69b4.svg)](PRICING.md)

一个用于 [Codex](https://chatgpt.com/codex) 的个人记账 skill：录入或批量导入收支，按分类 / 账户 / 时间区间汇总，查询结余，并生成 Markdown 报表。

## 特性

- **零依赖**：纯 Python 3 标准库，无需 `pip install`。XLSX 解析也是自己实现的，不依赖 openpyxl / pandas。
- **批量导入账单**：读 `.csv` / `.xlsx`，自动识别编码（UTF-8 / GBK）、表头位置与列含义，兼容「收入支出分列」和「带正负号金额」两类账单。
- **自动去重**：按 `日期 + 类型 + 金额 + 分类 + 账户 + 备注` 判重，同一份账单重复导入不会产生重复记录。
- **数据可迁移**：账本是一个 UTF-8（带 BOM）CSV，Excel、Numbers、脚本都能直接打开，中文不乱码。
- **多维统计**：按月份或任意区间，按分类、账户汇总收入与支出。
- **一键报表**：生成 Markdown 格式的账目报表，便于贴进笔记或继续加工。
- **本地优先**：所有数据保存在你自己的机器上，不联网、不上传。

## 定价

| 版本 | 价格 | 内容 |
| --- | --- | --- |
| **Free**（本仓库） | 免费 · MIT | 记账、账单导入、统计汇总、账户结余、Markdown 报表 |
| **Pro** | **$5 / 月** | 以上全部 ＋ 可视化 HTML 仪表盘、预算与超支提醒、优先支持 |

Pro 订阅：<https://ko-fi.com/hzexin5-star> ｜ 详见 [PRICING.md](PRICING.md)

> 只想记账的话，**免费版就够了**——功能完整、永久免费。

## 安装

把本仓库克隆到 Codex 的技能目录，**目录名保持 `bookkeeping`**：

```bash
git clone https://github.com/hzexin5-star/bookkeeping.git ~/.codex/skills/bookkeeping
```

Windows PowerShell：

```powershell
git clone https://github.com/hzexin5-star/bookkeeping.git "$env:USERPROFILE\.codex\skills\bookkeeping"
```

也可以下载 ZIP 解压到 `~/.codex/skills/bookkeeping`。之后 Codex 会以 `$bookkeeping` 列出该 skill。

## 快速试用

仓库自带两份示例账单（一份微信账单、一份银行流水），可以直接跑通全流程：

```bash
python scripts/ledger.py import examples/wechat-bill-sample.csv --account 微信 --dry-run
python scripts/ledger.py import examples/wechat-bill-sample.csv --account 微信
python scripts/ledger.py import examples/bank-bill-sample.csv
python scripts/ledger.py summary --month 2026-10
python scripts/ledger.py balance
```

账本默认写在当前目录的 `ledger.csv`，该文件已在 `.gitignore` 中排除，不会被提交。

## 使用

在 Codex 里直接用自然语言，例如：

- “记一笔，今天午饭 35.5，微信付的”
- “把这个月的微信账单导入进去”
- “统计一下这个月的收支”
- “看看我各账户的结余”
- “导出一份 10 月的账目报表”

也可以直接调用脚本：

```bash
python scripts/ledger.py add --type expense --category 餐饮 --amount 35.5 --account 微信 --note 午饭
python scripts/ledger.py import 账单.xlsx --sheet 微信支付 --account 微信 --dry-run
python scripts/ledger.py summary --month 2026-10
python scripts/ledger.py report --month 2026-10 --out 账目报表-2026-10.md
```

## 命令

| 命令 | 说明 |
| --- | --- |
| `add` | 新增一笔收支 |
| `import` | 从 CSV / XLSX 账单批量导入（自动识别列、自动去重） |
| `list` | 列出明细，可按区间 / 类型 / 分类 / 账户筛选 |
| `summary` | 按区间汇总总收入、总支出、净结余与分类占比 |
| `balance` | 按账户统计净结余 |
| `report` | 生成 Markdown 报表（可输出到文件） |
| `update <id>` | 修改指定编号的记录 |
| `delete <id>` | 删除指定编号的记录 |

`list` / `summary` / `balance` / `report` 通用筛选参数：

- `--month YYYY-MM`：按月份
- `--from` / `--to`：任意日期区间（含端点）
- `--type` / `--category` / `--account`：按类型、分类、账户筛选
- `--json`：以 JSON 输出，便于二次处理

用 `--ledger <path>` 指定账本文件，默认是当前目录下的 `ledger.csv`。

## 导入账单

```bash
# 先预览，确认列映射和金额方向无误
python scripts/ledger.py import 账单.xlsx --account 微信 --dry-run

# 确认无误后正式导入
python scripts/ledger.py import 账单.xlsx --account 微信
```

导入相关参数：

| 参数 | 说明 |
| --- | --- |
| `--sheet` | xlsx 工作表名称或序号（从 1 开始），默认第一个 |
| `--header-row N` | 表头所在行号（从 1 开始），默认自动识别 |
| `--map 字段=列名` | 手动指定列映射；可重复、可逗号分隔，值也支持 1 起始的列序号 |
| `--account` | 为本次导入的所有记录指定账户（覆盖文件里的账户列） |
| `--category` | 文件没有分类列时使用的默认分类 |
| `--positive-is` | 没有收支标志列时，正数金额算 `income` 还是 `expense`（默认 `income`） |
| `--dry-run` | 只预览、不写入 |
| `--limit` | 预览条数，默认 10 |

可映射字段：`date`、`amount`、`income_amount`、`expense_amount`、`type`、`category`、`account`、`note`。

支持的常见表头（不区分大小写与空格、括号）：

- 日期：`日期`、`交易日期`、`交易时间`、`记账日期`、`发生日期`、`date`
- 金额：`金额`、`交易金额`、`发生额`、`amount`
- 收支分列：`收入` / `支出`、`贷方发生额` / `借方发生额`
- 收支标志：`收支`、`收/支`、`收支类型`、`交易类型`、`借贷标志`
- 其它：`分类`、`账户`、`支付方式`、`摘要`、`备注`、`商品说明`

## 账本格式

| 列 | 含义 | 约定 |
| --- | --- | --- |
| `id` | 记录编号 | 脚本自增，用于修改 / 删除 |
| `date` | 日期 | `YYYY-MM-DD` |
| `type` | 方向 | `income` 或 `expense` |
| `category` | 分类 | 自由文本，例如 餐饮 / 交通 / 工资 |
| `amount` | 金额 | 正数，方向由 `type` 决定 |
| `account` | 账户 | 现金 / 微信 / 支付宝 / 银行卡…… |
| `note` | 备注 | 自由文本 |

```csv
id,date,type,category,amount,account,note
1,2026-10-01,income,工资,12000,银行卡,
2,2026-10-07,expense,餐饮,35.5,微信,午饭
```

## 仓库结构

```text
bookkeeping/
├── SKILL.md              # skill 主体指令
├── README.md             # 本文件
├── LICENSE               # MIT
├── agents/
│   └── openai.yaml       # UI 元数据与调用策略
├── examples/
│   ├── wechat-bill-sample.csv   # 微信账单示例
│   └── bank-bill-sample.csv     # 银行流水示例
├── scripts/
│   ├── ledger.py         # 记账 / 统计 / 导入 CLI
│   └── tabular.py        # 零依赖 CSV / XLSX 读取（标准库实现）
└── references/
    └── schema.md         # 字段与分类约定
```

## 发布到 GitHub

```bash
git init
git add .
git commit -m "feat: bookkeeping skill"
git branch -M main
git remote add origin https://github.com/hzexin5-star/bookkeeping.git
git push -u origin main
```

发布前记得把 `LICENSE` 里的版权署名换成你自己的名字。

## 许可

[MIT](LICENSE)