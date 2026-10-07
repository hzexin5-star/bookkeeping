# bookkeeping · 账目统计

一个用于 [Codex](https://chatgpt.com/codex) 的个人记账 skill：录入收支，按分类 / 账户 / 时间区间汇总，查询结余，并生成 Markdown 报表。

## 特性

- **零依赖**：纯 Python 3 标准库，无需 `pip install`。
- **数据可迁移**：账本是一个 UTF-8（带 BOM）CSV，Excel、Numbers、脚本都能直接打开，中文不乱码。
- **多维统计**：按月份或任意区间，按分类、账户汇总收入与支出。
- **一键报表**：生成 Markdown 格式的账目报表，便于贴进笔记或继续加工。
- **本地优先**：所有数据保存在你自己的机器上，不联网、不上传。

## 安装

把本仓库克隆到 Codex 的技能目录，**目录名保持 `bookkeeping`**：

```bash
git clone https://github.com/<you>/bookkeeping.git ~/.codex/skills/bookkeeping
```

Windows PowerShell：

```powershell
git clone https://github.com/<you>/bookkeeping.git "$env:USERPROFILE\.codex\skills\bookkeeping"
```

也可以下载 ZIP 解压到 `~/.codex/skills/bookkeeping`。之后 Codex 会以 `$bookkeeping` 列出该 skill。

## 使用

在 Codex 里直接用自然语言，例如：

- “记一笔，今天午饭 35.5，微信付的”
- “统计一下这个月的收支”
- “看看我各账户的结余”
- “导出一份 10 月的账目报表”

也可以直接调用脚本：

```bash
python scripts/ledger.py add --type expense --category 餐饮 --amount 35.5 --account 微信 --note 午饭
python scripts/ledger.py summary --month 2026-10
python scripts/ledger.py report --month 2026-10 --out 账目报表-2026-10.md
```

## 命令

| 命令 | 说明 |
| --- | --- |
| `add` | 新增一笔收支 |
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
├── scripts/
│   └── ledger.py         # 记账 / 统计 CLI
└── references/
    └── schema.md         # 字段与分类约定
```

## 发布到 GitHub

```bash
git init
git add .
git commit -m "feat: bookkeeping skill"
git branch -M main
git remote add origin https://github.com/<you>/bookkeeping.git
git push -u origin main
```

发布前记得把 `LICENSE` 里的版权署名换成你自己的名字。

## 许可

[MIT](LICENSE)