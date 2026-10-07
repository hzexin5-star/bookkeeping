---
name: bookkeeping
description: 记录和统计个人或家庭账目——录入收入/支出，按分类、账户、时间区间汇总，查询余额并生成 Markdown 报表。当用户要记账、统计收支、分析消费结构、核对账户余额或导出账目报表时使用。不适用于企业复式记账、报税或需要审计合规的正式财务核算。
---

# 账目统计

把用户口述或文本形式的收支记录，落成一份可持续追加的账本，并随时产出汇总与报表。

## 账本文件

- 账本是 UTF-8（带 BOM）的 CSV，默认路径为当前工作目录下的 `ledger.csv`。用户指定了路径、或工作目录里已经有账本时，沿用同一份，不要每次新建。
- 列固定为 `id,date,type,category,amount,account,note`：`type` 取 `income`/`expense`，`amount` 始终是正数，收支方向由 `type` 决定。
- 文件不存在时脚本会自动建好表头。**不要手工拼接或改写 CSV**，一律通过脚本读写。
- 默认币种为人民币（元）。用户提到其他币种时在报表中标注，并且不要把不同币种直接相加。

字段取值、分类命名等约定见 [references/schema.md](references/schema.md)。

## 工作流程

1. **解析**：把用户给的每笔收支整理成 `日期 / 类型 / 分类 / 金额 / 账户 / 备注`。日期缺省为今天，账户缺省留空，分类沿用用户原话——不要自行合并语义不同的项（例如“餐饮”与“买菜”保持区分）。
2. **写入**：用 `scripts/ledger.py add` 逐笔追加。
3. **统计**：用 `summary`、`balance`、`list`、`report` 生成结果，**直接引用脚本输出**，不要用脑算结果覆盖它。
4. **汇报**：用中文给出关键数字——收入、支出、结余、占比最高的分类；需要明细时给表格。

## 命令速查

```bash
python scripts/ledger.py add --type expense --category 餐饮 --amount 35.5 --account 微信 --note 午饭
python scripts/ledger.py list --month 2026-10
python scripts/ledger.py summary --month 2026-10
python scripts/ledger.py balance
python scripts/ledger.py report --month 2026-10 --out 账目报表-2026-10.md
python scripts/ledger.py update 3 --amount 15 --note 地铁
python scripts/ledger.py delete 4
```

`list` / `summary` / `balance` / `report` 通用筛选参数：`--month YYYY-MM`、`--from`、`--to`、`--type`、`--category`、`--account`；加 `--json` 得到机器可读输出。

## 约定

- 只记录用户明确给出的账目，不臆造金额、分类或日期。用户表述含糊（例如“这个月花了不少”）时，先问清楚，不要补全为具体数字。
- `update` / `delete` 属于修改既有数据，执行前先确认目标记录——用 `list` 让用户核对编号。
- 报表默认输出 Markdown；用户要 Excel 时，CSV 账本本身可直接用 Excel 打开。
- 这是轻量个人记账，不替代会计、报税或需要审计的正式财务核算。