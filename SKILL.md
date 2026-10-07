---
name: bookkeeping
description: 记录和统计个人或家庭账目——录入或批量导入收入/支出（CSV、XLSX 账单），按分类、账户、时间区间汇总，查询余额并生成 Markdown 报表。当用户要记账、导入银行或支付账单、统计收支、分析消费结构、核对账户余额或导出账目报表时使用。不适用于企业复式记账、报税或需要审计合规的正式财务核算。
---

# 账目统计

把用户口述的收支或现成的账单文件，落成一份可持续追加的账本，并随时产出汇总与报表。

## 账本文件

- 账本是 UTF-8（带 BOM）的 CSV，默认路径为当前工作目录下的 `ledger.csv`。用户指定了路径、或工作目录里已经有账本时，沿用同一份，不要每次新建。
- 列固定为 `id,date,type,category,amount,account,note`：`type` 取 `income`/`expense`，`amount` 始终是正数，收支方向由 `type` 决定。
- 文件不存在时脚本会自动建好表头。**不要手工拼接或改写 CSV**，一律通过脚本读写。
- 默认币种为人民币（元）。用户提到其他币种时在报表中标注，并且不要把不同币种直接相加。

字段取值、分类命名等约定见 [references/schema.md](references/schema.md)。

## 工作流程

1. **解析**：把用户给的每笔收支整理成 `日期 / 类型 / 分类 / 金额 / 账户 / 备注`。日期缺省为今天，账户缺省留空，分类沿用用户原话——不要自行合并语义不同的项（例如“餐饮”与“买菜”保持区分）。
2. **写入**：口述的单笔用 `add` 追加；现成的账单文件用 `import` 批量导入。
3. **统计**：用 `summary`、`balance`、`list`、`report` 生成结果，**直接引用脚本输出**，不要用脑算结果覆盖它。
4. **汇报**：用中文给出关键数字——收入、支出、结余、占比最高的分类；需要明细时给表格。

## 导入账单

支持 `.csv` 与 `.xlsx`（旧版 `.xls` 请先在 Excel 中另存）。脚本会自动完成：

- 识别文件编码（UTF-8 / GBK 等）与分隔符；
- 跳过银行导出的前置说明行，自动定位表头；
- 按列名识别日期、金额、收支、分类、账户、备注，兼容「收入/支出分列」和「带正负号的金额」两类常见格式；
- 按 `日期 + 类型 + 金额 + 分类 + 账户 + 备注` 去重，同一份账单重复导入不会产生重复记录。

自动识别不准时，用 `--map 字段=列名` 手动指定（支持列名或 1 起始的列序号）。没有收支标志列时用 `--positive-is expense` 把正数金额当作支出。

**导入是批量写入，先跑 `--dry-run` 让用户确认**，再执行正式导入。

## 命令速查

```bash
python scripts/ledger.py add --type expense --category 餐饮 --amount 35.5 --account 微信 --note 午饭

python scripts/ledger.py import 账单.csv --account 微信 --dry-run
python scripts/ledger.py import 账单.xlsx --sheet 微信支付 --account 微信
python scripts/ledger.py import 账单.csv --map date=交易时间,amount=金额 --positive-is expense

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
- `import` 批量导入前先预览；导入后把关键数字（新增/跳过重复/无法解析条数）回报给用户，并提示未分类或识别异常的记录。
- `update` / `delete` 属于修改既有数据，执行前先确认目标记录——用 `list` 让用户核对编号。
- 报表默认输出 Markdown；用户要 Excel 时，CSV 账本本身可直接用 Excel 打开。
- 这是轻量个人记账，不替代会计、报税或需要审计的正式财务核算。