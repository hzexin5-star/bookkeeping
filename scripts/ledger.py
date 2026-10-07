#!/usr/bin/env python3
"""Ledger CLI - record and summarize personal income/expenses.

Adapted for a CSV book that is safe to keep in git (UTF-8 with BOM so that
Excel on Windows opens Chinese text correctly).

Examples:
    python ledger.py add --type expense --category 餐饮 --amount 35.5 --account 微信 --note 午饭
    python ledger.py add --date 2026-10-01 --type income --category 工资 --amount 12000 --account 银行卡
    python ledger.py list --month 2026-10
    python ledger.py summary --month 2026-10
    python ledger.py balance
    python ledger.py report --month 2026-10 --out report.md
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

FIELDS = ["id", "date", "type", "category", "amount", "account", "note"]
DEFAULT_LEDGER = "ledger.csv"

TYPE_ALIASES = {
    "income": "income", "in": "income", "收入": "income", "收": "income", "+": "income",
    "expense": "expense", "out": "expense", "支出": "expense", "支": "expense",
    "花费": "expense", "消费": "expense", "-": "expense",
}


def normalize_type(value: str) -> str:
    key = (value or "").strip().lower()
    if key in TYPE_ALIASES:
        return TYPE_ALIASES[key]
    raise SystemExit(f"[错误] 无法识别类型: {value!r}（请用 income/expense 或 收入/支出）")


def parse_date(value: str | None) -> str:
    if not value or value.strip().lower() in {"today", "今天"}:
        return date.today().isoformat()
    if value.strip().lower() in {"yesterday", "昨天"}:
        return (date.today() - timedelta(days=1)).isoformat()
    raw = value.strip()
    today = date.today()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d", "%m-%d", "%m/%d"):
        try:
            parsed = datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
        if fmt in ("%m-%d", "%m/%d"):
            parsed = parsed.replace(year=today.year)
        return parsed.isoformat()
    raise SystemExit(f"[错误] 无法识别日期: {value!r}（请用 YYYY-MM-DD）")


def parse_amount(value) -> Decimal:
    cleaned = str(value).replace(",", "").replace("¥", "").replace("￥", "").strip()
    try:
        amount = Decimal(cleaned)
    except InvalidOperation:
        raise SystemExit(f"[错误] 无法识别金额: {value!r}")
    if amount == 0:
        raise SystemExit("[错误] 金额不能为 0")
    return abs(amount)


def fmt(amount: Decimal) -> str:
    return f"{amount:,.2f}"


def money_str(amount: Decimal) -> str:
    """Plain 2-decimal string, safe for machine-readable output."""
    return f"{amount:.2f}"


def cell_width(text: str) -> int:
    import unicodedata
    width = 0
    for ch in str(text):
        width += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return width


def pad(text: str, width: int, align: str = "left") -> str:
    text = "" if text is None else str(text)
    fill = " " * max(0, width - cell_width(text))
    return fill + text if align == "right" else text + fill


def row_amount(row: dict) -> Decimal:
    try:
        return Decimal(str(row.get("amount", "0") or "0"))
    except InvalidOperation:
        return Decimal(0)


def read_rows(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return [row for row in reader]


def write_rows(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDS)
        writer.writeheader()
        for row in rows:
            writer.writerow({key: row.get(key, "") for key in FIELDS})


def next_id(rows: list[dict]) -> int:
    ids = []
    for row in rows:
        try:
            ids.append(int(row.get("id", "0")))
        except (TypeError, ValueError):
            continue
    return max(ids, default=0) + 1


def filter_rows(rows: list[dict], args) -> list[dict]:
    result = []
    month = getattr(args, "month", None)
    from_date = getattr(args, "from_date", None)
    to_date = getattr(args, "to_date", None)
    rtype = getattr(args, "type", None)
    category = getattr(args, "category", None)
    account = getattr(args, "account", None)
    for row in rows:
        day = (row.get("date") or "")[:10]
        if month and not day.startswith(month):
            continue
        if from_date and day < from_date:
            continue
        if to_date and day > to_date:
            continue
        if rtype and row.get("type") != rtype:
            continue
        if category and row.get("category") != category:
            continue
        if account and row.get("account") != account:
            continue
        result.append(row)
    return result


def add_filters(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--month", help="按月份筛选，格式 YYYY-MM")
    parser.add_argument("--from", dest="from_date", help="起始日期 YYYY-MM-DD（含）")
    parser.add_argument("--to", dest="to_date", help="结束日期 YYYY-MM-DD（含）")
    parser.add_argument("--type", type=normalize_type, help="income 或 expense")
    parser.add_argument("--category", help="按分类精确筛选")
    parser.add_argument("--account", help="按账户精确筛选")


def totals(rows: list[dict]) -> tuple[Decimal, Decimal, Decimal]:
    income = sum((row_amount(r) for r in rows if r.get("type") == "income"), Decimal(0))
    expense = sum((row_amount(r) for r in rows if r.get("type") == "expense"), Decimal(0))
    return income, expense, income - expense


def group_sum(rows: list[dict], key: str, rtype: str) -> list[tuple[str, Decimal]]:
    buckets: dict[str, Decimal] = {}
    for row in rows:
        if row.get("type") != rtype:
            continue
        name = (row.get(key) or "(未填写)").strip() or "(未填写)"
        buckets[name] = buckets.get(name, Decimal(0)) + row_amount(row)
    return sorted(buckets.items(), key=lambda kv: kv[1], reverse=True)


def period_label(args) -> str:
    if getattr(args, "month", None):
        return args.month
    if getattr(args, "from_date", None) or getattr(args, "to_date", None):
        return f"{getattr(args, 'from_date', None) or '开始'} ~ {getattr(args, 'to_date', None) or '至今'}"
    return "全部"


def cmd_add(args) -> None:
    path = Path(args.ledger)
    rows = read_rows(path)
    entry = {
        "id": str(next_id(rows)),
        "date": parse_date(args.date),
        "type": normalize_type(args.type),
        "category": (args.category or "未分类").strip() or "未分类",
        "amount": str(parse_amount(args.amount)),
        "account": (args.account or "").strip(),
        "note": (args.note or "").strip(),
    }
    rows.append(entry)
    write_rows(path, rows)
    label = "收入" if entry["type"] == "income" else "支出"
    print(f"[已记录] #{entry['id']} {entry['date']} {label} {entry['category']} {fmt(row_amount(entry))} {entry['account']} {entry['note']}".rstrip())


def cmd_list(args) -> None:
    rows = filter_rows(read_rows(Path(args.ledger)), args)
    rows.sort(key=lambda r: (r.get("date", ""), r.get("id", "")))
    if args.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
        return
    if not rows:
        print("（没有符合条件的记录）")
        return
    income, expense, net = totals(rows)
    cols = [
        ("ID", 4, "right"), ("日期", 10, "left"), ("类型", 4, "left"),
        ("分类", 12, "left"), ("金额", 12, "right"), ("账户", 10, "left"), ("备注", 0, "left"),
    ]
    print("  ".join(pad(h, w, a) if w else h for h, w, a in cols))
    print("-" * 78)
    for row in rows:
        label = "收入" if row.get("type") == "income" else "支出"
        values = [
            (row.get("id", ""), 4, "right"),
            (row.get("date", ""), 10, "left"),
            (label, 4, "left"),
            (row.get("category", ""), 12, "left"),
            (fmt(row_amount(row)), 12, "right"),
            (row.get("account", ""), 10, "left"),
            (row.get("note", ""), 0, "left"),
        ]
        print("  ".join(pad(v, w, a) if w else str(v) for v, w, a in values))
    print("-" * 78)
    print(f"共 {len(rows)} 笔 | 收入 {fmt(income)} | 支出 {fmt(expense)} | 结余 {fmt(net)}")


def cmd_summary(args) -> None:
    rows = filter_rows(read_rows(Path(args.ledger)), args)
    income, expense, net = totals(rows)
    expense_by_category = group_sum(rows, "category", "expense")
    income_by_category = group_sum(rows, "category", "income")
    by_account = group_sum(rows, "account", "expense")
    if args.json:
        print(json.dumps({
            "period": period_label(args),
            "count": len(rows),
            "income": money_str(income),
            "expense": money_str(expense),
            "net": money_str(net),
            "expense_by_category": [[k, money_str(v)] for k, v in expense_by_category],
            "income_by_category": [[k, money_str(v)] for k, v in income_by_category],
            "expense_by_account": [[k, money_str(v)] for k, v in by_account],
        }, ensure_ascii=False, indent=2))
        return
    print(f"统计区间: {period_label(args)}   （共 {len(rows)} 笔）")
    print(f"总收入: {fmt(income)}")
    print(f"总支出: {fmt(expense)}")
    print(f"净结余: {fmt(net)}")
    print()
    print("[支出分类]")
    if not expense_by_category:
        print("  （无）")
    for name, value in expense_by_category:
        pct = (value / expense * 100) if expense else Decimal(0)
        print(f"  {name:<12} {fmt(value):>12}  {pct:5.1f}%")
    print()
    print("[收入分类]")
    if not income_by_category:
        print("  （无）")
    for name, value in income_by_category:
        pct = (value / income * 100) if income else Decimal(0)
        print(f"  {name:<12} {fmt(value):>12}  {pct:5.1f}%")
    if by_account:
        print()
        print("[支出账户]")
        for name, value in by_account:
            pct = (value / expense * 100) if expense else Decimal(0)
            print(f"  {name:<12} {fmt(value):>12}  {pct:5.1f}%")


def cmd_balance(args) -> None:
    rows = filter_rows(read_rows(Path(args.ledger)), args)
    accounts: dict[str, Decimal] = {}
    for row in rows:
        name = (row.get("account") or "(未填写)").strip() or "(未填写)"
        delta = row_amount(row) if row.get("type") == "income" else -row_amount(row)
        accounts[name] = accounts.get(name, Decimal(0)) + delta
    if args.json:
        print(json.dumps({k: money_str(v) for k, v in accounts.items()}, ensure_ascii=False, indent=2))
        return
    if not accounts:
        print("（没有符合条件的记录）")
        return
    print(f"账户结余（{period_label(args)}）")
    print("-" * 34)
    for name, value in sorted(accounts.items(), key=lambda kv: kv[1], reverse=True):
        print(f"  {name:<14} {fmt(value):>14}")
    total = sum(accounts.values(), Decimal(0))
    print("-" * 34)
    print(f"  {'合计':<14} {fmt(total):>14}")


def find_row(rows: list[dict], target: str) -> dict:
    for row in rows:
        if str(row.get("id", "")).strip() == str(target).strip():
            return row
    raise SystemExit(f"[错误] 找不到编号为 {target} 的记录")


def cmd_delete(args) -> None:
    path = Path(args.ledger)
    rows = read_rows(path)
    target = find_row(rows, args.id)
    rows = [r for r in rows if r is not target]
    write_rows(path, rows)
    print(f"[已删除] #{target.get('id')} {target.get('date')} {target.get('category')} {fmt(row_amount(target))}")


def cmd_update(args) -> None:
    path = Path(args.ledger)
    rows = read_rows(path)
    target = find_row(rows, args.id)
    if args.date is not None:
        target["date"] = parse_date(args.date)
    if args.type is not None:
        target["type"] = normalize_type(args.type)
    if args.category is not None:
        target["category"] = args.category
    if args.amount is not None:
        target["amount"] = str(parse_amount(args.amount))
    if args.account is not None:
        target["account"] = args.account
    if args.note is not None:
        target["note"] = args.note
    write_rows(path, rows)
    print(f"[已更新] #{target.get('id')} {target.get('date')} {target.get('category')} {fmt(row_amount(target))}")


def cmd_report(args) -> None:
    rows = filter_rows(read_rows(Path(args.ledger)), args)
    rows.sort(key=lambda r: (r.get("date", ""), r.get("id", "")))
    income, expense, net = totals(rows)
    label = period_label(args)
    lines: list[str] = []
    lines.append(f"# 账目报表 {label}")
    lines.append("")
    lines.append(f"- 记录笔数：{len(rows)}")
    lines.append(f"- 总收入：{fmt(income)}")
    lines.append(f"- 总支出：{fmt(expense)}")
    lines.append(f"- 净结余：{fmt(net)}")
    lines.append("")
    lines.append("## 支出分类")
    lines.append("")
    lines.append("| 分类 | 金额 | 占比 |")
    lines.append("| --- | ---: | ---: |")
    for name, value in group_sum(rows, "category", "expense"):
        pct = (value / expense * 100) if expense else Decimal(0)
        lines.append(f"| {name} | {fmt(value)} | {pct:.1f}% |")
    lines.append("")
    lines.append("## 收入分类")
    lines.append("")
    lines.append("| 分类 | 金额 | 占比 |")
    lines.append("| --- | ---: | ---: |")
    for name, value in group_sum(rows, "category", "income"):
        pct = (value / income * 100) if income else Decimal(0)
        lines.append(f"| {name} | {fmt(value)} | {pct:.1f}% |")
    lines.append("")
    lines.append("## 明细")
    lines.append("")
    lines.append("| ID | 日期 | 类型 | 分类 | 金额 | 账户 | 备注 |")
    lines.append("| ---: | --- | --- | --- | ---: | --- | --- |")
    for row in rows:
        label_type = "收入" if row.get("type") == "income" else "支出"
        lines.append(
            f"| {row.get('id', '')} | {row.get('date', '')} | {label_type} | "
            f"{row.get('category', '')} | {fmt(row_amount(row))} | {row.get('account', '')} | {row.get('note', '')} |"
        )
    text = "\n".join(lines) + "\n"
    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text, encoding="utf-8")
        print(f"[已生成报表] {out}")
    else:
        print(text, end="")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="ledger.py", description="记录与统计个人收支。")
    parser.add_argument("--ledger", default=DEFAULT_LEDGER, help=f"账本 CSV 路径（默认 {DEFAULT_LEDGER}）")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_add = sub.add_parser("add", help="新增一笔收支")
    p_add.add_argument("--date", help="日期 YYYY-MM-DD，缺省为今天")
    p_add.add_argument("--type", required=True, help="income/expense 或 收入/支出")
    p_add.add_argument("--category", help="分类，例如 餐饮/交通/工资")
    p_add.add_argument("--amount", required=True, help="金额，正数")
    p_add.add_argument("--account", help="账户，例如 微信/支付宝/银行卡/现金")
    p_add.add_argument("--note", help="备注")
    p_add.set_defaults(func=cmd_add)

    p_list = sub.add_parser("list", help="列出明细")
    add_filters(p_list)
    p_list.add_argument("--json", action="store_true", help="以 JSON 输出")
    p_list.set_defaults(func=cmd_list)

    p_sum = sub.add_parser("summary", help="按区间/分类汇总")
    add_filters(p_sum)
    p_sum.add_argument("--json", action="store_true", help="以 JSON 输出")
    p_sum.set_defaults(func=cmd_summary)

    p_bal = sub.add_parser("balance", help="按账户统计结余")
    add_filters(p_bal)
    p_bal.add_argument("--json", action="store_true", help="以 JSON 输出")
    p_bal.set_defaults(func=cmd_balance)

    p_del = sub.add_parser("delete", help="按编号删除一笔记录")
    p_del.add_argument("id", help="记录编号")
    p_del.set_defaults(func=cmd_delete)

    p_upd = sub.add_parser("update", help="按编号修改一笔记录")
    p_upd.add_argument("id", help="记录编号")
    p_upd.add_argument("--date")
    p_upd.add_argument("--type")
    p_upd.add_argument("--category")
    p_upd.add_argument("--amount")
    p_upd.add_argument("--account")
    p_upd.add_argument("--note")
    p_upd.set_defaults(func=cmd_update)

    p_rep = sub.add_parser("report", help="生成 Markdown 报表")
    add_filters(p_rep)
    p_rep.add_argument("--out", help="输出文件路径，缺省打印到终端")
    p_rep.set_defaults(func=cmd_report)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())