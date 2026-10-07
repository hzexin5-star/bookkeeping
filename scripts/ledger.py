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
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

_HERE = str(Path(__file__).resolve().parent)
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)
from tabular import read_table  # stdlib-only CSV/XLSX reader

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


# ---------------------------------------------------------------- 导入账单 --

HEADER_ALIASES = {
    "date": ["日期", "交易日期", "交易时间", "记账日期", "发生日期", "入账日期", "交易发生时间", "时间", "date", "transactiondate", "datetime"],
    "amount": ["金额", "交易金额", "发生额", "交易额", "金额元", "amount", "transactionamount"],
    "income_amount": ["收入", "收入金额", "贷方发生额", "贷方金额", "存入金额", "转入金额", "credit", "creditamount"],
    "expense_amount": ["支出", "支出金额", "借方发生额", "借方金额", "支取金额", "转出金额", "debit", "debitamount"],
    "type": ["收支", "收/支", "收支类型", "交易类型", "借贷标志", "借贷方向", "收付标志", "方向", "类型", "type", "direction"],
    "category": ["分类", "交易分类", "消费分类", "交易类别", "类别", "category"],
    "account": ["账户", "账户名称", "账号", "卡号", "支付方式", "付款方式", "收付款方式", "account", "card"],
    "note": ["备注", "摘要", "交易说明", "交易摘要", "说明", "商品说明", "商品", "附言", "note", "memo", "description", "remark"],
}

FIELD_ORDER = ("date", "amount", "income_amount", "expense_amount", "type", "category", "account", "note")
FIELD_ORDER_FALLBACK = ("date", "income_amount", "expense_amount", "amount", "type", "category", "account", "note")

FIELD_EXCLUDES = {
    "account": ("余额", "balance", "可用", "额度", "性质", "状态"),
    "amount": ("余额", "balance"),
    "income_amount": ("余额", "balance"),
    "expense_amount": ("余额", "balance"),
    "date": ("有效期", "expire"),
}

_INCOME_HINTS = ("收入", "收款", "存入", "转入", "入账", "credit", "income", "refund")
_EXPENSE_HINTS = ("支出", "支取", "消费", "付款", "转出", "出账", "debit", "expense", "payment")
_CURRENCY_NOISE = ("¥", "￥", "$", "€", "£", "人民币", "元", ",")


def normalize_header(text: str) -> str:
    return re.sub(r"[\s()（）\[\]【】:：*]+", "", str(text or "")).lower()


_NORMALIZED_ALIASES = {
    normalize_header(alias)
    for aliases in HEADER_ALIASES.values()
    for alias in aliases
}


def detect_header_row(table: list[list[str]], limit: int = 20) -> int | None:
    """Guess which row holds the column headers (0-based index)."""
    best_index: int | None = None
    best_score = 0
    for index, row in enumerate(table[:limit]):
        score = sum(
            1 for cell in row if normalize_header(cell) in _NORMALIZED_ALIASES
        )
        if score > best_score:
            best_index, best_score = index, score
    return best_index if best_score >= 2 else None


def resolve_column(spec: str, header: list[str]) -> int:
    text = str(spec).strip()
    if text.isdigit():
        index = int(text) - 1
        if 0 <= index < len(header):
            return index
        raise SystemExit(f"[错误] 列序号超出范围: {spec}")
    target = normalize_header(text)
    for index, cell in enumerate(header):
        if normalize_header(cell) == target:
            return index
    raise SystemExit(f"[错误] 找不到列: {spec}（可用列: {', '.join(c for c in header if c)}）")


def parse_map_args(items: list[str] | None) -> dict:
    overrides: dict[str, str] = {}
    for item in items or []:
        for chunk in str(item).split(","):
            chunk = chunk.strip()
            if not chunk:
                continue
            if "=" not in chunk:
                raise SystemExit(f"[错误] --map 格式应为 字段=列名，收到: {chunk}")
            field, value = chunk.split("=", 1)
            field = field.strip().lower()
            if field not in HEADER_ALIASES:
                raise SystemExit(f"[错误] 未知字段 {field}，可用: {', '.join(FIELD_ORDER)}")
            overrides[field] = value.strip()
    return overrides


def build_mapping(header: list[str], overrides: dict) -> dict:
    mapping: dict[str, int] = {}
    used: set[int] = set()
    normalized = [normalize_header(cell) for cell in header]

    for field in FIELD_ORDER:
        for alias in (normalize_header(a) for a in HEADER_ALIASES[field]):
            for index, key in enumerate(normalized):
                if index in used or not key or key != alias:
                    continue
                mapping[field] = index
                used.add(index)
                break
            if field in mapping:
                break

    for field in FIELD_ORDER_FALLBACK:
        if field in mapping:
            continue
        aliases = [normalize_header(a) for a in HEADER_ALIASES[field] if len(normalize_header(a)) >= 2]
        excludes = [normalize_header(x) for x in FIELD_EXCLUDES.get(field, ())]
        for alias in aliases:
            for index, key in enumerate(normalized):
                if index in used or not key:
                    continue
                if any(bad in key for bad in excludes):
                    continue
                if alias in key:
                    mapping[field] = index
                    used.add(index)
                    break
            if field in mapping:
                break

    for field, spec in overrides.items():
        index = resolve_column(spec, header)
        mapping[field] = index
        used.add(index)
    return mapping


def cell_at(row: list[str], index: int | None) -> str:
    if index is None or index < 0 or index >= len(row):
        return ""
    return str(row[index]).strip()


def try_amount(value) -> Decimal | None:
    text = str(value or "").strip()
    if not text:
        return None
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    for token in _CURRENCY_NOISE:
        text = text.replace(token, "")
    text = text.strip()
    if text.endswith("-"):
        negative = True
        text = text[:-1].strip()
    if text.startswith("-"):
        negative = True
        text = text[1:].strip()
    elif text.startswith("+"):
        text = text[1:].strip()
    if not text:
        return None
    try:
        amount = Decimal(text)
    except InvalidOperation:
        return None
    return -amount if negative else amount


def parse_type_hint(value) -> str | None:
    text = re.sub(r"\s+", "", str(value or "")).lower()
    if not text:
        return None
    if any(hint in text for hint in _EXPENSE_HINTS):
        return "expense"
    if any(hint in text for hint in _INCOME_HINTS):
        return "income"
    if text in ("支", "出", "借", "借方", "dr", "-"):
        return "expense"
    if text in ("收", "入", "贷", "贷方", "cr", "+"):
        return "income"
    return None


def parse_any_date(value) -> str | None:
    text = str(value or "").strip()
    if not text:
        return None
    if re.fullmatch(r"\d{5}(\.\d+)?", text):
        serial = float(text)
        return (datetime(1899, 12, 30) + timedelta(days=serial)).date().isoformat()
    cleaned = text.replace("年", "-").replace("月", "-").replace("日", " ")
    cleaned = cleaned.replace("/", "-").replace(".", "-")
    cleaned = re.split(r"[ T]", cleaned.strip())[0]
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%m-%d"):
        try:
            parsed = datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
        if fmt == "%m-%d":
            parsed = parsed.replace(year=date.today().year)
        return parsed.isoformat()
    return None


def dedup_key(entry: dict) -> tuple:
    try:
        amount = Decimal(str(entry.get("amount", "0"))).normalize()
    except InvalidOperation:
        amount = Decimal(0)
    return (
        str(entry.get("date", "")),
        str(entry.get("type", "")),
        str(amount),
        str(entry.get("category", "")),
        str(entry.get("account", "")),
        str(entry.get("note", "")),
    )


def row_to_entry(row: list[str], mapping: dict, args, positive_is: str) -> dict | None:
    date_value = parse_any_date(cell_at(row, mapping.get("date")))
    if not date_value:
        return None

    entry_type = parse_type_hint(cell_at(row, mapping.get("type")))
    amount: Decimal | None = None

    income_value = try_amount(cell_at(row, mapping.get("income_amount")))
    expense_value = try_amount(cell_at(row, mapping.get("expense_amount")))
    if income_value:
        entry_type, amount = "income", abs(income_value)
    elif expense_value:
        entry_type, amount = "expense", abs(expense_value)
    else:
        parsed = try_amount(cell_at(row, mapping.get("amount")))
        if parsed is None:
            return None
        amount = abs(parsed)
        if entry_type is None:
            if parsed == 0:
                return None
            positive = positive_is == "income"
            if parsed > 0:
                entry_type = "income" if positive else "expense"
            else:
                entry_type = "expense" if positive else "income"

    if amount is None or amount == 0:
        return None

    category = (cell_at(row, mapping.get("category")) or args.category or "未分类").strip() or "未分类"
    account = (args.account or cell_at(row, mapping.get("account"))).strip()
    note = cell_at(row, mapping.get("note")).strip()
    return {
        "date": date_value,
        "type": entry_type or "expense",
        "category": category,
        "amount": str(amount),
        "account": account,
        "note": note,
    }


def print_preview(entries: list[dict], limit: int = 10) -> None:
    if not entries:
        print("  （没有可导入的新记录）")
        return
    columns = [
        ("日期", 10, "left"), ("类型", 4, "left"), ("分类", 12, "left"),
        ("金额", 12, "right"), ("账户", 10, "left"), ("备注", 0, "left"),
    ]
    print("  " + "  ".join(pad(title, width, align) if width else title for title, width, align in columns))
    for entry in entries[:limit]:
        label = "收入" if entry["type"] == "income" else "支出"
        values = [
            (entry["date"], 10, "left"),
            (label, 4, "left"),
            (entry["category"], 12, "left"),
            (fmt(row_amount(entry)), 12, "right"),
            (entry["account"], 10, "left"),
            (entry["note"], 0, "left"),
        ]
        print("  " + "  ".join(pad(value, width, align) if width else str(value) for value, width, align in values))
    if len(entries) > limit:
        print(f"  ...（共 {len(entries)} 条，仅预览前 {limit} 条）")


def describe_mapping(mapping: dict, header: list[str], args) -> str:
    override_account = getattr(args, "account", None)
    parts = []
    for field, index in mapping.items():
        if field == "account" and override_account:
            parts.append(f"account→{override_account}（命令行 --account 指定）")
            continue
        name = header[index] if index < len(header) else str(index)
        parts.append(f"{field}→{name}")
    if override_account and "account" not in mapping:
        parts.append(f"account→{override_account}（命令行 --account 指定）")
    return "、".join(parts)


def cmd_import(args) -> None:
    source = Path(args.source)
    try:
        table = read_table(source, sheet=args.sheet)
    except (OSError, ValueError, KeyError) as exc:
        raise SystemExit(f"[错误] 读取失败: {exc}")
    if not table:
        raise SystemExit("[错误] 文件里没有数据")

    overrides = parse_map_args(args.map)
    header_index = (args.header_row - 1) if args.header_row else detect_header_row(table)
    if header_index is None:
        raise SystemExit("[错误] 无法识别表头：请用 --header-row N 指定表头行，或用 --map date=列名,amount=列名 手动指定")
    if header_index >= len(table):
        raise SystemExit(f"[错误] 表头行 {header_index + 1} 超出范围（文件共 {len(table)} 行）")

    header = table[header_index]
    mapping = build_mapping(header, overrides)
    if "date" not in mapping:
        raise SystemExit(f"[错误] 未能识别日期列，请用 --map date=列名 指定（可用列: {', '.join(c for c in header if c)}）")
    if "amount" not in mapping and "income_amount" not in mapping and "expense_amount" not in mapping:
        raise SystemExit(f"[错误] 未能识别金额列，请用 --map amount=列名 指定（可用列: {', '.join(c for c in header if c)}）")

    path = Path(args.ledger)
    existing = read_rows(path)
    seen = {dedup_key(row) for row in existing}
    start_id = next_id(existing)
    positive_is = args.positive_is

    fresh: list[dict] = []
    duplicates = 0
    invalid_rows: list[int] = []
    for offset, row in enumerate(table[header_index + 1:], start=header_index + 2):
        if not any(str(cell).strip() for cell in row):
            continue
        entry = row_to_entry(row, mapping, args, positive_is)
        if entry is None:
            invalid_rows.append(offset)
            continue
        key = dedup_key(entry)
        if key in seen:
            duplicates += 1
            continue
        seen.add(key)
        entry["id"] = str(start_id + len(fresh))
        fresh.append(entry)

    detected = describe_mapping(mapping, header, args)
    if args.dry_run:
        print(f"[预览] {source.name}：将新增 {len(fresh)} 条，跳过重复 {duplicates} 条，无法解析 {len(invalid_rows)} 条")
        print(f"  列映射：{detected}")
        print_preview(fresh, limit=args.limit)
        return

    if fresh:
        write_rows(path, existing + fresh)
    print(f"[导入完成] {source.name}：新增 {len(fresh)} 条，跳过重复 {duplicates} 条，无法解析 {len(invalid_rows)} 条")
    print(f"  列映射：{detected}")
    if invalid_rows:
        shown = ", ".join(str(n) for n in invalid_rows[:20])
        suffix = " ..." if len(invalid_rows) > 20 else ""
        print(f"  无法解析的行号（原文件行号）：{shown}{suffix}")
    if fresh:
        print_preview(fresh, limit=args.limit)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="ledger.py", description="记录与统计个人收支。")
    parser.add_argument("--ledger", default=DEFAULT_LEDGER, help=f"账本 CSV 路径（默认 {DEFAULT_LEDGER}）")

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--ledger", default=argparse.SUPPRESS, help="账本 CSV 路径（也可放在子命令之后）")

    sub = parser.add_subparsers(dest="cmd", required=True)

    p_add = sub.add_parser("add", parents=[common], help="新增一笔收支")
    p_add.add_argument("--date", help="日期 YYYY-MM-DD，缺省为今天")
    p_add.add_argument("--type", required=True, help="income/expense 或 收入/支出")
    p_add.add_argument("--category", help="分类，例如 餐饮/交通/工资")
    p_add.add_argument("--amount", required=True, help="金额，正数")
    p_add.add_argument("--account", help="账户，例如 微信/支付宝/银行卡/现金")
    p_add.add_argument("--note", help="备注")
    p_add.set_defaults(func=cmd_add)

    p_imp = sub.add_parser("import", parents=[common], help="从 CSV / XLSX 账单批量导入（自动去重）")
    p_imp.add_argument("source", help="账单文件路径（.csv / .xlsx）")
    p_imp.add_argument("--sheet", help="xlsx 工作表名称或序号（从 1 开始），缺省为第一个")
    p_imp.add_argument("--header-row", type=int, help="表头所在行号（从 1 开始），缺省自动识别")
    p_imp.add_argument("--map", action="append", help="指定列映射，如 --map date=交易时间 --map amount=金额；可重复、可逗号分隔")
    p_imp.add_argument("--account", help="为本次导入的所有记录指定账户（覆盖文件中的账户列）")
    p_imp.add_argument("--category", help="文件没有分类列时使用的默认分类")
    p_imp.add_argument("--positive-is", choices=("income", "expense"), default="income", help="没有收支标志列时，正数金额算收入还是支出（默认 income）")
    p_imp.add_argument("--dry-run", action="store_true", help="只预览，不写入账本")
    p_imp.add_argument("--limit", type=int, default=10, help="预览条数（默认 10）")
    p_imp.set_defaults(func=cmd_import)

    p_list = sub.add_parser("list", parents=[common], help="列出明细")
    add_filters(p_list)
    p_list.add_argument("--json", action="store_true", help="以 JSON 输出")
    p_list.set_defaults(func=cmd_list)

    p_sum = sub.add_parser("summary", parents=[common], help="按区间/分类汇总")
    add_filters(p_sum)
    p_sum.add_argument("--json", action="store_true", help="以 JSON 输出")
    p_sum.set_defaults(func=cmd_summary)

    p_bal = sub.add_parser("balance", parents=[common], help="按账户统计结余")
    add_filters(p_bal)
    p_bal.add_argument("--json", action="store_true", help="以 JSON 输出")
    p_bal.set_defaults(func=cmd_balance)

    p_del = sub.add_parser("delete", parents=[common], help="按编号删除一笔记录")
    p_del.add_argument("id", help="记录编号")
    p_del.set_defaults(func=cmd_delete)

    p_upd = sub.add_parser("update", parents=[common], help="按编号修改一笔记录")
    p_upd.add_argument("id", help="记录编号")
    p_upd.add_argument("--date")
    p_upd.add_argument("--type")
    p_upd.add_argument("--category")
    p_upd.add_argument("--amount")
    p_upd.add_argument("--account")
    p_upd.add_argument("--note")
    p_upd.set_defaults(func=cmd_update)

    p_rep = sub.add_parser("report", parents=[common], help="生成 Markdown 报表")
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