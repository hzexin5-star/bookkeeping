#!/usr/bin/env python3
"""Read CSV / XLSX tables using only the standard library.

Bank and payment-app exports are inconsistent: CSV files may be GBK encoded,
and XLSX files need real spreadsheet parsing. Both are handled here so that the
bookkeeping skill stays dependency-free (no openpyxl / pandas required).
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "gbk", "big5", "latin-1")
DELIMITERS = ",;\t|"


def read_table(path: Path | str, sheet: str | None = None) -> list[list[str]]:
    """Read *path* into a rectangular list of string cells."""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"找不到文件: {p}")
    suffix = p.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return read_xlsx(p, sheet)
    if suffix == ".xls":
        raise ValueError("不支持旧版 .xls，请先在 Excel 中另存为 .xlsx 或 CSV")
    return read_csv(p)


def read_csv(path: Path) -> list[list[str]]:
    raw = path.read_bytes()
    text = None
    for encoding in ENCODINGS:
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"无法识别文件编码: {path}")

    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=DELIMITERS)
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ","

    reader = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter)
    return [[cell.strip() for cell in row] for row in reader]


def read_xlsx(path: Path, sheet: str | None = None) -> list[list[str]]:
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        shared = _shared_strings(archive, names)
        sheet_path = _resolve_sheet(archive, names, sheet)
        root = ET.fromstring(archive.read(sheet_path))

    grid: list[list[str]] = []
    for row in root.iter(NS_MAIN + "row"):
        cells: dict[int, str] = {}
        for cell in row.findall(NS_MAIN + "c"):
            index = _column_index(cell.get("r"))
            cells[index] = _cell_text(cell, shared).strip()
        width = max(cells) + 1 if cells else 0
        grid.append([cells.get(i, "") for i in range(width)])
    return _rectangular(grid)


def _shared_strings(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    if "xl/sharedStrings.xml" not in names:
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.iter(NS_MAIN + "t"))
        for item in root.findall(NS_MAIN + "si")
    ]


def _column_index(reference: str | None) -> int:
    match = re.match(r"([A-Za-z]+)", reference or "")
    if not match:
        return 0
    index = 0
    for char in match.group(1).upper():
        index = index * 26 + (ord(char) - 64)
    return index - 1


def _cell_text(cell: ET.Element, shared: list[str]) -> str:
    kind = cell.get("t")
    if kind == "inlineStr":
        node = cell.find(NS_MAIN + "is")
        if node is None:
            return ""
        return "".join(text.text or "" for text in node.iter(NS_MAIN + "t"))
    value = cell.find(NS_MAIN + "v")
    if value is None or value.text is None:
        return ""
    if kind == "s":
        try:
            return shared[int(value.text)]
        except (ValueError, IndexError):
            return value.text
    return value.text


def _sheet_targets(archive: zipfile.ZipFile, names: set[str]) -> list[tuple[str, str]]:
    targets: list[tuple[str, str]] = []
    workbook = "xl/workbook.xml"
    if workbook not in names:
        return targets
    relationships: dict[str, str] = {}
    rel_path = "xl/_rels/workbook.xml.rels"
    if rel_path in names:
        rel_root = ET.fromstring(archive.read(rel_path))
        for rel in rel_root:
            relationships[rel.get("Id")] = rel.get("Target") or ""
    root = ET.fromstring(archive.read(workbook))
    for sheet in root.iter(NS_MAIN + "sheet"):
        target = relationships.get(sheet.get(NS_REL + "id"), "")
        if target.startswith("/"):
            resolved = target.lstrip("/")
        elif target.startswith("xl/"):
            resolved = target
        else:
            resolved = "xl/" + target.lstrip("/")
        targets.append((sheet.get("name") or "", resolved))
    return targets


def _resolve_sheet(archive: zipfile.ZipFile, names: set[str], sheet: str | None) -> str:
    targets = [path for _, path in _sheet_targets(archive, names) if path in names]
    if not targets:
        targets = sorted(n for n in names if n.startswith("xl/worksheets/sheet"))
    if not targets:
        raise ValueError("xlsx 中找不到工作表")
    if sheet is None:
        return targets[0]
    wanted = str(sheet).strip()
    if wanted.isdigit():
        position = int(wanted)
        if 1 <= position <= len(targets):
            return targets[position - 1]
        raise ValueError(f"工作表序号超出范围: {wanted}")
    for name, path in _sheet_targets(archive, names):
        if name == wanted and path in names:
            return path
    raise ValueError(f"找不到工作表: {wanted}")


def _rectangular(grid: list[list[str]]) -> list[list[str]]:
    width = max((len(row) for row in grid), default=0)
    return [row + [""] * (width - len(row)) for row in grid]