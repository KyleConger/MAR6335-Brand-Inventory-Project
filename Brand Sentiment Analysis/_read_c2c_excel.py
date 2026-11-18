"""Dump OnCloud_C2C_Brand_Exploratory.xlsx to JSON for upgrading drafts."""
from __future__ import annotations

import json
from pathlib import Path

import openpyxl

p = Path(__file__).resolve().parent / "OnCloud_C2C_Brand_Exploratory.xlsx"
out = Path(__file__).resolve().parent / "_c2c_excel_dump.json"
wb = openpyxl.load_workbook(p, data_only=True)
dump = {}
for name in wb.sheetnames:
    ws = wb[name]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append([("" if c is None else c) for c in row])
    dump[name] = rows
out.write_text(json.dumps(dump, ensure_ascii=False, indent=2), encoding="utf-8")
print("sheets", list(dump))
for k, rows in dump.items():
    print(k, "nrows", len(rows))
