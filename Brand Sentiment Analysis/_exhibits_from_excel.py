"""Regenerate OnCloud exhibits from OnCloud_C2C_Brand_Exploratory.xlsx."""
from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

BASE = Path(__file__).resolve().parent
XLSX = BASE / "OnCloud_C2C_Brand_Exploratory.xlsx"
OUT = BASE / "OnCloud_Exhibits"
OUT.mkdir(exist_ok=True)

wb = openpyxl.load_workbook(XLSX, data_only=True)

# Associations top 10
ws = wb["Key_Associations"]
assoc = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        continue
    if not row or not row[1]:
        continue
    assoc.append((str(row[1]), int(row[2])))
assoc = assoc[:10]

fig, ax = plt.subplots(figsize=(10, 5.5))
labels = [a[0] for a in assoc][::-1]
vals = [a[1] for a in assoc][::-1]
ax.barh(labels, vals, color="#1F4E79")
ax.set_xlabel("Frequency (On-relevant records with theme)")
ax.set_title("On Cloud C2C Association Frequencies (Excel mine)")
fig.tight_layout()
fig.savefig(OUT / "Exhibit01_Association_Frequency.png", dpi=160)
plt.close()

# Valence from association sheet aggregates — use Executive or compute from Key_Associations pos/neg
# Build simple stacked from top themes' pos/neg/mixed
pos = sum(int(r[5] or 0) for i, r in enumerate(ws.iter_rows(values_only=True)) if i and r[1])
neg = sum(int(r[6] or 0) for i, r in enumerate(ws.iter_rows(values_only=True)) if i and r[1])
mix = sum(int(r[7] or 0) for i, r in enumerate(ws.iter_rows(values_only=True)) if i and r[1])
# Note: these sum theme-level valence tags (not mutually exclusive across themes)
fig, ax = plt.subplots(figsize=(7, 5))
cats = ["positive", "mixed/neutral", "negative"]
vals = [pos, mix, neg]
ax.bar(cats, vals, color=["#2E7D32", "#90A4AE", "#C62828"])
ax.set_ylabel("Theme-tagged valence counts (summed across association themes)")
ax.set_title("C2C Valence Heuristic by Association Themes (Excel)")
fig.tight_layout()
fig.savefig(OUT / "Exhibit02_Valence_Breakdown.png", dpi=160)
plt.close()

# Competitors
ws = wb["Competitor_Comparisons"]
comps = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0 or not row or not row[0]:
        continue
    if str(row[0]).startswith("ANALYSIS") or row[0] == "":
        continue
    try:
        comps.append((str(row[0]), int(row[1])))
    except (TypeError, ValueError):
        continue
comps = comps[:8]
fig, ax = plt.subplots(figsize=(9, 5))
ax.barh([c[0] for c in comps][::-1], [c[1] for c in comps][::-1], color="#5B2C6F")
ax.set_xlabel("Co-mentions in On-relevant posts/comments")
ax.set_title("Competitor Co-Mentions (Excel C2C mine)")
fig.tight_layout()
fig.savefig(OUT / "Exhibit03_Competitor_CoMentions.png", dpi=160)
plt.close()

# Curated quotes from Quote_Bank (On-primary IDs)
keep_ids = {"Q01", "Q05", "Q06", "Q09", "Q11", "Q12", "Q14", "Q18", "Q19", "Q21"}
ws = wb["Quote_Bank"]
quotes = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        continue
    if row and row[0] in keep_ids:
        quotes.append(
            {
                "id": row[0],
                "quote": row[1],
                "valence": row[2],
                "score": row[3],
                "subreddit": row[5],
                "source_file": row[4],
                "permalink": row[8],
            }
        )

out_wb = Workbook()
qws = out_wb.active
qws.title = "Quote_Exhibit"
qws.append(["Exhibit_ID", "Subreddit", "Score", "Valence", "Quote", "Source_File", "Permalink"])
for c in qws[1]:
    c.font = Font(bold=True)
for q in quotes:
    qws.append([q["id"], q["subreddit"], q["score"], q["valence"], q["quote"], q["source_file"], q["permalink"]])
    qws.cell(row=qws.max_row, column=5).alignment = Alignment(wrap_text=True)
qws.column_dimensions["E"].width = 90

aws = out_wb.create_sheet("Association_Counts")
aws.append(["rank", "association_theme", "frequency_n"])
for i, (name, n) in enumerate(assoc, 1):
    aws.append([i, name, n])

cws = out_wb.create_sheet("Competitor_CoMentions")
cws.append(["competitor", "co_mention_n"])
for name, n in comps:
    cws.append([name, n])

mws = out_wb.create_sheet("Sample_Meta")
mws.append(["Metric", "Value"])
mws.append(["source", "OnCloud_C2C_Brand_Exploratory.xlsx"])
mws.append(["records_scanned", 298854])
mws.append(["on_relevant_hits", 12492])
mws.append(["hit_rate_pct", 4.18])

out_wb.save(OUT / "Exhibit04_Quote_and_Counts.xlsx")

meta = {
    "source": "OnCloud_C2C_Brand_Exploratory.xlsx",
    "associations_top10": assoc,
    "competitors_top": comps,
    "quotes_used": [q["id"] for q in quotes],
    "valence_theme_sum": {"positive": pos, "mixed_neutral": mix, "negative": neg},
}
(BASE / "_c2c_excel_exhibit_meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")
print("exhibits refreshed from Excel", meta["quotes_used"])
