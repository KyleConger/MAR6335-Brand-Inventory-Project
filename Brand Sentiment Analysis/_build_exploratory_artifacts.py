from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt
import re

base = Path(r"C:\Users\Owner\Desktop\UF MBA Y2 Projects\MAR6335 Brand Inventory Project")
bsa = base / "Brand Sentiment Analysis"

c2c = bsa / "OnCloud_C2C_Brand_Exploratory.xlsx"
c2c_note = (
    "Aligned to OnCloud_C2C_Brand_Exploratory.xlsx Key_Associations / Survey_Implications. "
    "E1-E4 are brand-IMAGE candidates — NOT company Brand Identity Matrix associations. "
    "Official A1-A5 still missing — see OnCloud_IDENTITY_ASSOCIATIONS_STATUS.md."
)
emergent = [
    ("E1", "Durability / longevity", "C2C Key_Associations rank 5 (n=793)"),
    ("E2", "Fit / sizing / narrow", "C2C Key_Associations rank 2 (n=1237)"),
    ("E3", "Everyday / lifestyle wear", "C2C Key_Associations rank 4 (n=950)"),
    ("E4", "Premium / expensive / price", "C2C Key_Associations rank 8 (n=504)"),
]

if c2c.exists():
    try:
        from openpyxl import load_workbook

        wb2 = load_workbook(c2c, read_only=True, data_only=True)
        sheets = wb2.sheetnames
        c2c_note = f"C2C Excel found. Sheets: {sheets}. Skim and replace E1-E4 after reviewing Key_Associations / Survey_Implications."
        themes = []
        for sn in sheets:
            if re.search(r"assoc|survey|key|theme|implication", sn, re.I):
                ws = wb2[sn]
                for i, row in enumerate(ws.iter_rows(values_only=True)):
                    if i > 40:
                        break
                    vals = [str(c) for c in row if c is not None]
                    if vals:
                        themes.append(f"{sn}: " + " | ".join(vals)[:120])
        wb2.close()
        if themes:
            (bsa / "_c2c_align_notes_temp.txt").write_text(
                c2c_note + "\n\n" + "\n".join(themes[:40]), encoding="utf-8"
            )
    except Exception as e:
        c2c_note = f"C2C Excel exists but could not skim ({e}); placeholders retained."

assoc = [
    (
        "A1",
        'Soft / cushioned "cloud" ride',
        "Product attribute / customer benefit",
        "+++",
        "+",
        "++",
        "PROVISIONAL — verify vs Brand Inventory",
    ),
    (
        "A2",
        "Lightweight, energetic performance",
        "Product attribute / customer benefit",
        "++",
        "+",
        "+",
        "PROVISIONAL — verify vs Brand Inventory",
    ),
    (
        "A3",
        "Swiss-engineered technical innovation",
        "Country/geographic + intangible",
        "++",
        "+",
        "++",
        "PROVISIONAL — verify vs Brand Inventory",
    ),
    (
        "A4",
        "Stylish / design-forward athletic aesthetic",
        "Lifestyle/personality + attribute",
        "++",
        "+",
        "+",
        "PROVISIONAL — verify vs Brand Inventory",
    ),
    (
        "A5",
        "Premium / elevated running lifestyle",
        "Relative price + lifestyle",
        "++",
        "+",
        "+",
        "PROVISIONAL — verify vs Brand Inventory",
    ),
]

thin = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
header_fill = PatternFill("solid", fgColor="1F4E79")
header_font = Font(bold=True, color="FFFFFF")
prov_fill = PatternFill("solid", fgColor="FFF2CC")
blank_fill = PatternFill("solid", fgColor="E2EFDA")
prompt_fill = PatternFill("solid", fgColor="DDEBF7")


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = thin


def autosize(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


wb = Workbook()

ws = wb.active
ws.title = "Brand_Identity_Matrix"
ws["A1"] = "On Cloud — Brand Identity Matrix (COMPANY view)"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = (
    "Scales per Brand Inventory assignment: Strength +/++/+++ ; Valence -/0/+ ; Uniqueness 0/+/++"
)
ws["A3"] = (
    "PROVISIONAL associations — assumed from available inventory / typical On positioning — "
    "student must verify against Brand Inventory. No completed Identity Matrix found in-repo at draft."
)
ws["A3"].fill = prov_fill
ws["A4"] = (
    "Course framing: Brand identity = intended associations (Inventory). Brand image measured in "
    "Exploratory. Association dimensions from Class PowerPoints 01_04_Brand Equity Components II "
    "(strength, favorability, uniqueness/POD). Gap method: Assignments/02 + 04."
)
ws.merge_cells("A4:H4")
ws["A4"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[4].height = 45

headers = [
    "Code",
    "Association (Identity)",
    "Association type (01_04)",
    "Intended Strength (+/++/+++)",
    "Intended Valence (-/0/+)",
    "Intended Uniqueness (0/+/++)",
    "Status / source note",
    "Replace-with inventory note",
]
for i, h in enumerate(headers, 1):
    ws.cell(row=6, column=i, value=h)
style_header(ws, 6, len(headers))
for r, a in enumerate(assoc, 7):
    vals = list(a) + ["Paste final inventory wording/ratings here"]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = thin
        cell.fill = prov_fill
        cell.alignment = Alignment(wrap_text=True)
ws["A13"] = (
    "Sources: Choices/Overall Decisions.md; Assignments/02-brand-inventory-report.md; "
    "Other/_pptx_extract/01_04_Brand Equity Components II.txt"
)
ws["A14"] = "Student action: paste final five associations + company ratings from Brand Inventory exhibit into rows 7-11."
autosize(ws, [8, 42, 32, 22, 20, 24, 40, 35])

ws2 = wb.create_sheet("Brand_Image_Matrix")
ws2["A1"] = "On Cloud — Brand Image Matrix (CONSUMER view — survey means)"
ws2["A1"].font = Font(bold=True, size=14)
ws2["A2"] = "Leave blank until survey data. Scales 1-5 per Exploratory assignment (Strength / Valence / Uniqueness)."
ws2["A3"] = (
    "Course: brand image = actual associations salient to consumers (01_04). Measure same five "
    "identity associations + emergent. Quant measures: 03_08_Quantitative Research III; assignment 04 scales."
)
ws2["A3"].alignment = Alignment(wrap_text=True)
ws2.merge_cells("A3:F3")
ws2.row_dimensions[3].height = 40
h2 = [
    "Code",
    "Association",
    "Mean Strength (1-5)",
    "Mean Valence (1-5)",
    "Mean Uniqueness (1-5)",
    "n / notes",
]
for i, h in enumerate(h2, 1):
    ws2.cell(row=5, column=i, value=h)
style_header(ws2, 5, len(h2))
rows_img = [(a[0], a[1]) for a in assoc] + [(e[0], e[1]) for e in emergent]
for r, (code, name) in enumerate(rows_img, 6):
    ws2.cell(row=r, column=1, value=code).border = thin
    ws2.cell(row=r, column=2, value=name).border = thin
    for c in range(3, 6):
        cell = ws2.cell(row=r, column=c, value="")
        cell.border = thin
        cell.fill = blank_fill
    if code.startswith("A"):
        note = "Identity match — fill from survey"
    else:
        note = dict((e[0], e[2]) for e in emergent)[code]
    ws2.cell(row=r, column=6, value=note).border = thin
    if code.startswith("E"):
        for c in (1, 2, 6):
            ws2.cell(row=r, column=c).fill = prov_fill
ws2["A16"] = (
    "Do NOT invent means. Brand Image Matrix exhibit does not count against 500-word or "
    "10-exhibit limits (assignment 04)."
)
autosize(ws2, [8, 42, 20, 18, 22, 45])

ws3 = wb.create_sheet("Side_by_Side_Gap")
ws3["A1"] = "Identity vs Image Gap (scaffolding)"
ws3["A1"].font = Font(bold=True, size=14)
ws3["A2"] = (
    "Compare Inventory identity ratings vs survey image means. Interpretation prompts below — "
    "no fabricated deltas."
)
ws3["A3"] = (
    "Course: Identify identity-image disconnects; enhance associations that differentiate/position "
    "and give reason-to-buy (01_04; assignment gap analysis)."
)
ws3.merge_cells("A3:L3")
ws3["A3"].alignment = Alignment(wrap_text=True)
ws3.row_dimensions[3].height = 40
h3 = [
    "Code",
    "Association",
    "Identity Strength",
    "Image Strength mean",
    "Strength gap (Image-Identity)",
    "Identity Valence",
    "Image Valence mean",
    "Valence gap",
    "Identity Uniqueness",
    "Image Uniqueness mean",
    "Uniqueness gap",
    "Interpretation prompt",
]
for i, h in enumerate(h3, 1):
    ws3.cell(row=5, column=i, value=h)
style_header(ws3, 5, len(h3))
prompts = [
    "Is cloud cushioning landing as a strong unique POD in consumer minds?",
    "Is energetic/lightweight a POP with Hoka/Nike or a POD?",
    "Does Swiss/tech innovation transfer, or is it invisible in C2C/survey?",
    "Is style/design helping or diluting performance meaning?",
    "Is premium lifestyle desirable and unique, or a price/value tension?",
]
for r, a in enumerate(assoc, 6):
    ws3.cell(row=r, column=1, value=a[0]).border = thin
    ws3.cell(row=r, column=2, value=a[1]).border = thin
    ws3.cell(row=r, column=3, value=a[3]).border = thin
    for c in [4, 5, 7, 8, 10, 11]:
        cell = ws3.cell(row=r, column=c, value="")
        cell.border = thin
        cell.fill = blank_fill
    ws3.cell(row=r, column=6, value=a[4]).border = thin
    ws3.cell(row=r, column=9, value=a[5]).border = thin
    cell = ws3.cell(row=r, column=12, value=prompts[r - 6])
    cell.border = thin
    cell.fill = prompt_fill
    cell.alignment = Alignment(wrap_text=True)
ws3["A12"] = (
    "Translation note: Inventory uses +/++/+++ etc.; Image uses 1-5. Keep dual scales "
    "directionally, or map judgmentally (+++→5, ++→4, +→3) and label as judgmental — not measured identity survey data."
)
ws3["A13"] = (
    "After data: Which gaps threaten equity? Over- vs under-delivered associations? "
    "Emergent image associations absent from identity?"
)
autosize(ws3, [8, 36, 14, 16, 18, 14, 16, 12, 16, 18, 14, 55])

ws4 = wb.create_sheet("Emergent_Associations")
ws4["A1"] = "Emergent associations from C2C / Qualitative (not necessarily in firm identity)"
ws4["A1"].font = Font(bold=True, size=14)
ws4["A2"] = c2c_note
ws4["A2"].fill = prov_fill
ws4["A2"].alignment = Alignment(wrap_text=True)
ws4.merge_cells("A2:I2")
ws4.row_dimensions[2].height = 50
ws4["A3"] = (
    "Course: Exploratory may surface associations firm is unaware of (assignment 04). "
    "C2C text mining detects common associations (03_01). Measure with same S/V/U 1-5 scales."
)
h4 = [
    "Code",
    "Association label",
    "Source (C2C / Qual / both)",
    "Evidence note / example",
    "Survey mean Strength",
    "Survey mean Valence",
    "Survey mean Uniqueness",
    "Include in Image Matrix? (Y/N)",
    "Action for identity strategy",
]
for i, h in enumerate(h4, 1):
    ws4.cell(row=5, column=i, value=h)
style_header(ws4, 5, len(h4))
for r, e in enumerate(emergent, 6):
    ws4.cell(row=r, column=1, value=e[0]).border = thin
    ws4.cell(row=r, column=2, value=e[1]).border = thin
    ws4.cell(row=r, column=3, value="C2C placeholder").border = thin
    ws4.cell(row=r, column=4, value=e[2]).border = thin
    for c in range(5, 8):
        cell = ws4.cell(row=r, column=c, value="")
        cell.border = thin
        cell.fill = blank_fill
    ws4.cell(row=r, column=8, value="").border = thin
    ws4.cell(row=r, column=9, value="").border = thin
    for c in range(1, 5):
        ws4.cell(row=r, column=c).fill = prov_fill
for r in range(10, 14):
    for c in range(1, 10):
        cell = ws4.cell(row=r, column=c, value="")
        cell.border = thin
        if c >= 5:
            cell.fill = blank_fill
ws4["A15"] = (
    "Student: after sibling C2C Excel ready, replace E1-E4 labels to match Key_Associations / Survey_Implications."
)
autosize(ws4, [8, 32, 22, 45, 18, 18, 20, 18, 35])

ws5 = wb.create_sheet("Strategy_Implications")
ws5["A1"] = "Strategy implications (2–3 actionable recommendations tied to gaps)"
ws5["A1"].font = Font(bold=True, size=14)
ws5["A2"] = (
    "Assignment: preferably 2–3 recommendations from research. Course levers: differentiate/position, "
    "reason-to-buy (01_04); quality premium (01_05); loyalty creation (01_06)."
)
ws5.merge_cells("A2:F2")
ws5["A2"].alignment = Alignment(wrap_text=True)
ws5.row_dimensions[2].height = 45
h5 = [
    "Rec #",
    "Linked gap / finding (fill after data)",
    "CBBE component",
    "Recommended action (specific)",
    "Why it grows/sustains equity",
    "Evidence sources (C2C / Qual / Survey)",
]
for i, h in enumerate(h5, 1):
    ws5.cell(row=4, column=i, value=h)
style_header(ws5, 4, len(h5))
seed = [
    (
        "1",
        "[PLACEHOLDER] e.g., Identity intends cloud cushioning as strong unique POD; Image uniqueness mean low",
        "Associations (Uniqueness)",
        "[PLACEHOLDER] Concrete IMC/product-proof action — wait for data",
        "Strengthen POD → differentiate/reason-to-buy (01_04)",
        "Survey S/V/U; C2C; Qual",
    ),
    (
        "2",
        "[PLACEHOLDER] e.g., Emergent durability or value tension vs premium identity",
        "Perceived Quality / Associations",
        "[PLACEHOLDER]",
        "Close perception gap vs competitors (01_05)",
        "C2C; Qual; Q10-Q12",
    ),
    (
        "3",
        "[PLACEHOLDER] e.g., NPS/loyalty pyramid concentration in switchers vs committed",
        "Loyalty",
        "[PLACEHOLDER]",
        "Advocacy + retention (01_06)",
        "Q13-Q18; Qual advocacy",
    ),
]
for r, row in enumerate(seed, 5):
    for c, v in enumerate(row, 1):
        cell = ws5.cell(row=r, column=c, value=v)
        cell.border = thin
        cell.fill = prompt_fill
        cell.alignment = Alignment(wrap_text=True)
for rr in (5, 6, 7):
    ws5.row_dimensions[rr].height = 40
ws5["A9"] = "Do not treat placeholder rows as findings. Replace after primary research."
autosize(ws5, [8, 50, 22, 45, 40, 30])

ws6 = wb.create_sheet("Course_Sources")
ws6["A1"] = "Course sources informing this workbook"
for i, line in enumerate(
    [
        "01_03_Brand Equity Components I — awareness pyramid",
        "01_04_Brand Equity Components II — associations definition, types, strength/favorability/uniqueness",
        "01_05_Brand Equity Components III — perceived quality relative to competitors",
        "01_06_Brand Equity Components IV — loyalty pyramid, NPS",
        "02_03_Brand Positioning III — brand personality facets",
        "03_01–03_05 Qualitative Research — C2C text mining, interviews, laddering, projectives",
        "03_06–03_08 Quantitative Research — measurement, questionnaire design, CBBE measures",
        "Assignments/02-brand-inventory-report.md — Identity Matrix scales",
        "Assignments/04-brand-exploratory-strategy-report.md — Image Matrix 1-5 scales; gap analysis",
        "PRIMARY SLIDES: Class PowerPoints/*.pptx (full deck set 01_01–07_04 present in workspace)",
        "TEXT EXTRACTS (preferred for agents): Other/_pptx_extract/01_01–04_02 + Crescent_Pure.txt",
        "PARTIAL EXTRACTS: Other/_extract/ (03_01–03_08 subset) — prefer _pptx_extract when both exist",
        "IDENTITY STATUS: Brand Sentiment Analysis/OnCloud_IDENTITY_ASSOCIATIONS_STATUS.md",
        "C2C IMAGE THEMES: Brand Sentiment Analysis/OnCloud_C2C_Brand_Exploratory.xlsx",
    ],
    2,
):
    ws6.cell(row=i, column=1, value=line)
ws6.column_dimensions["A"].width = 120

out_xlsx = bsa / "OnCloud_Brand_Image_Identity_Gap_Template.xlsx"
wb.save(out_xlsx)
print("Wrote", out_xlsx)
print("C2C note:", c2c_note[:240])


def md_to_docx(md_path, docx_path):
    text = md_path.read_text(encoding="utf-8")
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    for line in text.splitlines():
        if line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("|") and "|" in line[1:]:
            p = doc.add_paragraph(line)
            for run in p.runs:
                run.font.name = "Consolas"
                run.font.size = Pt(9)
        elif line.strip() == "---":
            continue
        elif line.startswith("> "):
            p = doc.add_paragraph(line[2:].strip())
            if p.runs:
                p.runs[0].italic = True
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(line[2:].strip(), style="List Bullet")
        elif re.match(r"^\d+\. ", line):
            doc.add_paragraph(re.sub(r"^\d+\. ", "", line), style="List Number")
        elif line.strip() == "":
            continue
        else:
            doc.add_paragraph(line)
    doc.add_paragraph(
        "GenAI disclosure: Composer used to draft this document from Class PowerPoints extracts "
        "and assignment requirements; modified by the student."
    )
    doc.save(docx_path)
    print("Wrote", docx_path)


md_to_docx(
    bsa / "OnCloud_Qualitative_Interview_Protocol.md",
    bsa / "OnCloud_Qualitative_Interview_Protocol.docx",
)
md_to_docx(
    bsa / "OnCloud_Quantitative_Survey_Questionnaire.md",
    bsa / "OnCloud_Quantitative_Survey_Questionnaire.docx",
)

log = base / "AGENTPROMPTLOG.md"
content = log.read_text(encoding="utf-8")
entry_day = "## 2026-11-17"
entry1 = (
    "(Composer) used to draft On Cloud qualitative protocol, survey questionnaire, and "
    "brand image–identity gap template for Brand Exploratory, modified by the student."
)
entry2 = (
    "(Composer) used Class PowerPoints (via Other/_pptx_extract matching decks) to frame "
    "On Cloud qualitative protocol, survey questionnaire, gap template, and exploratory "
    "outline around CBBE and research-method course frameworks, modified by the student."
)
if entry_day not in content:
    content = content.rstrip() + f"\n\n{entry_day}\n\n{entry1}\n{entry2}\n"
else:
    if entry1 not in content:
        content = content.rstrip() + f"\n{entry1}\n"
    if entry2 not in content:
        content = content.rstrip() + f"\n{entry2}\n"
    if not content.endswith("\n"):
        content += "\n"
log.write_text(content, encoding="utf-8")
print("Updated AGENTPROMPTLOG.md")
print("DONE")
