"""Convert OnCloud_C2C_Results_Draft.md to docx (simple)."""
from pathlib import Path

from docx import Document
from docx.shared import Pt

src = Path(__file__).resolve().parent / "OnCloud_C2C_Results_Draft.md"
out = Path(__file__).resolve().parent / "OnCloud_C2C_Results_Draft.docx"

text = src.read_text(encoding="utf-8")
doc = Document()
style = doc.styles["Normal"]
style.font.name = "Times New Roman"
style.font.size = Pt(12)

for line in text.splitlines():
    if line.startswith("# "):
        doc.add_heading(line[2:].strip(), level=1)
    elif line.startswith("## "):
        doc.add_heading(line[3:].strip(), level=2)
    elif line.startswith("### "):
        doc.add_heading(line[4:].strip(), level=3)
    elif line.strip() == "---":
        continue
    elif line.startswith("|") and "---" in line.replace("|", ""):
        continue
    else:
        p = doc.add_paragraph(line)
        p.paragraph_format.space_after = Pt(6)

doc.save(out)
print("wrote", out)
