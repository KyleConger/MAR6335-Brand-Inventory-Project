"""Refine On-specific quote exhibit; update summary + Excel quote sheet."""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = Path(__file__).resolve().parent
OUT_DIR = BASE / "OnCloud_Exhibits"
SUMMARY_PATH = BASE / "_c2c_sample_summary.json"

ON_PRODUCT = re.compile(
    r"\b("
    r"on\s*cloud(?:tec|monster|surfer|flow|swift|boom|vista|tilt|stratus|x|5|6)?|"
    r"on\s*running|"
    r"cloudtec|"
    r"cloudmonster|"
    r"cloudsurfer|"
    r"cloudflow|"
    r"cloudswift|"
    r"cloudboom|"
    r"lightspray"
    r")\b",
    re.I,
)
# Titles that are clearly about another brand as primary subject
OTHER_PRIMARY = re.compile(
    r"^(first\s+run|review|check[- ]?in)?[:\s]*("
    r"nike|hoka|brooks|asics|new\s*balance|saucony|adidas|altra|topo|puma|"
    r"mount\s*to\s*coast|dynafit|dynafish|li[- ]?ning|salomon|under\s*armour|"
    r"skechers|mizuno|hoka|"
    r"alphafly|pegasus|vomero|vaporfly|superblast|novablast|megablast|"
    r"clifton|bondi|ghost|glycerin|rebel|paramount"
    r")\b",
    re.I,
)
POS = re.compile(
    r"\b(love|loved|great|amazing|excellent|best|favorite|recommend|comfortable|"
    r"comfy|perfect|awesome|solid|worth|impressed|dream|want)\b",
    re.I,
)
NEG = re.compile(
    r"\b(hate|hated|terrible|awful|worst|disappoint|uncomfortable|overpriced|"
    r"wear\s*out|returned|junk|regret|stiff|numb|disdain|glorified|overhyped|"
    r"trap\s*rocks|avoid)\b",
    re.I,
)

FILES = [
    ("r_onrunning_comments.jsonl", 11319),
    ("r_runningshoes_comments.jsonl", 31315),
    ("r_runningshoes_posts.jsonl", 3964),
    ("r_askrunningshoegeeks_comments.jsonl", 40000),
    ("r_askrunningshoegeeks_posts.jsonl", 15000),
    ("r_runningfashion_comments.jsonl", 30000),
    ("r_sneakers_comments.jsonl", 20000),
    ("r_sneakers_posts.jsonl", 12150),
]


def valence(text: str) -> str:
    p, n = len(POS.findall(text)), len(NEG.findall(text))
    if p > n:
        return "positive"
    if n > p:
        return "negative"
    if p and n:
        return "mixed"
    return "neutral"


def theme_tags(text: str) -> list[str]:
    tags = []
    checks = [
        ("cushioning_cloud_ride", r"cushion|cloudtec|soft|plush|bounce|springy|pods"),
        ("style_aesthetic", r"style|stylish|look|aesthetic|fashion|design|suburbia"),
        ("durability", r"durab|last|wear\s*out|miles|longevity"),
        ("fit_sizing", r"\bfit\b|sizing|narrow|wide|numb|toe"),
        ("value_price", r"price|expensive|overpriced|worth|value|\$\d+"),
        ("running_performance", r"run|race|tempo|distance|marathon"),
        ("comfort", r"comfort|comfy|comfortable"),
        ("lifestyle_crossover", r"lifestyle|office|suburbia|everyday|casual|coworkers"),
        ("swiss_innovation_tech", r"swiss|tech|innovation|engineered|lightspray"),
    ]
    for name, pat in checks:
        if re.search(pat, text, re.I):
            tags.append(name)
    return tags


def is_on_primary(title: str, body: str, source: str) -> bool:
    text = f"{title}\n{body}".strip()
    if not ON_PRODUCT.search(text):
        return False
    # Brand community comments: allow if product mentioned
    if "onrunning" in source.lower():
        return True
    title_l = (title or "").strip()
    if title_l and ON_PRODUCT.search(title_l):
        return True
    # Reject clear other-brand-first titles unless On also in title
    if title_l and OTHER_PRIMARY.search(title_l) and not ON_PRODUCT.search(title_l):
        return False
    # Body must mention On product near start or multiple times
    matches = list(ON_PRODUCT.finditer(text))
    if not matches:
        return False
    # first match within first 180 chars OR title empty + strong product name
    if matches[0].start() <= 180:
        return True
    if len(matches) >= 2:
        return True
    return False


def main() -> None:
    quotes = []
    for fname, cap in FILES:
        path = BASE / fname
        if not path.exists():
            continue
        n = 0
        with path.open(encoding="utf-8", errors="replace") as fh:
            for line in fh:
                if n >= cap:
                    break
                n += 1
                try:
                    o = json.loads(line)
                except json.JSONDecodeError:
                    continue
                title = o.get("title") if isinstance(o.get("title"), str) else ""
                body_parts = []
                for k in ("selftext", "body"):
                    v = o.get(k)
                    if isinstance(v, str) and v not in ("", "[deleted]", "[removed]"):
                        body_parts.append(v)
                body = " ".join(body_parts).strip()
                text = f"{title} {body}".strip()
                if len(text) < 80:
                    continue
                if not is_on_primary(title, body, fname):
                    continue
                try:
                    score = int(o.get("score") or 0)
                except (TypeError, ValueError):
                    score = 0
                if score < 2:
                    continue
                quotes.append(
                    {
                        "subreddit": o.get("subreddit") or fname,
                        "score": score,
                        "quote": re.sub(r"\s+", " ", text)[:340],
                        "valence": valence(text),
                        "associations": theme_tags(text),
                        "source_file": fname,
                        "title": title[:120],
                    }
                )

    quotes.sort(key=lambda x: x["score"], reverse=True)
    seen = set()
    picked = []
    for q in quotes:
        key = q["quote"][:110].lower()
        if key in seen:
            continue
        seen.add(key)
        picked.append(q)
        if len(picked) >= 12:
            break

    summary = json.loads(SUMMARY_PATH.read_text(encoding="utf-8"))
    summary["selected_quotes"] = [
        {k: v for k, v in q.items() if k != "title"} for q in picked
    ]
    summary["quote_filter_note"] = (
        "Quotes restricted to On product as primary subject (title/early mention or "
        "r/ONrunning community); score>=2; length>=80."
    )
    SUMMARY_PATH.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    xlsx = OUT_DIR / "Exhibit04_Quote_and_Counts.xlsx"
    wb = load_workbook(xlsx)
    ws = wb["Quote_Exhibit"]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for i, q in enumerate(picked, 1):
        ws.append(
            [
                f"Q{i}",
                q["subreddit"],
                q["score"],
                q["valence"],
                ", ".join(q["associations"]),
                "",
                q["quote"],
                q["source_file"],
            ]
        )
        ws.cell(row=i + 1, column=7).alignment = Alignment(wrap_text=True)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(xlsx)

    for i, q in enumerate(picked, 1):
        print(f"--- Q{i} score={q['score']} {q['subreddit']} {q['valence']} title={q['title']!r}")
        print(q["quote"])
        print()
    print("picked", len(picked), "from", len(quotes))


if __name__ == "__main__":
    main()
