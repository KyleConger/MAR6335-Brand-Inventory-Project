"""Sample C2C text-mining for On Cloud Brand Exploratory.
Documents sample sizes; extracts real quotes/counts only.
"""
from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

BASE = Path(__file__).resolve().parent
OUT_DIR = BASE / "OnCloud_Exhibits"
OUT_DIR.mkdir(parents=True, exist_ok=True)
SUMMARY_PATH = BASE / "_c2c_sample_summary.json"

# Brand mention patterns (avoid bare "on" alone)
BRAND_RE = re.compile(
    r"\b("
    r"on\s*cloud(?:tec|monster|surfer|flow|swift|boom|vista|tilt|stratus|x)?|"
    r"on\s*running|"
    r"cloudtec|"
    r"cloudmonster|"
    r"cloudsurfer|"
    r"cloudflow|"
    r"cloudswift|"
    r"cloudboom|"
    r"(?<![A-Za-z])ons?(?=\s+(?:shoes?|shoe|cloud|running|cloudtec))"
    r")\b",
    re.I,
)
ON_TOKEN = re.compile(r"(?<![A-Za-z0-9])On(?![A-Za-z0-9])")  # capitalized On

COMPETITORS = {
    "Hoka": re.compile(r"\bhoka(?:s|'s)?\b", re.I),
    "Nike": re.compile(r"\bnike(?:s|'s)?\b", re.I),
    "Brooks": re.compile(r"\bbrooks\b", re.I),
    "Asics": re.compile(r"\basics\b", re.I),
    "New Balance": re.compile(r"\bnew\s*balance\b|\bnb\b", re.I),
    "Saucony": re.compile(r"\bsaucony\b", re.I),
    "Adidas": re.compile(r"\badidas\b", re.I),
    "Altra": re.compile(r"\baltra\b", re.I),
    "Salomon": re.compile(r"\bsalomon\b", re.I),
    "Under Armour": re.compile(r"\bunder\s*armour\b|\bua\b", re.I),
}

ASSOCIATIONS = {
    "cushioning_cloud_ride": [
        r"cushion", r"cloudtec", r"soft", r"plush", r"cloud\s*feel", r"bounce", r"springy"
    ],
    "lightweight_performance": [
        r"light(?:weight)?", r"energetic", r"propulsion", r"speed", r"tempo", r"race"
    ],
    "swiss_innovation_tech": [
        r"swiss", r"engineered", r"tech(?:nology)?", r"innovation", r"rocker", r"plate"
    ],
    "style_aesthetic": [
        r"style|stylish|look(?:s|ing)?|aesthetic|fashion|clean|design|pretty|ugly|colorway"
    ],
    "premium_lifestyle": [
        r"premium|expensive|pricey|costly|overpriced|\$\d+|lifestyle|fashion"
    ],
    "durability": [r"durab|last(?:ed|s|ing)?|wear\s*out|fell\s*apart|longevity|miles"],
    "fit_sizing": [r"\bfit\b|sizing|narrow|wide|size\s*up|size\s*down|toe\s*box"],
    "comfort": [r"comfort|comfortable|comfy"],
    "value_price": [r"worth|value|deal|sale|discount|cheap|price"],
    "running_performance": [r"\brun(?:ning|ner|s)?\b|marathon|race|tempo|easy\s*day|long\s*run"],
}

POS = re.compile(
    r"\b(love|loved|great|amazing|excellent|best|favorite|favourite|recommend|"
    r"comfortable|comfy|perfect|awesome|solid|fantastic|happy|impressed|"
    r"worth|outstanding|superb)\b",
    re.I,
)
NEG = re.compile(
    r"\b(hate|hated|terrible|awful|worst|disappoint|uncomfortable|narrow|"
    r"overpriced|expensive|wear\s*out|fell\s*apart|returned|return|junk|"
    r"meh|mediocre|avoid|regret|pain|blister|cheap\s*feel)\b",
    re.I,
)

# Sample caps per file (reasonable exploratory sample)
FILE_CAPS = {
    "r_onrunning_comments.jsonl": 25000,  # brand community — denser relevance
    "r_runningshoes_comments.jsonl": 40000,
    "r_runningshoes_posts.jsonl": 15000,
    "r_askrunningshoegeeks_comments.jsonl": 40000,
    "r_askrunningshoegeeks_posts.jsonl": 15000,
    "r_runningfashion_comments.jsonl": 30000,
    "r_sneakers_comments.jsonl": 35000,
    "r_sneakers_posts.jsonl": 15000,
}


def text_of(obj: dict) -> str:
    parts = []
    for k in ("title", "selftext", "body"):
        v = obj.get(k)
        if isinstance(v, str) and v and v not in ("[deleted]", "[removed]"):
            parts.append(v)
    return "\n".join(parts)


def is_brand_hit(text: str, source: str) -> bool:
    if BRAND_RE.search(text):
        return True
    # In r_onrunning, treat most non-trivial comments as brand-context community talk
    if "onrunning" in source.lower():
        if len(text.strip()) >= 40 and not text.strip().lower().startswith(("http", "www")):
            # still require some footwear/brand cue OR capitalized On
            if ON_TOKEN.search(text) or re.search(
                r"shoe|cloud|run|pair|miles|fit|cushion|hoka|nike|brooks", text, re.I
            ):
                return True
    return False


def valence(text: str) -> str:
    p = len(POS.findall(text))
    n = len(NEG.findall(text))
    if p > n:
        return "positive"
    if n > p:
        return "negative"
    if p and n:
        return "mixed"
    return "neutral"


def assoc_hits(text: str) -> list[str]:
    hits = []
    for name, patterns in ASSOCIATIONS.items():
        for pat in patterns:
            if re.search(pat, text, re.I):
                hits.append(name)
                break
    return hits


def competitor_hits(text: str) -> list[str]:
    return [name for name, rx in COMPETITORS.items() if rx.search(text)]


def main() -> None:
    scanned = Counter()
    hits = []
    assoc_counter = Counter()
    competitor_counter = Counter()
    valence_counter = Counter()
    subreddit_hits = Counter()
    quote_candidates = []

    for fname, cap in FILE_CAPS.items():
        path = BASE / fname
        if not path.exists():
            continue
        n = 0
        with path.open(encoding="utf-8", errors="replace") as fh:
            for line in fh:
                if n >= cap:
                    break
                n += 1
                scanned[fname] += 1
                try:
                    obj = json.loads(line)
                except json.JSONDecodeError:
                    continue
                text = text_of(obj)
                if not text or len(text) < 15:
                    continue
                if not is_brand_hit(text, fname):
                    continue
                sub = obj.get("subreddit") or fname.replace(".jsonl", "")
                score = obj.get("score") or 0
                try:
                    score = int(score)
                except (TypeError, ValueError):
                    score = 0
                v = valence(text)
                a = assoc_hits(text)
                c = competitor_hits(text)
                for x in a:
                    assoc_counter[x] += 1
                for x in c:
                    competitor_counter[x] += 1
                valence_counter[v] += 1
                subreddit_hits[str(sub)] += 1
                rec = {
                    "source_file": fname,
                    "subreddit": str(sub),
                    "score": score,
                    "valence": v,
                    "associations": a,
                    "competitors": c,
                    "text": text.strip()[:500],
                    "permalink": obj.get("permalink") or obj.get("url") or "",
                    "id": obj.get("id") or "",
                }
                hits.append(rec)
                # Prefer longer, scored quotes with clear brand + theme
                if len(text) >= 60 and score >= 1:
                    quote_candidates.append(rec)

    # Deduplicate near-identical texts
    seen = set()
    unique_hits = []
    for h in hits:
        key = re.sub(r"\s+", " ", h["text"].lower())[:180]
        if key in seen:
            continue
        seen.add(key)
        unique_hits.append(h)

    quote_candidates.sort(key=lambda r: (r["score"], len(r["text"])), reverse=True)
    # Select diverse quotes across themes
    selected_quotes = []
    used_themes = set()
    for q in quote_candidates:
        themes = tuple(sorted(q["associations"][:2])) or ("general",)
        # allow multiple but diversify
        if themes in used_themes and len(selected_quotes) >= 6:
            continue
        # skip very short / noise
        t = q["text"].replace("\n", " ").strip()
        if len(t) < 70:
            continue
        if not BRAND_RE.search(t) and "onrunning" not in q["source_file"]:
            continue
        selected_quotes.append(
            {
                "subreddit": q["subreddit"],
                "score": q["score"],
                "valence": q["valence"],
                "associations": q["associations"],
                "competitors": q["competitors"],
                "quote": t[:320],
                "source_file": q["source_file"],
            }
        )
        used_themes.add(themes)
        if len(selected_quotes) >= 12:
            break

    summary = {
        "method_note": (
            "Reasonable stratified sample of Reddit JSONL dumps; brand-hit filter applied. "
            "Excel OnCloud_C2C_Brand_Exploratory.xlsx was not available at analysis time."
        ),
        "lines_scanned_by_file": dict(scanned),
        "total_lines_scanned": int(sum(scanned.values())),
        "brand_relevant_hits": len(unique_hits),
        "raw_hits_before_dedupe": len(hits),
        "valence_counts": dict(valence_counter),
        "association_counts": dict(assoc_counter.most_common()),
        "competitor_co_mention_counts": dict(competitor_counter.most_common()),
        "subreddit_hit_counts": dict(subreddit_hits.most_common()),
        "selected_quotes": selected_quotes,
    }
    SUMMARY_PATH.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print("Wrote", SUMMARY_PATH)
    print("scanned", summary["total_lines_scanned"], "hits", summary["brand_relevant_hits"])
    print("associations", summary["association_counts"])
    print("competitors", summary["competitor_co_mention_counts"])
    print("valence", summary["valence_counts"])
    print("quotes", len(selected_quotes))

    # --- Exhibits ---
    # Association bar chart
    labels = list(assoc_counter.keys())
    vals = [assoc_counter[k] for k in labels]
    # sort
    order = sorted(range(len(labels)), key=lambda i: vals[i], reverse=True)
    labels = [labels[i].replace("_", " ") for i in order]
    vals = [vals[i] for i in order]

    fig, ax = plt.subplots(figsize=(10, 5.5))
    ax.barh(labels[::-1], vals[::-1], color="#1F4E79")
    ax.set_xlabel("Mentions in brand-relevant sample comments/posts")
    ax.set_title("On Cloud C2C Association Frequencies (Reddit sample)")
    fig.tight_layout()
    fig.savefig(OUT_DIR / "Exhibit01_Association_Frequency.png", dpi=160)
    plt.close(fig)

    # Valence pie/bar
    v_order = ["positive", "mixed", "neutral", "negative"]
    v_vals = [valence_counter.get(k, 0) for k in v_order]
    fig, ax = plt.subplots(figsize=(7, 5))
    colors = ["#2E7D32", "#F9A825", "#90A4AE", "#C62828"]
    ax.bar(v_order, v_vals, color=colors)
    ax.set_ylabel("Count of brand-relevant hits")
    ax.set_title("C2C Lexical Valence Breakdown (Reddit sample)")
    fig.tight_layout()
    fig.savefig(OUT_DIR / "Exhibit02_Valence_Breakdown.png", dpi=160)
    plt.close(fig)

    # Competitor co-mentions
    c_labels = [k for k, _ in competitor_counter.most_common(10)]
    c_vals = [competitor_counter[k] for k in c_labels]
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.barh(c_labels[::-1], c_vals[::-1], color="#5B2C6F")
    ax.set_xlabel("Co-mentions with On/On Cloud context")
    ax.set_title("Competitor Co-Mentions in On-Relevant Reddit Talk")
    fig.tight_layout()
    fig.savefig(OUT_DIR / "Exhibit03_Competitor_CoMentions.png", dpi=160)
    plt.close(fig)

    # Quote exhibit Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Quote_Exhibit"
    headers = ["Exhibit_ID", "Subreddit", "Score", "Valence", "Associations", "Competitors", "Quote", "Source_File"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for i, q in enumerate(selected_quotes, 1):
        ws.append(
            [
                f"Q{i}",
                q["subreddit"],
                q["score"],
                q["valence"],
                ", ".join(q["associations"]),
                ", ".join(q["competitors"]),
                q["quote"],
                q["source_file"],
            ]
        )
        ws.cell(row=i + 1, column=7).alignment = Alignment(wrap_text=True)
    ws.column_dimensions["G"].width = 80
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["E"].width = 28

    ws2 = wb.create_sheet("Association_Counts")
    ws2.append(["Association_Theme", "Count"])
    for k, v in assoc_counter.most_common():
        ws2.append([k, v])

    ws3 = wb.create_sheet("Valence_Counts")
    ws3.append(["Valence", "Count"])
    for k in v_order:
        ws3.append([k, valence_counter.get(k, 0)])

    ws4 = wb.create_sheet("Competitor_CoMentions")
    ws4.append(["Competitor", "CoMention_Count"])
    for k, v in competitor_counter.most_common():
        ws4.append([k, v])

    ws5 = wb.create_sheet("Sample_Meta")
    ws5.append(["Metric", "Value"])
    ws5.append(["total_lines_scanned", summary["total_lines_scanned"]])
    ws5.append(["brand_relevant_hits", summary["brand_relevant_hits"]])
    for f, n in scanned.items():
        ws5.append([f"scanned::{f}", n])

    xlsx_path = OUT_DIR / "Exhibit04_Quote_and_Counts.xlsx"
    wb.save(xlsx_path)
    print("Wrote exhibits to", OUT_DIR)


if __name__ == "__main__":
    main()
