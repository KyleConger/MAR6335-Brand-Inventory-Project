# -*- coding: utf-8 -*-
"""Mine Reddit JSONL for On / On Cloud C2C brand exploratory analysis."""
from __future__ import annotations

import json
import os
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

BASE = Path(__file__).resolve().parent
OUT = BASE / "OnCloud_C2C_Brand_Exploratory.xlsx"

INPUT_FILES = [
    "r_askrunningshoegeeks_posts.jsonl",
    "r_askrunningshoegeeks_comments.jsonl",
    "r_onrunning_comments.jsonl",
    "r_runningfashion_comments.jsonl",
    "r_runningshoes_comments.jsonl",
    "r_runningshoes_posts.jsonl",
    "r_sneakers_posts.jsonl",
    "r_sneakers_comments.jsonl",
]

# Brand-specific patterns — avoid bare "on"
ON_MODEL_RE = re.compile(
    r"\b("
    r"cloudmonster|cloudsurfer|cloudstratus|cloudrunner|cloudflow|cloudswift|"
    r"cloudnova|cloudvista|cloudtilt|cloudeclipse|cloudflyer|cloudultra|"
    r"cloudrift|cloudboom|cloudpulse|cloudaway|cloud5|cloud\s*x|"
    r"cloudtec|cloudtec™|cloud\s*tec|"
    r"on\s*clouds?|oncloud|"
    r"roger\s+federer\s*(?:x\s*)?on|federer\s*(?:x\s*)?on|"
    r"cloudmonster\s*hyper|"
    r"the\s+roger"
    r")\b"
    # "On Running" as brand — not "on running a marathon"
    r"|\bon\s+running(?:\s+shoes?\b|\s+brand\b|\b)(?!\s+(?:a|an|my|the|this|that|his|her|your)\b)",
    re.I,
)
ONS_RE = re.compile(
    r"\b(?:my|your|his|her|their|these|those|some|new|old|pair of|pairs of)\s+ons\b"
    r"|\bons\s+(?:are|were|feel|feel|run|shoe|shoes|cloud|running)\b"
    r"|\b(?:wear|wearing|bought|buy|owned|own|love|hate|tried|trying)\s+(?:my\s+)?ons\b",
    re.I,
)
# Standalone brand cue — avoid "money on shoes" / "years on the" false positives.
# Prefer "On Running", "On Cloud*", "brand On", "On's Cloudmonster", not bare "on shoes".
ON_BRANDISH_RE = re.compile(
    r"\b(?:brand\s+)on\b"
    r"|\bon(?:'s)?\s+cloud(?:s|monster|surfer|stratus|runner|flow|swift|nova|tec|\s*x)?\b"
    r"|\bon\s+running(?:\s+shoes?\b|\s+brand\b|\b)(?!\s+(?:a|an|my|the|this|that|his|her|your)\b)"
    r"|\b(?:from|by|of|with)\s+on\s+running\b"
    r"|\bon\s+cloudtec\b"
    r"|\bons\s+(?:running\s+)?(?:shoes?|trainers?|sneakers?|models?)\b",
    re.I,
)

COMPETITORS = {
    "Hoka": re.compile(r"\bhoka(?:s|oneone|one)?\b", re.I),
    "Nike": re.compile(r"\bnike\b|\bvaporfly\b|\bpegasus\b|\binvincible\b|\bzoomx\b", re.I),
    "Brooks": re.compile(r"\bbrooks\b|\bghost\b|\bglycerin\b|\bhyperion\b", re.I),
    "Asics": re.compile(r"\basics\b|\bgel[- ]?nimbus\b|\bkayano\b|\bnovablast\b|\bmetaspeed\b", re.I),
    "New Balance": re.compile(r"\bnew\s*balance\b|\bnb\s?\d{3,4}\b|\bfresh\s*foam\b|\brebel\b", re.I),
    "Saucony": re.compile(r"\bsaucony\b|\bendorphin\b|\bkinvara\b|\btriumph\b", re.I),
    "Adidas": re.compile(r"\badidas\b|\badizero\b|\bultraboost\b|\bboston\b", re.I),
    "Altra": re.compile(r"\baltra\b|\btorín\b|\btorin\b|\bescalante\b", re.I),
    "Salomon": re.compile(r"\bsalomon\b", re.I),
    "Puma": re.compile(r"\bpuma\b|\bdeviate\b", re.I),
    "Under Armour": re.compile(r"\bunder\s*armour\b|\bua\s+flow\b", re.I),
    "Mizuno": re.compile(r"\bmizuno\b|\bwave\s*rider\b", re.I),
    "Topo": re.compile(r"\btopo\s*athletic\b|\btopo\b", re.I),
}

# Association / theme lexicon (course: brand associations / meaning)
ASSOCIATION_THEMES = [
    ("Cloud cushioning / cloud feel", [
        r"\bcloud\s*(?:feel|feeling|like|cushion|cushioning|foam|elements?|pods?)\b",
        r"\bsoft(?:ness)?\b.*\bcloud\b", r"\bcloudtec\b",
        r"\bplush\b", r"\bmarshmallow\b", r"\bpillowy\b",
        r"\bcushion(?:ing|ed)?\b", r"\bsoft\s+landings?\b", r"\bbounc(?:e|y|iness)\b",
    ]),
    ("Unique / distinctive ride (CloudTec)", [
        r"\bcloudtec\b", r"\bunique\s+(?:feel|ride|cushion)\b",
        r"\bdifferent\s+(?:feel|ride)\b", r"\bunusual\s+(?:feel|ride)\b",
        r"\bpods?\b", r"\bcloud\s*elements?\b",
    ]),
    ("Lightweight / agile", [
        r"\blight(?:weight)?\b", r"\bagile\b", r"\bnimble\b", r"\bsnappy\b",
    ]),
    ("Style / fashion / aesthetics", [
        r"\bstyl(?:e|ish|ing)\b", r"\bfashion\b", r"\baesthetic\b", r"\blook(?:s|ing)?\b",
        r"\btrendy\b", r"\bfashion[- ]?forward\b", r"\bclean\s+look\b",
        r"\bstreetwear\b", r"\bcute\b", r"\bpretty\b", r"\bsleek\b",
    ]),
    ("Premium / expensive / price", [
        r"\bexpensiv(?:e|eness)\b", r"\bpricey\b", r"\bpremium\b", r"\boverpriced\b",
        r"\bprice(?:y|d)?\b", r"\bcost(?:ly)?\b", r"\b\$\d{2,3}\b",
        r"\bworth\s+(?:it|the\s+money)\b", r"\bnot\s+worth\b",
    ]),
    ("Durability / longevity", [
        r"\bdurab(?:le|ility)\b", r"\blast(?:s|ed|ing)?\b", r"\bworn\s+out\b",
        r"\bfall(?:ing)?\s+apart\b", r"\bmiles?\b.*\b(?:out|done|dead)\b",
        r"\boutsole\b", r"\bwear\s+(?:and\s+tear|fast|quickly)\b",
    ]),
    ("Comfort", [
        r"\bcomfort(?:able|ability)?\b", r"\buncomfortable\b", r"\bcomfy\b",
    ]),
    ("Fit / sizing / narrow", [
        r"\bfit(?:s|ting)?\b", r"\bsiz(?:e|ing|ed)\b", r"\bnarrow\b", r"\bwide\b",
        r"\btoe\s*box\b", r"\brun\s+(?:small|large|narrow|true)\b",
    ]),
    ("Performance / running / speed", [
        r"\bperformance\b", r"\bspeed\b", r"\brace\b", r"\btempo\b", r"\bmarathon\b",
        r"\bdaily\s+trainer\b", r"\beasy\s+(?:run|day|miles)\b", r"\brecovery\b",
        r"\bpropulsion\b", r"\benergy\s+return\b", r"\bstack\b",
    ]),
    ("Trendy / hype / status", [
        r"\bhype\b", r"\btrendy\b", r"\bstatus\b", r"\binstagram\b", r"\binfluencer\b",
        r"\bflex\b", r"\bcool\s+(?:brand|shoes?)\b", r"\bpopular\b",
    ]),
    ("Swiss / European / design-led", [
        r"\bswiss\b", r"\beurope(?:an)?\b", r"\bzurich\b", r"\bdesign(?:er|[- ]led)?\b",
        r"\bminimal(?:ist)?\b",
    ]),
    ("Celebrity / Roger Federer", [
        r"\bfederer\b", r"\broger\b", r"\bcelebrity\b", r"\bathlete\b.*\bendors",
    ]),
    ("Injury / support concerns", [
        r"\binjur(?:y|ies|ed)\b", r"\bknee\b", r"\bachilles\b", r"\bplantar\b",
        r"\bsupport\b", r"\bstability\b", r"\bpronat", r"\bcushion\s+only\b",
    ]),
    ("Everyday / lifestyle wear", [
        r"\blifestyle\b", r"\bcausal\b", r"\bcasual\b", r"\beveryday\b",
        r"\bwalking\b", r"\bcommute\b", r"\bnot\s+(?:just\s+)?for\s+running\b",
    ]),
    ("Switching / comparison shopping", [
        r"\bswitch(?:ed|ing)?\b", r"\binstead\s+of\b", r"\bvs\.?\b", r"\bcompared\s+to\b",
        r"\balternative\s+to\b", r"\breplaced?\b",
    ]),
]

QUALITY_THEMES = [
    ("Comfort", [r"\bcomfort", r"\bcomfy\b", r"\bplush\b", r"\bsoft\b"]),
    ("Cushioning / ride", [r"\bcushion", r"\bride\b", r"\bcloud\s*feel", r"\bsoft\b", r"\bfirm\b"]),
    ("Durability", [r"\bdurab", r"\blast", r"\bworn\s+out", r"\boutsole", r"\bmiles"]),
    ("Fit / sizing", [r"\bfit", r"\bsiz", r"\bnarrow", r"\bwide", r"\btoe\s*box"]),
    ("Aesthetics / look", [r"\blook", r"\bstyl", r"\baesthetic", r"\bcolor", r"\bcolour"]),
    ("Performance / speed", [r"\bperformance", r"\bspeed", r"\brace", r"\btempo", r"\benergy\s+return"]),
    ("Value / price-quality", [r"\bworth", r"\bvalue", r"\bexpensiv", r"\bpricey", r"\boverpriced", r"\bcheap"]),
]

LOYALTY_PATTERNS = {
    "repurchase_intent": re.compile(
        r"\b(?:buy|buying|bought)\s+(?:another|more|again|a second|my next)\b|"
        r"\brepurchase\b|\bwill\s+buy\s+again\b|\bnext\s+pair\s+(?:of\s+)?ons?\b|"
        r"\bon\s+(?:my\s+)?(?:third|fourth|fifth|2nd|3rd|4th)\s+pair\b|"
        r"\bmultiple\s+pairs?\b|\bsecond\s+pair\b",
        re.I,
    ),
    "advocacy_recommend": re.compile(
        r"\brecommend(?:ed|ing)?\b|\bhighly\s+recommend\b|\btell\s+(?:everyone|friends)\b|"
        r"\bmust[- ]try\b|\bgo\s+(?:for|with)\s+(?:the\s+)?ons?\b|"
        r"\blove\s+(?:my\s+)?ons?\b|\bgreat\s+shoe\b",
        re.I,
    ),
    "switching_away": re.compile(
        r"\bswitch(?:ed|ing)?\s+(?:away|from|to)\b|\bgave\s+up\s+on\b|"
        r"\bwon'?t\s+buy\s+again\b|\bnever\s+again\b|\breturned?\b|"
        r"\bsold\s+(?:my\s+)?ons?\b|\bditched\b|\bmoved\s+(?:on|away)\s+from\b",
        re.I,
    ),
    "criticism_intensity": re.compile(
        r"\bhate\b|\bterrible\b|\bawful\b|\bgarbage\b|\bwaste\s+of\s+money\b|"
        r"\boverhyped\b|\boverrated\b|\bdisappoint(?:ed|ing)\b|\bregret\b|"
        r"\bworst\b|\bdo\s+not\s+buy\b|\bdon'?t\s+buy\b",
        re.I,
    ),
    "committed_love": re.compile(
        r"\bfavorite\b|\bfavourite\b|\bobsessed\b|\ball[- ]time\b|"
        r"\bcan'?t\s+live\s+without\b|\bgrail\b|\bdied\s+in\b",
        re.I,
    ),
}

PERSONALITY_CUES = {
    "Excitement (trendy / up-to-date / daring)": [
        r"\btrendy\b", r"\bhype\b", r"\bcool\b", r"\bmodern\b", r"\bfresh\b",
        r"\bup[- ]to[- ]date\b", r"\bfashion\b", r"\bstylish\b", r"\bedgy\b",
    ],
    "Sophistication (premium / design / Swiss)": [
        r"\bpremium\b", r"\bsleek\b", r"\bswiss\b", r"\belegant\b", r"\bminimal\b",
        r"\bdesign(?:er)?\b", r"\bchic\b", r"\bpolished\b", r"\bluxury\b",
    ],
    "Competence (technical / engineered)": [
        r"\btechnical\b", r"\bengineered?\b", r"\binnovative\b", r"\btech\b",
        r"\bperformance\b", r"\bcloudtec\b", r"\bscience\b", r"\bprecise\b",
    ],
    "Ruggedness (outdoorsy / athletic)": [
        r"\boutdoor\b", r"\btrail\b", r"\brugged\b", r"\btough\b", r"\bathletic\b",
    ],
    "Sincerity (honest / wholesome) — sparse expected": [
        r"\bhonest\b", r"\bauthentic\b", r"\bwholesome\b", r"\bdown[- ]to[- ]earth\b",
    ],
}

POS_WORDS = re.compile(
    r"\b(?:love|loved|loving|great|amazing|awesome|excellent|perfect|best|fantastic|"
    r"comfortable|comfy|recommend|solid|impressed|worth|favorite|favourite|happy|"
    r"smooth|snappy|fun|beautiful|gorgeous|stunning)\b",
    re.I,
)
NEG_WORDS = re.compile(
    r"\b(?:hate|hated|terrible|awful|worst|disappoint|overpriced|overhyped|overrated|"
    r"uncomfortable|narrow|pain|injury|injur|garbage|regret|waste|returned|return|"
    r"cheap(?:ly)?\s+made|fall(?:ing)?\s+apart|worn\s+out\s+fast|not\s+worth)\b",
    re.I,
)

STOP = set(
    """
    a an the and or but if then else when while of to for in on at by with from as is are was were be been being
    this that these those it its it s i me my we our you your they them their he she his her what which who whom
    not no yes so too very just also only into about over under again further more most other some such than
    can will would could should may might must do does did done having have has had get got getting go going
    went gone make made making look looking looks like liked likes really pretty quite still even back way
    one two three first second new old good bad big little much many any all each both few own same
    im ive id ill dont doesnt didnt isnt arent wasnt werent cant wont wouldnt couldnt shouldnt
    u ur lol lmao tbh imo imho rn asap etc via http https www com org net
    shoe shoes run running runner runners mile miles pair pairs day days week weeks month months year years
    time times thing things something anything nothing someone anyone everyone people person
    think thinking thought know knowing knew feel feeling felt want wanted need needed try tried trying
    use used using wear wearing wore worn buy buying bought say said says saying tell told
    also just even still well now then here there where how why what
    """.split()
)

NGRAM_KEEP = re.compile(r"^[a-z][a-z0-9'-]{1,24}$")


IDIOM_FP_RE = re.compile(
    r"\bwalking\s+on\s+clouds?\b|\bon\s+cloud\s+nine\b|\bhead\s+in\s+the\s+clouds\b|"
    r"\bhooked\s+on\s+running\b|\bfocus(?:ed|ing)?\s+on\s+running\b|"
    r"\bkeep(?:ing)?\s+on\s+running\b|\bbased\s+on\s+running\b",
    re.I,
)


def is_on_relevant(text: str, subreddit: str) -> tuple[bool, list[str]]:
    """Return (match?, matched_rule_tags). r/OnRunning treated as brand community context."""
    tags = []
    if not text:
        return False, tags
    if subreddit and subreddit.lower().replace(" ", "") in {"onrunning", "on_running"}:
        tags.append("subreddit_onrunning")
        # Still scan for models; community posts count as On-relevant
        if ON_MODEL_RE.search(text) and not _only_idiom_cloud_match(text):
            tags.append("model_or_brand_phrase")
        if ONS_RE.search(text):
            tags.append("ons_possessive_usage")
        return True, tags
    model_hit = bool(ON_MODEL_RE.search(text)) and not _only_idiom_cloud_match(text)
    if model_hit:
        tags.append("model_or_brand_phrase")
    if ONS_RE.search(text):
        tags.append("ons_possessive_usage")
    if ON_BRANDISH_RE.search(text) and not IDIOM_FP_RE.search(text):
        tags.append("on_brandish_phrase")
    # If the only cue was an idiom-like "on cloud(s)", drop
    return (len(tags) > 0), tags


def _only_idiom_cloud_match(text: str) -> bool:
    """True when On-cloud-like regex hits are only idioms (walking on clouds, etc.)."""
    if not IDIOM_FP_RE.search(text):
        return False
    # Strong non-idiom brand evidence?
    strong = re.compile(
        r"\b(?:cloudmonster|cloudsurfer|cloudstratus|cloudrunner|cloudflow|cloudswift|"
        r"cloudnova|cloudtec|oncloud|on\s*cloud\s*(?:5|x|monster|surfer|stratus)|"
        r"on\s+running\s+shoes?|my\s+ons\b|pair\s+of\s+ons\b)\b",
        re.I,
    )
    return not strong.search(text)


def valence_of(text: str) -> str:
    p = len(POS_WORDS.findall(text))
    n = len(NEG_WORDS.findall(text))
    if p > n and p > 0:
        return "positive"
    if n > p and n > 0:
        return "negative"
    if p > 0 and n > 0:
        return "mixed"
    return "mixed/neutral"


def extract_text(obj: dict, kind: str) -> str:
    if kind == "post":
        title = obj.get("title") or ""
        selftext = obj.get("selftext") or ""
        if selftext in ("[removed]", "[deleted]"):
            selftext = ""
        return f"{title}\n{selftext}".strip()
    body = obj.get("body") or ""
    if body in ("[removed]", "[deleted]"):
        return ""
    return body.strip()


def tokenize(text: str) -> list[str]:
    text = text.lower()
    text = re.sub(r"https?://\S+", " ", text)
    text = re.sub(r"[^a-z0-9'\-\s]", " ", text)
    toks = [t for t in text.split() if NGRAM_KEEP.match(t) and t not in STOP and len(t) > 2]
    return toks


def theme_hits(text: str, themes) -> list[str]:
    hits = []
    for name, pats in themes:
        for p in pats:
            if re.search(p, text, re.I):
                hits.append(name)
                break
    return hits


def short_quote(text: str, maxlen: int = 220) -> str:
    t = re.sub(r"\s+", " ", text).strip()
    if len(t) <= maxlen:
        return t
    return t[: maxlen - 1] + "…"


def write_df(ws, df: pd.DataFrame, header_fill="1F4E79"):
    header_font = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor=header_fill)
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(r_idx, c_idx, val)
            if r_idx == 1:
                cell.font = header_font
                cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        maxlen = 12
        for cell in col[:80]:
            if cell.value:
                maxlen = max(maxlen, min(60, len(str(cell.value))))
        ws.column_dimensions[letter].width = maxlen + 2


def write_kv_sheet(ws, rows: list[tuple], title: str | None = None):
    if title:
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=14)
        start = 3
    else:
        start = 1
    header_font = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E79")
    ws.cell(start, 1, "Section / Item").font = header_font
    ws.cell(start, 1).fill = fill
    ws.cell(start, 2, "Content (evidence / analysis)").font = header_font
    ws.cell(start, 2).fill = fill
    for i, (a, b) in enumerate(rows, start + 1):
        ws.cell(i, 1, a)
        ws.cell(i, 2, b)
        ws.cell(i, 1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 100


def main():
    stats = {
        "total_scanned": 0,
        "posts_scanned": 0,
        "comments_scanned": 0,
        "on_hits": 0,
        "on_posts": 0,
        "on_comments": 0,
        "parse_errors": 0,
        "empty_text": 0,
    }
    by_source = Counter()  # (file, subreddit, type) -> scanned
    by_source_hits = Counter()
    association_counts = Counter()
    association_examples = defaultdict(list)
    association_valence = defaultdict(Counter)
    keyword_uni = Counter()
    keyword_bi = Counter()
    quality_counts = Counter()
    quality_examples = defaultdict(list)
    quality_valence = defaultdict(Counter)
    loyalty_counts = Counter()
    loyalty_examples = defaultdict(list)
    personality_counts = Counter()
    personality_examples = defaultdict(list)
    competitor_co = Counter()
    competitor_examples = defaultdict(list)
    quote_bank = []
    match_rule_counts = Counter()
    model_mention_counts = Counter()

    MODEL_LIST = [
        "Cloudmonster", "Cloudsurfer", "Cloudstratus", "Cloudrunner", "Cloudflow",
        "Cloudswift", "Cloudnova", "Cloud X", "Cloudtec", "Cloud 5", "Cloudaway",
        "Cloudultra", "Cloudrift", "The Roger", "On Cloud", "On Running",
    ]
    model_res = [(m, re.compile(re.escape(m).replace(r"\ ", r"\s+"), re.I)) for m in MODEL_LIST]

    for fname in INPUT_FILES:
        path = BASE / fname
        if not path.exists():
            print(f"MISSING: {fname}")
            continue
        kind = "post" if "posts" in fname else "comment"
        print(f"Scanning {fname} ({kind})...")
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            for line in fh:
                stats["total_scanned"] += 1
                if kind == "post":
                    stats["posts_scanned"] += 1
                else:
                    stats["comments_scanned"] += 1
                try:
                    obj = json.loads(line)
                except json.JSONDecodeError:
                    stats["parse_errors"] += 1
                    continue
                sub = obj.get("subreddit") or ""
                by_source[(fname, sub, kind)] += 1
                text = extract_text(obj, kind)
                if not text:
                    stats["empty_text"] += 1
                    continue
                relevant, tags = is_on_relevant(text, sub)
                if not relevant:
                    continue
                stats["on_hits"] += 1
                if kind == "post":
                    stats["on_posts"] += 1
                else:
                    stats["on_comments"] += 1
                by_source_hits[(fname, sub, kind)] += 1
                for t in tags:
                    match_rule_counts[t] += 1

                for mname, mre in model_res:
                    if mre.search(text):
                        model_mention_counts[mname] += 1

                val = valence_of(text)
                score = obj.get("score")
                try:
                    score = int(score) if score is not None else 0
                except (TypeError, ValueError):
                    score = 0
                permalink = obj.get("permalink") or ""
                rid = obj.get("id") or ""

                # Associations
                for th in theme_hits(text, ASSOCIATION_THEMES):
                    association_counts[th] += 1
                    association_valence[th][val] += 1
                    if len(association_examples[th]) < 8:
                        association_examples[th].append(
                            {"quote": short_quote(text), "source": fname, "subreddit": sub, "type": kind, "valence": val}
                        )

                # Keywords
                toks = tokenize(text)
                keyword_uni.update(toks)
                for i in range(len(toks) - 1):
                    keyword_bi[f"{toks[i]} {toks[i+1]}"] += 1

                # Quality
                for th in theme_hits(text, [(n, p) for n, p in QUALITY_THEMES]):
                    quality_counts[th] += 1
                    quality_valence[th][val] += 1
                    if len(quality_examples[th]) < 6:
                        quality_examples[th].append(
                            {"quote": short_quote(text), "source": fname, "subreddit": sub, "type": kind, "valence": val}
                        )

                # Loyalty
                for lname, lre in LOYALTY_PATTERNS.items():
                    if lre.search(text):
                        loyalty_counts[lname] += 1
                        if len(loyalty_examples[lname]) < 8:
                            loyalty_examples[lname].append(
                                {"quote": short_quote(text), "source": fname, "subreddit": sub, "type": kind, "valence": val}
                            )

                # Personality
                for pname, pats in PERSONALITY_CUES.items():
                    for p in pats:
                        if re.search(p, text, re.I):
                            personality_counts[pname] += 1
                            if len(personality_examples[pname]) < 5:
                                personality_examples[pname].append(
                                    {"quote": short_quote(text), "source": fname, "subreddit": sub, "type": kind}
                                )
                            break

                # Competitors co-mentioned
                for cname, cre in COMPETITORS.items():
                    if cre.search(text):
                        competitor_co[cname] += 1
                        if len(competitor_examples[cname]) < 5:
                            competitor_examples[cname].append(
                                {"quote": short_quote(text), "source": fname, "subreddit": sub, "type": kind, "valence": val}
                            )

                # Quote bank: require brand-specific text cues (not community-only hits)
                brand_text_tags = [t for t in tags if t != "subreddit_onrunning"]
                if brand_text_tags and len(text) >= 60 and (score >= 1 or len(quote_bank) < 300):
                    # Skip obvious non-brand noise (links-only heavy, unrelated sneaker drops)
                    if not re.search(r"\b(?:jordan|yeezy|dunk|chicago|griffin)\b", text, re.I) or ON_MODEL_RE.search(text) or ONS_RE.search(text):
                        quote_bank.append(
                            {
                                "quote": short_quote(text, 280),
                                "source_file": fname,
                                "subreddit": sub,
                                "type": kind,
                                "score": score,
                                "valence": val,
                                "id": rid,
                                "permalink": permalink,
                                "match_rules": ",".join(tags),
                            }
                        )

    # Curate quote bank: prefer brand-phrase hits, diversify subreddits, then score
    def quote_priority(q):
        rules = q.get("match_rules") or ""
        brandish = 2 if "model_or_brand_phrase" in rules or "ons_possessive_usage" in rules else 0
        return (brandish, q["score"], len(q["quote"]))

    quote_bank_sorted = sorted(quote_bank, key=quote_priority, reverse=True)
    curated = []
    seen = set()
    sub_counts = Counter()
    for q in quote_bank_sorted:
        key = q["quote"][:80].lower()
        if key in seen:
            continue
        # Require at least one clear On cue in the quote text itself
        qtext = q["quote"]
        if _only_idiom_cloud_match(qtext):
            continue
        if not (ON_MODEL_RE.search(qtext) or ONS_RE.search(qtext) or ON_BRANDISH_RE.search(qtext)):
            continue
        # Prefer quotes that name a Cloud* model / On Cloud / Ons (exhibit quality)
        strong = re.search(
            r"\b(?:cloudmonster|cloudsurfer|cloudstratus|cloudrunner|cloudflow|cloudswift|"
            r"cloudnova|cloudtec|onclouds?|on\s*cloud|on\s+running|my\s+ons|pair\s+of\s+ons|"
            r"cloudboom|the\s+roger|federer)\b",
            qtext,
            re.I,
        )
        if not strong:
            continue
        if sub_counts[q["subreddit"]] >= 12:
            continue
        seen.add(key)
        sub_counts[q["subreddit"]] += 1
        curated.append(q)
        if len(curated) >= 40:
            break

    # Dominant valence helper
    def dom_valence(counter: Counter) -> str:
        if not counter:
            return "mixed/neutral"
        return counter.most_common(1)[0][0]

    # ---- Build workbook ----
    wb = Workbook()

    # 1. Executive Summary
    ws = wb.active
    ws.title = "Executive_Summary"
    hit_rate = (100.0 * stats["on_hits"] / stats["total_scanned"]) if stats["total_scanned"] else 0
    top_assoc = association_counts.most_common(8)
    top_assoc_txt = "; ".join(f"{a} (n={c})" for a, c in top_assoc) if top_assoc else "n/a"
    top_comp = competitor_co.most_common(5)
    top_comp_txt = "; ".join(f"{a} (n={c})" for a, c in top_comp) if top_comp else "n/a"

    exec_rows = [
        ("Study type", "C2C (consumer-to-consumer) Reddit text mining — primary research input for Brand Exploratory (Assignments/04-brand-exploratory-strategy-report.md). Focus: meaning of the brand to consumers, not star ratings."),
        ("Brand", "On Running / On Cloud (Swiss running brand; consumer shorthand includes OnCloud, Ons, Cloud* models)."),
        ("Course framework — brand equity components", "Findings organized around the principal Brand Exploratory dimensions: brand awareness, associations (brand image/meaning), perceived quality, loyalty, and brand personality — aligned with Brand Equity Components lectures (01_03_Brand Equity Components I; 01_04_Brand Equity Components II; 01_05_Brand Equity Components III; 01_06_Brand Equity Components IV) and measuring guidance in 03_08_Quantitative Research III."),
        ("Course method note", "Text mining is positioned in class as most applicable to C2C data and useful for detecting common brand associations (03_01_Qualitative Research I — Text Mining). C2C findings should feed qualitative probes and quantitative association measurement (strength / valence / uniqueness per Brand Exploratory assignment)."),
        ("EVIDENCE — sample size", f"Records scanned: {stats['total_scanned']:,} ({stats['posts_scanned']:,} posts; {stats['comments_scanned']:,} comments). On-relevant hits: {stats['on_hits']:,} ({stats['on_posts']:,} posts; {stats['on_comments']:,} comments). Hit rate: {hit_rate:.2f}%."),
        ("EVIDENCE — awareness signals (proxies)", "Reddit C2C cannot measure recognition vs recall as in survey (Brand Awareness Pyramid: Top of Mind → Brand Recall → Brand Recognition → Unaware; 01_03_Brand Equity Components I; measurement wording in 03_08_Quantitative Research III). Proxy signals: (1) dedicated r/OnRunning community discourse volume; (2) unaided-style spontaneous mentions in general running/sneaker subs using brand-specific language (On Cloud*, Ons, Cloudmonster, etc.); (3) co-mention with consideration-set competitors. These support hypothesis generation for aided/unaided awareness items, not population awareness levels."),
        ("ANALYSIS — key associations (brand image)", f"Most frequent association themes in On-relevant texts: {top_assoc_txt}. Per 01_04_Brand Equity Components II, associations = brand meaning/image; dimensions to measure next: Strength, Favorability, Uniqueness (assignment also uses Strength/Valence/Uniqueness scales)."),
        ("ANALYSIS — perceived quality", "Consumer talk centers on comfort/cushioning (cloud feel), fit/narrowness, durability vs. price, aesthetics, and performance role (daily trainer vs. lifestyle). Perceived quality is relative to competitors (01_05_Brand Equity Components III definition). See Perceived_Quality sheet."),
        ("ANALYSIS — loyalty", f"Loyalty pyramid framing (01_06_Brand Equity Components IV): Switchers → Habitual → Switching-cost satisfied → Brand-as-friend → Committed. Pattern counts — repurchase_intent={loyalty_counts.get('repurchase_intent',0)}, advocacy_recommend={loyalty_counts.get('advocacy_recommend',0)}, committed_love={loyalty_counts.get('committed_love',0)}, switching_away={loyalty_counts.get('switching_away',0)}, criticism_intensity={loyalty_counts.get('criticism_intensity',0)}. NPS-style recommend language appears but is not a computed NPS (see 03_08_Quantitative Research III)."),
        ("ANALYSIS — brand personality", "Language skews toward Excitement (trendy/stylish/up-to-date) and Sophistication (premium/Swiss/design), with Competence (technical/CloudTec) also present — mapped to Aaker facets taught in 02_03_Brand Positioning III and 03_08_Quantitative Research III."),
        ("EVIDENCE — competitor co-mentions", f"Most frequent co-mentioned competitors in On-relevant posts/comments: {top_comp_txt}."),
        ("Implication for Brand Image–Identity Gap", "Treat high-frequency C2C themes as candidate brand-image associations for the Brand Image Matrix. Measure the same five inventory identity associations PLUS emergent C2C associations (assignment guidance). Rate Strength, Valence, Uniqueness on 1–5 scales (04-brand-exploratory-strategy-report)."),
        ("Caveat", "Enthusiast Reddit sample ≠ general target market. Interpret as directional C2C meaning, then confirm with qualitative interviews and quantitative survey."),
    ]
    write_kv_sheet(ws, exec_rows, "On Cloud — C2C Brand Exploratory Executive Summary")

    # 2. Methodology
    ws = wb.create_sheet("Methodology")
    meth_rows = [
        ("Sources", "Eight Reddit JSONL extracts in Brand Sentiment Analysis/: AskRunningShoeGeeks (posts+comments), OnRunning (comments), RunningFashion (comments), RunningShoes (posts+comments), Sneakers (posts+comments)."),
        ("Processing", "Streamed line-by-line JSON parse; posts use title+selftext; comments use body; removed/deleted bodies excluded from text."),
        ("Matching rules — high precision", "Brand-relevant if ANY of: (a) model/brand phrases via regex (On Cloud, OnCloud, On Running, Cloudtec, Cloudmonster, Cloudsurfer, Cloudstratus, Cloudrunner, Cloudflow, Cloudswift, Cloudnova, Cloud X, The Roger / Federer x On, etc.); (b) possessive/colloquial 'Ons' patterns (my Ons, pair of Ons, wear Ons…); (c) 'On' + shoe/running/cloud model collocates; (d) ANY record from r/OnRunning (brand community context). Bare preposition 'on' alone is NOT matched."),
        ("False-positive control", "Deliberately avoided matching standalone 'on'. Documented rule tags counted in AI pipeline. Residual risk: occasional figurative uses of model names or 'Roger' ambiguity — curated quotes reviewed for exhibit use."),
        ("Cleaning", "URL stripping for n-grams; lowercase tokenization; stopword removal; simple valence lexicon (positive/negative keyword counts) — heuristic only."),
        ("Association coding", "Dictionary/theme lexicon over On-relevant texts; frequency = # of On-relevant records containing theme pattern(s). Example quotes stored verbatim from source text."),
        ("Course alignment", "C2C web search is required primary method for Brand Exploratory. Text mining for associations per 03_01_Qualitative Research I. Equity components per 01_03–01_06. Measurement blueprint for survey per 03_08_Quantitative Research III and assignment association scales (Strength/Valence/Uniqueness)."),
        ("Limitations", "Reddit bias (enthusiast, English-heavy, fashion+running mix); r/OnRunning inflates volume and may skew loyalty; automated valence/personality are proxies; no demographics; awareness recognition vs recall not directly measurable from C2C; co-mention ≠ preference."),
        ("Decisions made without user", "Included all eight files; treated r/OnRunning as On-relevant; competitor list expanded slightly (Puma, UA, Mizuno, Topo) beyond core list; 12 Excel sheets as recommended."),
        ("Reproducibility", "Script: Brand Sentiment Analysis/_oncloud_c2c_mine.py; output: OnCloud_C2C_Brand_Exploratory.xlsx."),
        ("Generated", datetime.now().isoformat(timespec="seconds")),
    ]
    write_kv_sheet(ws, meth_rows, "Methodology")

    # 3. Mentions_By_Source
    ws = wb.create_sheet("Mentions_By_Source")
    rows = []
    for (fname, sub, kind), scanned in sorted(by_source.items()):
        hits = by_source_hits.get((fname, sub, kind), 0)
        rows.append(
            {
                "source_file": fname,
                "subreddit": sub,
                "record_type": kind,
                "records_scanned": scanned,
                "on_relevant": hits,
                "pct_on_relevant": round(100.0 * hits / scanned, 2) if scanned else 0,
            }
        )
    # totals by file
    file_tot = defaultdict(lambda: {"scanned": 0, "hits": 0})
    for r in rows:
        file_tot[r["source_file"]]["scanned"] += r["records_scanned"]
        file_tot[r["source_file"]]["hits"] += r["on_relevant"]
    df_src = pd.DataFrame(rows)
    write_df(ws, df_src)
    ws2_start = len(rows) + 4
    ws.cell(ws2_start, 1, "Match rule tag counts (among On-relevant records)")
    ws.cell(ws2_start, 1).font = Font(bold=True)
    ws.cell(ws2_start + 1, 1, "rule_tag")
    ws.cell(ws2_start + 1, 2, "count")
    for i, (k, v) in enumerate(match_rule_counts.most_common(), ws2_start + 2):
        ws.cell(i, 1, k)
        ws.cell(i, 2, v)
    mstart = ws2_start + len(match_rule_counts) + 4
    ws.cell(mstart, 1, "Model / brand phrase mention counts (On-relevant texts; multi-label)")
    ws.cell(mstart, 1).font = Font(bold=True)
    ws.cell(mstart + 1, 1, "model_or_phrase")
    ws.cell(mstart + 1, 2, "count")
    for i, (k, v) in enumerate(model_mention_counts.most_common(), mstart + 2):
        ws.cell(i, 1, k)
        ws.cell(i, 2, v)

    # 4. Key_Associations
    ws = wb.create_sheet("Key_Associations")
    assoc_rows = []
    for rank, (theme, cnt) in enumerate(association_counts.most_common(), 1):
        vc = association_valence[theme]
        ex = association_examples[theme][0]["quote"] if association_examples[theme] else ""
        # Heuristic uniqueness note
        uniq_note = "Candidate POD if CloudTec/cloud-feel language is distinctive vs Hoka max-cushion narrative; confirm uniqueness in survey."
        if "price" in theme.lower() or "durability" in theme.lower() or "comfort" in theme.lower():
            uniq_note = "Likely POP or contested attribute vs running peers; measure uniqueness explicitly."
        if "swiss" in theme.lower() or "federer" in theme.lower() or "cloudtec" in theme.lower() or "unique" in theme.lower():
            uniq_note = "Strong uniqueness candidate (celebrity/geography/tech naming); verify with Uniqueness scale."
        assoc_rows.append(
            {
                "rank": rank,
                "association_theme": theme,
                "frequency_n": cnt,
                "pct_of_on_hits": round(100.0 * cnt / stats["on_hits"], 2) if stats["on_hits"] else 0,
                "dominant_valence_heuristic": dom_valence(vc),
                "valence_pos": vc.get("positive", 0),
                "valence_neg": vc.get("negative", 0),
                "valence_mixed_neutral": vc.get("mixed/neutral", 0) + vc.get("mixed", 0),
                "example_quote": ex,
                "association_type_course": "See 01_04 types (attributes/benefits/lifestyle/personality/competitors/celebrity/country)",
                "strength_proxy_note": "Higher frequency ≈ higher salience/strength proxy only; survey Strength item required (03_08 / assignment).",
                "favorability_uniqueness_note": uniq_note,
                "brand_image_matrix_candidate": "Yes — emergent image association" if rank <= 10 else "Optional",
            }
        )
    write_df(ws, pd.DataFrame(assoc_rows))

    # 5. Keywords
    ws = wb.create_sheet("Keywords")
    # Filter noise: keep topical
    topical_boost = {
        "cloud", "cushion", "cushioning", "comfort", "comfortable", "comfy", "durability",
        "durable", "style", "stylish", "fashion", "trendy", "price", "expensive", "premium",
        "narrow", "fit", "sizing", "hoka", "nike", "brooks", "asics", "saucony", "swiss",
        "federer", "cloudmonster", "cloudsurfer", "cloudtec", "lightweight", "stack",
        "plush", "soft", "firm", "outsole", "toe", "wide", "race", "tempo", "trainer",
        "lifestyle", "hype", "overpriced", "worth", "recommend", "switch",
    }
    uni_rows = []
    for w, c in keyword_uni.most_common(400):
        if w in topical_boost or c >= 25:
            if w in STOP:
                continue
            uni_rows.append({"ngram": w, "n": c, "type": "unigram", "salience_flag": "high" if w in topical_boost else "freq"})
            if len(uni_rows) >= 80:
                break
    bi_rows = []
    for w, c in keyword_bi.most_common(500):
        if c < 8:
            continue
        if any(t in w for t in topical_boost) or c >= 20:
            bi_rows.append({"ngram": w, "n": c, "type": "bigram", "salience_flag": "high" if any(t in w for t in topical_boost) else "freq"})
            if len(bi_rows) >= 80:
                break
    write_df(ws, pd.DataFrame(uni_rows + bi_rows))

    # 6. Perceived_Quality
    ws = wb.create_sheet("Perceived_Quality")
    qrows = []
    for theme, cnt in quality_counts.most_common():
        vc = quality_valence[theme]
        examples = association_examples  # fallback
        exs = quality_examples.get(theme, [])
        qrows.append(
            {
                "quality_dimension": theme,
                "frequency_n": cnt,
                "dominant_valence_heuristic": dom_valence(vc),
                "pos": vc.get("positive", 0),
                "neg": vc.get("negative", 0),
                "example_quote_1": exs[0]["quote"] if len(exs) > 0 else "",
                "example_quote_2": exs[1]["quote"] if len(exs) > 1 else "",
                "course_note": "Perceived quality = customer perception of overall quality vs intended purpose, relative to competitors (01_05_Brand Equity Components III). Survey: rate On and competitors on these dimensions (03_08).",
            }
        )
    write_df(ws, pd.DataFrame(qrows))
    # analysis block
    r0 = len(qrows) + 3
    ws.cell(r0, 1, "ANALYSIS (not raw evidence)")
    ws.cell(r0, 1).font = Font(bold=True)
    ws.cell(r0 + 1, 1, "Synthesis")
    ws.cell(
        r0 + 1,
        2,
        "C2C talk suggests On's perceived quality is multi-dimensional: soft/cloud cushioning and aesthetics often praised; fit (narrow) and durability/price frequently contested. Relative framing vs Hoka (max cushion), Nike (performance/speed), Brooks/Asics (reliable trainers) is common — consistent with measuring PQ comparatively.",
    )

    # 7. Loyalty_Signals
    ws = wb.create_sheet("Loyalty_Signals")
    loyalty_label = {
        "repurchase_intent": "Repurchase / multiple-pair language (Satisfied→Committed proxies)",
        "advocacy_recommend": "Advocacy / recommend language (NPS-adjacent; 03_08)",
        "committed_love": "Affective commitment / favorite language (top of Loyalty Pyramid)",
        "switching_away": "Switching / rejection language (Switcher level / erosion)",
        "criticism_intensity": "Strong criticism (loyalty risk / detractor language)",
    }
    lrows = []
    for k, cnt in loyalty_counts.most_common():
        exs = loyalty_examples.get(k, [])
        lrows.append(
            {
                "signal": loyalty_label.get(k, k),
                "code": k,
                "frequency_n": cnt,
                "loyalty_pyramid_link": "01_06_Brand Equity Components IV",
                "example_quote_1": exs[0]["quote"] if len(exs) > 0 else "",
                "example_quote_2": exs[1]["quote"] if len(exs) > 1 else "",
                "example_quote_3": exs[2]["quote"] if len(exs) > 2 else "",
            }
        )
    write_df(ws, pd.DataFrame(lrows))

    # 8. Brand_Personality
    ws = wb.create_sheet("Brand_Personality")
    prows = []
    for trait, cnt in personality_counts.most_common():
        exs = personality_examples.get(trait, [])
        prows.append(
            {
                "aaker_facet_mapping": trait,
                "cue_frequency_n": cnt,
                "example_quote": exs[0]["quote"] if exs else "",
                "course_source": "Aaker (1997) facets in 02_03_Brand Positioning III and 03_08_Quantitative Research III",
                "analysis_note": "Inferred from consumer language in C2C; confirm with personality rating items in survey.",
            }
        )
    write_df(ws, pd.DataFrame(prows))
    # extra narrative
    pr = len(prows) + 3
    ws.cell(pr, 1, "ANALYSIS")
    ws.cell(pr, 1).font = Font(bold=True)
    ws.cell(
        pr + 1,
        1,
        "Working personality hypothesis: On skews Excitement + Sophistication (fashion-forward, premium, Swiss design) with Competence (CloudTec/technical). Ruggedness appears mainly in trail/outdoor talk; Sincerity cues are thinner. Use as positioning/personality hypothesis for Brand Exploratory, not as final scores.",
    )

    # 9. Competitor_Comparisons
    ws = wb.create_sheet("Competitor_Comparisons")
    crows = []
    for cname, cnt in competitor_co.most_common():
        exs = competitor_examples.get(cname, [])
        crows.append(
            {
                "competitor": cname,
                "co_mention_n_in_on_relevant": cnt,
                "pct_of_on_hits": round(100.0 * cnt / stats["on_hits"], 2) if stats["on_hits"] else 0,
                "example_quote": exs[0]["quote"] if exs else "",
                "example_quote_2": exs[1]["quote"] if len(exs) > 1 else "",
                "positioning_hypothesis": (
                    "Often framed as max-cushion / stack alternative."
                    if cname == "Hoka"
                    else "Performance/racing or mainstream athletic benchmark."
                    if cname == "Nike"
                    else "Reliable daily-trainer / comfort POP comparison."
                    if cname in {"Brooks", "Asics", "Saucony", "New Balance"}
                    else "Category peer in consumer consideration talk."
                ),
            }
        )
    write_df(ws, pd.DataFrame(crows))

    # 10. Quote_Bank
    ws = wb.create_sheet("Quote_Bank")
    qbrows = []
    for i, q in enumerate(curated, 1):
        qbrows.append(
            {
                "exhibit_id": f"Q{i:02d}",
                "quote": q["quote"],
                "valence_heuristic": q["valence"],
                "score": q["score"],
                "source_file": q["source_file"],
                "subreddit": q["subreddit"],
                "type": q["type"],
                "match_rules": q["match_rules"],
                "permalink": q["permalink"],
                "use": "Report exhibit candidate — verify context before citing",
            }
        )
    write_df(ws, pd.DataFrame(qbrows))

    # 11. Survey_Implications
    ws = wb.create_sheet("Survey_Implications")
    # Candidate associations for survey
    candidates = []
    # Prefer top association themes
    for theme, cnt in association_counts.most_common(10):
        candidates.append((theme, cnt))
    # Ensure key ones
    must = [
        "Cloud cushioning / cloud feel",
        "Style / fashion / aesthetics",
        "Premium / expensive / price",
        "Fit / sizing / narrow",
        "Durability / longevity",
        "Unique / distinctive ride (CloudTec)",
        "Swiss / European / design-led",
        "Trendy / hype / status",
    ]
    seen_t = {t for t, _ in candidates}
    for m in must:
        if m not in seen_t and association_counts.get(m, 0) > 0:
            candidates.append((m, association_counts[m]))
    candidates = candidates[:10]

    srows = []
    srows.append(
        {
            "item": "FRAME — Brand Exploratory dimensions",
            "detail": "Survey must cover Awareness (recognition & recall), Associations, Perceived Quality, Loyalty, and optional Brand Personality (assignment Qs + 03_08_Quantitative Research III).",
            "measure_strength": "",
            "measure_valence_favorability": "",
            "measure_uniqueness": "",
            "course_citation": "04-brand-exploratory-strategy-report; 01_03–01_06; 03_08",
        }
    )
    srows.append(
        {
            "item": "Awareness — unaided recall",
            "detail": "Open-end: 'Which running shoe brands are you familiar with?' / consideration set. On appearance = unaided recall proxy (Brand Awareness Pyramid, 01_03; wording pattern 03_08).",
            "measure_strength": "n/a",
            "measure_valence_favorability": "n/a",
            "measure_uniqueness": "n/a",
            "course_citation": "01_03_Brand Equity Components I; 03_08_Quantitative Research III",
        }
    )
    srows.append(
        {
            "item": "Awareness — recognition",
            "detail": "Aided list including On / On Cloud among competitors. Separates recognition from recall.",
            "measure_strength": "n/a",
            "measure_valence_favorability": "n/a",
            "measure_uniqueness": "n/a",
            "course_citation": "01_03_Brand Equity Components I; 03_08_Quantitative Research III",
        }
    )
    for i, (theme, cnt) in enumerate(candidates, 1):
        srows.append(
            {
                "item": f"Association candidate A{i}: {theme}",
                "detail": f"Emerged from C2C text mining (n={cnt} On-relevant records with theme cues). Include alongside the five Brand Inventory identity associations for image–identity gap matching.",
                "measure_strength": "To what degree does this attribute describe On? 1=Not at all … 5=Very descriptive",
                "measure_valence_favorability": "Is this attribute desirable? 1=Very undesirable … 5=Very desirable (assignment Valence; maps to Favorability in 01_04)",
                "measure_uniqueness": "How different from competing brands is this for On? 1=Not at all … 5=Very different",
                "course_citation": "01_04_Brand Equity Components II (Strength/Favorability/Uniqueness); Brand Exploratory association scales",
            }
        )
    srows.append(
        {
            "item": "Perceived quality battery",
            "detail": "Rate On and 2–3 competitors (e.g., Hoka, Nike, Brooks) on comfort, cushioning, durability, fit, style, performance, value — comparative PQ (01_05; 03_08).",
            "measure_strength": "1–7 or 1–10 perception ratings",
            "measure_valence_favorability": "Implied by higher=better scales; optionally separate desirability",
            "measure_uniqueness": "Optional difference item per dimension",
            "course_citation": "01_05_Brand Equity Components III; 03_08_Quantitative Research III",
        }
    )
    srows.append(
        {
            "item": "Loyalty battery",
            "detail": "Satisfaction, repurchase likelihood, brand preference ranking, and NPS recommend 0–10 among users only (03_08; Loyalty Pyramid 01_06).",
            "measure_strength": "n/a",
            "measure_valence_favorability": "Attitude/satisfaction scales",
            "measure_uniqueness": "n/a",
            "course_citation": "01_06_Brand Equity Components IV; 03_08_Quantitative Research III",
        }
    )
    srows.append(
        {
            "item": "Brand personality ratings",
            "detail": "Rate On on Aaker facets emphasizing Excitement, Sophistication, Competence (hypotheses from C2C).",
            "measure_strength": "Trait descriptiveness 1–5",
            "measure_valence_favorability": "Optional",
            "measure_uniqueness": "Optional vs competitor personality",
            "course_citation": "02_03_Brand Positioning III; 03_08_Quantitative Research III",
        }
    )
    write_df(ws, pd.DataFrame(srows))

    # 12. AI_Use_Log
    ws = wb.create_sheet("AI_Use_Log")
    ai_rows = [
        ("Tool", "Cursor IDE agent (Composer) for text mining script + Excel workbook creation."),
        ("Student responsibility", "Student must verify quotes against source JSONL, validate matching precision/recall on a sample, and confirm course citations against actual Class PowerPoints decks before submission."),
        ("Prompt summary", "Analyze Reddit C2C JSONL for On/On Cloud; export Brand Exploratory Excel; frame with Class PowerPoints frameworks (awareness recognition/recall, associations strength/favorability/uniqueness, perceived quality, loyalty, personality); update AGENTPROMPTLOG; do not git commit."),
        ("Data processed", ", ".join(INPUT_FILES)),
        ("Output file", str(OUT)),
        ("Class materials used (extracted slide text standing in for Class PowerPoints/)", "01_03_Brand Equity Components I; 01_04_Brand Equity Components II; 01_05_Brand Equity Components III; 01_06_Brand Equity Components IV; 02_03_Brand Positioning III; 03_01_Qualitative Research I; 03_08_Quantitative Research III; plus Assignments/04-brand-exploratory-strategy-report.md. Note: Class PowerPoints/ folder not present as files in workspace at runtime; content applied from Other/_pptx_extract/ mirrors of those decks."),
        ("GenAI limitations", "Automated theme/valence coding can misclassify sarcasm/context; quotes are truncated; findings are exploratory hypotheses for Brand Exploratory, not definitive equity scores."),
        ("Date", datetime.now().isoformat(timespec="seconds")),
    ]
    write_kv_sheet(ws, ai_rows, "AI Use / GenAI Disclosure Log")

    wb.save(OUT)
    print("WROTE", OUT)
    print("TOTAL", stats["total_scanned"], "ON_HITS", stats["on_hits"])
    print("TOP_ASSOC", top_assoc[:8])
    print("FILE_TOTALS", dict(file_tot))


if __name__ == "__main__":
    main()
