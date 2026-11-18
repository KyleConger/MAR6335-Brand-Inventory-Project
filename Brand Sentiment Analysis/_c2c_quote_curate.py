"""Write curated On-primary quotes into summary + Exhibit04."""
from __future__ import annotations

import json
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

BASE = Path(__file__).resolve().parent
OUT = BASE / "OnCloud_Exhibits" / "Exhibit04_Quote_and_Counts.xlsx"
SUMMARY = BASE / "_c2c_sample_summary.json"

PICKED = [
    {
        "subreddit": "runningshoes",
        "score": 84,
        "valence": "mixed",
        "associations": ["lifestyle_crossover", "running_performance", "style_aesthetic"],
        "quote": (
            "First Run: On Cloudmonster 3 Hyper They've finally done it. On has finally made a shoe "
            "that I actually want to run in. I have always been gleeful in my disdain for On, and my "
            "belief that they are just \"glorified Sketchers for people who work in offices but need "
            "their coworkers and/or customers to know they are athletic outside of work.\""
        ),
        "source_file": "r_runningshoes_posts.jsonl",
    },
    {
        "subreddit": "runningshoes",
        "score": 65,
        "valence": "mixed",
        "associations": ["running_performance", "durability", "swiss_innovation_tech"],
        "quote": (
            "First run: ON LightSpray Cloudmonster 3 Hyper I finally retired my Nike's Pegasus 41s "
            "after dragging them through over 2000km (~1250 miles) over the last eight months... "
            "switching to the LightSpray Cloudmonster 3 Hyper was obviously a night and day dif"
        ),
        "source_file": "r_runningshoes_posts.jsonl",
    },
    {
        "subreddit": "runningshoes",
        "score": 53,
        "valence": "neutral",
        "associations": ["lifestyle_crossover", "style_aesthetic"],
        "quote": "Ain't nobody hyping On Cloud 5s. It's just the default lifestyle shoe of suburbia.",
        "source_file": "r_runningshoes_comments.jsonl",
    },
    {
        "subreddit": "AskRunningShoeGeeks",
        "score": 49,
        "valence": "mixed",
        "associations": ["fit_sizing", "running_performance", "cushioning_cloud_ride"],
        "quote": (
            "On LS cloudmonster 3 hyper My underfoot gets numb, like bloodflow is bring restricted "
            "when running in these pair of shoes. Read online that its true to size but if its too wide, "
            "the recommendation would be to size up? Or totally forget about this shoe. Ps: the ride "
            "is really amazing until the numbness ruins it."
        ),
        "source_file": "r_askrunningshoegeeks_posts.jsonl",
    },
    {
        "subreddit": "runningshoes",
        "score": 39,
        "valence": "negative",
        "associations": ["cushioning_cloud_ride", "running_performance", "lifestyle_crossover"],
        "quote": (
            "What is the most overhyped shoe of all time? I'll start: On Cloud 5s. They are everywhere "
            "right now, but the pods constantly trap rocks and they feel way too stiff for actual "
            "distance running. Honorable mention goes to the Hoka Clifton 9..."
        ),
        "source_file": "r_runningshoes_posts.jsonl",
    },
    {
        "subreddit": "runningshoes",
        "score": 37,
        "valence": "positive",
        "associations": ["running_performance", "durability"],
        "quote": (
            "200 mile check-in: On Cloudmonster 3 Hyper (do we have a new \"champion shoe\" of 2026?) "
            "I ran a half marathon race in the Cloudmonster 3 Hypers this morning, and while doing so "
            "crossed the 200 mile threshold with these shoes."
        ),
        "source_file": "r_runningshoes_posts.jsonl",
    },
    {
        "subreddit": "runningshoes",
        "score": 30,
        "valence": "positive",
        "associations": ["comfort", "fit_sizing", "running_performance", "durability"],
        "quote": (
            "On Cloudboom Strike - Underrated I am on my third pair of On Cloudboom Strike... "
            "The shoe is super comfortable and has a good fit for people with regular/wide feet. "
            "I am usually using them for half marathons, 10k or 5k races. My first pair is almost "
            "at 600km..."
        ),
        "source_file": "r_runningshoes_posts.jsonl",
    },
    {
        "subreddit": "AskRunningShoeGeeks",
        "score": 17,
        "valence": "negative",
        "associations": ["value_price", "durability"],
        "quote": (
            "Why does it have to be on cloud? The brand is rarely recommended here. Not because "
            "there's anything necessarily wrong with it, but they tend to be more expensive and "
            "less durable than your average asics/adidas/nike etc"
        ),
        "source_file": "r_askrunningshoegeeks_comments.jsonl",
    },
    {
        "subreddit": "ONrunning",
        "score": 15,
        "valence": "neutral",
        "associations": ["lifestyle_crossover", "running_performance"],
        "quote": (
            "The only thing I'll say is that this model is not for running. On Cloud 5/6 etc are "
            "for walking or wearing to work. On make actual road and trail running shoes such as "
            "Cloudboom, Cloudultra, etc. You need to buy shoes for the correct activity"
        ),
        "source_file": "r_onrunning_comments.jsonl",
    },
]

summary = json.loads(SUMMARY.read_text(encoding="utf-8"))
summary["selected_quotes"] = PICKED
summary["quote_filter_note"] = (
    "Manually curated On-primary quotes from score-ranked brand-strict candidates; "
    "false-positive competitor reviews removed."
)
SUMMARY.write_text(json.dumps(summary, indent=2), encoding="utf-8")

wb = load_workbook(OUT)
ws = wb["Quote_Exhibit"]
if ws.max_row > 1:
    ws.delete_rows(2, ws.max_row - 1)
for i, q in enumerate(PICKED, 1):
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
wb.save(OUT)
print("curated", len(PICKED), "quotes")
