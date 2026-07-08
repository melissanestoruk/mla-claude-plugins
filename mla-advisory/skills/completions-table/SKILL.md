---
name: completions-table
description: Builds an MLA Completions Table from a HubSpot CRM export of projects completing in a market — PROJECT, TYPE (CON/WF), DEVELOPER, LAUNCH, SALES, UNITS, EST. COMPLETION (quarter), sorted soonest completion first with a TOTAL row. Word output is the table; PowerPoint output is TWO charts — slide 1 the table, slide 2 a Sold/Unsold by completion-quarter bar chart rebuilt from the same pull. Use this skill whenever a HubSpot pull of completing/completions projects is dropped and someone asks for a completions table, upcoming completions, completions chart, yearly completions, or "run the completions table". NOT for actively selling comp tables (active-table), pipeline launches (upcoming-projects-table), unit mix, offerings, or amenity matrices — separate skills.
version: 2026-07-07 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Completions Table

Turn a HubSpot export of completing projects into the MLA Completions Table — and, in PowerPoint, the companion Sold/Unsold by quarter bar chart — cloned from the locked gold templates.

## Step 1: Identify the inputs

1. **The HubSpot pull** — CSV export with Name, Developer, Construction Material, Contract Writing Date, Units Sold, Total Units, Completion Date.
2. **The format** — Word or PowerPoint. **Always ask** unless stated. PowerPoint produces both slides (table + bar chart) in one file.
3. **Title** (PPT) — e.g. "Upcoming Coquitlam Completions"; applied to both slides.

## Step 2: Run the builder

```
python scripts/build_completions_table.py --csv <pull.csv> --format word|ppt \
    [--title "<Slide Title>"] --output <Completions Table.docx|.pptx>
```

The script (stdlib only):
- TYPE from Construction Material: Concrete → CON, Wood Frame → WF.
- LAUNCH from Contract Writing Date ("Sep 2024"); missing dates render "TBD".
- EST. COMPLETION as a quarter ("Q3 2026") from Completion Date.
- Sorts by completion date, soonest first.
- Word UNITS column shows "178 (63%)" (percent sold); PPT shows the plain number.
- TOTAL row sums sales and units (with overall percent in Word).
- PPT slide 2: rewrites the bar chart's cached data — categories become the pull's completion quarters, Sold = sum of Units Sold per quarter, Unsold = sum of (Total − Sold). The chart's external workbook link is not needed; the cache drives the render.

## Step 3: Verify before delivering

1. Row count = pull row count (script prints it).
2. Spot-check one project's sold/total/quarter against the CSV.
3. PPT: chart bar totals per quarter = sum of that quarter's UNITS column in slide 1.

Deliver with a one-line note on file location. One canonical file per artifact — no version suffixes.

## Known behaviours and manual steps

- Asterisk footnotes and chart annotations (e.g. "*Accounts for towers 3, 4, 6 and 7") are manual post-generation notes — the skill blanks the gold's examples.
- The Word gold carries a trailing blank landscape page (template structure) — delete manually if the deliverable is standalone.
- PPT warns above 16 projects (slide 1 table gets tall); filter the pull if needed.
- PPT slide footers come from the deck master — self-correct on paste.

## Files

- `templates/completions_table_template.docx` — gold Word
- `templates/completions_table_template.pptx` — gold PowerPoint (table slide + chart slide)
- `scripts/build_completions_table.py` — the builder
- `examples/hubspot_pull_example.csv` — real pull (Coquitlam 2026 completions, 16 projects)
