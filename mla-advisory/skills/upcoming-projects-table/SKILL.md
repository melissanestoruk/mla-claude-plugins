---
name: upcoming-projects-table
description: Builds an MLA Upcoming Project Table from a HubSpot CRM export of pipeline/upcoming projects — the standard Advisory table of projects launching soon in a market. One layout for for-sale and rental product. Word columns; PROJECT, DEVELOPER/APPLICANT, UNITS, STATUS/TIMING. PowerPoint adds a STORIES column. STATUS/TIMING combines pipeline stage and estimated launch ("Approved / Est. Oct 2026"); rows sort soonest launch first with a TOTAL units row. Use this skill whenever a HubSpot pull of upcoming, pipeline, in-planning, or approved-but-unlaunched projects is dropped and someone asks for an upcoming project table, upcoming launches table, pipeline table, or "run the upcoming table". NOT for actively selling projects (active-table), unit mix tables, offerings, completions, or amenity matrices — separate skills.
version: 2026-07-07 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Upcoming Projects Table

Turn a HubSpot export of pipeline projects into the MLA Upcoming Project Table, cloned from the locked gold templates. The market is stated in the title/header — there is no per-row location column.

## Step 1: Identify the inputs

1. **The HubSpot pull** — CSV export, normally pre-filtered by market (e.g. upcoming highrise launches in one market).
2. **The format** — Word or PowerPoint. **Always ask** unless stated.
3. **Title** (PPT only) — e.g. "Upcoming Concrete Condominium – Coquitlam". The Word table has no title paragraph; it gets pasted under a heading in the host document.

If the pull covers many markets and PowerPoint is requested, suggest filtering to one market first — more than ~13 rows overflows the slide (the script warns but still builds).

## Step 2: Run the builder

```
python scripts/build_upcoming_table.py --csv <pull.csv> --format word|ppt \
    [--title "<Slide Title>"] --output <Upcoming Project Table.docx|.pptx>
```

The script (stdlib only):
- Maps columns flexibly: Name → PROJECT, Developer → DEVELOPER/APPLICANT, Total Units → UNITS, Storeys → STORIES (PPT only), Contract Writing Date + Project pipeline stage → STATUS/TIMING.
- STATUS/TIMING: "Approved / Est. Oct 2026" or "In Planning / Est. Mar 2027" ("Approved/Upcoming" displays as "Approved").
- Sorts by estimated launch date, soonest first.
- Blank developers ("-" or empty) and blank storeys render as "-".
- TOTAL row sums units.

## Step 3: Verify before delivering

1. Row count = pull row count (script prints it).
2. Spot-check one row's units and date against the CSV.
3. TOTAL = sum of the units column.

Deliver with a one-line note on file location. One canonical file per artifact — no version suffixes.

## Known behaviours and manual steps

- **Asterisks and footnotes** (e.g. `**Leasehold`) are manual post-generation annotations — the skill never adds them, and it blanks the gold's example footnotes.
- **No location column by design**: the golds had street addresses, but HubSpot pulls don't carry them and the market is obvious from the table's title/header (decision 2026-07-07).
- Long pulls: Word handles any length (multi-page table); PPT warns above 15 rows.
- PPT slide footer text comes from the deck master — self-corrects on paste.

## Files

- `templates/upcoming_table_template.docx` — gold Word (location column removed)
- `templates/upcoming_table_template.pptx` — gold PowerPoint slide (address column removed)
- `scripts/build_upcoming_table.py` — the builder
- `examples/hubspot_pull_example.csv` — real pull (upcoming highrise launches, 103 projects)
