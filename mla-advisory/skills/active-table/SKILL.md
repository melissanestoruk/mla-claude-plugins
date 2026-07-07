---
name: active-table
description: Builds an MLA Active Chart (competition analysis table) from a HubSpot CRM export — the standard Advisory comparable table of actively selling or actively leasing projects. Two variants; FOR SALE (PROJECT, DEVELOPER, LAUNCH, AVG SF, PPSF, SOLD, RLSD, TOTAL with % sold) outputs Word or PowerPoint; RENTAL (PROJECT, DEVELOPER, MARKET, TYPE, STOREYS, OCCUPANCY, AVG SF, AVG RENT, INITIAL PSF, LEASED, LEASES/MONTH, TOTAL with % leased) outputs Word only. Use this skill whenever a HubSpot pull (CSV export) of comparable projects is dropped and someone asks for an active table, active chart, competition analysis table, comp table of actively selling projects, or "run the active table" — even if they don't name the skill. NOT for unit mix tables (unit-mix-table), upcoming project tables, offerings tables, completions tables, or amenity matrices — those are separate skills.
version: 2026-07-07 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Active Table

Turn a HubSpot CRM export of comparable projects into a formatted MLA Active Chart, cloned from the locked gold templates. The script does all the math and formatting — never rebuild the table by hand.

## Step 1: Identify the inputs

You need three things from the user:

1. **The HubSpot pull** — a CSV export dropped into chat or a file path.
2. **The variant** — for sale or rental. Infer from the CSV columns (a rental pull has rent/lease columns; a for-sale pull has Units Sold / PPSF). If ambiguous, ask.
3. **The format** — Word or PowerPoint. **Always ask which one they want** unless they already said. Rental is Word-only; if PowerPoint is requested for rental, say there is no rental PPT gold template yet and produce Word.

Also ask for (or infer from the pull/filename) the **chart title**, e.g. "Coquitlam Condominium Competition Analysis". If the user gives a market name only, compose the title as "[Market] [Product] Competition Analysis".

If the pull is pasted as text rather than attached as a file, save it to a temp CSV first, preserving the header row exactly.

## Step 2: Run the builder

```
python scripts/build_active_chart.py \
    --csv <pull.csv> \
    --variant forsale|rental \
    --format word|ppt \
    --title "<Chart Title>" \
    --output <Project Name Active Chart.docx|.pptx>
```

The script (stdlib only, no pip installs):
- Matches CSV columns flexibly (e.g. "Contract Writing Date" → LAUNCH, "Avg Unit Size Estimate" → AVG SF, "PPSF Estimate" → PPSF). It errors listing found headers if a required column is missing — show that error to the user and ask which column maps.
- Sorts rows by launch date (for sale) or occupancy date (rental), newest first.
- Formats: launch/occupancy as "Mar 2026"; PPSF/rent as $1,100; PSF rent as $5.30; blanks as "-".
- TOTAL column per row: `total (sold÷total %)` — "166 (57%)", or "(-%)" when sold/leased is blank.
- TOTAL row: sums Sold/RLSD/Total (for sale) or sum Leased / average Leases-per-month / sum Total (rental), with the overall percent.
- Word: clones the gold template rows so banding, borders, Montserrat Light, and the MLA header/footer are preserved untouched.
- PPT: fills the gold slide's table, swaps the title, and keeps only the generic `*Approximate PPSF...` footnote.

## Step 3: Verify before delivering

1. Row count in the output = row count in the pull.
2. Spot-check one row's numbers against the CSV.
3. Confirm the TOTAL row percent = total sold ÷ total units.

Deliver the file with a one-line note on where it was saved. One canonical file per artifact — no version suffixes in filenames.

## Known behaviours and manual steps

- **Asterisks and project-specific footnotes** (e.g. `**Sales Centre Closed.`) are added MANUALLY by the operator after generation. The skill never adds or invents them.
- **TYPE column (PPT for-sale and rental)**: filled from a Construction Type / Building Type / Type column if the pull has one, otherwise "-" for manual fill.
- **PPT slide footer** (e.g. "Burnaby Market Analysis | May 2026") comes from the slide master of whatever deck the slide lands in — it corrects itself on paste; ignore it in the standalone file.
- **Estimated values** in rental pulls (text like "Est. $2,930") pass through as-is and are excluded from the TOTAL/AVG math.
- The rental column mapping is drafted against expected HubSpot property names (Submarket, Occupancy Date, Avg Rent Estimate, Initial PSF, Units Leased, Leases Per Month). First real rental pull may need mapping tweaks — see `_private/overlay.md`.

## Files

- `templates/active_chart_forsale_template.docx` — gold Word, for-sale chart (portrait)
- `templates/active_chart_rental_template.docx` — gold Word, rental chart (landscape)
- `templates/active_chart_forsale.pptx` — gold PowerPoint slide, for-sale chart
- `templates/active_charts_gold.docx` — original combined gold document (reference only)
- `scripts/build_active_chart.py` — the builder
- `examples/hubspot_pull_forsale_example.csv` — real HubSpot export (Coquitlam)
- `examples/hubspot_pull_rental_example.csv` — synthetic rental pull matching the expected schema
