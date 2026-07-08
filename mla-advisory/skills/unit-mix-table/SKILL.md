---
name: unit-mix-table
description: Builds an MLA Unit Mix Breakdown Table (comparables table, one row per project with unit types, counts, mix percents, and size ranges) from per-project floorplan screenshots — typically Avesdo/NHSLive "Floorplan Data" panels. Outputs Word or PowerPoint, same layout for sale and rental/condo product. Use this skill whenever floorplan-data screenshots are dropped for one or more projects and someone asks for a unit mix table, unit mix breakdown, unit mix comparables, or "run the unit mix table". NOT the same as unit-mix-summary (which aggregates pasted unit mix data across projects into one statistical summary table) — this skill produces the per-project breakdown table; its output can feed unit-mix-summary. Also NOT for active tables, offerings, completions, upcoming, or amenity matrices — separate skills.
version: 2026-07-07 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Unit Mix Table

Turn per-project floorplan screenshots into the MLA Unit Mix Breakdown Table — one table row per project, with unit types, counts, mix percents, and size ranges stacked inside each row. Cloned from the locked gold templates; the script does the math and formatting.

## Step 1: Collect the inputs

1. **One screenshot per project** — usually the Avesdo/NHSLive "FLOORPLAN DATA" panel (columns like Plan Type, Bths, # Stalls, Rlsd, %, Sold, Unsold, Min SF, Max SF, Min $, Max $). Multiple screenshots may arrive in one message.
2. **Project name and developer for each screenshot** — the panel usually doesn't show them. If not given, ask.
3. **Format** — Word or PowerPoint. **Always ask** unless already stated.
4. **Title** — e.g. "Unit Mix Comparables – Concrete Condominium" (Word heading and PPT title). For Word, also compose the intro sentence (see Step 3).

## Step 2: Extract the data from each screenshot (vision)

For each floorplan row, read:

- **type** — the Plan Type label, then normalize: expand abbreviations ("1 Bd" → "1 BED", "1+Den" → "1 BED + DEN", "2+Den" → "2 BED + DEN", "3+Den" → "3 BED + DEN"). Keep compound labels faithful ("2 Bed+Den+Flex" → "2 BED + DEN + FLEX", "PH – 2 BED", "TH 2 BED"). Everything renders uppercase.
- **count** — use the **Rlsd (released) column** as the unit count unless the user says to use a different column (e.g. total units). This matches how the gold tables were built.
- **min_sf / max_sf** — the Min SF and Max SF columns. Ignore stalls, baths, and all $ columns — pricing is not part of this table.

Skip rows with zero/blank counts unless the user wants them. Preserve the screenshot's row order within each project, and the order projects were provided in.

**Show the extracted numbers to the user in a quick text table before building** — vision extraction of screenshots must be verifiable. Then write the JSON input file:

```json
{
  "title": "Unit Mix Comparables – Concrete Condominium",
  "intro": "The table below highlights ...",
  "projects": [
    {"name": "Vesa at Little Mountain", "developer": "Holborn Developments",
     "units": [{"type": "1 Bed", "count": 25, "min_sf": 495, "max_sf": 592}]}
  ]
}
```

## Step 3: Compose title and intro

- Title pattern: `Unit Mix Comparables – [Product Type]` (Word) / `[Product Type] Unit Mix` works too for PPT — follow what the user asks for.
- Word intro boilerplate (swap the bracketed parts):
  > The table below highlights the size ranges, unit mixes, and pricing for active [product type]s in [market]. The following trends observed are aimed to provide context to the unit mix programming of the most recently selling projects in this market.
  Add the price-sheet caveat sentence only if the user says some pricing was unavailable.
- `intro` is Word-only; PPT has no intro paragraph.

## Step 4: Run the builder

```
python scripts/build_unit_mix_table.py --data <data.json> --format word|ppt --output <Unit Mix Table.docx|.pptx>
```

The script computes per-type mix percent (round half up; **"<1%"** for anything under 1 percent), the bold `N TOTAL` and `100%` lines, and en-dash size ranges (`495 – 592`; single number when min = max). Project names bold over developer names, all caps.

## Step 5: Verify before delivering

1. Type count per project in output = rows extracted from that screenshot.
2. Spot-check one project's counts and SF against the screenshot.
3. Percent column ends in a bold 100%; count column ends in the correct total.

Deliver with a one-line note on the file location. One canonical file per artifact — no version suffixes.

## Known behaviours

- Same table layout for for-sale and rental/condo product — no variants.
- Word column order: PROJECT, UNIT TYPE, UNIT COUNT, UNIT MIX, SIZE RANGE. PPT order: PROJECT, FLOOR PLAN, SF, UNIT COUNT, UNIT MIX % — the script handles the swap.
- The Word gold renders as 2 pages (floating table layout); that matches the original template, not a bug.
- PPT slide footer text comes from the deck master — self-corrects when pasted into a project deck.
- This skill's output is a valid input to the existing `unit-mix-summary` skill (aggregated statistical summary across the comp set).

## Files

- `templates/unit_mix_table_template.docx` — gold Word (landscape, Plain-style banded table)
- `templates/unit_mix_table_template.pptx` — gold PowerPoint slide
- `scripts/build_unit_mix_table.py` — the builder (stdlib only)
- `examples/example_input.json` — sample input built from a real Avesdo floorplan screenshot
