---
name: unit-mix-table
description: Builds an MLA Unit Mix Breakdown Table (comparables table, one row per project with unit types, counts, mix percents, and size ranges) from per-project floorplan screenshots — typically Avesdo/NHSLive "Floorplan Data" panels. Outputs Word or PowerPoint for the standard 5-column table; a rental variant (--variant rental, Word only) builds the 7-column "Unit Mix Comparables - Purpose-Built Rental" chart with gross rent and rental PSF. Use this skill whenever floorplan-data screenshots are dropped for one or more projects and someone asks for a unit mix table, unit mix breakdown, unit mix comparables, or "run the unit mix table". NOT the same as unit-mix-summary (which aggregates pasted unit mix data across projects into one statistical summary table) — this skill produces the per-project breakdown table; its output can feed unit-mix-summary. Also NOT for active tables, offerings, completions, upcoming, or amenity matrices — separate skills.
version: 2026-11-10 v2
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

- Two variants: `forsale` (default, the 5-column table, Word or PPT) and `rental` (7-column, Word only, see below).
- Word column order: PROJECT, UNIT TYPE, UNIT COUNT, UNIT MIX, SIZE RANGE. PPT order: PROJECT, FLOOR PLAN, SF, UNIT COUNT, UNIT MIX % — the script handles the swap.
- The Word gold renders as 2 pages (floating table layout); that matches the original template, not a bug.
- PPT slide footer text comes from the deck master — self-corrects when pasted into a project deck.
- This skill's output is a valid input to the existing `unit-mix-summary` skill (aggregated statistical summary across the comp set).

## Rental variant (purpose-built rental chart)

Use when the comp set is actively leasing purpose-built rental projects and the chart needs rents and rental PSF (for example the 888 Cambie / Vancouver West rental unit mix chart). Word only; `--format ppt --variant rental` exits with an error.

Columns: PROJECT, FLOORPLAN, UNITS, UNIT MIX, SQFT, GROSS RENT RANGE, GROSS RENTAL PSF. One row per project; one paragraph per floorplan line in each cell; UNITS ends with a bold `N TOTAL`, UNIT MIX with a bold `100%`. Project name bold over developer, all caps.

Input JSON (see `examples/example_input_rental.json`). `rent` and `psf` are free-text strings used exactly as given (`"$2,400+"`, `"$2,071 - $2,150"`, `"-"`):

```json
{"title": "Unit Mix Comparables - Purpose-Built Rental",
 "intro": "...",
 "projects": [{"name": "W71", "developer": "Birmingham & Wood",
   "units": [{"type": "1 Bed", "count": 25, "min_sf": 446, "max_sf": 570,
              "rent": "$2,000 - $2,350", "psf": "$4.12 - $4.48"}]}]}
```

Run:

```
python -I scripts/build_unit_mix_table.py --data <rental.json> --format word --variant rental --output <Rental Unit Mix Chart.docx>
```

Intro boilerplate (swap the market; keep the last sentence only when initial lease-up ranges were used):
> The table below highlights the size ranges, unit mixes, and rental rates for actively leasing purpose-built rental developments in [market]. The following trends observed are aimed at providing context to the unit mix programming of the most recently leasing projects in this market. Where current asking rents were unavailable for select unit types due to limited availability, pricing ranges from the project's initial lease-up program were used based on publicly available information.

Pricing conventions:
- Use what is posted. Where only "from" pricing exists, show `start+` for rent and PSF+ = start rent divided by the row's minimum SF.
- Where current asking rents are unavailable, use the project's initial lease-up ranges (and keep the intro sentence saying so).
- A PSF range is min rent / min SF to max rent / max SF, shown low to high.
- SQFT is `min - max` or a single number, with thousands separators. Use plain hyphens everywhere (the brand gate rejects en and em dashes), including the title.

Layout rules (built into the script):
- **Each page is its own chart.** The heading repeats at the top of every page, the intro paragraph appears on page 1 only, and each page has its own table with its own header row. Page 2+ headings carry pageBreakBefore. Projects are grouped greedily by line count (plan lines plus one TOTAL line per project); capacities are the constants `CAP_FIRST_PAGE = 26` (page 1, with intro) and `CAP_NEXT_PAGE = 31` at the top of the rental block in the script. If a project count or font change makes a page spill, adjust them and re-render.
- **Banding starts grey on every page.** The header row is white, so project rows alternate F2F2F2, FFFFFF, F2F2F2 and restart grey at the top of each page. Fills are explicit (`w:shd clear/auto/fill`) with themeFill and themeFillShade removed; the source chart's "white" rows carried themeFillShade=F2, which Word draws grey.
- Body rows carry no fixed row height and have 85 dxa top and bottom cell margins.
- The landscape section is the document's final section (no trailing blank portrait page).
- Rows keep the order given in the JSON (newest occupancy first in the 888 Cambie chart).
- Stdlib only (no python-docx).

Verify the rental output: export to PDF with Word (`ExportAsFixedFormat`, format 17), render the pages, and check every page is landscape with the heading, the intro only on page 1, a grey first row on each page then alternating, no blank page, and no table spilling across a page break. The 11-project 888 Cambie set renders as 3 pages.

## Files

- `templates/unit_mix_table_template.docx` — gold Word (landscape, Plain-style banded table)
- `templates/unit_mix_table_template.pptx` — gold PowerPoint slide
- `templates/unit_mix_table_rental_template.docx` — gold rental Word (landscape, 7 columns, one prototype row, grey fill)
- `scripts/build_unit_mix_table.py` — the builder (stdlib only; `--variant forsale|rental`)
- `examples/example_input.json` — sample input built from a real Avesdo floorplan screenshot
- `examples/example_input_rental.json` — rental variant sample (W71, Maple+41, Westward)
