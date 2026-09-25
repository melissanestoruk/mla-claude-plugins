---
name: unit-mix-overview
description: Builds an MLA Unit Mix Overview page (Word, landscape) - the transposed aggregate summary with plan types as columns (STUDIO through PH) and INVENTORY, MIN, MAX, AVERAGE, MEDIAN as rows, one or more tables per page (e.g. actively leasing vs recently stabilized), each with a blurb and a grouping footnote. Use this skill whenever someone asks for a unit mix overview, unit mix summary page, plan-type summary table, aggregate unit mix stats across a comp set, or says "make this page based on those unit mix breakdowns" - even if they don't name the skill. It aggregates the per-project output of the unit-mix-table skill. NOT the same as unit-mix-summary (portrait per-type statistical table, different layout) or unit-mix-table (per-project breakdown rows).
version: 2026-09-25 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Unit Mix Overview

Aggregate per-project unit mix breakdowns (typically the unit-mix-table skill's output) into the transposed overview page: plan types across the top, five stat rows down the side, one table per comp set. The script does the math and the Word COM generator does the formatting - your job is the grouping and the prose.

## Step 1: Group the plan types (judgment step - do this carefully)

Collect every plan row (count, min SF, max SF) from the source breakdowns, then map each into a column group. The locked MLA Advisory grouping rules:

- **All penthouse plans combine into one PH column**, whatever the bed count, including PH + den layouts. Penthouse identity beats every other rule (a PH-3+Den goes to PH, not 3 BED).
- **Flex layouts are categorized with dens** (1 BED + FLEX joins 1 BED + DEN), except where the junior rule below takes the row first.
- **Two-bedroom, one-bath layouts - including flex variants - are junior two-bedrooms** (JR 2 BED).
- **Studio + den plans fold into STUDIO.**
- **Low-volume den variants fold into their base type** when they'd make a near-empty column (e.g. 3 BED + DEN into 3 BED); keep them separate only when there's real volume.
- **Townhome units are excluded** from the chart entirely.
- Honour any comp-set exclusions the user has made elsewhere in the document (e.g. projects under 50 units) so the pages agree with each other.

Keep both tables on a page to the same column set when possible - it makes them comparable at a glance.

## Step 2: Write the input JSON

One `sections` entry per table. `groups` maps each column to its plan rows as `[count, min_sf, max_sf]`. Set `expected_total` to the comp set's known unit count (minus exclusions) - the script refuses to run if the grouping doesn't sum to it, which catches dropped or double-counted rows.

```json
{
  "title": "Active and Stabilized Developments - Unit Mix Overview",
  "intro": "The following charts contain unit mix comparables from ...",
  "sections": [
    {
      "label": "UNIT MIX SUMMARY: ACTIVELY LEASING CONCRETE RENTALS*",
      "blurb": "The table below provides an overview of ...",
      "footnote": "*Certain plan types have been grouped together for simplicity: ...",
      "order": ["STUDIO", "1 BED", "1 BED + DEN", "JR 2 BED", "2 BED", "2 BED + DEN", "3 BED", "PH"],
      "expected_total": 1041,
      "groups": {"STUDIO": [[10, 394, 430], [67, 370, 411]], "1 BED": [[31, 455, 455]]}
    }
  ]
}
```

## Step 3: Run the two scripts

```
python scripts/aggregate_overview.py --data grouped.json --output render.json
powershell -File scripts/generate_overview.ps1 -DataPath render.json -OutputPath "<Name> Unit Mix Overview.docx"
```

The aggregator computes: INVENTORY as count share (round half up, `<1%` under 1 percent), MIN/MAX as pooled extremes, AVERAGE as the count-weighted average of plan midpoints, MEDIAN as the unweighted median of pooled min/max values - averages and medians rounded to the nearest 5, matching the unit-mix-summary methodology. The generator renders Butler title, Montserrat Light intro/blurbs, banded stat rows, the light-blue MEDIAN row, and a footnote per table. No yellow anywhere, hyphens never en dashes. Word must be installed (COM).

## Step 4: Blurbs and footnotes must describe what actually happened

The footnote is a factual record of the grouping, not boilerplate. After grouping, walk each clause and confirm it happened in this data set (if there was exactly one 3 BED + DEN plan, say "the three-bedroom + den layout", singular; if projects were excluded by size, the footnote ends "...Townhome units and projects under 50 units have been excluded from the chart."). The blurb states market-level observations - concentration by plan type, how sizes step up - and each claim must be checkable against the printed table. When the underlying comp set changes, re-verify both.

## Step 5: Verify before delivering

1. Each section's grouped total = expected_total (the script enforces this - never lower expected_total to make it pass; find the missing row).
2. Spot-check one column's MIN/MAX against the source breakdown.
3. Every blurb and footnote claim reconciles with the final table.

One canonical file per artifact - no version suffixes. Deliver with a one-line note on location.

## Files

- `scripts/aggregate_overview.py` - stats aggregator (stdlib only)
- `scripts/generate_overview.ps1` - Word COM page generator (landscape)
- `examples/example_input.json` - the Vancouver West active + stabilized page this skill was built from
