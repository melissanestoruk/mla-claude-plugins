---
name: stabilized-table
description: Builds an MLA Stabilized Rental Chart (Word, landscape) - the Advisory comparable table of stabilized purpose-built rental projects with columns PROJECT, DEVELOPER, MARKET, STOREYS, LEASING PERIOD, LEASES/MONTH, AVG SF, AVG RENT, INITIAL GROSS PPSF, UNITS, plus a companion writeup paragraph on lease-up velocity. Use this skill whenever someone asks for a stabilized rental chart, stabilized comps, a "recently stabilized" table, lease-up velocity comparison, leases per month by project, or drops a HubSpot pull or screenshot of stabilized/leased-up rental projects - even if they don't name the skill. NOT for actively leasing projects (active-table skill) or unit mix breakdowns (unit-mix-table skill).
version: 2026-09-25 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Stabilized Table

Turn a HubSpot pull or screenshot of stabilized rental comps into the MLA Stabilized Rental Chart - one row per project with its leasing period, calculated leasing velocity, and initial program stats. Cloned from the active-table rental gold template; the script does the math and formatting.

## Step 1: Collect the inputs

Per project: name, developer, submarket, storeys, **lease start and lease-up dates** (or month-level periods), average unit SF, average initial rent, initial gross PSF, and total units. Sources are HubSpot pulls, Avesdo screenshots, or values the user pastes or edits by hand - the user's own edited values always win over source data.

Rounding, when building from raw source data: SF to the nearest 5, rents to the nearest 5. If the user hands you already-rounded or already-chosen values, use them verbatim.

Sort rows by lease start date, newest first, unless the user orders them.

## Step 2: The leases-per-month convention (this is the one that matters)

**Leases/month = total units ÷ the INCLUSIVE month count of the printed leasing period.** "Jul 2025 - May 2026" is 11 months (count Jul through May on your fingers); 175 units ÷ 11 = 16/month.

Never compute velocity from day-level source dates when the chart prints month-level periods. The reason: a client reading the chart will multiply the printed period by the printed velocity, and if those don't reconcile the whole table loses credibility. The printed table must be self-consistent - that is the locked MLA Advisory standard. If the user later edits a leasing period, recompute that row's velocity from the new printed period (the script does this automatically on rebuild).

Exclusions: if the user drops projects (e.g. "anything under 50 units"), remove their rows, recompute the units TOTAL, and state the exclusion rule explicitly in the writeup.

## Step 3: Write the input JSON and run the builder

```json
{
  "title": "Recently Stabilized",
  "writeup": "The table below provides an overview of ...",
  "rows": [
    {"project": "The Raven", "developer": "Westbank", "market": "Point Grey",
     "storeys": 14, "period_start": "Oct 2025", "period_end": "Jun 2026",
     "avg_sf": 570, "avg_rent": 2680, "psf": 5.10, "units": 131}
  ]
}
```

```
python scripts/build_stabilized_chart.py --data input.json --output "<Market> Stabilized Rental Chart.docx"
```

The script computes leases/month per the convention above (pass `leases_month` on a row only to override), formats currency and the period, sums the UNITS total, handles any row count, and preserves the gold template's banding, borders, and fonts. It reuses helpers from the active-table skill, which must be installed.

## Step 4: Compose the writeup

The writeup lives above the table and every quantitative claim in it must divide out from the printed table. Structure:

1. Scope sentence: market, product type, time window covered by the lease-up dates, and any exclusion rule (e.g. "projects under 50 units were excluded from the comparable set").
2. Lease-up duration observation: name the larger projects and their duration range in months, from the printed periods.
3. Outlier note if one exists: a project far outside the pack (e.g. a 3-month lease-up) gets named with its actual figures, not glossed as "smaller" unless it truly is small.
4. Velocity summary: "Excluding [named outliers], leasing velocity ranged from approximately X to Y units per month, averaging approximately Z units per month."
5. Implied pace: "At this average pace, a 100-unit project would be expected to reach stabilization in approximately N months." (N = 100 ÷ the unrounded average.)

**Verify before delivering**: recompute the range, average, and implied pace from the final leases/month column and check each writeup number against them. Any edit to the table re-triggers this check. Use hyphens, never en dashes.

## Step 5: Verify before delivering

1. Row count in output = comp set size; TOTAL = sum of the units column.
2. Spot-check one row: printed period month count × printed leases/month ≈ printed units.
3. Every writeup number reconciles with the table (Step 4 check).

One canonical file per artifact - no version suffixes. Deliver with a one-line note on location.

## Files

- `templates/stabilized_chart_template.docx` - gold Word template (landscape; copy of the active-table rental gold, columns re-mapped by the script)
- `scripts/build_stabilized_chart.py` - the builder (stdlib only; imports helpers from the active-table skill)
- `examples/example_input.json` - the Vancouver West comp set this skill was built from
