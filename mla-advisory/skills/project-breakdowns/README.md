# project-breakdowns

Builds a presale unit pricing analysis Excel workbook from floor plan images and price sheet data. The "Project Breakdown" deliverable used by MLA Advisory.

## What it does

Reads building plan thumbnail images to count units per type, combines with pricing data, and outputs a two-sheet SUMPRODUCT-weighted pricing summary workbook (.xlsx).

## When to use

Trigger on: "project breakdown for [project]", "pricing summary", "count units from floor plans", "weighted average PPSF", "pricing analysis Excel", "PPSF breakdown", or when floor plan images are provided alongside pricing or square footage data. Also triggers on "do a sumproduct" on pricing or size data.

## Outputs

- `[ProjectName]_Pricing_Summary.xlsx`
  - **Summary** sheet — one row per unit type, Weighted Averages footer block: SUMPRODUCT(avg of min/max sizes & start/top-end prices, counts) / total count; Average PSF = weighted avg price ÷ weighted avg SF
  - **Plan Detail** sheet — one row per individual plan name

## Bundled script

`scripts/build_summary.py` — accepts `--data <json>` and `--output <xlsx>`. Requires `openpyxl`.

## Owner

Ryan Lalonde — MLA Advisory
