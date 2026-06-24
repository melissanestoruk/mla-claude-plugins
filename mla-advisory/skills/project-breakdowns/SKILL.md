---
name: project-breakdowns
version: 2026-06-24 v1
status: active
shareable: true
owner: ryan-lalonde
owner_team: mla-advisory
description: >
  Builds a presale unit pricing analysis Excel workbook from floor plan images
  and price sheet data — the "Project Breakdown" deliverable used by MLA Advisory.
  Always use this skill when someone uploads floor plan images or a price sheet
  and asks for a pricing breakdown, unit type analysis, or PPSF calculation for
  a presale project. Triggers on any of these phrases: "project breakdown for
  [project]", "pricing summary", "count units from floor plans", "how many units
  per type", "weighted average PPSF", "pricing analysis Excel", "PPSF breakdown",
  "price per square foot chart", "unit type chart", or when floor plan images are
  provided alongside pricing or square footage data. Also triggers when someone
  says "do a sumproduct" on pricing or size data.
---

# Project Breakdowns

Produces a formatted Excel pricing analysis for a presale residential project. Reads building plan thumbnail images to count units per type, combines with pricing data from a price sheet or user input, and outputs a SUMPRODUCT-weighted pricing summary workbook.

## Inputs — gather before building

| Input | Source | Notes |
|---|---|---|
| Floor plan images | Uploaded by user | One per plan type; each has a small building thumbnail in the corner |
| Price sheet | Uploaded image or user input | Starting prices per unit type or plan name |
| Project name | User | Used for file naming |
| Unit types | Floor plans or user | e.g. Studio, 1 Bed, 1 Bed+Den, Jr. 2 Bed, 2 Bed/2 Bath, Jr. 3 Bed, 3 Bed |
| Level grouping rules | User (ask if unclear) | e.g. "Levels 3 and 4 share the same layout" means each highlighted position counts ×2 for those levels |
| Top-end price | User (ask if not on price sheet) | Estimated upper range for each unit type |

If the user has already supplied some of this in the conversation, use it — don't re-ask.

## Step 1 — Count units from floor plan thumbnails

Each floor plan page has a small building thumbnail (corner inset) showing the full plate with the plan's units highlighted — typically a tan, bronze, or gold fill. Units NOT of that type are shown unfilled (white or grey).

**Counting approach:**

1. For each floor plan image, look at the small thumbnail carefully
2. Count the highlighted (filled) cells — those are the units of that plan type
3. Check the level labels. If a note says two or more levels share the same layout (e.g. "L3 & L4" or "Floors 5–8"), multiply the per-floor count accordingly
4. Sum across all levels shown

**When thumbnails are ambiguous:** work level-by-level, list your count for each level, and ask the user to confirm before proceeding. A quick table like this works well:

```
Plan A — counted units:
  L2:  2 units
  L3:  2 units  (L3 & L4 share layout → ×2 = 4 units)
  L5–L8: 2 units × 4 floors → 8
  Total: 14 units — does this look right?
```

**When a stacking plan is provided:** use it directly instead of thumbnails.

**If the user gives you the counts directly:** use them without re-counting.

## Step 2 — Confirm before building

Before writing the Excel, show the user a confirmation table:

```
Unit Type      | Plan(s)            | Count
Studio         | Studio A, Studio B | 9
1 Bed          | A, A1, Ab          | 18
1 Bed + Den    | B10                | 3
Jr. 2 Bed      | A17                | 33
2 Bed / 2 Bath | ...                | 12
Jr. 3 Bed      | ...                | 14
3 Bed          | ...                | 5
────────────────────────────────────────
TOTAL                               | 94
```

Ask: "Do these counts look right before I build the Excel?" Only proceed once confirmed.

## Step 3 — Build the Excel

Use the bundled script `scripts/build_summary.py`. Prepare a JSON data file and call the script:

```bash
python <skill_dir>/scripts/build_summary.py \
  --data /tmp/project_data.json \
  --output "/path/to/ProjectName_Pricing_Summary.xlsx"
```

The JSON data format:

```json
{
  "project_name": "Harth at Tamanawis Park Newton",
  "rows": [
    {
      "unit_type": "Studio",
      "plan_names": "Studio A, Studio B",
      "count": 9,
      "min_sf": 366,
      "max_sf": 374,
      "starting_price": 499900,
      "top_end_price": 549900
    }
  ]
}
```

`top_end_price` is optional — pass null if unknown; the cell will be left blank.

Save the output file to the user's workspace folder. Present it when done.

## SUMPRODUCT formula logic

The weighted average PPSF at the bottom of the Summary sheet uses:

```
=IF(SUM(counts)>0, SUMPRODUCT(counts, starting_ppsf) / SUM(counts), "-")
```

Where `starting_ppsf` per row = Starting Price ÷ Min SF.

The formula is live in Excel — if the user updates counts or prices, the weighted average recalculates automatically. Count cells (column C) are highlighted in yellow to signal they are editable inputs.

## When bash is unavailable

If the shell is not accessible, write the openpyxl build code inline and run it in the available Python environment, or write the file directly. The bundled script is a convenience wrapper — what matters is delivering a correctly structured .xlsx with live formulas.

## Output

**Filename:** `[ProjectName]_Pricing_Summary.xlsx`  
**Location:** User's workspace folder  
**Sheets:**
- **Summary** — one row per unit type, SUMPRODUCT weighted avg PPSF in footer row
- **Plan Detail** — one row per individual plan name (more granular breakdown)
