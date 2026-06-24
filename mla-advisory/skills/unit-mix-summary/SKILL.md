---
name: unit-mix-summary
description: Generates a unit mix summary table Word document for multi-family real estate competitive analysis. Takes pasted project data (project name, unit types, unit counts, size ranges in SF) for one or more projects and produces a formatted .docx table with UNIT TYPE, %-INVENTORY, MIN-SIZE (SF), MAX-SIZE (SF), AVERAGE-SIZE, and MEDIAN-SIZE aggregated across all projects. Use this skill whenever unit mix data from multiple projects is pasted and a summary comparison table, competitive set analysis, or unit mix breakdown is needed — even if the user doesn't say "unit mix summary" explicitly.
version: 2026-06-18 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Unit Mix Summary

Aggregate multi-family unit mix data from multiple projects into a single formatted Word document summary table.

## Step 1: Parse the Input

The user will paste one or more project tables. For each project extract:
- Project name
- For each unit type row: **unit type label**, **unit count**, **size range** (min SF – max SF)

Ignore price range, unit mix %, developer name — those are not needed.

If a size is a single number (no range), treat min = max = that number.

---

## Step 2: Normalize Unit Types

Map every input label to a standard type using the rules below. Apply FLEX = DEN everywhere.

| Standard Type | Maps From |
|---|---|
| STUDIO | Studio, Bachelor |
| JR 1-BED | Jr 1 Bed, Junior 1 Bed, Jr. 1 Bedroom |
| 1-BED | 1 Bed, One Bedroom, 1B |
| 1-BED + DEN | 1 Bed + Den, 1 Bed Den, 1 Bed + Flex, 1B+D, 1B+F |
| JR 2-BED | Jr 2 Bed, Junior 2 Bed |
| 2-BED | 2 Bed, Two Bedroom, 2B, 2 Bed (1 Bath), 2 Bed (2 Bath) |
| 2-BED + DEN | 2 Bed + Den, 2 Bed + Flex, 2B+D, 2B+F |
| 3-BED | 3 Bed, Three Bedroom, 3B, 3 Bed 3 Bath |
| 3-BED + DEN | 3 Bed + Den, 3 Bed + Flex, 3B+D, 3 Bed + Den + Rooftop |
| 4-BED | 4 Bed, Four Bedroom |
| TH [type] | Any TH prefix: TH 2 Bed → TH 2-BED, TH 3 Bed → TH 3-BED, TH 2 Bed + Den → TH 2-BED + DEN, TH 3 Bed+ → TH 3-BED |
| GH [type] | Any GH prefix: GH 2 Bed → GH 2-BED |
| PH [type] | Any PH prefix OR SKYHOME: PH 1 Bed → PH 1-BED, PH 2 Bed → PH 2-BED, PH 3 Bed → PH 3-BED, PH 3 Bed + Den → PH 3-BED + DEN, SKYHOME (no bed type) → PH |

**Key rules:**
- FLEX and DEN are the same — always combine into "+ DEN"
- PH and SKYHOME are the same group; keep the bed type suffix if given
- "3 BED, 3 BATH" and "3 BED +" normalize to 3-BED
- Anything not mappable: keep label as-is, sort to the end

---

## Step 3: Calculate Statistics

For each unit type entry, compute **midpoint** = (min_size + max_size) / 2.

Aggregate across ALL projects for each normalized unit type:

| Column | Formula |
|---|---|
| %-INVENTORY | sum(count for this type) ÷ total units across all projects, formatted as whole-number % |
| MIN-SIZE (SF) | minimum of all min_sizes for this type |
| MAX-SIZE (SF) | maximum of all max_sizes for this type |
| AVERAGE-SIZE | sum(count × midpoint) ÷ sum(count) — rounded to nearest whole number |
| MEDIAN-SIZE | unweighted median: pool ALL min and max size values from every project entry for this type (ignore unit counts), sort them, take the middle value — rounded to nearest whole number. Matches the Excel template method: `=MEDIAN(min_col:max_col)` across all project rows. |

---

## Step 4: Sort Rows

Use this display order:

1. STUDIO
2. JR 1-BED
3. 1-BED
4. 1-BED + DEN
5. JR 2-BED
6. 2-BED
7. 2-BED + DEN
8. 3-BED
9. 3-BED + DEN
10. 4-BED
11. GH types (ascending bed count)
12. TH types (ascending bed count)
13. PH types (ascending bed count)
14. Any remaining unlisted types

---

## Step 5: Generate the Word Document

Run the bundled PowerShell script (no installs needed — uses Microsoft Word COM):

```powershell
$json = '<json_string>'
& "scripts/generate_table.ps1" -DataJson $json -OutputPath "Unit Mix Summary.docx"
```

JSON format expected by the script:
```json
{
  "rows": [
    {
      "unit_type": "STUDIO",
      "pct_inventory": "3%",
      "min_size": 353,
      "max_size": 478,
      "avg_size": 367,
      "median_size": 416
    }
  ]
}
```

The script path is relative to the skill directory. Use the full absolute path when calling from a different working directory:
`C:\Users\MelissaNestoruk\.claude\skills\mla-advisory\unit-mix-summary\scripts\generate_table.ps1`

Save to the Desktop: `"$([Environment]::GetFolderPath('Desktop'))\Unit Mix Summary.docx"` — resolve this path at runtime using PowerShell. Tell the user the full path.

---

## Step 6: Confirm and Offer Adjustments

After delivering the file, briefly confirm what projects were included and how many total units were counted. Offer to re-run if the user wants to add, remove, or re-categorize any projects.
