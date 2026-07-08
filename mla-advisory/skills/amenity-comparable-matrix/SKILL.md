---
name: amenity-comparable-matrix
description: Builds an MLA Amenity Comparables Matrix (Word ONLY) from a HubSpot CRM export with a Project Amenities column — the checkmark matrix of 21 canonical amenities (grouped FITNESS & WELLNESS / ENTERTAINMENT / CONVENIENCE) across a competitive set, with a TOTAL % column, STRATA FEES row, and per-project amenity-count TOTAL row. Use this skill whenever a HubSpot pull with amenity lists is dropped and someone asks for an amenity matrix, amenity comparables, amenity comparison, amenities table, or "run the amenity matrix". This is the only repeatable-tables skill with NO PowerPoint output. NOT for active tables, unit mix, upcoming, offerings, or completions — separate skills.
version: 2026-07-07 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Amenity Comparable Matrix

Turn a HubSpot export (Name, Project Amenities, Strata Fee per SqFt.) into the MLA Amenity Comparables Matrix — checkmarks with blue shading against the locked 21-row amenity taxonomy. **Word only.**

## Step 1: Identify the inputs

1. **The HubSpot pull** — CSV with a semicolon-separated "Project Amenities" column and "Strata Fee per SqFt.".
2. **Title/intro** — defaults exist in the template; swap the intro's market/product wording to match the comp set.

## Step 2: Map amenities to the canonical rows (Claude does this, not the script)

The matrix rows are FIXED — the template's 21 amenities. Map each HubSpot amenity term to exactly one canonical row:

| HubSpot term (examples) | Canonical row |
|---|---|
| Fitness Centre, Gymnasium | FITNESS CENTRE |
| Yoga Area | YOGA/BARRE ROOM |
| Steam / Sauna | SAUNA/STEAM ROOM |
| Hot Tub, Pool, Cold Plunge | POOL/HOT TUB/COLD PLUNGE |
| Indoor Social Lounge, Party Room w/ Kitchen, Amenity Clubhouse | ENTERTAINMENT LOUNGE + KITCHEN |
| Shared BBQ Area, Outdoor Social Space/Terrace, Rooftop Space | OUTDOOR LOUNGE WITH BBQ |
| Outdoor Dining (explicit) | OUTDOOR DINING |
| Games/Flex Room | GAMES ROOM |
| Soundproof Music Room | MUSIC ROOM |
| Childrens Play Area (Outdoor), Playground | CHILDREN'S PLAY AREA |
| Firepits | FIREPITS |
| Outdoor Garden Plots | COMMUNITY GARDEN/PARK SPACE |
| Sports Simulator | SPORTS SIMULATOR |
| Bocce Ball Area, Sports Court, Putting Green | SPORTS COURT /BOCCE/PUTTING GREEN |
| Parcel Delivery Room | PARCEL LOCKERS |
| Concierge | CONCIERGE |
| Guest Suite | GUEST SUITE |
| Dog Run, Pet Wash | DOG RUN/PET WASH STATION |
| Car Wash | CAR WASH |
| Co-working Space, Boardroom/Meeting Room | CO-WORKING SPACE |
| Bike/Tool Workshop, Bike Repair | BIKE/TOOL WORKSHOP |

Terms with no canonical row (e.g. EV Car Charging) are dropped. **Show the user the full mapping — including anything dropped or ambiguous — before building.** New unmapped terms: ask the user which row they belong to (or whether to drop).

Then write the JSON input:

```json
{
  "title": "Amenity Comparables Matrix",
  "intro": "Below is an amenity matrix for ...",
  "projects": [
    {"name": "Prima", "strata": 0.60,
     "amenities": ["FITNESS CENTRE", "YOGA/BARRE ROOM", "ENTERTAINMENT LOUNGE + KITCHEN"]}
  ]
}
```

Amenity strings must match the canonical labels exactly (the script validates and errors on unknowns).

## Step 3: Run the builder

```
python scripts/build_amenity_matrix.py --data <data.json> --output <Amenity Matrix.docx>
```

The script (stdlib only):
- Reads the canonical rows from the template itself.
- Checked cells get the gold's blue shading + checkmark; unchecked cells stay clear.
- TOTAL % column = share of ALL projects with that amenity (same values in every chunk).
- STRATA FEES row ($0.60 format, "-" if missing); TOTAL row per project "9/21 (43%)".
- More than 7 projects chunk into additional matrices, one per page, keeping the pull's project order.

## Step 4: Verify before delivering

1. Project count across matrices = pull row count.
2. Spot-check one project's checkmarks against its amenity list in the CSV.
3. A 100%-common amenity should be checked in every column; 0% rows stay empty.

Deliver with a one-line note on file location. One canonical file per artifact — no version suffixes.

## Known behaviours

- **Word only** — no PowerPoint gold exists for this table.
- The CONVENIENCE* footnote about unadvertised amenities is part of the template and stays.
- The taxonomy is locked; adding a new amenity row means updating the gold template, not the script.

## Files

- `templates/amenity_matrix_template.docx` — gold Word (landscape, 21-row taxonomy)
- `scripts/build_amenity_matrix.py` — the builder
- `examples/hubspot_pull_example.csv` — real pull (13 projects)
- `examples/example_input.json` — the mapped JSON built from that pull
