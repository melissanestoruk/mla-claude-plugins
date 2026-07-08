---
name: offerings-table
description: Builds an MLA Offering Table (Competition Offering / Rental Offering) from a HubSpot CRM export — the transposed Advisory comparables table where projects are COLUMNS and offering attributes are ROWS. FOR SALE rows; DEPOSIT, REALTOR COMMISSION, REALTOR BONUS, PURCHASER INCENTIVE, ASSIGNMENT FEE, STRATA ($/SF), COMPLETION, PARKING/STORAGE, APPLIANCES, LATEST SALES ACTIVITY. RENTAL rows; INCENTIVE, PET POLICY, PARKING RATIO, PARKING, AMENITIES, A/C, FLOORING, STORAGE, APPLIANCES, LATEST MESSAGING. Outputs Word or PowerPoint; auto-chunks big comp sets (5 projects per Word table, 4 per PPT slide). Use this skill whenever a HubSpot pull with offering/incentive data (deposit structure, commissions, buyer incentives, strata fees, parking/storage costs, appliances) is dropped and someone asks for an offering table, offerings table, competition offering, sales offerings breakdown, rental offering, or "run the offerings table". NOT for active tables, unit mix, upcoming, completions, or amenity matrices — separate skills.
version: 2026-07-07 v1
status: draft
shareable: true
owner: melissa-nestoruk
---

# Offerings Table

Turn a HubSpot export with offering data into the MLA Offering Table — transposed layout, projects as columns, attributes as rows — cloned from the locked gold templates.

## Step 1: Identify the inputs

1. **The HubSpot pull** — CSV export with offering columns (Deposit %, Deposit Structure, Realtor Commission, Buyer Incentive, Assignment Policy, Strata Fee per SqFt., Completion Date, Parking/Storage allocations and costs, Appliance Package/Sizing).
2. **The variant** — forsale or rental. Infer from columns (rental pulls have pet policy / parking ratio / flooring). Ask if ambiguous.
3. **The format** — Word or PowerPoint. **Always ask** unless stated.
4. **Title** — e.g. "Competition Offering – Brentwood Concrete" (Word heading and PPT title) or "Rental Offering".

## Step 2: Run the builder

```
python scripts/build_offering_table.py --csv <pull.csv> --variant forsale|rental \
    --format word|ppt --title "<Title>" --output <Offering Table.docx|.pptx>
```

The script (stdlib only):
- Projects keep the pull's row order (no re-sorting).
- Composes DEPOSIT as "10%: $10K at writing, 5% in 7 days" (Deposit % + structure).
- Composes PARKING/STORAGE from the four allocation/cost columns, with costs as "$45K/stall" / "$5K/locker".
- Composes APPLIANCES as "Package; Sizing".
- STRATA formatted "$0.47"; COMPLETION as "Dec 2027"; blanks render "-".
- **LATEST SALES ACTIVITY is always "-"** — HubSpot doesn't carry it; filled manually after.
- **Chunking**: more than 5 projects (Word) or 4 (PPT) automatically produces additional tables (each on its own page — the gold table is a floating table) or additional slides. Surplus columns in the last chunk are trimmed and the rest stretch to full width.

## Step 3: Verify before delivering

1. Project count across all tables/slides = pull row count (script prints it).
2. Spot-check one project's deposit and strata against the CSV.
3. Word: no overlapping tables (each chunk on its own page). PPT: slide count = ceil(projects ÷ 4).

Deliver with a one-line note on file location. One canonical file per artifact — no version suffixes.

## Known behaviours and manual steps

- LATEST SALES ACTIVITY (for-sale) is manual — always "-" in output.
- Asterisk footnotes are manual post-generation annotations; the gold's example footnote was removed from the templates.
- The rental column mapping is drafted against expected HubSpot property names (Rental Incentive, Pet Policy, Parking Ratio, Parking Cost, Amenities, A/C, Flooring, Storage Cost, Appliance Package, Latest Messaging) — verify on the first real rental pull (see `_private/overlay.md`).
- PPT slide footers come from the deck master — self-correct on paste.

## Files

- `templates/offering_forsale_template.docx` / `.pptx` — gold for-sale (landscape / slide)
- `templates/offering_rental_template.docx` / `.pptx` — gold rental
- `scripts/build_offering_table.py` — the builder
- `examples/hubspot_pull_forsale_example.csv` — real pull (Brentwood concrete, 9 projects)
- `examples/hubspot_pull_rental_example.csv` — synthetic rental pull matching the expected schema
