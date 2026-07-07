# active-table

MLA Advisory's Active Chart generator. Drop a HubSpot CRM export of comparable projects into Claude and say "run the active table" — get back the standard competition analysis table as a Word document or a PowerPoint slide, cloned from the locked gold templates.

## Variants

| Variant | Columns | Output |
|---|---|---|
| For sale | PROJECT, DEVELOPER, LAUNCH, AVG SF, PPSF, SOLD, RLSD, TOTAL (% sold) | Word or PowerPoint |
| Rental | PROJECT, DEVELOPER, MARKET, TYPE, STOREYS, OCCUPANCY, AVG SF, AVG RENT, INITIAL PSF, LEASED, LEASES/MONTH, TOTAL (% leased) | Word only |

## Usage

In Claude Code, attach the HubSpot CSV export and say:

> Run the active table for Coquitlam — Word please.

Claude asks Word or PowerPoint if you don't say, runs `scripts/build_active_chart.py`, and returns the finished file. Rows sort newest launch/occupancy first; totals and percentages are computed for you.

Project-specific asterisk footnotes (`**Sales Centre Closed.` etc.) are added manually after generation — the skill leaves project names clean.

## Owner

Melissa Nestoruk, Product Development Specialist. Part of the Advisory repeatable-tables series (active-table, unit-mix-table, upcoming-projects-table, offerings-table, completions-table, amenity-comparable-matrix).
