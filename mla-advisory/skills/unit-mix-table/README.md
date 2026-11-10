# unit-mix-table

MLA Advisory's Unit Mix Breakdown Table generator. Drop one floorplan screenshot per project (Avesdo/NHSLive "Floorplan Data" panel) into Claude, give it the project names, and say "run the unit mix table" — get back the comparables table as Word or PowerPoint, one row per project with unit types, counts, mix percents, and size ranges.

Not the same as `unit-mix-summary`: this produces the per-project breakdown table; the summary skill aggregates statistics across the comp set (and can consume this table's data).

## Usage

> Here are floorplan screenshots for Vesa (Holborn) and Latitude (Transca) — unit mix table in PowerPoint.

Claude extracts the rows (counts from the Rlsd column), shows you the extracted numbers for a quick check, then builds the file from the locked gold templates. Mix percents are computed (with the "<1%" convention), TOTAL lines are bold, size ranges use en dashes.

## Rental variant

For actively leasing purpose-built rental comps, run the builder with `--variant rental` (Word only) and a JSON that carries `rent` and `psf` strings per floorplan line. It builds the 7-column "Unit Mix Comparables - Purpose-Built Rental" chart, one self-contained chart per landscape page, with row banding starting grey on each page. See SKILL.md and `examples/example_input_rental.json`.

## Owner

Melissa Nestoruk, Product Development Specialist. Part of the Advisory repeatable-tables series (active-table, unit-mix-table, upcoming-projects-table, offerings-table, completions-table, amenity-comparable-matrix).
