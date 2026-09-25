#!/usr/bin/env python3
"""Build an MLA Stabilized Rental Chart (Word, landscape) from a JSON input.

Clone-and-swap on the gold template in ../templates/ (the active-table rental gold
with TYPE/LEASED dropped and a LEASING PERIOD column added). Brand formatting is
never rebuilt, only text is swapped.

Input JSON:
{
  "title": "Recently Stabilized",
  "writeup": "The table below provides an overview of ...",   // optional intro paragraph
  "rows": [
    {"project": "The Raven", "developer": "Westbank", "market": "Point Grey",
     "storeys": 14, "period_start": "Oct 2025", "period_end": "Jun 2026",
     "avg_sf": 570, "avg_rent": 2680, "psf": 5.10, "units": 131}
  ]
}

LEASES/MONTH is computed as units / inclusive month count of the printed period
(Jul 2025 - May 2026 = 11 months), so every row reconciles with what the reader
sees in the LEASING PERIOD column. Pass "leases_month" on a row to override.
Rows may pass avg_sf/avg_rent/psf as numbers (formatted here) or as final strings.

Usage:
  python build_stabilized_chart.py --data input.json --output "Chart.docx"

Stdlib only.
"""
import argparse
import copy
import json
import os
import sys
from datetime import datetime
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
TEMPLATE = os.path.join(HERE, '..', 'templates', 'stabilized_chart_template.docx')

# reuse the active-table zip/XML plumbing if that skill is installed; otherwise inline copies
try:
    sys.path.insert(0, os.path.join(HERE, '..', '..', 'active-table', 'scripts'))
    from build_active_chart import (W, load_part, save_with_part, parse_keep_ns,
                                    serialize_restore_root, docx_set_cell,
                                    docx_para_set_text)
except ImportError:
    sys.exit('ERROR: this skill reuses helpers from the active-table skill '
             '(~/.claude/skills/active-table/scripts/build_active_chart.py); install it first.')

# columns kept from the 12-column gold template (drops TYPE and LEASED)
KEEP = [0, 1, 2, 4, 5, 10, 6, 7, 8, 11]
HEADERS = ['PROJECT', 'DEVELOPER', 'MARKET', 'STOREYS', 'LEASING PERIOD',
           'LEASES/MONTH', 'AVG SF', 'AVG RENT', 'INITIAL GROSS PPSF', 'UNITS']
WIDTHS = [1699, 1682, 1270, 987, 2301, 1156, 804, 1055, 1196, 1356]  # sum 13506


def parse_month(s):
    return datetime.strptime(s.strip(), '%b %Y')


def inclusive_months(start, end):
    a, b = parse_month(start), parse_month(end)
    return (b.year - a.year) * 12 + (b.month - a.month) + 1


def fmt_int(v):
    return v if isinstance(v, str) else f'{int(round(v)):,}'


def fmt_money0(v):
    return v if isinstance(v, str) else f'${int(round(v)):,}'


def fmt_money2(v):
    return v if isinstance(v, str) else f'${v:,.2f}'


def build(data, output):
    rows = []
    for r in data['rows']:
        months = inclusive_months(r['period_start'], r['period_end'])
        lpm = r.get('leases_month')
        if lpm is None:
            lpm = int(r['units'] / months + 0.5)
        rows.append([
            r['project'], r['developer'], r['market'], str(r['storeys']),
            f"{r['period_start']} - {r['period_end']}", str(lpm),
            fmt_int(r['avg_sf']), fmt_money0(r['avg_rent']), fmt_money2(r['psf']),
            f"{int(r['units']):,}",
        ])
        print(f"{r['project']}: {months} months, {r['units']} units -> {lpm}/month")
    total_units = sum(int(r['units']) for r in data['rows'])

    xml = load_part(TEMPLATE, 'word/document.xml')
    root = parse_keep_ns(xml)
    body = root.find(W + 'body')

    text_paras = body.findall(W + 'p')
    docx_para_set_text(text_paras[0], data.get('title', 'Recently Stabilized'))
    if data.get('writeup') and len(text_paras) > 1:
        p2 = text_paras[1]
        docx_para_set_text(p2, data['writeup'])
        # template's second paragraph is Butler 22pt; restyle to body copy
        run = p2.find(W + 'r')
        rPr = run.find(W + 'rPr')
        if rPr is None:
            rPr = ET.Element(W + 'rPr')
            run.insert(0, rPr)
        for tag in ('rFonts', 'b', 'sz', 'szCs', 'color'):
            el = rPr.find(W + tag)
            if el is not None:
                rPr.remove(el)
        fonts = ET.SubElement(rPr, W + 'rFonts')
        fonts.set(W + 'ascii', 'Montserrat Light')
        fonts.set(W + 'hAnsi', 'Montserrat Light')
        for tag in ('sz', 'szCs'):
            ET.SubElement(rPr, W + tag).set(W + 'val', '19')  # 9.5pt

    tbl = next(iter(body.iter(W + 'tbl')))

    grid = tbl.find(W + 'tblGrid')
    cols = grid.findall(W + 'gridCol')
    for c in cols:
        grid.remove(c)
    for w in WIDTHS:
        gc = copy.deepcopy(cols[0])
        gc.set(W + 'w', str(w))
        grid.append(gc)

    def set_width(tc, w):
        tc.find(W + 'tcPr').find(W + 'tcW').set(W + 'w', str(w))

    trs = tbl.findall(W + 'tr')
    header, data_trs, total_tr = trs[0], trs[1:-1], trs[-1]
    # template ships 10 data rows: clone the last banded pair to grow, pop to shrink
    while len(data_trs) < len(rows):
        clone = copy.deepcopy(data_trs[len(data_trs) - 2])  # keeps banding alternation
        tbl.insert(list(tbl).index(total_tr), clone)
        data_trs.append(clone)
    while len(data_trs) > len(rows):
        tbl.remove(data_trs.pop())

    for tr, texts in [(header, HEADERS)] + list(zip(data_trs, rows)):
        tcs = tr.findall(W + 'tc')
        kept = [tcs[i] for i in KEEP] if len(tcs) == 12 else tcs
        for tc in tcs:
            tr.remove(tc)
        for tc, w, text in zip(kept, WIDTHS, texts):
            set_width(tc, w)
            docx_set_cell(tc, text)
            tr.append(tc)

    tcs = total_tr.findall(W + 'tc')
    label, last = tcs[0], tcs[-1]
    for tc in tcs[1:-1]:
        total_tr.remove(tc)
    label.find(W + 'tcPr').find(W + 'gridSpan').set(W + 'val', '9')
    set_width(label, sum(WIDTHS[:-1]))
    set_width(last, WIDTHS[-1])
    docx_set_cell(label, 'TOTAL')
    docx_set_cell(last, f'{total_units:,}')

    save_with_part(TEMPLATE, output, 'word/document.xml',
                   serialize_restore_root(root, xml, 'document'))
    print(f'Wrote {output} | total units {total_units:,}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    with open(args.data, encoding='utf-8') as f:
        data = json.load(f)
    if not data.get('rows'):
        sys.exit('ERROR: no rows in input JSON')
    build(data, args.output)


if __name__ == '__main__':
    main()
