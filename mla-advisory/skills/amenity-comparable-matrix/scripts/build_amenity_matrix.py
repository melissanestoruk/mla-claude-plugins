#!/usr/bin/env python3
"""Build an MLA Amenity Comparables Matrix (Word only) from mapped project data.

Clone-and-swap on the gold template in ../templates/. The matrix rows (21
canonical amenities in three category bands) are read from the template itself —
the input JSON must reference those exact row labels. Mapping HubSpot's free-text
amenity names to canonical rows is done by Claude (per SKILL.md) before calling
this script.

Input JSON:
{
  "title": "Amenity Comparables Matrix",          // optional
  "intro": "Below is an amenity matrix for ...",  // optional
  "projects": [
    {"name": "Prima", "strata": 0.60,
     "amenities": ["FITNESS CENTRE", "YOGA/BARRE ROOM", ...]}
  ]
}

Checked cells get the gold's blue shading + checkmark; the TOTAL % column is
computed across ALL projects; STRATA FEES and per-project TOTAL rows are filled.
More than 7 projects chunk into additional matrices (7 columns each), one per
page.

Usage:
  python build_amenity_matrix.py --data data.json --output "Amenity Matrix.docx"
"""
import argparse
import copy
import json
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = os.path.join(HERE, '..', 'templates')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
CHUNK = 7


def cell_text(tc):
    return ' '.join(t.text or '' for t in tc.iter(W + 't')).strip()


def parse_keep_ns(xml_text):
    for prefix, uri in re.findall(r'xmlns:([\w-]+)="([^"]+)"', xml_text):
        ET.register_namespace(prefix, uri)
    return ET.fromstring(xml_text)


def serialize_restore_root(root, orig_xml):
    out = ET.tostring(root, encoding='unicode')
    new_tag = re.search(r'<[\w:]*document[^>]*>', out).group(0)
    orig_tag = re.search(r'<[\w:]*document[^>]*>', orig_xml).group(0)
    missing = [m.group(0) for m in re.finditer(r'xmlns:[\w-]+="[^"]+"', orig_tag)
               if m.group(0) not in new_tag]
    if missing:
        out = out.replace(new_tag, new_tag[:-1].rstrip() + ' ' + ' '.join(missing) + '>', 1)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out


def set_para_text(p, text):
    runs = p.findall(W + 'r')
    if runs:
        first = runs[0]
        for extra in runs[1:]:
            p.remove(extra)
    else:
        first = ET.SubElement(p, W + 'r')
    t = first.find(W + 't')
    if t is None:
        t = ET.SubElement(first, W + 't')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


def set_cell_text(tc, text):
    paras = tc.findall(W + 'p')
    for extra in paras[1:]:
        tc.remove(extra)
    set_para_text(paras[0], text)


def replace_cell(tr, old_tc, proto_tc):
    """Swap a cell for a checked/unchecked prototype, keeping the column width."""
    new = copy.deepcopy(proto_tc)
    old_w = old_tc.find(W + 'tcPr').find(W + 'tcW')
    new_w = new.find(W + 'tcPr').find(W + 'tcW')
    new_w.set(W + 'w', old_w.get(W + 'w'))
    new_w.set(W + 'type', old_w.get(W + 'type') or 'dxa')
    idx = list(tr).index(old_tc)
    tr.remove(old_tc)
    tr.insert(idx, new)


def strata_txt(v):
    if v is None or v == '':
        return '-'
    try:
        return f'${float(v):.2f}'
    except (TypeError, ValueError):
        return str(v)


def analyze_table(tbl):
    """Classify template rows: header, category bands, amenity rows, strata, total."""
    info = {'amenities': {}, 'categories': [], 'header': None, 'strata': None, 'total': None}
    for tr in tbl.findall(W + 'tr'):
        tcs = tr.findall(W + 'tc')
        label = cell_text(tcs[0]).upper()
        if len(tcs) == 1:
            info['categories'].append(tr)
        elif label == 'AMENITY':
            info['header'] = tr
        elif label == 'STRATA FEES':
            info['strata'] = tr
        elif label == 'TOTAL':
            info['total'] = tr
        else:
            info['amenities'][label] = tr
    return info


def build(data, output):
    template = os.path.join(TEMPLATES, 'amenity_matrix_template.docx')
    with zipfile.ZipFile(template) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    root = parse_keep_ns(xml)
    body = root.find(W + 'body')

    text_paras = [p for p in body.findall(W + 'p')
                  if any((t.text or '').strip() for t in p.iter(W + 't'))]
    if data.get('title') and text_paras:
        set_para_text(text_paras[0], data['title'])
    if data.get('intro') and len(text_paras) > 1:
        set_para_text(text_paras[1], data['intro'])

    gold_tbl = next(iter(body.iter(W + 'tbl')))
    gold_info = analyze_table(gold_tbl)
    canonical = list(gold_info['amenities'].keys())
    n_amenities = len(canonical)
    full_cols = len(gold_tbl.find(W + 'tblGrid').findall(W + 'gridCol'))  # 9
    max_projects = full_cols - 2

    projects = data['projects']
    for p in projects:
        p['set'] = {a.upper().strip() for a in p.get('amenities', [])}
        unknown = p['set'] - set(canonical)
        if unknown:
            sys.exit(f"ERROR: unknown amenity rows for {p['name']}: {sorted(unknown)}\n"
                     f"Canonical rows: {canonical}")

    # prototype checked/unchecked cells: scan amenity rows for a shaded and an
    # unshaded project cell
    proto_checked = proto_unchecked = None
    for tr in gold_info['amenities'].values():
        for tc in tr.findall(W + 'tc')[1:-1]:
            shd = tc.find(W + 'tcPr').find(W + 'shd')
            filled = shd is not None and (shd.get(W + 'fill') or 'auto') != 'auto'
            if filled and proto_checked is None:
                proto_checked = copy.deepcopy(tc)
            if not filled and proto_unchecked is None:
                proto_unchecked = copy.deepcopy(tc)
        if proto_checked is not None and proto_unchecked is not None:
            break
    if proto_checked is None or proto_unchecked is None:
        sys.exit('ERROR: could not find checked/unchecked prototype cells in template')

    chunks = [projects[i:i + max_projects] for i in range(0, len(projects), max_projects)]
    proto_tbl = copy.deepcopy(gold_tbl)

    def fill_table(tbl, chunk):
        info = analyze_table(tbl)
        n = len(chunk)
        surplus = max_projects - n
        if surplus:
            grid = tbl.find(W + 'tblGrid')
            cols = grid.findall(W + 'gridCol')
            removed = sum(int(c.get(W + 'w')) for c in cols[1 + n:-1])
            for c in cols[1 + n:-1]:
                grid.remove(c)
            for c in grid.findall(W + 'gridCol')[1:-1]:
                c.set(W + 'w', str(int(c.get(W + 'w')) + removed // n))
            for tr in tbl.findall(W + 'tr'):
                tcs = tr.findall(W + 'tc')
                if len(tcs) == 1:
                    span = tcs[0].find(W + 'tcPr').find(W + 'gridSpan')
                    span.set(W + 'val', str(n + 2))
                else:
                    for tc in tcs[1 + n:-1]:
                        tr.remove(tc)
        # header
        tcs = info['header'].findall(W + 'tc')
        for j, p in enumerate(chunk):
            set_cell_text(tcs[1 + j], p['name'].upper())
        # amenity rows: swap checked/unchecked prototypes + TOTAL % across ALL projects
        for label, tr in analyze_table(tbl)['amenities'].items():
            tcs = tr.findall(W + 'tc')
            for j, p in enumerate(chunk):
                proto = proto_checked if label in p['set'] else proto_unchecked
                replace_cell(tr, tcs[1 + j], proto)
                if label in p['set']:
                    set_cell_text(tr.findall(W + 'tc')[1 + j], '✓')
                else:
                    set_cell_text(tr.findall(W + 'tc')[1 + j], '')
            have = sum(1 for p in projects if label in p['set'])
            pct = round(have / len(projects) * 100) if projects else 0
            set_cell_text(tr.findall(W + 'tc')[-1], f'{pct}%')
        # strata + total rows
        info2 = analyze_table(tbl)
        tcs = info2['strata'].findall(W + 'tc')
        for j, p in enumerate(chunk):
            set_cell_text(tcs[1 + j], strata_txt(p.get('strata')))
        set_cell_text(tcs[-1], '')
        tcs = info2['total'].findall(W + 'tc')
        for j, p in enumerate(chunk):
            k = len(p['set'])
            set_cell_text(tcs[1 + j], f'{k}/{n_amenities} ({round(k / n_amenities * 100)}%)')
        set_cell_text(tcs[-1], '')

    fill_table(gold_tbl, chunks[0])
    anchor = list(body)[-2]     # last content element before body sectPr
    for chunk in chunks[1:]:
        # a floating table needs an anchor paragraph on its own page: a
        # pageBreakBefore paragraph (a break RUN pushes the floating table one
        # page too far), then the table, then a plain paragraph it anchors to
        pbreak = ET.Element(W + 'p')
        ppr = ET.SubElement(pbreak, W + 'pPr')
        ET.SubElement(ppr, W + 'pageBreakBefore')
        t2 = copy.deepcopy(proto_tbl)
        fill_table(t2, chunk)
        tail = ET.Element(W + 'p')
        idx = list(body).index(anchor) + 1
        body.insert(idx, pbreak)
        body.insert(idx + 1, t2)
        body.insert(idx + 2, tail)
        anchor = tail

    with zipfile.ZipFile(template) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, blob in items:
            if info.filename == 'word/document.xml':
                blob = serialize_restore_root(root, xml).encode('utf-8')
            zout.writestr(info, blob)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True, help='JSON input file')
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    with open(args.data, encoding='utf-8') as f:
        data = json.load(f)
    if not data.get('projects'):
        sys.exit('ERROR: no projects in input JSON')
    build(data, args.output)
    print(f"Wrote {args.output} ({len(data['projects'])} projects)")


if __name__ == '__main__':
    main()
