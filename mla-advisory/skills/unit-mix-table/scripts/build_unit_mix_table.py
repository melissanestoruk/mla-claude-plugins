#!/usr/bin/env python3
"""Build an MLA Unit Mix Breakdown Table (Word or PowerPoint) from extracted
per-project floorplan data.

Clone-and-swap on the gold templates in ../templates/. Input is a JSON file
(typically written by Claude after reading Avesdo/NHSLive floorplan screenshots):

{
  "title": "Unit Mix Comparables - Concrete Condominium",   // Word heading / PPT title
  "intro": "The table below highlights ...",                 // Word only; omit to keep template text
  "projects": [
    {
      "name": "VESA AT LITTLE MOUNTAIN",
      "developer": "HOLBORN DEVELOPMENTS",
      "units": [
        {"type": "1 BED", "count": 25, "min_sf": 495, "max_sf": 592},
        {"type": "1 BED + DEN", "count": 17, "min_sf": 549, "max_sf": 617}
      ]
    }
  ]
}

Usage:
  python build_unit_mix_table.py --data data.json --format word --output out.docx
  python build_unit_mix_table.py --data data.json --format ppt  --output out.pptx
  python build_unit_mix_table.py --data rental.json --format word --variant rental --output out.docx

--variant rental (Word only) builds the 7-column "Unit Mix Comparables - Purpose-Built
Rental" chart: PROJECT, FLOORPLAN, UNITS, UNIT MIX, SQFT, GROSS RENT RANGE, GROSS RENTAL PSF.
Each unit line carries free-text "rent" and "psf" strings, e.g.
  {"type": "1 BED", "count": 25, "min_sf": 446, "max_sf": 570,
   "rent": "$2,000 - $2,350", "psf": "$4.12 - $4.48"}
Every page is its own chart (heading repeated, blurb on page 1 only, own header row).

Stdlib only. One table row per project; unit mix percents are computed
(round half up, '<1%' for anything under 1 percent), TOTAL lines are bold.
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
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
# plain hyphen: the MLA office-output-lock gate rejects em/en dashes in document text
EN_DASH = '-'


# ------------------------------------------------------------- computation

def pct_label(count, total):
    if not total:
        return '-'
    p = count / total * 100
    if 0 < p < 1:
        return '<1%'
    return f'{int(p + 0.5)}%'


def size_label(min_sf, max_sf):
    lo = f'{int(min_sf):,}' if min_sf else ''
    hi = f'{int(max_sf):,}' if max_sf else ''
    if lo and hi and lo != hi:
        return f'{lo} {EN_DASH} {hi}'
    return lo or hi or '-'


def project_lines(proj):
    """Return the five cell line-lists for one project row.
    Each line is (text, kind) with kind in {'name','regular','total','blank'}."""
    units = proj['units']
    total = sum(u.get('count') or 0 for u in units)
    types = [(u['type'].upper(), 'regular') for u in units] + [('', 'blank')]
    counts = [(f"{u.get('count') or 0:,}", 'regular') for u in units] + \
             [(f'{total:,} TOTAL', 'total')]
    mixes = [(pct_label(u.get('count') or 0, total), 'regular') for u in units] + \
            [('100%', 'total')]
    sizes = [(size_label(u.get('min_sf'), u.get('max_sf')), 'regular') for u in units] + \
            [('', 'blank')]
    project = [(proj['name'].upper(), 'name'), ((proj.get('developer') or '').upper(), 'regular')]
    return project, types, counts, mixes, sizes


# ------------------------------------------------------- shared zip plumbing

def load_part(path, part):
    with zipfile.ZipFile(path) as z:
        return z.read(part).decode('utf-8')


def save_with_part(src, dst, part, new_xml):
    with zipfile.ZipFile(src) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == part:
                data = new_xml.encode('utf-8')
            zout.writestr(info, data)


def parse_keep_ns(xml_text):
    for prefix, uri in re.findall(r'xmlns:([\w-]+)="([^"]+)"', xml_text):
        ET.register_namespace(prefix, uri)
    return ET.fromstring(xml_text)


def serialize_restore_root(root, orig_xml, root_local):
    """Merge original root xmlns decls (Ignorable-only ones ET drops) into output."""
    out = ET.tostring(root, encoding='unicode')
    new_tag = re.search(r'<[\w:]*%s[^>]*>' % root_local, out).group(0)
    orig_tag = re.search(r'<[\w:]*%s[^>]*>' % root_local, orig_xml).group(0)
    missing = [m.group(0) for m in re.finditer(r'xmlns:[\w-]+="[^"]+"', orig_tag)
               if m.group(0) not in new_tag]
    if missing:
        merged = new_tag[:-1].rstrip() + ' ' + ' '.join(missing) + '>'
        out = out.replace(new_tag, merged, 1)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out


# ------------------------------------------------------------ Word building

def docx_prototypes(tc):
    """From a prototype cell pick paragraphs to clone: regular, bold, blank."""
    protos = {'regular': None, 'total': None, 'name': None, 'blank': None}
    for p in tc.findall(W + 'p'):
        has_text = any((t.text or '').strip() for t in p.iter(W + 't'))
        r = p.find(W + 'r')
        bold = False
        if r is not None:
            rPr = r.find(W + 'rPr')
            bold = rPr is not None and rPr.find(W + 'b') is not None
        if not has_text:
            protos['blank'] = copy.deepcopy(p) if protos['blank'] is None else protos['blank']
        elif bold:
            key = 'name' if protos['name'] is None and p is tc.find(W + 'p') else 'total'
            protos[key] = copy.deepcopy(p) if protos[key] is None else protos[key]
        else:
            protos['regular'] = copy.deepcopy(p) if protos['regular'] is None else protos['regular']
    # fallbacks
    base = next(v for v in (protos['regular'], protos['total'], protos['name']) if v is not None)
    for k in protos:
        if protos[k] is None:
            protos[k] = copy.deepcopy(base)
    return protos


def docx_para_with_text(proto, text):
    p = copy.deepcopy(proto)
    runs = p.findall(W + 'r')
    if not text:
        for r in runs:
            p.remove(r)
        return p
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
    return p


def docx_fill_cell(tc, lines, protos):
    for p in tc.findall(W + 'p'):
        tc.remove(p)
    # cell must keep tcPr first; paragraphs append after it
    for text, kind in lines:
        tc.append(docx_para_with_text(protos[kind], text))


def build_word(data, output):
    template = os.path.join(TEMPLATES, 'unit_mix_table_template.docx')
    xml = load_part(template, 'word/document.xml')
    root = parse_keep_ns(xml)
    body = root.find(W + 'body')

    text_paras = [p for p in body.findall(W + 'p')
                  if any((t.text or '').strip() for t in p.iter(W + 't'))]
    if data.get('title') and text_paras:
        docx_para_set_text(text_paras[0], data['title'])
    if data.get('intro') and len(text_paras) > 1:
        docx_para_set_text(text_paras[1], data['intro'])

    tbl = next(iter(body.iter(W + 'tbl')))
    trs = tbl.findall(W + 'tr')
    proto_row = trs[1]
    cell_protos = [docx_prototypes(tc) for tc in proto_row.findall(W + 'tc')]
    # project-name paragraph is the bold first paragraph of cell 0
    first_p = proto_row.findall(W + 'tc')[0].find(W + 'p')
    cell_protos[0]['name'] = copy.deepcopy(first_p)

    for tr in trs[1:]:
        tbl.remove(tr)
    for idx, proj in enumerate(data['projects']):
        tr = copy.deepcopy(proto_row)
        tcs = tr.findall(W + 'tc')
        for tc, lines, protos in zip(tcs, project_lines(proj), cell_protos):
            docx_fill_cell(tc, lines, protos)
        # alternate row banding: odd projects white, even grey (gold look)
        fill = 'FFFFFF' if idx % 2 == 0 else 'F2F2F2'
        for tc in tcs:
            tc_pr = tc.find(W + 'tcPr')
            shd = tc_pr.find(W + 'shd') if tc_pr is not None else None
            if shd is not None:
                shd.set(W + 'fill', fill)
        tbl.append(tr)

    save_with_part(template, output, 'word/document.xml',
                   serialize_restore_root(root, xml, 'document'))


def docx_para_set_text(p, text):
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


# ----------------------------------------------------- Word rental variant

# Table lines per page (plan lines + 1 TOTAL line per project). Calibrated against the
# 888 Cambie chart: page 1 carries the blurb, later pages do not. Verify with a real render.
CAP_FIRST_PAGE = 26
CAP_NEXT_PAGE = 31


def rental_cell_lines(proj):
    """Return (total, six column line-lists) for one rental project row.
    Columns after PROJECT: FLOORPLAN, UNITS, UNIT MIX, SQFT, GROSS RENT RANGE, GROSS RENTAL PSF."""
    units = proj['units']
    total = sum(u.get('count') or 0 for u in units)
    return total, [
        [u['type'].upper() for u in units],
        [f"{u.get('count') or 0:,}" for u in units],
        [pct_label(u.get('count') or 0, total) for u in units],
        [size_label(u.get('min_sf'), u.get('max_sf')) for u in units],
        [(u.get('rent') or '-') for u in units],
        [(u.get('psf') or '-') for u in units],
    ]


def set_row_fill(tr, fill):
    """Explicit fill only. Theme attributes are removed: themeFillShade=F2 makes Word
    draw 'white' rows grey."""
    for tc in tr.findall(W + 'tc'):
        shd = tc.find(W + 'tcPr').find(W + 'shd')
        for a in ('themeFill', 'themeFillShade', 'themeFillTint'):
            shd.attrib.pop(W + a, None)
        shd.set(W + 'val', 'clear')
        shd.set(W + 'color', 'auto')
        shd.set(W + 'fill', fill)


def _clear_runs(p):
    for r in p.findall(W + 'r'):
        p.remove(r)
    return p


def build_word_rental(data, output):
    template = os.path.join(TEMPLATES, 'unit_mix_table_rental_template.docx')
    xml = load_part(template, 'word/document.xml')
    root = parse_keep_ns(xml)
    body = root.find(W + 'body')

    title_p = body.findall(W + 'p')[0]
    intro_p = body.findall(W + 'p')[1]
    if data.get('title'):
        docx_para_set_text(title_p, data['title'])
    if data.get('intro'):
        docx_para_set_text(intro_p, data['intro'])

    tbl = body.find(W + 'tbl')
    trs = tbl.findall(W + 'tr')
    proto_row = trs[1]
    for tr in trs[1:]:
        tbl.remove(tr)
    tbl_blank = copy.deepcopy(tbl)            # header-only table for pages 2+
    spacer_p = copy.deepcopy(body.findall(W + 'p')[-1])   # tiny paragraph after the table

    # prototypes: first paragraph = plan line, last paragraph = tail (TOTAL / 100% / blank)
    pcells = proto_row.findall(W + 'tc')
    name_proto = copy.deepcopy(pcells[0].findall(W + 'p')[0])
    dev_proto = copy.deepcopy(pcells[0].findall(W + 'p')[1])
    line_protos = [copy.deepcopy(tc.findall(W + 'p')[0]) for tc in pcells[1:]]
    tail_protos = [copy.deepcopy(tc.findall(W + 'p')[-1]) for tc in pcells[1:]]

    rows = []  # (lines_on_page, tr)
    for proj in data['projects']:
        total, cols = rental_cell_lines(proj)
        tr = copy.deepcopy(proto_row)
        tcs = tr.findall(W + 'tc')
        for p in tcs[0].findall(W + 'p'):
            tcs[0].remove(p)
        pn = copy.deepcopy(name_proto)
        docx_para_set_text(pn, proj['name'].upper())
        tcs[0].append(pn)
        pd = copy.deepcopy(dev_proto)
        dev = (proj.get('developer') or '').upper()
        if dev:
            docx_para_set_text(pd, dev)
        else:
            _clear_runs(pd)
        tcs[0].append(pd)
        for ci, (tc, lines) in enumerate(zip(tcs[1:], cols)):
            for p in tc.findall(W + 'p'):
                tc.remove(p)
            for text in lines:
                np_ = copy.deepcopy(line_protos[ci])
                docx_para_set_text(np_, text)
                tc.append(np_)
            tail = copy.deepcopy(tail_protos[ci])
            if ci == 1:
                docx_para_set_text(tail, f'{total:,} TOTAL')
            elif ci == 2:
                docx_para_set_text(tail, '100%')
            else:
                _clear_runs(tail)
            tc.append(tail)
        rows.append((len(proj['units']) + 1, tr))

    # greedy pagination by line count
    pages, cur, used = [], [], 0
    for n, tr in rows:
        cap = CAP_FIRST_PAGE if not pages else CAP_NEXT_PAGE
        if cur and used + n > cap:
            pages.append(cur)
            cur, used = [], 0
        cur.append(tr)
        used += n
    pages.append(cur)

    anchor = tbl
    for pi, page_rows in enumerate(pages):
        if pi == 0:
            t = tbl
        else:
            h = copy.deepcopy(title_p)
            ppr = h.find(W + 'pPr')
            if ppr is None:
                ppr = ET.Element(W + 'pPr')
                h.insert(0, ppr)
            ppr.insert(0, ET.Element(W + 'pageBreakBefore'))
            t = copy.deepcopy(tbl_blank)
            # order after the previous table's spacer paragraph: heading, table, spacer
            at = list(body).index(anchor) + 2   # skip the previous table's spacer paragraph
            body.insert(at, h)
            body.insert(at + 1, t)
            body.insert(at + 2, copy.deepcopy(spacer_p))
        for i, tr in enumerate(page_rows):
            set_row_fill(tr, 'FFFFFF' if i % 2 else 'F2F2F2')   # header row is white, so each page starts grey
            trpr = tr.find(W + 'trPr')
            if trpr is not None:
                for hgt in trpr.findall(W + 'trHeight'):
                    trpr.remove(hgt)
            t.append(tr)
        anchor = t

    save_with_part(template, output, 'word/document.xml',
                   serialize_restore_root(root, xml, 'document'))
    return len(pages)


# ------------------------------------------------------- PowerPoint building

def pptx_prototypes(tc):
    protos = {'regular': None, 'total': None, 'name': None, 'blank': None}
    tx = tc.find(A + 'txBody')
    first = tx.find(A + 'p')
    for p in tx.findall(A + 'p'):
        runs = p.findall(A + 'r')
        has_text = any((t.text or '').strip() for t in p.iter(A + 't'))
        bold = bool(runs) and runs[0].find(A + 'rPr') is not None \
            and runs[0].find(A + 'rPr').get('b') == '1'
        if not has_text:
            protos['blank'] = copy.deepcopy(p) if protos['blank'] is None else protos['blank']
        elif bold:
            key = 'name' if p is first else 'total'
            protos[key] = copy.deepcopy(p) if protos[key] is None else protos[key]
        else:
            protos['regular'] = copy.deepcopy(p) if protos['regular'] is None else protos['regular']
    base = next(v for v in (protos['regular'], protos['total'], protos['name']) if v is not None)
    for k in protos:
        if protos[k] is None:
            protos[k] = copy.deepcopy(base)
    return protos


def pptx_para_with_text(proto, text):
    p = copy.deepcopy(proto)
    runs = p.findall(A + 'r')
    if not text:
        for r in runs:
            p.remove(r)
        return p
    if runs:
        first = runs[0]
        for extra in runs[1:]:
            p.remove(extra)
        first.find(A + 't').text = text
    else:
        end = p.find(A + 'endParaRPr')
        r = ET.Element(A + 'r')
        if end is not None:
            rPr = copy.deepcopy(end)
            rPr.tag = A + 'rPr'
            r.append(rPr)
        t = ET.SubElement(r, A + 't')
        t.text = text
        p.insert(list(p).index(end) if end is not None else 0, r)
    return p


def pptx_fill_cell(tc, lines, protos):
    tx = tc.find(A + 'txBody')
    for p in tx.findall(A + 'p'):
        tx.remove(p)
    for text, kind in lines:
        tx.append(pptx_para_with_text(protos[kind], text))


def build_ppt(data, output):
    template = os.path.join(TEMPLATES, 'unit_mix_table_template.pptx')
    part = 'ppt/slides/slide1.xml'
    xml = load_part(template, part)
    root = parse_keep_ns(xml)

    if data.get('title'):
        tbl_texts = set()
        for tbl in root.iter(A + 'tbl'):
            for t in tbl.iter(A + 't'):
                tbl_texts.add(id(t))
        for t in root.iter(A + 't'):
            if id(t) not in tbl_texts and (t.text or '').strip():
                t.text = data['title']
                break

    tbl = next(iter(root.iter(A + 'tbl')))
    trs = tbl.findall(A + 'tr')
    proto_row = trs[1]
    cell_protos = [pptx_prototypes(tc) for tc in proto_row.findall(A + 'tc')]

    for tr in trs[1:]:
        tbl.remove(tr)
    for proj in data['projects']:
        project, types, counts, mixes, sizes = project_lines(proj)
        # PPT column order: PROJECT | FLOOR PLAN | SF | UNIT COUNT | UNIT MIX %
        ordered = [project, types, sizes, counts, mixes]
        tr = copy.deepcopy(proto_row)
        tcs = tr.findall(A + 'tc')
        for tc, lines, protos in zip(tcs, ordered, cell_protos):
            pptx_fill_cell(tc, lines, protos)
        tbl.append(tr)

    save_with_part(template, output, part, serialize_restore_root(root, xml, 'sld'))


# ------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True, help='JSON input file')
    ap.add_argument('--format', choices=['word', 'ppt'], required=True)
    ap.add_argument('--output', required=True)
    ap.add_argument('--variant', choices=['forsale', 'rental'], default='forsale',
                    help='forsale (default): 5-column table, Word or PPT. '
                         'rental: 7-column purpose-built rental chart, Word only.')
    args = ap.parse_args()
    if args.variant == 'rental' and args.format != 'word':
        sys.exit('ERROR: --variant rental is Word only (the rental chart has no PowerPoint '
                 'version). Re-run with --format word.')

    with open(args.data, encoding='utf-8') as f:
        data = json.load(f)
    if not data.get('projects'):
        sys.exit('ERROR: no projects in input JSON')
    for proj in data['projects']:
        if not proj.get('units'):
            sys.exit(f"ERROR: project {proj.get('name')} has no units")

    if args.variant == 'rental':
        n = build_word_rental(data, args.output)
        print(f'Rental chart: {len(data["projects"])} projects across {n} page(s)')
    elif args.format == 'word':
        build_word(data, args.output)
    else:
        build_ppt(data, args.output)
    print(f'Wrote {args.output}')


if __name__ == '__main__':
    main()
