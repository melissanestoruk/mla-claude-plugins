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
EN_DASH = '–'


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
    for proj in data['projects']:
        tr = copy.deepcopy(proto_row)
        tcs = tr.findall(W + 'tc')
        for tc, lines, protos in zip(tcs, project_lines(proj), cell_protos):
            docx_fill_cell(tc, lines, protos)
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
    args = ap.parse_args()

    with open(args.data, encoding='utf-8') as f:
        data = json.load(f)
    if not data.get('projects'):
        sys.exit('ERROR: no projects in input JSON')
    for proj in data['projects']:
        if not proj.get('units'):
            sys.exit(f"ERROR: project {proj.get('name')} has no units")

    if args.format == 'word':
        build_word(data, args.output)
    else:
        build_ppt(data, args.output)
    print(f'Wrote {args.output}')


if __name__ == '__main__':
    main()
