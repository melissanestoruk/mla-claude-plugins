#!/usr/bin/env python3
"""Build an MLA Upcoming Project Table (Word or PowerPoint) from a HubSpot CRM
export of upcoming/pipeline projects.

Clone-and-swap on the gold templates in ../templates/ (location column removed —
the market is stated in the table's title/header, not per row).

Word columns: PROJECT | DEVELOPER/APPLICANT | UNITS | STATUS/TIMING
PPT columns:  PROJECT | DEVELOPER/APPLICANT | STORIES | UNITS | STATUS/TIMING

STATUS/TIMING = pipeline stage + estimated launch, e.g. "Approved / Est. Oct 2026".
Rows sort by estimated launch date, soonest first. TOTAL row sums units.

Usage:
  python build_upcoming_table.py --csv pull.csv --format word --output out.docx
  python build_upcoming_table.py --csv pull.csv --format ppt --title "Upcoming Concrete Condominium" --output out.pptx

Stdlib only. Same table for for-sale and rental product.
"""
import argparse
import copy
import csv
import os
import re
import sys
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = os.path.join(HERE, '..', 'templates')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

PPT_ROW_WARN = 15

STAGE_DISPLAY = {
    'approved/upcoming': 'Approved',
    'in planning': 'In Planning',
}


# ---------------------------------------------------------------- CSV parsing

def norm(s):
    return re.sub(r'[^a-z0-9]+', ' ', (s or '').lower()).strip()


def find_col(headers, *candidates):
    normed = {h: norm(h) for h in headers}
    for cand in candidates:
        for h, n in normed.items():
            if n == cand:
                return h
    for cand in candidates:
        for h, n in normed.items():
            if cand in n:
                return h
    return None


def to_num(val):
    s = str(val or '').strip().replace('$', '').replace(',', '')
    if not s or s in ('-',):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(val):
    s = str(val or '').strip()
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y', '%b %Y', '%B %Y'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def fmt_int(val):
    n = to_num(val)
    return f'{int(round(n)):,}' if n is not None else '-'


def status_label(stage, date_val):
    stage = (stage or '').strip()
    disp = STAGE_DISPLAY.get(stage.lower(), stage) or ''
    d = parse_date(date_val)
    est = f'Est. {d.strftime("%b %Y")}' if d else ''
    if disp and est:
        return f'{disp} / {est}'
    return disp or est or '-'


def read_rows(csv_path):
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        raw = [r for r in reader if any((v or '').strip() for v in r.values())]
        headers = reader.fieldnames or []

    cmap = {
        'project': find_col(headers, 'name', 'project name', 'project'),
        'developer': find_col(headers, 'developer'),
        'date': find_col(headers, 'contract writing date', 'launch date', 'est launch'),
        'storeys': find_col(headers, 'storeys', 'stories', 'storey'),
        'units': find_col(headers, 'total units', 'units'),
        'stage': find_col(headers, 'project pipeline stage', 'pipeline stage', 'stage', 'status'),
    }
    missing = [k for k in ('project', 'developer', 'date', 'units') if not cmap.get(k)]
    if missing:
        sys.exit(f'ERROR: could not find CSV columns for: {", ".join(missing)}\n'
                 f'CSV headers found: {", ".join(headers)}')

    rows = []
    for r in raw:
        row = {k: (r.get(h, '') if h else '') for k, h in cmap.items()}
        row['project'] = (row['project'] or '').strip() or '-'
        row['developer'] = (row['developer'] or '').strip().strip('-').strip() or '-'
        rows.append(row)

    rows.sort(key=lambda r: parse_date(r['date']) or datetime.max)
    return rows


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
    out = ET.tostring(root, encoding='unicode')
    new_tag = re.search(r'<[\w:]*%s[^>]*>' % root_local, out).group(0)
    orig_tag = re.search(r'<[\w:]*%s[^>]*>' % root_local, orig_xml).group(0)
    missing = [m.group(0) for m in re.finditer(r'xmlns:[\w-]+="[^"]+"', orig_tag)
               if m.group(0) not in new_tag]
    if missing:
        out = out.replace(new_tag, new_tag[:-1].rstrip() + ' ' + ' '.join(missing) + '>', 1)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out


# ------------------------------------------------------------ Word building

def docx_set_cell(tc, text):
    paras = tc.findall(W + 'p')
    p = paras[0]
    for extra in paras[1:]:
        tc.remove(extra)
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


def build_word(rows, output):
    template = os.path.join(TEMPLATES, 'upcoming_table_template.docx')
    xml = load_part(template, 'word/document.xml')
    root = parse_keep_ns(xml)
    tbl = next(iter(root.iter(W + 'tbl')))
    trs = tbl.findall(W + 'tr')
    header, data_rows, total_row = trs[0], trs[1:-1], trs[-1]
    proto = copy.deepcopy(data_rows[0])
    for tr in data_rows:
        tbl.remove(tr)

    insert_at = list(tbl).index(total_row)
    for r in rows:
        tr = copy.deepcopy(proto)
        tcs = tr.findall(W + 'tc')
        texts = [r['project'], r['developer'], fmt_int(r['units']),
                 status_label(r['stage'], r['date'])]
        for tc, text in zip(tcs, texts):
            docx_set_cell(tc, text)
        tbl.insert(insert_at, tr)
        insert_at += 1

    tcs = total_row.findall(W + 'tc')
    docx_set_cell(tcs[1], fmt_int(sum(to_num(r['units']) or 0 for r in rows)))

    save_with_part(template, output, 'word/document.xml',
                   serialize_restore_root(root, xml, 'document'))


# -------------------------------------------------------- PowerPoint building

def pptx_set_cell(tc, text):
    tx = tc.find(A + 'txBody')
    paras = tx.findall(A + 'p')
    p = paras[0]
    for extra in paras[1:]:
        tx.remove(extra)
    runs = p.findall(A + 'r')
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


def build_ppt(rows, title, output):
    if len(rows) > PPT_ROW_WARN:
        print(f'WARNING: {len(rows)} projects is a lot for one slide — '
              f'consider filtering the pull by market/date first.', file=sys.stderr)
    template = os.path.join(TEMPLATES, 'upcoming_table_template.pptx')
    part = 'ppt/slides/slide1.xml'
    xml = load_part(template, part)
    root = parse_keep_ns(xml)

    P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
    tbl_texts = set()
    for tbl in root.iter(A + 'tbl'):
        for t in tbl.iter(A + 't'):
            tbl_texts.add(id(t))
    # swap the title; blank the project-specific '*' footnote shapes (manual notes)
    title_done = False
    for tx in root.iter(P + 'txBody'):
        text = ''.join(t.text or '' for t in tx.iter(A + 't'))
        if any(id(t) in tbl_texts for t in tx.iter(A + 't')):
            continue
        if text.strip().startswith('*'):
            for t in tx.iter(A + 't'):
                t.text = ''
        elif text.strip() and title and not title_done:
            ts = [t for t in tx.iter(A + 't')]
            ts[0].text = title
            for t in ts[1:]:
                t.text = ''
            title_done = True

    tbl = next(iter(root.iter(A + 'tbl')))
    trs = tbl.findall(A + 'tr')
    header, data_rows, total_row = trs[0], trs[1:-1], trs[-1]
    proto = copy.deepcopy(data_rows[0])
    for tr in data_rows:
        tbl.remove(tr)

    insert_at = list(tbl).index(total_row)
    for r in rows:
        tr = copy.deepcopy(proto)
        tcs = tr.findall(A + 'tc')
        texts = [r['project'], r['developer'], fmt_int(r.get('storeys')),
                 fmt_int(r['units']), status_label(r['stage'], r['date'])]
        for tc, text in zip(tcs, texts):
            pptx_set_cell(tc, text)
        tbl.insert(insert_at, tr)
        insert_at += 1

    # TOTAL row: [merged label, hMerge, hMerge, UNITS, STATUS]
    tcs = total_row.findall(A + 'tc')
    pptx_set_cell(tcs[-2], fmt_int(sum(to_num(r['units']) or 0 for r in rows)))

    save_with_part(template, output, part, serialize_restore_root(root, xml, 'sld'))


# ------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--csv', required=True)
    ap.add_argument('--format', choices=['word', 'ppt'], required=True)
    ap.add_argument('--title', default='', help='PPT slide title (Word table has no title)')
    ap.add_argument('--output', required=True)
    args = ap.parse_args()

    rows = read_rows(args.csv)
    if args.format == 'word':
        build_word(rows, args.output)
    else:
        build_ppt(rows, args.title, args.output)
    print(f'Wrote {args.output} ({len(rows)} projects)')


if __name__ == '__main__':
    main()
