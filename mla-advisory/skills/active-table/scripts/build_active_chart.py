#!/usr/bin/env python3
"""Build an MLA Active Chart (for-sale or rental) from a HubSpot CRM export.

Clone-and-swap on the gold templates in ../templates/ — brand formatting is never
rebuilt, only the data rows are replaced.

Usage:
  python build_active_chart.py --csv pull.csv --variant forsale --format word \
      --title "Coquitlam Condominium Competition Analysis" --output out.docx
  python build_active_chart.py --csv pull.csv --variant forsale --format ppt \
      --title "..." --output out.pptx

Stdlib only. Word output for both variants; PowerPoint output for forsale only.
"""
import argparse
import copy
import csv
import io
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


# ---------------------------------------------------------------- CSV parsing

def norm(s):
    return re.sub(r'[^a-z0-9]+', ' ', (s or '').lower()).strip()


def find_col(headers, *candidates):
    """Return the header whose normalized form contains any candidate (in order)."""
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
    """Parse a number out of '$1,100', '810', '55%'. Return None if not numeric."""
    if val is None:
        return None
    s = str(val).strip().replace('$', '').replace(',', '').replace('%', '')
    if not s or s in ('-', 'n/a', 'na', 'tbd'):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def fmt_month(val):
    """'3/14/2026' -> 'Mar 2026'. Pass through anything already non-date-like."""
    s = str(val or '').strip()
    if not s:
        return '-'
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%d-%b-%Y', '%B %d, %Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%b %Y')
        except ValueError:
            continue
    return s  # already 'Mar 2026' or similar


def parse_month(val):
    s = str(val or '').strip()
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%d-%b-%Y', '%B %d, %Y', '%b %Y', '%B %Y'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return datetime.min  # unparseable dates sort last


def fmt_int(val):
    n = to_num(val)
    return f'{int(round(n)):,}' if n is not None else '-'


def fmt_money0(val):
    n = to_num(val)
    if n is None:
        s = str(val or '').strip()
        return s if s else '-'  # keep 'Est. $2,930' style text as-is
    return f'${int(round(n)):,}'


def fmt_money2(val):
    n = to_num(val)
    if n is None:
        s = str(val or '').strip()
        return s if s else '-'
    return f'${n:,.2f}'


def fmt_total(total, numerator):
    """'166 (57%)' — percent of total sold/leased; '(-%)' when numerator blank."""
    t = to_num(total)
    if t is None:
        return '-'
    n = to_num(numerator)
    pct = f'{round(n / t * 100)}%' if (n is not None and t) else '-%'
    return f'{int(round(t)):,} ({pct})'


def read_rows(csv_path, variant):
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        raw = [r for r in reader if any((v or '').strip() for v in r.values())]
        headers = reader.fieldnames or []

    def col(*cands):
        return find_col(headers, *cands)

    if variant == 'forsale':
        cmap = {
            'project': col('name', 'project name', 'project'),
            'developer': col('developer'),
            'launch': col('contract writing date', 'launch date', 'launch'),
            'avg_sf': col('avg unit size estimate', 'avg unit size', 'avg sf', 'average size'),
            'ppsf': col('ppsf estimate', 'ppsf'),
            'sold': col('units sold', 'sold'),
            'released': col('units released', 'released', 'rlsd'),
            'total': col('total units', 'total'),
            'type': col('construction type', 'building type', 'type'),
        }
        required = ['project', 'developer', 'launch', 'avg_sf', 'ppsf', 'sold', 'released', 'total']
    else:
        cmap = {
            'project': col('name', 'project name', 'project'),
            'developer': col('developer'),
            'market': col('submarket', 'market'),
            'type': col('construction type', 'building type', 'type'),
            'storeys': col('storeys', 'stories', 'storey'),
            'occupancy': col('occupancy date', 'occupancy'),
            'avg_sf': col('avg unit size estimate', 'avg unit size', 'avg sf', 'average size'),
            'avg_rent': col('avg rent estimate', 'avg rent', 'average rent', 'rent'),
            'initial_psf': col('initial psf', 'psf estimate', 'ppsf'),
            'leased': col('units leased', 'leased'),
            'leases_month': col('leases per month', 'leases month', 'per month'),
            'total': col('total units', 'total'),
        }
        required = ['project', 'developer', 'market', 'occupancy', 'avg_sf',
                    'avg_rent', 'initial_psf', 'leased', 'total']

    missing = [k for k in required if not cmap.get(k)]
    if missing:
        sys.exit(f'ERROR: could not find CSV columns for: {", ".join(missing)}\n'
                 f'CSV headers found: {", ".join(headers)}')

    rows = []
    for r in raw:
        rows.append({k: (r.get(h, '') if h else '') for k, h in cmap.items()})

    sort_key = 'launch' if variant == 'forsale' else 'occupancy'
    rows.sort(key=lambda r: parse_month(r[sort_key]), reverse=True)
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
    """Parse XML, registering all root namespaces so ET round-trips prefixes."""
    for prefix, uri in re.findall(r'xmlns:([\w-]+)="([^"]+)"', xml_text):
        ET.register_namespace(prefix, uri)
    return ET.fromstring(xml_text)


def serialize_restore_root(root, orig_xml, root_local):
    """ET drops xmlns decls that are only referenced via mc:Ignorable; merge the
    original root's declarations into the generated root tag (never remove any)."""
    out = ET.tostring(root, encoding='unicode')
    new_tag = re.search(r'<[\w:]*%s[^>]*>' % root_local, out).group(0)
    orig_tag = re.search(r'<[\w:]*%s[^>]*>' % root_local, orig_xml).group(0)
    missing = [m.group(0) for m in re.finditer(r'xmlns:[\w-]+="[^"]+"', orig_tag)
               if m.group(0) not in new_tag]
    if missing:
        merged = new_tag[:-1].rstrip() + ' ' + ' '.join(missing) + '>'
        out = out.replace(new_tag, merged, 1)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out


# ------------------------------------------------------------- Word building

def docx_set_cell(tc, text):
    """Put text in a table cell, preserving the prototype's run formatting."""
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


def build_word(csv_path, variant, title, output):
    template = os.path.join(TEMPLATES, f'active_chart_{variant}_template.docx')
    rows = read_rows(csv_path, variant)
    xml = load_part(template, 'word/document.xml')
    root = parse_keep_ns(xml)
    body = root.find(W + 'body')

    # title: first paragraph with text
    if title:
        for p in body.findall(W + 'p'):
            if any((t.text or '').strip() for t in p.iter(W + 't')):
                docx_para_set_text(p, title)
                break

    tbl = body.find(W + 'tbl')
    if tbl is None:  # table may be nested after paragraphs
        tbl = next(iter(body.iter(W + 'tbl')))
    trs = tbl.findall(W + 'tr')
    header, data_rows, total_row = trs[0], trs[1:-1], trs[-1]
    proto = [copy.deepcopy(data_rows[0]),
             copy.deepcopy(data_rows[1] if len(data_rows) > 1 else data_rows[0])]
    for tr in data_rows:
        tbl.remove(tr)

    if variant == 'forsale':
        cells_of = lambda r: [r['project'], r['developer'], fmt_month(r['launch']),
                              fmt_int(r['avg_sf']), fmt_money0(r['ppsf']),
                              fmt_int(r['sold']), fmt_int(r['released']),
                              fmt_total(r['total'], r['sold'])]
        totals = [fmt_int(sum(to_num(r['sold']) or 0 for r in rows)),
                  fmt_int(sum(to_num(r['released']) or 0 for r in rows)),
                  fmt_total(sum(to_num(r['total']) or 0 for r in rows),
                            sum(to_num(r['sold']) or 0 for r in rows))]
    else:
        cells_of = lambda r: [r['project'], r['developer'], r['market'],
                              (r.get('type') or '-').strip() or '-',
                              fmt_int(r.get('storeys')), fmt_month(r['occupancy']),
                              fmt_int(r['avg_sf']), fmt_money0(r['avg_rent']),
                              fmt_money2(r['initial_psf']), fmt_int(r['leased']),
                              fmt_int(r.get('leases_month')),
                              fmt_total(r['total'], r['leased'])]
        lm = [to_num(r.get('leases_month')) for r in rows]
        lm = [v for v in lm if v is not None]
        totals = [fmt_int(sum(to_num(r['leased']) or 0 for r in rows)),
                  fmt_int(sum(lm) / len(lm)) if lm else '-',
                  fmt_total(sum(to_num(r['total']) or 0 for r in rows),
                            sum(to_num(r['leased']) or 0 for r in rows))]

    insert_at = list(tbl).index(total_row)
    for i, r in enumerate(rows):
        tr = copy.deepcopy(proto[i % 2])
        tcs = tr.findall(W + 'tc')
        for tc, text in zip(tcs, cells_of(r)):
            docx_set_cell(tc, text)
        tbl.insert(insert_at, tr)
        insert_at += 1

    # total row: first cell is the merged label, last N cells are the sums
    tcs = total_row.findall(W + 'tc')
    for tc, text in zip(tcs[-len(totals):], totals):
        docx_set_cell(tc, text)

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
        # empty prototype cell: build a run, inheriting endParaRPr formatting
        end = p.find(A + 'endParaRPr')
        r = ET.Element(A + 'r')
        if end is not None:
            rPr = copy.deepcopy(end)
            rPr.tag = A + 'rPr'
            r.append(rPr)
        t = ET.SubElement(r, A + 't')
        t.text = text
        p.insert(list(p).index(end) if end is not None else 0, r)


def build_ppt(csv_path, title, output):
    template = os.path.join(TEMPLATES, 'active_chart_forsale.pptx')
    rows = read_rows(csv_path, 'forsale')
    part = 'ppt/slides/slide1.xml'
    xml = load_part(template, part)
    root = parse_keep_ns(xml)

    # title
    if title:
        for t in root.iter(A + 't'):
            if 'Competition Analysis' in (t.text or ''):
                t.text = title
                break

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
        texts = [r['project'], r['developer'], (r.get('type') or '-').strip() or '-',
                 fmt_month(r['launch']), fmt_int(r['avg_sf']), fmt_money0(r['ppsf']),
                 fmt_int(r['sold']), fmt_int(r['released']),
                 fmt_total(r['total'], r['sold'])]
        for tc, text in zip(tcs, texts):
            pptx_set_cell(tc, text)
        tbl.insert(insert_at, tr)
        insert_at += 1

    tcs = total_row.findall(A + 'tc')
    sold_sum = sum(to_num(r['sold']) or 0 for r in rows)
    pptx_set_cell(tcs[0], 'TOTAL')
    for i in range(1, 6):
        pptx_set_cell(tcs[i], '')
    pptx_set_cell(tcs[6], fmt_int(sold_sum))
    pptx_set_cell(tcs[7], fmt_int(sum(to_num(r['released']) or 0 for r in rows)))
    pptx_set_cell(tcs[8], fmt_total(sum(to_num(r['total']) or 0 for r in rows), sold_sum))

    # footnotes: keep the generic '*Approximate PPSF...' line, drop the
    # project-specific '**'/'***' lines (those are added manually per deliverable);
    # shape text bodies are p:txBody (presentationml), table cells a:txBody
    P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
    for tx in list(root.iter(P + 'txBody')) + list(root.iter(A + 'txBody')):
        paras = tx.findall(A + 'p')
        for p in paras:
            text = ''.join(t.text or '' for t in p.iter(A + 't'))
            if text.strip().startswith('**'):
                tx.remove(p)

    save_with_part(template, output, part, serialize_restore_root(root, xml, 'sld'))


# ------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--csv', required=True)
    ap.add_argument('--variant', choices=['forsale', 'rental'], required=True)
    ap.add_argument('--format', choices=['word', 'ppt'], required=True)
    ap.add_argument('--title', default='')
    ap.add_argument('--output', required=True)
    args = ap.parse_args()

    if args.format == 'ppt' and args.variant == 'rental':
        sys.exit('ERROR: rental Active Chart is Word-only (no PowerPoint gold template yet).')

    if args.format == 'word':
        build_word(args.csv, args.variant, args.title, args.output)
    else:
        build_ppt(args.csv, args.title, args.output)
    print(f'Wrote {args.output}')


if __name__ == '__main__':
    main()
