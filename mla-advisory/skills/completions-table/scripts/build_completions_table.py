#!/usr/bin/env python3
"""Build an MLA Completions Table from a HubSpot CRM export of projects
completing in a market.

Clone-and-swap on the gold templates in ../templates/.

Word: one table — PROJECT | TYPE | DEVELOPER | LAUNCH | SALES | UNITS | EST. COMPLETION
      (UNITS shows "178 (63%)" — percent sold; TOTAL row sums sales + units).
PPT:  slide 1 — the same table (UNITS as a plain number);
      slide 2 — the Sold/Unsold by quarter bar chart, cached values rewritten
      from the pull (categories = completion quarters, Sold = sum of Units Sold,
      Unsold = sum of Total Units − Units Sold).

Rows sort by completion date, soonest first. TYPE: Concrete→CON, Wood Frame→WF.
Launch dates missing from the pull render "TBD" (matches the gold).

Usage:
  python build_completions_table.py --csv pull.csv --format word --output out.docx
  python build_completions_table.py --csv pull.csv --format ppt --title "Upcoming Condominium Completions" --output out.pptx

Stdlib only.
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
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

PPT_ROW_WARN = 16

TYPE_MAP = {'concrete': 'CON', 'wood frame': 'WF', 'con+wf (mixed use)': 'CON+WF'}


# ---------------------------------------------------------------- CSV parsing

def norm(s):
    return re.sub(r'[^a-z0-9]+', ' ', (s or '').lower()).strip()


def find_col(headers, *candidates):
    normed = {h: norm(h) for h in headers}
    for cand in candidates:
        nc = norm(cand)
        for h, n in normed.items():
            if n == nc:
                return h
    for cand in candidates:
        nc = norm(cand)
        for h, n in normed.items():
            if nc in n:
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
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def fmt_int(val):
    n = to_num(val)
    return f'{int(round(n)):,}' if n is not None else '-'


def quarter_label(val):
    d = parse_date(val)
    if not d:
        return str(val or 'TBD').strip() or 'TBD'
    return f'Q{(d.month - 1) // 3 + 1} {d.year}'


def units_with_pct(total, sold):
    t = to_num(total)
    if t is None:
        return '-'
    s = to_num(sold)
    pct = f'{round(s / t * 100)}%' if (s is not None and t) else '-%'
    return f'{int(round(t)):,} ({pct})'


def read_rows(csv_path):
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        raw = [r for r in reader if any((v or '').strip() for v in r.values())]
        headers = reader.fieldnames or []

    cmap = {
        'project': find_col(headers, 'name', 'project name', 'project'),
        'material': find_col(headers, 'construction material', 'construction type'),
        'developer': find_col(headers, 'developer'),
        'launch': find_col(headers, 'contract writing date', 'launch date'),
        'sold': find_col(headers, 'units sold', 'sold'),
        'total': find_col(headers, 'total units', 'units'),
        'complete': find_col(headers, 'completion date', 'completion'),
    }
    missing = [k for k in ('project', 'developer', 'sold', 'total', 'complete') if not cmap.get(k)]
    if missing:
        sys.exit(f'ERROR: could not find CSV columns for: {", ".join(missing)}\n'
                 f'CSV headers found: {", ".join(headers)}')

    rows = []
    for r in raw:
        g = lambda k: (r.get(cmap[k], '') if cmap.get(k) else '')
        d = parse_date(g('launch'))
        rows.append({
            'project': g('project').strip(),
            'type': TYPE_MAP.get(norm(g('material')), g('material').strip() or '-'),
            'developer': g('developer').strip() or '-',
            'launch': d.strftime('%b %Y') if d else 'TBD',
            'sold': g('sold'),
            'total': g('total'),
            'complete_date': parse_date(g('complete')),
            'quarter': quarter_label(g('complete')),
        })
    rows.sort(key=lambda r: r['complete_date'] or datetime.max)
    return rows


def quarter_series(rows):
    """Ordered (quarter, sold_sum, unsold_sum) triples for the bar chart."""
    order, agg = [], {}
    for r in rows:
        q = r['quarter']
        if q not in agg:
            agg[q] = [0, 0]
            order.append(q)
        s = to_num(r['sold']) or 0
        t = to_num(r['total']) or 0
        agg[q][0] += s
        agg[q][1] += max(t - s, 0)
    return [(q, int(agg[q][0]), int(agg[q][1])) for q in order]


# ------------------------------------------------------- shared zip plumbing

def load_part(path, part):
    with zipfile.ZipFile(path) as z:
        return z.read(part).decode('utf-8')


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


def repack(template, output, replacements):
    with zipfile.ZipFile(template) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename in replacements:
                data = replacements[info.filename].encode('utf-8')
            zout.writestr(info, data)


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
    template = os.path.join(TEMPLATES, 'completions_table_template.docx')
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
        texts = [r['project'], r['type'], r['developer'], r['launch'],
                 fmt_int(r['sold']), units_with_pct(r['total'], r['sold']), r['quarter']]
        for tc, text in zip(tcs, texts):
            docx_set_cell(tc, text)
        tbl.insert(insert_at, tr)
        insert_at += 1

    sold_sum = sum(to_num(r['sold']) or 0 for r in rows)
    total_sum = sum(to_num(r['total']) or 0 for r in rows)
    tcs = total_row.findall(W + 'tc')
    docx_set_cell(tcs[1], fmt_int(sold_sum))
    docx_set_cell(tcs[2], units_with_pct(total_sum, sold_sum))

    repack(template, output,
           {'word/document.xml': serialize_restore_root(root, xml, 'document')})


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


def swap_title_blank_notes(root, title):
    """Replace the first non-table text shape with the title; blank the rest
    (example-specific footnotes/annotations are manual per deliverable)."""
    tbl_texts = set()
    for tbl in root.iter(A + 'tbl'):
        for t in tbl.iter(A + 't'):
            tbl_texts.add(id(t))
    title_done = False
    for tx in root.iter(P + 'txBody'):
        ts = [t for t in tx.iter(A + 't')]
        if not ts or any(id(t) in tbl_texts for t in ts):
            continue
        text = ''.join(t.text or '' for t in ts).strip()
        if not text:
            continue
        if not title_done:
            if title:
                ts[0].text = title
                for t in ts[1:]:
                    t.text = ''
            title_done = True
        else:
            for t in ts:
                t.text = ''


def set_cache(ser, tag_ref, cache_tag, f_text, values, fmt=None):
    """Rewrite a c:cat/c:val cache: c:f range, ptCount, and pt values."""
    holder = ser.find(tag_ref[0])
    ref = holder.find(tag_ref[1])
    ref.find(C + 'f').text = f_text
    cache = ref.find(cache_tag)
    for pt in cache.findall(C + 'pt'):
        cache.remove(pt)
    cache.find(C + 'ptCount').set('val', str(len(values)))
    for i, v in enumerate(values):
        pt = ET.SubElement(cache, C + 'pt', {'idx': str(i)})
        ET.SubElement(pt, C + 'v').text = str(v)


def build_ppt(rows, title, output):
    if len(rows) > PPT_ROW_WARN:
        print(f'WARNING: {len(rows)} projects is a lot for one slide — '
              f'consider filtering the pull first.', file=sys.stderr)
    template = os.path.join(TEMPLATES, 'completions_table_template.pptx')

    # ---- slide 1: the table
    part1 = 'ppt/slides/slide1.xml'
    xml1 = load_part(template, part1)
    root1 = parse_keep_ns(xml1)
    swap_title_blank_notes(root1, title)
    tbl = next(iter(root1.iter(A + 'tbl')))
    trs = tbl.findall(A + 'tr')
    header, data_rows, total_row = trs[0], trs[1:-1], trs[-1]
    proto = copy.deepcopy(data_rows[0])
    for tr in data_rows:
        tbl.remove(tr)
    insert_at = list(tbl).index(total_row)
    for r in rows:
        tr = copy.deepcopy(proto)
        tcs = tr.findall(A + 'tc')
        texts = [r['project'], r['type'], r['developer'], r['launch'],
                 fmt_int(r['sold']), fmt_int(r['total']), r['quarter']]
        for tc, text in zip(tcs, texts):
            pptx_set_cell(tc, text)
        tbl.insert(insert_at, tr)
        insert_at += 1
    tcs = total_row.findall(A + 'tc')
    pptx_set_cell(tcs[-3], fmt_int(sum(to_num(r['sold']) or 0 for r in rows)))
    pptx_set_cell(tcs[-2], fmt_int(sum(to_num(r['total']) or 0 for r in rows)))
    pptx_set_cell(tcs[-1], '')

    # ---- slide 2: the Sold/Unsold by quarter bar chart
    part2 = 'ppt/slides/slide2.xml'
    xml2 = load_part(template, part2)
    root2 = parse_keep_ns(xml2)
    swap_title_blank_notes(root2, title)

    chart_part = 'ppt/charts/chart1.xml'
    cxml = load_part(template, chart_part)
    croot = parse_keep_ns(cxml)
    quarters = quarter_series(rows)
    cats = [q for q, _, _ in quarters]
    n = len(cats)
    sers = list(croot.iter(C + 'ser'))
    for si, ser in enumerate(sers):
        vals = [s for _, s, _ in quarters] if si == 0 else [u for _, _, u in quarters]
        set_cache(ser, (C + 'cat', C + 'strRef'), C + 'strCache',
                  f'Sheet2!$A$22:$A${21 + n}', cats)
        col = 'B' if si == 0 else 'C'
        set_cache(ser, (C + 'val', C + 'numRef'), C + 'numCache',
                  f'Sheet2!${col}$22:${col}${21 + n}', vals)

    repack(template, output, {
        part1: serialize_restore_root(root1, xml1, 'sld'),
        part2: serialize_restore_root(root2, xml2, 'sld'),
        chart_part: serialize_restore_root(croot, cxml, 'chartSpace'),
    })


# ------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--csv', required=True)
    ap.add_argument('--format', choices=['word', 'ppt'], required=True)
    ap.add_argument('--title', default='', help='PPT slide title (both slides)')
    ap.add_argument('--output', required=True)
    args = ap.parse_args()

    rows = read_rows(args.csv)
    if not rows:
        sys.exit('ERROR: no rows found in the CSV')
    if args.format == 'word':
        build_word(rows, args.output)
    else:
        build_ppt(rows, args.title, args.output)
    print(f'Wrote {args.output} ({len(rows)} projects)')


if __name__ == '__main__':
    main()
