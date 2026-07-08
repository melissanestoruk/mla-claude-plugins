#!/usr/bin/env python3
"""Build an MLA Offering Table (Word or PowerPoint) from a HubSpot CRM export.

Transposed layout: projects are COLUMNS, offering attributes are ROWS.
Clone-and-swap on the gold templates in ../templates/.

Variants:
  forsale — rows: DEPOSIT, REALTOR COMMISSION, REALTOR BONUS/INCENTIVE,
            PURCHASER INCENTIVE, ASSIGNMENT FEE, STRATA ($/SF), COMPLETION,
            PARKING/STORAGE, APPLIANCES, LATEST SALES ACTIVITY (manual, "-")
  rental  — rows: INCENTIVE, PET POLICY, PARKING RATIO, PARKING, AMENITIES,
            A/C, FLOORING, STORAGE, APPLIANCES, LATEST MESSAGING

Chunking: Word fits 5 projects per table (extra tables stack below);
PPT fits 4 projects per slide (extra slides are appended automatically).

Usage:
  python build_offering_table.py --csv pull.csv --variant forsale --format word --output out.docx
  python build_offering_table.py --csv pull.csv --variant forsale --format ppt --title "Competition Offering – Concrete" --output out.pptx

Stdlib only. Projects keep the pull's row order.
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
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

WORD_CHUNK = 5
PPT_CHUNK = 4


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


def clean(val):
    s = str(val or '').strip()
    return s if s and s not in ('-',) else ''


def to_num(val):
    s = clean(val).replace('$', '').replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None


def money_k(val):
    n = to_num(val)
    if not n:
        return ''
    if n >= 1000:
        k = n / 1000
        return f'${k:g}K'
    return f'${n:g}'


def pct_txt(val):
    n = to_num(val)
    if n is None:
        return clean(val)
    if n <= 1:
        n *= 100
    return f'{n:g}%'


def fmt_month(val):
    s = clean(val)
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y'):
        try:
            return datetime.strptime(s, fmt).strftime('%b %Y')
        except ValueError:
            continue
    return s or '-'


def dash(s):
    return clean(s) or '-'


def compose_deposit(pct, structure):
    p, st = pct_txt(pct), clean(structure)
    if p and st:
        return f'{p}: {st}'
    return p or st or '-'


def compose_parking_storage(p_alloc, p_cost, s_alloc, s_cost):
    park = clean(p_alloc)
    pc = money_k(p_cost)
    if pc:
        park = f'{park} ({pc}/stall)'.strip() if park else f'{pc}/stall'
    stor = clean(s_alloc)
    sc = money_k(s_cost)
    if sc:
        stor = f'{stor} ({sc}/locker)'.strip() if stor else f'{sc}/locker'
    if park and stor:
        return f'{park} / {stor}'
    return park or stor or '-'


def compose_appliances(pkg, sizing):
    pkg, sizing = clean(pkg), clean(sizing)
    if pkg and sizing:
        return f'{pkg}; {sizing}'
    return pkg or sizing or '-'


def strata_txt(val):
    n = to_num(val)
    return f'${n:.2f}' if n is not None else dash(val)


def read_projects(csv_path, variant):
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        raw = [r for r in reader if any((v or '').strip() for v in r.values())]
        headers = reader.fieldnames or []

    def col(*cands):
        return find_col(headers, *cands)

    def get(r, h):
        return r.get(h, '') if h else ''

    projects = []
    if variant == 'forsale':
        c = {
            'name': col('name', 'project'),
            'dep_pct': col('deposit %', 'deposit pct'),
            'dep_str': col('deposit structure'),
            'comm': col('realtor commission'),
            'bonus': col('realtor incentive', 'realtor bonus'),
            'incent': col('buyer incentive', 'purchaser incentive'),
            'assign': col('assignment policy', 'assignment fee'),
            'strata': col('strata fee per sqft', 'strata'),
            'complete': col('completion date', 'completion'),
            'p_alloc': col('parking allocation'),
            'p_cost': col('parking stall cost'),
            's_alloc': col('storage allocation'),
            's_cost': col('storage locker cost'),
            'app_pkg': col('appliance package'),
            'app_size': col('appliance sizing'),
        }
        if not c['name']:
            sys.exit(f'ERROR: no project name column. Headers: {", ".join(headers)}')
        for r in raw:
            projects.append({
                'name': clean(get(r, c['name'])).upper(),
                'rows': [
                    compose_deposit(get(r, c['dep_pct']), get(r, c['dep_str'])),
                    dash(get(r, c['comm'])),
                    dash(get(r, c['bonus'])),
                    dash(get(r, c['incent'])),
                    dash(get(r, c['assign'])),
                    strata_txt(get(r, c['strata'])),
                    fmt_month(get(r, c['complete'])),
                    compose_parking_storage(get(r, c['p_alloc']), get(r, c['p_cost']),
                                            get(r, c['s_alloc']), get(r, c['s_cost'])),
                    compose_appliances(get(r, c['app_pkg']), get(r, c['app_size'])),
                    '-',   # LATEST SALES ACTIVITY — manual fill
                ],
            })
    else:
        c = {
            'name': col('name', 'project'),
            'incent': col('rental incentive', 'incentive'),
            'pets': col('pet policy', 'pets'),
            'ratio': col('parking ratio'),
            'parking': col('parking cost', 'parking'),
            'amenities': col('amenities', 'amenity'),
            'ac': col('a c', 'air conditioning', 'cooling'),
            'floor': col('flooring'),
            'storage': col('storage cost', 'storage'),
            'app': col('appliance package', 'appliances'),
            'msg': col('latest messaging', 'messaging'),
        }
        if not c['name']:
            sys.exit(f'ERROR: no project name column. Headers: {", ".join(headers)}')
        for r in raw:
            projects.append({
                'name': clean(get(r, c['name'])).upper(),
                'rows': [
                    dash(get(r, c['incent'])),
                    dash(get(r, c['pets'])),
                    dash(get(r, c['ratio'])),
                    dash(get(r, c['parking'])),
                    dash(get(r, c['amenities'])),
                    dash(get(r, c['ac'])),
                    dash(get(r, c['floor'])),
                    dash(get(r, c['storage'])),
                    dash(get(r, c['app'])),
                    dash(get(r, c['msg'])),
                ],
            })
    return projects


def chunked(lst, n):
    return [lst[i:i + n] for i in range(0, len(lst), n)]


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


def fill_word_table(tbl, chunk, full_cols):
    """Fill a transposed table: header row gets project names, attribute rows get
    values. Trim surplus project columns, stretching the survivors."""
    trs = tbl.findall(W + 'tr')
    n = len(chunk)
    surplus = full_cols - n
    if surplus:
        grid = tbl.find(W + 'tblGrid')
        cols = grid.findall(W + 'gridCol')
        removed = sum(int(c.get(W + 'w')) for c in cols[1 + n:])
        for c in cols[1 + n:]:
            grid.remove(c)
        extra = removed // n
        for c in grid.findall(W + 'gridCol')[1:]:
            c.set(W + 'w', str(int(c.get(W + 'w')) + extra))
        for tr in trs:
            tcs = tr.findall(W + 'tc')
            rem_pct = 0
            for tc in tcs[1 + n:]:
                wv = tc.find(W + 'tcPr').find(W + 'tcW')
                if wv is not None and wv.get(W + 'type') == 'pct':
                    rem_pct += int(wv.get(W + 'w'))
                tr.remove(tc)
            if rem_pct:
                for tc in tr.findall(W + 'tc')[1:]:
                    wv = tc.find(W + 'tcPr').find(W + 'tcW')
                    if wv is not None and wv.get(W + 'type') == 'pct':
                        wv.set(W + 'w', str(int(wv.get(W + 'w')) + rem_pct // n))
    for tr_i, tr in enumerate(tbl.findall(W + 'tr')):
        tcs = tr.findall(W + 'tc')
        for j, proj in enumerate(chunk):
            text = proj['name'] if tr_i == 0 else proj['rows'][tr_i - 1]
            docx_set_cell(tcs[1 + j], text)


def build_word(projects, variant, title, output):
    template = os.path.join(TEMPLATES, f'offering_{variant}_template.docx')
    xml = load_part(template, 'word/document.xml')
    root = parse_keep_ns(xml)
    body = root.find(W + 'body')

    if title:
        for p in body.findall(W + 'p'):
            if any((t.text or '').strip() for t in p.iter(W + 't')):
                docx_set_cell_like_para(p, title)
                break

    tbl = next(iter(body.iter(W + 'tbl')))
    full_cols = len(tbl.find(W + 'tblGrid').findall(W + 'gridCol')) - 1
    proto = copy.deepcopy(tbl)
    chunks = chunked(projects, WORD_CHUNK)

    fill_word_table(tbl, chunks[0], full_cols)
    # extra chunks: the gold table is a floating table, so stacked copies at the
    # same anchor overlap — give each chunk its own page via a page-break para,
    # inserted after the last content element (the caption)
    anchor = list(body)[-2]          # last element before the body sectPr
    for chunk in chunks[1:]:
        pbreak = ET.Element(W + 'p')
        pr = ET.SubElement(pbreak, W + 'r')
        ET.SubElement(pr, W + 'br', {W + 'type': 'page'})
        t2 = copy.deepcopy(proto)
        fill_word_table(t2, chunk, full_cols)
        idx = list(body).index(anchor) + 1
        body.insert(idx, pbreak)
        body.insert(idx + 1, t2)
        anchor = t2

    with zipfile.ZipFile(template) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == 'word/document.xml':
                data = serialize_restore_root(root, xml, 'document').encode('utf-8')
            zout.writestr(info, data)


def docx_set_cell_like_para(p, text):
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
        end = p.find(A + 'endParaRPr')
        r = ET.Element(A + 'r')
        if end is not None:
            rPr = copy.deepcopy(end)
            rPr.tag = A + 'rPr'
            r.append(rPr)
        t = ET.SubElement(r, A + 't')
        t.text = text
        p.insert(list(p).index(end) if end is not None else 0, r)


def fill_ppt_slide(root, chunk, full_cols, title):
    if title:
        tbl_texts = set()
        for tbl in root.iter(A + 'tbl'):
            for t in tbl.iter(A + 't'):
                tbl_texts.add(id(t))
        for t in root.iter(A + 't'):
            if id(t) not in tbl_texts and (t.text or '').strip():
                t.text = title
                break
    tbl = next(iter(root.iter(A + 'tbl')))
    n = len(chunk)
    surplus = full_cols - n
    if surplus:
        grid = tbl.find(A + 'tblGrid')
        cols = grid.findall(A + 'gridCol')
        removed = sum(int(c.get('w')) for c in cols[1 + n:])
        for c in cols[1 + n:]:
            grid.remove(c)
        for c in grid.findall(A + 'gridCol')[1:]:
            c.set('w', str(int(c.get('w')) + removed // n))
        for tr in tbl.findall(A + 'tr'):
            for tc in tr.findall(A + 'tc')[1 + n:]:
                tr.remove(tc)
    for tr_i, tr in enumerate(tbl.findall(A + 'tr')):
        tcs = tr.findall(A + 'tc')
        for j, proj in enumerate(chunk):
            text = proj['name'] if tr_i == 0 else proj['rows'][tr_i - 1]
            pptx_set_cell(tcs[1 + j], text)


def build_ppt(projects, variant, title, output):
    template = os.path.join(TEMPLATES, f'offering_{variant}_template.pptx')
    with zipfile.ZipFile(template) as zin:
        parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
        infos = list(zin.infolist())

    slide_part = next(n for n in parts if re.match(r'ppt/slides/slide\d+\.xml$', n))
    slide_num = int(re.search(r'slide(\d+)\.xml', slide_part).group(1))
    xml = parts[slide_part].decode('utf-8')
    chunks = chunked(projects, PPT_CHUNK)

    root = parse_keep_ns(xml)
    tbl = next(iter(root.iter(A + 'tbl')))
    full_cols = len(tbl.find(A + 'tblGrid').findall(A + 'gridCol')) - 1

    # first chunk fills the template slide
    fill_ppt_slide(root, chunks[0], full_cols, title)
    parts[slide_part] = serialize_restore_root(root, xml, 'sld').encode('utf-8')

    # extra chunks become duplicated slides
    pres = parts['ppt/presentation.xml'].decode('utf-8')
    pres_rels = parts['ppt/_rels/presentation.xml.rels'].decode('utf-8')
    ctypes = parts['[Content_Types].xml'].decode('utf-8')
    new_names = []
    for k, chunk in enumerate(chunks[1:], start=1):
        new_num = 100 + k                                    # avoid collisions
        new_part = f'ppt/slides/slide{new_num}.xml'
        root_k = parse_keep_ns(xml)
        fill_ppt_slide(root_k, chunk, full_cols, title)
        parts[new_part] = serialize_restore_root(root_k, xml, 'sld').encode('utf-8')
        # duplicated slides keep only the slideLayout rel — sharing the original
        # slide's notesSlide corrupts the package
        src_rels = parts[f'ppt/slides/_rels/slide{slide_num}.xml.rels'].decode('utf-8')
        src_rels = re.sub(r'<Relationship [^>]*Type="[^"]*/notesSlide"[^>]*/>', '', src_rels)
        parts[f'ppt/slides/_rels/slide{new_num}.xml.rels'] = src_rels.encode('utf-8')
        new_names.extend([new_part, f'ppt/slides/_rels/slide{new_num}.xml.rels'])
        rid = f'rIdX{k}'
        pres_rels = pres_rels.replace(
            '</Relationships>',
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
            f'officeDocument/2006/relationships/slide" Target="slides/slide{new_num}.xml"/>'
            '</Relationships>')
        max_id = max(int(m) for m in re.findall(r'<p:sldId id="(\d+)"', pres))
        pres = pres.replace('</p:sldIdLst>',
                            f'<p:sldId id="{max_id + 1}" r:id="{rid}"/></p:sldIdLst>')
        ctypes = ctypes.replace(
            '</Types>',
            f'<Override PartName="/{new_part}" ContentType="application/vnd.'
            f'openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    parts['ppt/presentation.xml'] = pres.encode('utf-8')
    parts['ppt/_rels/presentation.xml.rels'] = pres_rels.encode('utf-8')
    parts['[Content_Types].xml'] = ctypes.encode('utf-8')

    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, parts[info.filename])
        for name in new_names:
            zout.writestr(name, parts[name])


# ------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--csv', required=True)
    ap.add_argument('--variant', choices=['forsale', 'rental'], required=True)
    ap.add_argument('--format', choices=['word', 'ppt'], required=True)
    ap.add_argument('--title', default='')
    ap.add_argument('--output', required=True)
    args = ap.parse_args()

    projects = read_projects(args.csv, args.variant)
    if not projects:
        sys.exit('ERROR: no projects found in the CSV')
    if args.format == 'word':
        build_word(projects, args.variant, args.title, args.output)
    else:
        build_ppt(projects, args.variant, args.title, args.output)
    print(f'Wrote {args.output} ({len(projects)} projects)')


if __name__ == '__main__':
    main()
