#!/usr/bin/env python3
"""Aggregate grouped unit mix data into the transposed Unit Mix Overview render JSON.

Input JSON (grouping is Claude's job per SKILL.md; this script only does the math):
{
  "title": "Active and Stabilized Developments - Unit Mix Overview",
  "intro": "The following charts contain ...",
  "sections": [
    {
      "label": "UNIT MIX SUMMARY: ACTIVELY LEASING CONCRETE RENTALS*",
      "blurb": "The table below provides an overview of ...",
      "footnote": "*Certain plan types have been grouped together for simplicity: ...",
      "order": ["STUDIO", "1 BED", "..."],          // optional; defaults to insertion order
      "expected_total": 1041,                        // optional cross-check
      "groups": {"STUDIO": [[10, 394, 430], [67, 370, 411]], ...}  // [count, min_sf, max_sf] per plan row
    }
  ]
}

Stats per the locked MLA Advisory methodology (same as unit-mix-summary):
- INVENTORY: count share, round half up, "<1%" below 1 percent
- MIN / MAX: pooled extremes, unrounded
- AVERAGE: count-weighted average of plan midpoints, rounded to nearest 5
- MEDIAN: unweighted median of the pooled min/max values, rounded to nearest 5

Usage: python aggregate_overview.py --data grouped.json --output render.json
"""
import argparse
import json
import statistics


def r5(x):
    return int(round(x / 5.0) * 5)


def build_columns(groups, order, expected_total=None):
    total = sum(c for rows in groups.values() for c, _, _ in rows)
    if expected_total is not None and total != expected_total:
        raise SystemExit(f'ERROR: grouped total {total} != expected {expected_total} - '
                         'a plan row is missing or double-counted')
    cols = []
    for t in order or list(groups):
        rows = groups[t]
        n = sum(c for c, _, _ in rows)
        pct = n / total * 100
        wavg = sum(c * (mn + mx) / 2 for c, mn, mx in rows) / n
        pool = [v for _, mn, mx in rows for v in (mn, mx)]
        cols.append({
            'type': t,
            'pct': '<1%' if pct < 1 else f'{int(pct + 0.5)}%',
            'min': f'{min(mn for _, mn, _ in rows):,}',
            'max': f'{max(mx for _, _, mx in rows):,}',
            'avg': f'{r5(wavg):,}',
            'med': f'{r5(statistics.median(pool)):,}',
            'n': n,
        })
    return total, cols


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    with open(args.data, encoding='utf-8') as f:
        data = json.load(f)

    out = {'title': data['title'], 'intro': data['intro'], 'sections': []}
    for sec in data['sections']:
        total, cols = build_columns(sec['groups'], sec.get('order'), sec.get('expected_total'))
        print(f"{sec['label']}: {total:,} units")
        for c in cols:
            print('  %-14s n=%-4d %4s min=%6s max=%6s avg=%6s med=%6s' %
                  (c['type'], c['n'], c['pct'], c['min'], c['max'], c['avg'], c['med']))
        out['sections'].append({'label': sec['label'], 'blurb': sec['blurb'],
                                'footnote': sec['footnote'], 'columns': cols})

    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=2)
    print('wrote', args.output)


if __name__ == '__main__':
    main()
