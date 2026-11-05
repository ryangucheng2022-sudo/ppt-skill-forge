#!/usr/bin/env python3
"""derive_whitelist.py — derive a font-size whitelist candidate from a fingerprint.

Three-layer rule (validated against a production enterprise corpus):
  1. Quarter-grid sizes (multiples of 0.25pt) at or above the floor, seen at
     least --min-count times, enter the whitelist.
  2. Non-quarter sizes are auto-fit artifacts (org charts and auto-shrink
     runs hug the underside of a declared size). They never enter the
     whitelist; they are reported with their nearest declared neighbor.
  3. Sub-floor quarter sizes and singleton sizes become DECISION items, not
     silent inclusions: a human confirms them (asset furniture often needs a
     relaxed asset-mode floor; singletons may be one-off mistakes).

The output is a CANDIDATE in the same schema as the production
font_sizes.json, plus a decision list. A derived whitelist becomes a
standard only after the decision list is signed off — and after checking
corpus coverage by page type (display sizes for covers/chapters never
appear in a corpus that has no covers/chapters).

Usage: derive_whitelist.py FINGERPRINT.json --out font_sizes.json
       [--decisions decisions.md] [--min-count 2] [--floor 8.0]
"""
import argparse, json
from collections import Counter

def is_quarter(x):
    return abs(x * 4 - round(x * 4)) < 0.01

def main():
    ap = argparse.ArgumentParser(description='Derive a font-size whitelist candidate from a fingerprint')
    ap.add_argument('fingerprint')
    ap.add_argument('--out', required=True)
    ap.add_argument('--decisions')
    ap.add_argument('--min-count', type=int, default=2)
    ap.add_argument('--floor', type=float, default=8.0)
    a = ap.parse_args()

    fp = json.load(open(a.fingerprint, encoding='utf-8'))
    hist = {float(k): v for k, v in fp['font_sizes'].items()}
    allowed, singletons, subfloor, fitted = [], [], [], []
    for size, count in sorted(hist.items()):
        if not is_quarter(size):
            fitted.append((size, count))
        elif size < a.floor:
            subfloor.append((size, count))
        elif count >= a.min_count:
            allowed.append(size)
        else:
            singletons.append((size, count))

    def nearest(size):
        return min(allowed, key=lambda s: abs(s - size)) if allowed else None

    cfg = {
        'allowed_pt': allowed,
        'comment': ('DERIVED CANDIDATE by ppt-skill-forge from observed usage; not a standard until the '
                    'decision list is signed off. Quarter-grid sizes seen >= %d times at/above the floor. '
                    'Non-quarter sizes are treated as auto-fit values (warning only). Corpus coverage by '
                    'page type must be checked: display sizes absent from the corpus cannot be derived.' % a.min_count),
        'floor_error_pt': a.floor,
        'small_report_pt': a.floor + 1.0,
        'template_table_floor_pt': a.floor,
    }
    json.dump(cfg, open(a.out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    total = sum(hist.values())
    covered = sum(c for s, c in hist.items() if s in allowed)
    lines = ['# Whitelist derivation — decision list', '',
             f'Corpus: {fp.get("total_slides", "?")} slides, {total} sized runs, {len(hist)} distinct sizes.',
             f'Derived whitelist: {len(allowed)} sizes, directly covering {covered} runs ({covered/total:.1%}); '
             f'auto-fit values account for {sum(c for _, c in fitted)} runs and are tolerated by design.', '',
             '## Items needing a human decision', '']
    if subfloor:
        lines.append(f'- **Sub-floor sizes** (below {a.floor}pt): ' +
                     ', '.join(f'{s}pt ×{c}' for s, c in subfloor) +
                     '. If these are template furniture (footnotes, page numbers), set an asset-mode floor at or below '
                     f'{min(s for s, _ in subfloor)}pt for baseline checks instead of whitelisting them.')
    if singletons:
        lines.append('- **Singleton sizes** (seen once, at/above floor): ' +
                     ', '.join(f'{s}pt' for s, _ in singletons) +
                     '. One-off usage — confirm whether each is a deliberate display size or a mistake before whitelisting.')
    near = {}
    for s, c in fitted:
        n = nearest(s)
        if n is not None and 0 < n - s <= 0.6:
            near.setdefault(n, []).append((s, c))
    if near:
        lines.append('- **Auto-fit clusters** (no action expected): ' +
                     '; '.join(f'{n}pt ← ' + ', '.join(f'{s}×{c}' for s, c in v) for n, v in sorted(near.items())) +
                     '. Values hugging the underside of a declared size are auto-shrink artifacts.')
    if len(lines) <= 6:
        lines.append('- None. The corpus maps cleanly onto the derived whitelist.')
    lines += ['', '## Corpus coverage warning', '',
              'Sizes never observed cannot be derived. Verify the corpus contains covers, chapter '
              'dividers and closing pages; display sizes for those page types must be added from the '
              'style guide or the templates themselves before this candidate becomes the standard.']
    if a.decisions:
        open(a.decisions, 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
    print(json.dumps({'allowed': len(allowed), 'direct_coverage': round(covered / total, 3),
                      'fitted_runs': sum(c for _, c in fitted), 'subfloor': subfloor,
                      'singletons': [s for s, _ in singletons]}, ensure_ascii=False))

if __name__ == '__main__':
    main()
