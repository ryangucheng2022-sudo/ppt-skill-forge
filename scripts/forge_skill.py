#!/usr/bin/env python3
"""forge_skill.py — generate an enterprise PPT skill skeleton.

Inputs: a fingerprint (extract_fingerprint.py), a derived whitelist
(derive_whitelist.py) plus its decision list, the company's asset files,
and a company name. Output: a self-contained skill package —

  OUT/
    SKILL.md                       (from templates/generated-skill.md)
    scripts/font_sizes.json        (the derived candidate)
    scripts/check_geometry.py      (generic checker, name-neutralized)
    scripts/check_content.py       (generic checker, name-neutralized)
    references/style-fingerprint.md
    references/derivation-decisions.md
    assets/*.pptx                  (copies of the source assets)

The checkers' logic is copied byte-for-byte except program-name strings;
all company specificity lives in the config, per the forge method.

Usage: forge_skill.py --company NAME --fingerprint F.json --whitelist W.json
       --decisions D.md --assets-dir DIR --out DIR
"""
import argparse, json, os, shutil, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

def render_fingerprint_md(fp, company):
    inv = fp.get('inventory', [])
    lines = [f'# {company} — style fingerprint (observed usage)', '',
             f"Files: {len(inv)} · Slides: {fp.get('total_slides', '?')}", '']
    lines.append('## Inventory')
    for i in inv:
        lines.append(f"- {i['file']}: {i['slides']} slides, {i['slide_in'][0]}×{i['slide_in'][1]}in")
    lines += ['', '## Fonts by run share', '',
              'Latin: ' + ', '.join(f'{n} ×{c}' for n, c in fp.get('fonts_latin', [])[:5]),
              'East-Asian: ' + ', '.join(f'{n} ×{c}' for n, c in fp.get('fonts_ea', [])[:5]),
              '', '## Solid fills by use', '',
              ', '.join(f'{rgb} ×{c}' for rgb, c in fp.get('fill_colors', [])[:10]),
              '', '## Text colors by run', '',
              ', '.join(f'{rgb} ×{c}' for rgb, c in fp.get('text_colors', [])[:10]),
              '', '## Corner furniture (position clusters, inches)', '']
    for (l, t, w, h), c in fp.get('corner_furniture', [])[:5]:
        lines.append(f'- left {l}, top {t}, {w}×{h} — on {c} slides')
    th = fp.get('theme', {})
    if th:
        lines += ['', '## Declared theme (for contrast — usage outranks it)', '',
                  f"Colors: {th.get('colors', {})}", f"Fonts: {th.get('fonts', {})}"]
    return '\n'.join(lines) + '\n'

def main():
    ap = argparse.ArgumentParser(description='Generate an enterprise PPT skill skeleton')
    ap.add_argument('--company', required=True)
    ap.add_argument('--fingerprint', required=True)
    ap.add_argument('--whitelist', required=True)
    ap.add_argument('--decisions')
    ap.add_argument('--assets-dir')
    ap.add_argument('--out', required=True)
    a = ap.parse_args()

    fp = json.load(open(a.fingerprint, encoding='utf-8'))
    os.makedirs(os.path.join(a.out, 'scripts'), exist_ok=True)
    os.makedirs(os.path.join(a.out, 'references'), exist_ok=True)
    written = []

    # SKILL.md from template
    tpl = open(os.path.join(ROOT, 'templates', 'generated-skill.md'), encoding='utf-8').read()
    inv = fp.get('inventory', [])
    canvas = f"{inv[0]['slide_in'][0]}×{inv[0]['slide_in'][1]}in" if inv else 'unknown'
    fonts = ', '.join(n for n, _ in fp.get('fonts_latin', [])[:3]) or 'unknown'
    colors = ', '.join(rgb for rgb, _ in fp.get('fill_colors', [])[:6]) or 'unknown'
    cf = fp.get('corner_furniture', [])
    corner = (f'cluster at left {cf[0][0][0]}, top {cf[0][0][1]}, {cf[0][0][2]}×{cf[0][0][3]}in '
              f'on {cf[0][1]} slides — verify against the official asset before treating as spec') if cf else 'none detected'
    skill = (tpl.replace('{{COMPANY}}', a.company).replace('{{CANVAS}}', canvas)
                .replace('{{FONTS}}', fonts).replace('{{COLORS}}', colors).replace('{{CORNER}}', corner))
    open(os.path.join(a.out, 'SKILL.md'), 'w', encoding='utf-8').write(skill)
    written.append('SKILL.md')

    # whitelist + checkers
    shutil.copy(a.whitelist, os.path.join(a.out, 'scripts', 'font_sizes.json'))
    written.append('scripts/font_sizes.json')
    for name in ('check_geometry.py', 'check_content.py'):
        src = open(os.path.join(ROOT, 'assets', 'checkers', name), encoding='utf-8').read()
        src = src.replace('Enterprise PPT geometry checker', f'{a.company} PPT geometry checker')
        src = src.replace('Enterprise PPT content-dimension check', f'{a.company} PPT content-dimension check')
        src = src.replace('Enterprise PPT 版面几何校验器', f'{a.company} PPT 版面几何校验器')
        src = src.replace('Enterprise PPT 内容维度机检', f'{a.company} PPT 内容维度机检')
        open(os.path.join(a.out, 'scripts', name), 'w', encoding='utf-8').write(src)
        written.append(f'scripts/{name}')

    # references
    open(os.path.join(a.out, 'references', 'style-fingerprint.md'), 'w', encoding='utf-8').write(
        render_fingerprint_md(fp, a.company))
    written.append('references/style-fingerprint.md')
    if a.decisions and os.path.isfile(a.decisions):
        shutil.copy(a.decisions, os.path.join(a.out, 'references', 'derivation-decisions.md'))
        written.append('references/derivation-decisions.md')

    # assets
    if a.assets_dir:
        os.makedirs(os.path.join(a.out, 'assets'), exist_ok=True)
        for f in sorted(os.listdir(a.assets_dir)):
            if f.endswith('.pptx'):
                shutil.copy(os.path.join(a.assets_dir, f), os.path.join(a.out, 'assets', f))
                written.append(f'assets/{f}')
    print(json.dumps({'company': a.company, 'out': a.out, 'written': written}, ensure_ascii=False, indent=1))

if __name__ == '__main__':
    main()
