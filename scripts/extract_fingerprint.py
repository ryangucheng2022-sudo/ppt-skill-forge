#!/usr/bin/env python3
"""ppt-forge P0 prototype — extract a style fingerprint from PPTX assets.

Reads a set of .pptx files and produces:
  - inventory: per-file slide count, slide size, role guess
  - theme: color scheme + font scheme (from the first file's theme parts)
  - font_sizes: histogram of explicit run font sizes (pt)
  - fonts: latin + east-asian typefaces by run count
  - fill_colors: solid-fill RGB distribution
  - furniture: recurring small shapes in the top-left corner zone
    (position-clustered across slides — the "corner mark" pattern)

Usage: extract_fingerprint.py OUT.json FILE.pptx [FILE.pptx ...]
Pure python-pptx + stdlib. Charts' internal text is not traversed (v0 limit).
"""
import json, sys, zipfile
from collections import Counter
from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from lxml import etree

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
EMU_IN = 914400

def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # group
            yield from iter_shapes(sh.shapes)

def ea_typeface(run):
    rPr = run._r.find('a:rPr', NS)
    if rPr is None: return None
    ea = rPr.find('a:ea', NS)
    return ea.get('typeface') if ea is not None else None

def theme_info(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if 'theme/theme' in n]
        if not names: return out
        root = etree.fromstring(z.read(names[0]))
    clr = root.find('.//a:clrScheme', NS)
    if clr is not None:
        colors = {}
        for child in clr:
            tag = etree.QName(child).localname
            srgb = child.find('a:srgbClr', NS)
            sysc = child.find('a:sysClr', NS)
            colors[tag] = srgb.get('val') if srgb is not None else (sysc.get('lastClr') if sysc is not None else None)
        out['colors'] = colors
    fs = root.find('.//a:fontScheme', NS)
    if fs is not None:
        fonts = {}
        for part in ('majorFont', 'minorFont'):
            el = fs.find(f'a:{part}', NS)
            if el is not None:
                latin = el.find('a:latin', NS); ea = el.find('a:ea', NS)
                fonts[part] = {'latin': latin.get('typeface') if latin is not None else None,
                               'ea': ea.get('typeface') if ea is not None else None}
        out['fonts'] = fonts
    return out

def fingerprint(paths):
    sizes, fonts_latin, fonts_ea, fills, theme_fills, text_colors = Counter(), Counter(), Counter(), Counter(), Counter(), Counter()
    corner = Counter()
    inventory = []
    total_slides = 0
    for path in paths:
        prs = Presentation(path)
        slides = len(prs.slides.__iter__.__self__._sldIdLst)
        total_slides += slides
        inventory.append({'file': path.split('/')[-1], 'slides': slides,
                          'slide_in': [round(prs.slide_width / EMU_IN, 3), round(prs.slide_height / EMU_IN, 3)]})
        for slide in prs.slides:
            for sh in iter_shapes(slide.shapes):
                # corner furniture: small shapes in the top-left 1.2in x 1.2in zone
                try:
                    if sh.width and sh.height and sh.width < 1.2 * EMU_IN and sh.height < 1.2 * EMU_IN \
                       and sh.left is not None and sh.top is not None and sh.left < 0.6 * EMU_IN and sh.top < 0.8 * EMU_IN:
                        corner[(round(sh.left / EMU_IN, 2), round(sh.top / EMU_IN, 2),
                                round(sh.width / EMU_IN, 2), round(sh.height / EMU_IN, 2))] += 1
                except Exception:
                    pass
                try:
                    if sh.fill and sh.fill.type == MSO_FILL.SOLID:
                        c = sh.fill.fore_color
                        if c.type is not None and str(c.type) == 'MSO_THEME_COLOR.NOT_THEME_COLOR (0)' or (c.rgb is not None):
                            fills[str(c.rgb)] += 1
                        elif c.theme_color is not None:
                            theme_fills[str(c.theme_color)] += 1
                except Exception:
                    pass
                frames = []
                if sh.has_text_frame: frames.append(sh.text_frame)
                if getattr(sh, 'has_table', False) and sh.has_table:
                    for row in sh.table.rows:
                        for cell in row.cells: frames.append(cell.text_frame)
                for tf in frames:
                    for para in tf.paragraphs:
                        for run in para.runs:
                            if run.font.size is not None:
                                sizes[round(run.font.size.pt, 2)] += 1
                            if run.font.name: fonts_latin[run.font.name] += 1
                            try:
                                if run.font.color and run.font.color.rgb is not None:
                                    text_colors[str(run.font.color.rgb)] += 1
                            except Exception:
                                pass
                            ea = ea_typeface(run)
                            if ea: fonts_ea[ea] += 1
    return {'inventory': inventory, 'total_slides': total_slides,
            'theme': theme_info(paths[0]) if paths else {},
            'font_sizes': dict(sorted(sizes.items())), 'fonts_latin': fonts_latin.most_common(10),
            'fonts_ea': fonts_ea.most_common(10), 'fill_colors': fills.most_common(15),
            'theme_fill_uses': theme_fills.most_common(5), 'text_colors': text_colors.most_common(10),
            'corner_furniture': [(list(k), v) for k, v in corner.most_common(5)]}

if __name__ == '__main__':
    out_path, files = sys.argv[1], sys.argv[2:]
    fp = fingerprint(files)
    json.dump(fp, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    total_runs = sum(fp['font_sizes'].values())
    print(f"files={len(files)} slides={fp['total_slides']} sized_runs={total_runs} distinct_sizes={len(fp['font_sizes'])}")
