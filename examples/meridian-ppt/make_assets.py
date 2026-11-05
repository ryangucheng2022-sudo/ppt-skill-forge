#!/usr/bin/env python3
"""Generate the fictional Meridian Office Furniture asset corpus.

Deterministic, python-pptx only. The corpus deliberately plants the
phenomena the forge must handle:

- a corner-furniture spec (navy square at 0.35/0.32in, 0.28in) repeated on
  body pages;
- a size system: 40 cover titles (x2 files), 20/16/12/10.5/10 regulars;
- singleton display sizes: 54 (chapter numeral), 24 (chapter title),
  28 (closing) — each used once, so derivation must hold them as decisions;
- sub-floor footnotes at 7pt;
- auto-fit values 11.41 / 9.85 in the org chart (non-quarter sizes);
- a near-brand color variant: one title in 203A66 instead of brand 1F3864;
- the theme is left at the Office default — the brand lives in direct
  formatting, as in most real companies.

Usage: make_assets.py OUT_DIR   (writes meridian-template.pptx, meridian-library.pptx)
"""
import os, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

NAVY = RGBColor(0x1F, 0x38, 0x64)
NAVY_VARIANT = RGBColor(0x20, 0x3A, 0x66)
SLATE = RGBColor(0x44, 0x54, 0x6A)
ACCENT = RGBColor(0xED, 0x7D, 0x31)
DARK = RGBColor(0x40, 0x40, 0x40)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD = RGBColor(0xF2, 0xF2, 0xF2)
FOOT = RGBColor(0x80, 0x80, 0x80)

def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def box(slide, x, y, w, h, fill=None):
    sh = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.line.fill.background()
    if fill is not None:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    else:
        sh.fill.background()
    return sh

def text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT):
    """runs: list of (txt, size, color, bold, font)"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    for i, (txt, size, color, bold, font) in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = txt
        r.font.size = Pt(size); r.font.color.rgb = color; r.font.bold = bold; r.font.name = font
    return tb

def corner(slide):
    box(slide, 0.35, 0.32, 0.28, 0.28, NAVY)

def title_bar(slide, txt, color=NAVY):
    corner(slide)
    text(slide, 0.8, 0.22, 9.0, 0.6, [(txt, 16, color, True, 'Calibri')])

def footnote(slide, txt='Source: Meridian internal review, FY2026. Fictional sample data.'):
    text(slide, 0.8, 7.08, 11.0, 0.3, [(txt, 7, FOOT, False, 'Calibri')])

def new_deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    return prs

# ---------------- template ----------------
prs = new_deck()
s = blank(prs)  # cover
box(s, 0, 0, 13.333, 7.5, NAVY)
box(s, 0.9, 2.02, 1.6, 0.09, ACCENT)
text(s, 0.9, 2.3, 11.5, 1.2, [('Meridian Office Furniture', 40, WHITE, True, 'Calibri Light')])
text(s, 0.9, 3.5, 11.5, 0.7, [('Annual Business Review 2026', 18, WHITE, False, 'Calibri')])
s = blank(prs)  # TOC
text(s, 0.9, 0.7, 8.0, 0.7, [('Contents', 20, NAVY, True, 'Calibri')])
for i, item in enumerate(['1  Market and customers', '2  Product lines', '3  Operations', '4  Financials']):
    text(s, 0.9, 1.7 + i * 0.62, 9.0, 0.5, [(item, 12, DARK, False, 'Calibri')])
s = blank(prs)  # chapter (singletons 54 / 24)
text(s, 0.9, 2.2, 3.0, 1.6, [('01', 54, NAVY, True, 'Calibri Light')])
text(s, 0.9, 3.9, 10.0, 0.9, [('Market and customers', 24, SLATE, False, 'Calibri Light')])
s = blank(prs)  # body
title_bar(s, 'Meridian at a glance')
for i in range(3):
    box(s, 0.8 + i * 4.1, 1.4, 3.8, 2.2, CARD)
    text(s, 1.05 + i * 4.1, 1.65, 3.3, 1.8,
         [(f'Pillar {i+1}', 12, NAVY, True, 'Calibri'),
          ('Workspace systems, seating and storage for project-based offices.', 12, DARK, False, 'Calibri')])
text(s, 0.8, 4.1, 11.7, 1.6,
     [('How to read this deck', 10.5, SLATE, True, 'Calibri'),
      ('Figures are fictional and serve the template only. Sizes and colors follow the Meridian deck standard.', 10.5, DARK, False, 'Calibri')])
footnote(s)
s = blank(prs)  # closing (singleton 28)
box(s, 0, 0, 13.333, 7.5, NAVY)
text(s, 0.9, 2.9, 11.5, 1.0, [('Thank you', 28, WHITE, True, 'Calibri Light')], PP_ALIGN.CENTER)
text(s, 0.9, 4.1, 11.5, 0.5, [('hello@meridian-office.example', 12, WHITE, False, 'Calibri')], PP_ALIGN.CENTER)
prs.save(os.path.join(sys.argv[1], 'meridian-template.pptx'))

# ---------------- library ----------------
prs = new_deck()
s = blank(prs)  # second cover-weight title (makes 40 a regular)
box(s, 0, 0, 13.333, 7.5, NAVY)
text(s, 0.9, 2.6, 11.5, 1.2, [('Meridian Page Library', 40, WHITE, True, 'Calibri Light')])
text(s, 0.9, 3.9, 11.5, 0.6, [('Reusable pages for business reviews', 18, WHITE, False, 'Calibri')])
s = blank(prs)  # table page
title_bar(s, 'Product lines by region')
tbl_shape = s.shapes.add_table(4, 4, Inches(0.8), Inches(1.3), Inches(11.7), Inches(3.4))
tbl = tbl_shape.table
data = [['Region', 'Desks', 'Seating', 'Storage'],
        ['North', '1,240', '860', '410'], ['South', '980', '720', '365'], ['West', '1,105', '790', '388']]
for r, row in enumerate(data):
    for c, val in enumerate(row):
        cell = tbl.cell(r, c)
        cell.text = ''
        p = cell.text_frame.paragraphs[0]; run = p.add_run(); run.text = val
        run.font.size = Pt(12 if r == 0 else 10)
        run.font.name = 'Calibri'; run.font.bold = (r == 0)
        run.font.color.rgb = NAVY if r == 0 else DARK
footnote(s)
s = blank(prs)  # org chart (fitted sizes)
title_bar(s, 'Organization (auto-fit text)')
levels = [(5.37, 1.5, 'General Manager', 12), (2.2, 3.1, 'Product', 11.41), (5.37, 3.1, 'Sales', 11.41),
          (8.55, 3.1, 'Operations', 11.41), (3.6, 4.9, 'Design Studio', 9.85), (6.9, 4.9, 'Supply Chain', 9.85)]
for x, y, label, size in levels:
    box(s, x, y, 2.6, 0.85, CARD)
    text(s, x + 0.1, y + 0.22, 2.4, 0.45, [(label, size, DARK, False, 'Calibri')], PP_ALIGN.CENTER)
text(s, 0.8, 6.2, 11.7, 0.4, [('Box text auto-shrinks to fit; the recorded sizes are fit results, not standards.', 12, SLATE, False, 'Calibri')])
s = blank(prs)  # two-column body, title in the near-brand variant
title_bar(s, 'Regional notes', color=NAVY_VARIANT)
for col in range(2):
    x = 0.8 + col * 6.0
    text(s, x, 1.3, 5.6, 4.6,
         [('Highlights', 12, NAVY, True, 'Calibri'),
          ('Showroom refresh completed in two pilot cities; project pipeline reviewed monthly.', 12, DARK, False, 'Calibri'),
          ('Watch-outs', 10.5, SLATE, True, 'Calibri'),
          ('Lead times on imported fittings remain the main delivery risk.', 10.5, DARK, False, 'Calibri')])
footnote(s)
prs.save(os.path.join(sys.argv[1], 'meridian-library.pptx'))
print('meridian assets written to', sys.argv[1])
