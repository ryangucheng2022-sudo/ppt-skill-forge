#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Enterprise PPT 内容维度机检（Python 版）。

与 macOS 版 verify_bps_ppt_mac.mjs 的 --check-content 同口径移植：
从 slide XML 统计逐页内容指标（形状数、带文本形状数、正文字数、图片数、
原生表格／图表有无、最大字号），按页型启发式剔除封面/目录/章节/收尾页后，
对正文页套用阈值档案判定内容密度风险。

定位：**咨询性**。默认只输出 CONTENT ADVISORY，不影响退出码；
--strict-content 时咨询项判失败（exit 1）。机检只提示风险，
不能替代内容判断，提示项须回到逐页信息规划核实。

用法：
  python check_content.py <deck.pptx> [--profile A|B] [--json report.json]
      [--strict-content]

退出码：0 通过（含仅咨询项）；1 --strict-content 下有咨询项；2 IO/用法错误。
"""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
P = "{%s}" % NS["p"]
A = "{%s}" % NS["a"]
C = "{%s}" % NS["c"]

# 与 mac 版 CONTENT_PROFILES 一致
PROFILES = {
    # A 管理汇报风：宽口径，按技能演示稿校准
    "A": {"min_shapes": 6, "min_text_chars": 60, "require_picture": False},
    # B 宣讲交流风：按 references/townhall_style_branch.md §四 正文页 ≥40 形状
    "B": {"min_shapes": 40, "min_text_chars": 350, "require_picture": True},
}
CONTENT_STRUCT_SHAPES = 15      # ≥15 形状视为结构化图示（证明对象）
SHAPE_TAGS = {f"{P}sp", f"{P}pic", f"{P}graphicFrame",
              f"{P}grpSp", f"{P}cxnSp", f"{P}contentPart"}


def _slide_paths(zf):
    names = [n for n in zf.namelist()
             if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    return sorted(names, key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[-1]).group(1)))


def _text_of(node):
    return "".join(t.text or "" for t in node.iter(f"{A}t"))


def inspect_slide(xml_bytes):
    root = etree.fromstring(xml_bytes)
    sp_tree = root.find(f".//{P}spTree")
    if sp_tree is None:
        sp_tree = root
    shapes = [el for el in sp_tree.iter() if el.tag in SHAPE_TAGS]
    graphic_frames = [el for el in shapes if el.tag == f"{P}graphicFrame"]
    native_charts = sum(
        1 for gf in graphic_frames
        if gf.find(f".//{C}chart") is not None
        or any("chart" in (gd.get("uri") or "") for gd in gf.iter(f"{A}graphicData")))
    text_shapes = sum(
        1 for el in shapes
        if el.tag in (f"{P}sp", f"{P}graphicFrame") and _text_of(el).strip())
    raw_text = _text_of(sp_tree)
    sizes = [int(rp.get("sz")) / 100.0
             for rp in sp_tree.iter(f"{A}rPr") if rp.get("sz")]
    return {
        "shapes": len(shapes),
        "text_shapes": text_shapes,
        "text_chars": len(re.sub(r"\s+", "", raw_text)),
        "text_sample": re.sub(r"\s+", " ", raw_text).strip(),
        "pictures": sum(1 for el in shapes if el.tag == f"{P}pic"),
        "native_tables": sum(1 for _ in sp_tree.iter(f"{A}tbl")),
        "native_charts": native_charts,
        "max_font_pt": max(sizes) if sizes else 0.0,
    }


def classify_page(m, position, total):
    """页型启发式：封面=首页；收尾=末页且有联系方式；目录=文字以“目录”开头；
    章节=文字≤80 字且最大字号≥40pt；其余为正文页。"""
    if position == 1:
        return "cover"
    text = m["text_sample"]
    has_contact = bool(re.search(r"[\w.+-]+@[\w-]+\.[\w.]+", text)) or (
        re.search(r"www\.", text, re.I) and re.search(r"(电话|地址|热线)", text))
    if position == total and has_contact:
        return "closing"
    if text.startswith("目录"):
        return "contents"
    if m["text_chars"] <= 80 and m["max_font_pt"] >= 40:
        return "chapter"
    return "body"


def layout_signature(m):
    pic = "p0" if m["pictures"] == 0 else "p1" if m["pictures"] <= 2 else "p2"
    shp = ("s0" if m["shapes"] < 15 else "s1" if m["shapes"] < 40
           else "s2" if m["shapes"] < 80 else "s3")
    txt = "t0" if m["text_chars"] < 200 else "t1" if m["text_chars"] < 500 else "t2"
    return ("T" if m["native_tables"] > 0 else "-") + \
           ("C" if m["native_charts"] > 0 else "-") + pic + shp + txt


def median(values):
    if not values:
        return 0
    s = sorted(values)
    mid = len(s) // 2
    return s[mid] if len(s) % 2 else round((s[mid - 1] + s[mid]) / 2, 1)


def analyse(pptx_path, profile="A"):
    prof = PROFILES[profile]
    slides, advisories = [], []
    with zipfile.ZipFile(pptx_path) as zf:
        paths = _slide_paths(zf)
        total = len(paths)
        for i, name in enumerate(paths, 1):
            m = inspect_slide(zf.read(name))
            m["slide"] = i
            m["page_type"] = classify_page(m, i, total)
            m["layout_sig"] = layout_signature(m)
            m["proof"] = m["shapes"] >= CONTENT_STRUCT_SHAPES
            slides.append(m)
    body = [s for s in slides if s["page_type"] == "body"]
    for s in body:
        if s["shapes"] < prof["min_shapes"]:
            advisories.append(
                f"p{s['slide']} shapes {s['shapes']} below profile {profile} "
                f"body-page floor {prof['min_shapes']} (content density risk)")
        if s["text_chars"] < prof["min_text_chars"]:
            advisories.append(
                f"p{s['slide']} text {s['text_chars']} chars below profile "
                f"{profile} floor {prof['min_text_chars']} (content may be too thin)")
        if prof["require_picture"] and s["pictures"] == 0:
            advisories.append(
                f"p{s['slide']} no picture on body page (profile {profile} "
                f"requires image-text pairing)")
    sigs = {s["layout_sig"] for s in body}
    if len(body) >= 3 and len(sigs) == 1:
        advisories.append(
            f"all {len(body)} body pages share one layout signature "
            f"{next(iter(sigs))} (page-type monotony risk)")
    return {
        "deck": str(pptx_path),
        "profile": profile,
        "total_slides": total,
        "body_pages": len(body),
        "body_shapes_min": min((s["shapes"] for s in body), default=0),
        "body_shapes_median": median([s["shapes"] for s in body]),
        "body_chars_min": min((s["text_chars"] for s in body), default=0),
        "body_chars_median": median([s["text_chars"] for s in body]),
        "layout_signatures": sorted(sigs),
        "layout_diversity": len(sigs),
        "advisories": advisories,
        "slides": slides,
    }


def main():
    ap = argparse.ArgumentParser(description="Enterprise PPT content-dimension check")
    ap.add_argument("pptx")
    ap.add_argument("--profile", default="A", choices=["A", "B"])
    ap.add_argument("--json", default="")
    ap.add_argument("--strict-content", action="store_true")
    args = ap.parse_args()

    p = Path(args.pptx)
    if not p.exists():
        print(f"ERROR: file not found: {p}", file=sys.stderr)
        return 2
    try:
        r = analyse(p, args.profile)
    except Exception as e:                                  # noqa: BLE001
        print(f"ERROR: cannot parse deck: {e}", file=sys.stderr)
        return 2

    print(f"CONTENT CHECK (profile {r['profile']}"
          f"{', strict' if args.strict_content else ', advisory only'}): "
          f"{r['total_slides']} slides, {r['body_pages']} body pages")
    print(f"  body shapes  min {r['body_shapes_min']} / median "
          f"{r['body_shapes_median']}")
    print(f"  body chars   min {r['body_chars_min']} / median "
          f"{r['body_chars_median']}")
    print(f"  layout diversity {r['layout_diversity']} "
          f"({', '.join(r['layout_signatures']) or '-'})")
    for s in r["slides"]:
        print(f"  p{s['slide']:<2d} {s['page_type']:<8s} shapes={s['shapes']:<3d} "
              f"text_shapes={s['text_shapes']:<3d} chars={s['text_chars']:<4d} "
              f"pics={s['pictures']} tbl={s['native_tables']} "
              f"chart={s['native_charts']} max_pt={s['max_font_pt']:g} "
              f"proof={'yes' if s['proof'] else 'no'}")
    if r["advisories"]:
        for a in r["advisories"]:
            print(f"  CONTENT ADVISORY: {a}")
    else:
        print("  CONTENT ADVISORY: none")

    if args.json:
        Path(args.json).write_text(
            json.dumps(r, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Report {args.json}")

    if args.strict_content and r["advisories"]:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
