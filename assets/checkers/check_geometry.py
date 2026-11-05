#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Enterprise PPT 版面几何校验器（检查层，不改生成行为）。

借鉴来源：PPT Master（hugohe3/ppt-master，MIT）svg_quality_checker 的
检查项与闸门思想（bounds 量化溢出拦截、字号 role 声明制、质量报告前置），
基于 python-pptx 形状树实现。

检查项：
  1. bounds    形状（含组内递归、旋转后包围盒）超出页面边界（容差 --tol，
               默认 0.02in）→ error，逐边报溢出量。
  2. intrusion 非文字形状与文字承载形状「部分交叠」（既非完全包含、亦非被
               完全包含）且交叠面积 > 文字形状面积 5% → error。
               前置条件（v1.4.0 修正）：只有绘制在文字形状**之上**（z 序更大）
               的非文字形状才判定侵入——位于文字下方的卡片/面板/色块是背景板，
               不构成遮挡。修复前把官方资产的「卡片托文字」判成 44%～94% 侵入。
               豁免：整页背景/面板容器（完全包含文字形状）、连接线与细线
               （任一边 < 0.04in）、表格与图表（容器类）、组形状本体
               （只查其子形状）、带文字的形状之间不互查。
  3. font      全册 run 字号对照技能内白名单（font_sizes.json，声明制）：
               - 干净字号（0.25pt 整数倍）不在白名单 → error；
               - 字号 < 8pt（floor_error_pt）：干净值 → error；非干净值
                 （自动适配计算值）→ warning 单列（生成器各调用点 minimum
                 为 7.5–8.5，属既有拟合行为）；template_table 页（--spec
                 可判时）下限同为 8pt（模板页豁免范围）；
               - 非 0.25 整数倍字号 = orgchart 自动适配的连续拟合值
                 → warning（fitted），单列清单；
               - < 9pt 字号一律单列 small_sizes 清单（模板页标注豁免）；
               - 字号继承未显式设置 → warning（聚合计数）。
               图表部件（ppt/charts/*.xml）内的 sz 同步收集、同一规则判定。

用法：
  python3 check_geometry.py <deck.pptx> [--spec deck.json] [--tol 0.02]
      [--config font_sizes.json] [--json report.json] [--strict-warnings]

退出码：0 无 error；1 有 error（--strict-warnings 时 warning 亦判 1）；2 IO/用法错误。
报告：人读摘要打到 stdout；--json 写结构化报告（含 postflight 统计块）。
"""
import argparse
import json
import math
import re
import sys
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

EMU_IN = 914400
MIN_DIM_IN = 0.06          # 细线/连接线/发丝分隔条豁免阈值（0.06in≈1.5mm）
CONTAIN_EPS_IN = 0.01      # 包含判定容差
OVERLAP_RATIO = 0.05       # 侵入判定：交叠面积 / 文字形状面积
OVERLAP_MIN_AREA = 0.05    # 侵入判定：交叠绝对面积下限（in²），低于此为擦边
PAGE_AREA_RATIO_BLEED = 0.5  # 非文字形状占页面面积 ≥50% 时越界视为满版溢出
LAYOUT_FURNITURE_HINTS = ("页脚", "灯片编号", "幻灯片编号", "日期",
                          "footer", "slide number", "date placeholder")


# ---------------------------------------------------------------- geometry

def _f(v):
    return None if v is None else Emu(v).inches


def shape_bbox(shape, xform=None):
    """Effective bbox (x0, y0, x1, y1) in inches, slide coordinates.

    xform: composed affine (sx, sy, ox, oy) mapping shape-local EMU inches
    to slide inches: slide = ox + x * sx. Composed across all ancestor
    groups (v1.4.0); before that only the immediate parent was applied,
    which misplaced shapes in nested groups.
    Rotation (clockwise deg) is applied around the shape centre.
    Returns None when geometry is unavailable (e.g. freeform-less connector).
    """
    try:
        left, top, width, height = (_f(shape.left), _f(shape.top),
                                    _f(shape.width), _f(shape.height))
    except Exception:
        return None
    if None in (left, top, width, height):
        return None
    if xform is not None:
        sx, sy, ox, oy = xform
        left, top = ox + left * sx, oy + top * sy
        width, height = width * sx, height * sy
    rot = 0.0
    try:
        rot = float(shape.rotation or 0.0)
    except Exception:
        rot = 0.0
    if rot % 360:
        cx, cy = left + width / 2, top + height / 2
        th = math.radians(rot)
        c, s = abs(math.cos(th)), abs(math.sin(th))
        w2 = width * c + height * s
        h2 = width * s + height * c
        left, top, width, height = cx - w2 / 2, cy - h2 / 2, w2, h2
    return (left, top, left + width, top + height)


def compose_xform(shape, parent=None):
    """Compose parent transform with this group's child->parent mapping.

    PPTX group semantics: a child offset is expressed in the group's child
    coordinate space (a:chOff/a:chExt) and must be mapped through
    off + (child - chOff) * (ext / chExt) before the parent transform.
    """
    try:
        left, top = _f(shape.left), _f(shape.top)
        width, height = _f(shape.width), _f(shape.height)
    except Exception:
        return parent
    if None in (left, top, width, height):
        return parent
    sp = shape.element.find(
        './/{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
    if sp is None:
        return parent
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    ch_off = sp.find(f'{ns}chOff')
    ch_ext = sp.find(f'{ns}chExt')
    if ch_off is None or ch_ext is None:
        return parent
    try:
        chx = Emu(int(ch_off.get('x'))).inches
        chy = Emu(int(ch_off.get('y'))).inches
        cex = Emu(int(ch_ext.get('cx'))).inches
        cey = Emu(int(ch_ext.get('cy'))).inches
    except Exception:
        return parent
    if cex == 0 or cey == 0:
        return parent
    psx, psy, pox, poy = parent if parent else (1.0, 1.0, 0.0, 0.0)
    # group origin in slide space
    gx, gy = pox + left * psx, poy + top * psy
    # scale of child space -> slide space
    gsx, gsy = psx * (width / cex), psy * (height / cey)
    # slide = gx + (child - chOff) * gsx = (gx - chOff*gsx) + child*gsx
    return (gsx, gsy, gx - chx * gsx, gy - chy * gsy)


def iter_flat(shapes, xform=None):
    """Yield (shape, bbox) for every leaf shape, groups expanded.

    Nested groups compose their transforms down from the slide root.
    """
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_flat(sh.shapes, compose_xform(sh, xform))
        else:
            yield sh, shape_bbox(sh, xform)


def has_visible_text(shape):
    if getattr(shape, 'has_text_frame', False) and shape.has_text_frame:
        return bool(shape.text_frame.text.strip())
    return False


def contains(outer, inner, eps=CONTAIN_EPS_IN):
    return (outer[0] <= inner[0] + eps and outer[1] <= inner[1] + eps
            and outer[2] >= inner[2] - eps and outer[3] >= inner[3] - eps)


def overlap_area(a, b):
    dx = min(a[2], b[2]) - max(a[0], b[0])
    dy = min(a[3], b[3]) - max(a[1], b[1])
    return dx * dy if dx > 0 and dy > 0 else 0.0


def is_layout_furniture(shape):
    """页脚/页码/日期占位符属版式家具，不参与正文侵入判定。"""
    try:
        if shape.is_placeholder:
            ph = shape.placeholder_format.type
            if ph is not None:
                # 12=FOOTER 13=SLIDE_NUMBER 14=DATE（python-pptx 枚举值）
                if int(ph) in (12, 13, 14):
                    return True
    except Exception:
        pass
    name = (shape.name or "").lower()
    return any(h in name for h in LAYOUT_FURNITURE_HINTS)


def subtree_has_text(shape):
    """组形状递归判断子树内是否有可见文字。"""
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return any(subtree_has_text(c) for c in shape.shapes)
    return has_visible_text(shape)


def subtree_bbox(shape, xform=None):
    """叶子形状返回自身 bbox；组返回全部子叶子的并集。"""
    if shape.shape_type != MSO_SHAPE_TYPE.GROUP:
        return shape_bbox(shape, xform)
    boxes = [bb for _s, bb in iter_flat(shape.shapes,
                                        compose_xform(shape, xform))]
    boxes = [b for b in boxes if b]
    if not boxes:
        return None
    return (min(b[0] for b in boxes), min(b[1] for b in boxes),
            max(b[2] for b in boxes), max(b[3] for b in boxes))


def top_layers(slide, tol):
    """顶层图层拆分：文字承载体 vs 非文字装饰。

    侵入判定只看顶层图层（v1.4.0）。组内子形状是设计上的组合图形
    （如 islide 项目符号、图标簇），组内元素互相叠压属正常构图，
    逐叶子互查会产生大量误报。旋转形状的包围盒是过近似，
    同样不参与判定（真实遮挡由渲染目检兜底）。

    返回 (texts, decos, rotated_skipped)，元素为 (shape, bbox, z)。
    """
    texts, decos, rotated = [], [], 0
    for z, sh in enumerate(slide.shapes):
        bb = subtree_bbox(sh)
        if bb is None:
            continue
        st = sh.shape_type
        if subtree_has_text(sh):
            if not is_layout_furniture(sh):
                texts.append((sh, bb, z))
            continue
        if st in (MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART,
                  MSO_SHAPE_TYPE.LINE) or 'CONNECTOR' in str(st):
            continue
        if getattr(sh, 'has_table', False) and sh.has_table:
            continue
        if getattr(sh, 'has_chart', False) and sh.has_chart:
            continue
        try:
            rot = abs(float(sh.rotation or 0.0)) % 360
        except Exception:
            rot = 0.0
        if rot > 0.5 and abs(rot - 360) > 0.5:
            rotated += 1
            continue
        w, h = bb[2] - bb[0], bb[3] - bb[1]
        if w >= MIN_DIM_IN and h >= MIN_DIM_IN:
            decos.append((sh, bb, z))
    return texts, decos, rotated


# ---------------------------------------------------------------- fonts

def _para_default_size(para):
    """Paragraph-level default run size (a:pPr/a:defRPr@sz), in pt."""
    p = para._p
    pPr = p.find('{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    if pPr is not None:
        defRPr = pPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
        if defRPr is not None and defRPr.get('sz'):
            return int(defRPr.get('sz')) / 100.0
    return None


def collect_frame_fonts(tf, where, out):
    for para in tf.paragraphs:
        fallback = _para_default_size(para)
        for run in para.runs:
            if not run.text:
                continue
            size = run.font.size.pt if run.font.size is not None else fallback
            out.append((size, where, run.text[:24]))


def is_clean(size_pt):
    return abs(size_pt * 4 - round(size_pt * 4)) <= 0.04


# ---------------------------------------------------------------- main

def check_deck(pptx_path, config, spec_types=None, tol=0.02):
    prs = Presentation(str(pptx_path))
    page_w, page_h = _f(prs.slide_width), _f(prs.slide_height)
    errors, warnings = [], []
    stats = {"pages": len(prs.slides._sldIdLst), "shapes": 0,
             "text_frames": 0, "tables": 0, "charts": 0, "pictures": 0}
    fitted, small, undeclared, below_floor, below_floor_fitted = [], [], [], [], []
    bleed = []
    rotated_decos = 0
    inherited_runs = 0
    font_seen = {}

    for idx, slide in enumerate(prs.slides, 1):
        flat_raw = list(iter_flat(slide.shapes))
        # z 序 = 文档顺序（首个为最底层）。组内子形状按展开顺序递增，
        # 足以区分「卡片在下、文字在上」这类背景板关系。
        flat = [(sh, bb, z) for z, (sh, bb) in enumerate(flat_raw)]
        texts, decos = [], []
        for sh, bb, z in flat:
            stats["shapes"] += 1
            st = sh.shape_type
            if getattr(sh, 'has_table', False) and sh.has_table:
                stats["tables"] += 1
            if getattr(sh, 'has_chart', False) and sh.has_chart:
                stats["charts"] += 1
            if st == MSO_SHAPE_TYPE.PICTURE:
                stats["pictures"] += 1
            if bb is None:
                continue
            # ---- bounds
            over = []
            if bb[0] < -tol:
                over.append(f"左 {-bb[0]:.3f}in")
            if bb[1] < -tol:
                over.append(f"上 {-bb[1]:.3f}in")
            if bb[2] > page_w + tol:
                over.append(f"右 {bb[2] - page_w:.3f}in")
            if bb[3] > page_h + tol:
                over.append(f"下 {bb[3] - page_h:.3f}in")
            if over:
                area = (bb[2] - bb[0]) * (bb[3] - bb[1])
                w0, h0 = bb[2] - bb[0], bb[3] - bb[1]
                has_txt = has_visible_text(sh)
                if not has_txt and (
                        area >= PAGE_AREA_RATIO_BLEED * page_w * page_h
                        or min(w0, h0) < MIN_DIM_IN):
                    # 满版背景/大图，或发丝分割线：有意出血/拉出页边
                    bleed.append(
                        f"p{idx} 形状「{sh.name}」满版溢出：{'、'.join(over)}"
                        f"（占页面 {area / (page_w * page_h) * 100:.0f}%，"
                        f"按有意出血处理）")
                else:
                    errors.append(
                        f"[bounds] p{idx} 形状「{sh.name}」越界：{'、'.join(over)}"
                        f"（bbox {bb[0]:.2f},{bb[1]:.2f}–{bb[2]:.2f},{bb[3]:.2f}，"
                        f"页面 {page_w:.2f}×{page_h:.2f}）")
            # ---- classify for intrusion (顶层图层，见下)
            if has_visible_text(sh):
                stats["text_frames"] += 1
        # ---- intrusion（v1.4.0：按顶层图层判定）
        texts, decos, rot_skip = top_layers(slide, tol)
        rotated_decos += rot_skip
        for tsh, tbb, tz in texts:
            tarea = max((tbb[2] - tbb[0]) * (tbb[3] - tbb[1]), 1e-9)
            for dsh, dbb, dz in decos:
                if dsh is tsh:
                    continue
                if dz < tz:
                    continue  # 位于文字下方 = 背景板/卡片，不构成遮挡
                oa = overlap_area(tbb, dbb)
                if oa <= 0:
                    continue
                if contains(dbb, tbb) or contains(tbb, dbb):
                    continue  # 容器包含/被包含：版式结构，非侵入
                ratio = oa / tarea
                if ratio > OVERLAP_RATIO and oa >= OVERLAP_MIN_AREA:
                    errors.append(
                        f"[intrusion] p{idx} 非文字形状「{dsh.name}」与文字形状"
                        f"「{tsh.name}」部分交叠，交叠面积 {oa:.2f}in²"
                        f"（占文字形状 {ratio * 100:.1f}%，阈值 5% 且 "
                        f"≥{OVERLAP_MIN_AREA}in²）")
        # ---- fonts (slide text + tables)
        runs = []
        for sh, bb, _z in flat:
            if getattr(sh, 'has_text_frame', False) and sh.has_text_frame:
                collect_frame_fonts(sh.text_frame, f"p{idx}「{sh.name}」", runs)
            if getattr(sh, 'has_table', False) and sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        collect_frame_fonts(cell.text_frame,
                                            f"p{idx} 表格「{sh.name}」", runs)
        page_type = (spec_types or {}).get(idx)
        for size, where, snippet in runs:
            if size is None:
                inherited_runs += 1
                continue
            font_seen[round(size, 2)] = font_seen.get(round(size, 2), 0) + 1
            if size < config["floor_error_pt"] and not is_clean(size):
                below_floor_fitted.append((idx, size, where, snippet))
            elif size < config["floor_error_pt"]:
                below_floor.append((idx, size, where, snippet, page_type))
            elif is_clean(size) and size not in config["allowed_pt"]:
                undeclared.append((idx, size, where, snippet))
            elif not is_clean(size):
                fitted.append((idx, size, where, snippet))
            if size < config["small_report_pt"]:
                small.append((idx, size, where, snippet, page_type))

    # ---- chart-part fonts
    chart_sizes = []
    try:
        with zipfile.ZipFile(pptx_path) as z:
            for name in z.namelist():
                if re.match(r'ppt/charts/chart\d+\.xml$', name):
                    xml = z.read(name).decode('utf-8', 'ignore')
                    for m in re.finditer(r'\bsz="(\d+)"', xml):
                        chart_sizes.append(int(m.group(1)) / 100.0)
    except Exception:
        pass
    chart_undeclared = sorted({s for s in chart_sizes
                               if s >= config["floor_error_pt"]
                               and is_clean(s) and s not in config["allowed_pt"]})
    chart_below = sorted({s for s in chart_sizes
                          if s < config["floor_error_pt"] and is_clean(s)})
    chart_below_fitted = sorted({s for s in chart_sizes
                                 if s < config["floor_error_pt"] and not is_clean(s)})
    for s in chart_below_fitted:
        below_floor_fitted.append((0, s, "图表部件", ""))
    for s in chart_undeclared:
        undeclared.append((0, s, "图表部件", ""))
    for s in chart_below:
        below_floor.append((0, s, "图表部件", "", None))

    for idx, size, where, snippet, _pt in below_floor:
        errors.append(f"[font] {where} 字号 {size:g}pt 低于下限 "
                      f"{config['floor_error_pt']:g}pt（「{snippet}」）")
    for idx, size, where, snippet in undeclared:
        errors.append(f"[font] {where} 字号 {size:g}pt 未在白名单申报"
                      f"（「{snippet}」）；新字号须先登记 font_sizes.json")
    if below_floor_fitted:
        agg = {}
        for idx, size, where, snippet in below_floor_fitted:
            agg.setdefault(round(size, 2), []).append(where)
        warnings.append("[font] 低于 8pt 的自动适配字号（计算值，非声明值，"
                        "仅提示）：" + "；".join(
                            f"{s:g}pt ×{len(w)}" for s, w in sorted(agg.items())))
    if fitted:
        agg = {}
        for idx, size, where, snippet in fitted:
            agg.setdefault(round(size, 2), []).append(idx)
        warnings.append("[font] 自动适配拟合字号（orgchart _fit_size 连续值，"
                        "仅提示）：" + "；".join(
                            f"{s:g}pt ×{len(p)}（页 {sorted(set(p))}）"
                            for s, p in sorted(agg.items())))
    if inherited_runs:
        warnings.append(f"[font] {inherited_runs} 个 run 字号为继承值（未显式 "
                        f"设置），未纳入白名单判定")
    if bleed:
        warnings.append("[bounds] 满版出血（有意溢出，非缺陷）：\n  - "
                        + "\n  - ".join(bleed))
    if rotated_decos:
        warnings.append(f"[intrusion] {rotated_decos} 个顶层旋转装饰形状未参与"
                        f"侵入判定（旋转包围盒为过近似，遮挡请以渲染目检为准）")
    small_notes = []
    for idx, size, where, snippet, ptype in small:
        tag = "（template_table 模板页豁免范围）" if ptype == "template_table" else ""
        small_notes.append(f"p{idx} {size:g}pt {where}「{snippet}」{tag}")
    if small_notes:
        warnings.append("[font] <9pt 字号单列清单：\n  - " + "\n  - ".join(small_notes))

    report = {
        "deck": str(pptx_path),
        "page_size_in": [round(page_w, 3), round(page_h, 3)],
        "errors": errors,
        "warnings": warnings,
        "font_sizes_seen_pt": {str(k): v for k, v in sorted(font_seen.items())},
        "postflight": {
            "pages": stats["pages"], "shapes": stats["shapes"],
            "text_frames": stats["text_frames"], "tables": stats["tables"],
            "charts": stats["charts"], "pictures": stats["pictures"],
            "geometry_gate": "FAIL" if errors else "PASS",
        },
    }
    return report


def main():
    ap = argparse.ArgumentParser(description="Enterprise PPT geometry checker")
    ap.add_argument("pptx")
    ap.add_argument("--spec", default="", help="spec JSON（可选，用于判定页型，"
                                                 "如 template_table 豁免标注）")
    ap.add_argument("--tol", type=float, default=0.02)
    ap.add_argument("--config", default=str(Path(__file__).with_name("font_sizes.json")))
    ap.add_argument("--json", default="")
    ap.add_argument("--strict-warnings", action="store_true")
    ap.add_argument("--min-pt", type=float, default=0.0,
                    help="覆盖字号下限（默认取配置 floor_error_pt=8）。"
                         "仅用于资产基线核对等特殊场景，交付稿勿改")
    ap.add_argument("--asset-mode", action="store_true",
                    help="官方资产/种子基线模式：字号下限放宽到 7pt。"
                         "仅用于检查 assets/ 下的参考文件，不代表交付标准")
    args = ap.parse_args()

    pptx_path = Path(args.pptx)
    if not pptx_path.exists():
        print(f"ERROR: file not found: {pptx_path}", file=sys.stderr)
        return 2
    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    config["allowed_pt"] = [float(x) for x in config["allowed_pt"]]
    mode = "delivery"
    if args.asset_mode:
        # 官方资产里 7pt 是密排组织芯片的既有取值，属参考基线而非交付标准
        config["floor_error_pt"] = 7.0
        if 7.0 not in config["allowed_pt"]:
            config["allowed_pt"] = sorted(config["allowed_pt"] + [7.0])
        mode = "asset-baseline"
    if args.min_pt:
        config["floor_error_pt"] = args.min_pt
        mode = "override"

    spec_types = None
    if args.spec:
        spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
        spec_types = {1: "cover"}
        for i, page in enumerate(spec.get("pages") or [], 2):
            spec_types[i] = page.get("type", "evidence")

    report = check_deck(pptx_path, config, spec_types, tol=args.tol)
    report["mode"] = mode
    report["floor_error_pt"] = config["floor_error_pt"]
    pf = report["postflight"]
    print(f"GEOMETRY CHECK {pptx_path.name}: "
          f"{'FAIL' if report['errors'] else 'PASS'} "
          f"（{pf['pages']} 页 / {pf['shapes']} 形状 / {pf['text_frames']} 文本帧 / "
          f"{pf['tables']} 表 / {pf['charts']} 图表 / {pf['pictures']} 图片）")
    if mode != "delivery":
        print(f"  MODE  {mode}：字号下限 {config['floor_error_pt']:g}pt"
              f"（非交付标准，仅用于资产基线与调试）")
    for e in report["errors"]:
        print(f"  ERROR {e}")
    for w in report["warnings"]:
        print(f"  WARN  {w}")
    if args.json:
        Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=2),
                                   encoding="utf-8")
        print(f"Report {args.json}")
    if report["errors"]:
        return 1
    if args.strict_warnings and report["warnings"]:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
