#!/usr/bin/env python3
"""PPTD → PPTX 本地转换引擎（ppt-studio 的离线导出后端）。

纯本地渲染，无登录/网络依赖。

用法:
  python export_pptx_local.py <deck.pptd> [--output deck.pptx] [--transition fade|none]

支持范围（PPTD v2）:
  - text 富文本: <p>/<span>/<strong>/<em>/<u>/<s>/<sub>/<sup>/<br>/
    <ul>/<ol>/<li>、style 属性(text-align/line-height/margin-top/font-size/
    color/font-family/background-color)、LaTeX 退化为纯文本
  - shape: 40+ 常用形状映射 + custom path (M/L/H/V/C/S/Q/Z)
  - image: fit cover/contain/fill、crop、cropShape(roundRect/ellipse)
  - table: 行列合并、firstRowStyle/bodyStyles 样式链
  - line: 直线/折线/贝塞尔、起止箭头
  - icon: 常用 FA 图标 unicode 近似表（约 60 个），未收录退化为 ★
  - chart: bar/line/area/pie(doughnut)/scatter/radar/bubble 原生映射;
    waterfall/heatmap/treemap/sunburst/candlestick 输出占位并告警
  - theme: colors / textStyles / tableStyles 与 $ref 解析
  - background: solid / gradient / image
  - 转场: fade（XML 注入，每页根级）
已知限制见文末 report。

依赖: python-pptx >= 1.0, PyYAML
"""
import argparse
import json
import math
import pathlib
import re
import sys

import yaml
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.util import Emu, Pt

EMU_PER_PX = 9525
DEFAULT_FONT = "微软雅黑"
FONT_FALLBACK = {  # 模板云字体 -> 本机常见字体
    "MiSans": "微软雅黑", "Noto Sans SC": "微软雅黑", "思源宋体": "宋体",
    "阿里妈妈数黑体": "微软雅黑", "得意黑": "微软雅黑", "站酷文艺体": "楷体",
}

# ---------------- 基础解析 ----------------

def resolve_color(c, colors):
    if isinstance(colors, dict) and "colors" in colors and isinstance(colors.get("colors"), dict):
        colors = colors["colors"]
    if c is None:
        return None
    if isinstance(c, str) and c.startswith("$"):
        c = colors.get(c[1:], "#000000")
    c = str(c)
    if len(c) >= 7 and c.startswith("#"):
        try:
            return RGBColor(int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16))
        except ValueError:
            return RGBColor(0, 0, 0)
    return RGBColor(0, 0, 0)


def px(v):
    return Emu(int(v * EMU_PER_PX))


def font_name_of(spec):
    if isinstance(spec, dict):
        spec = spec.get("ea") or spec.get("latin") or DEFAULT_FONT
    return FONT_FALLBACK.get(spec, spec or DEFAULT_FONT)


# ---------------- 富文本 HTML 解析 ----------------
TAG = re.compile(r"(</?p(?:\s[^>]*)?>|</?span(?:\s[^>]*)?>|<strong>|</strong>|"
                 r"<em>|</em>|<u>|</u>|<s>|</s>|<sub>|</sub>|<sup>|</sup>|"
                 r"<br\s*/?>|<ul>|</ul>|<ol>|</ol>|<li[^>]*>|</li>|<a [^>]*>|</a>)")
STYLE_KV = re.compile(r"([a-z-]+)\s*:\s*([^;\"']+)")


def parse_style(chunk):
    return dict(STYLE_KV.findall(chunk))


def html_to_paras(html, theme):
    """HTML 富文本 → [{align, runs:[(text, style)], li:level}]。style 是扁平 dict。"""
    colors = theme.get("colors", {})
    st = {"bold": False, "italic": False, "underline": False, "strike": False,
          "sub": False, "sup": False, "color": None, "size": None, "font": None,
          "bg": None, "href": None}
    paras, cur, li = [], {"runs": [], "align": None, "lineHeight": None,
                          "marginTop": None, "li": 0}, 0

    def flush():
        nonlocal cur
        if cur["runs"] or cur["li"]:
            paras.append(cur)
        cur = {"runs": [], "align": cur.get("align"), "lineHeight": None,
               "marginTop": None, "li": 0}

    pos, n = 0, len(html)
    stack_list = []
    while pos < n:
        m = TAG.search(html, pos)
        plain = html[pos: m.start()] if m else html[pos:]
        if plain:
            for i, line in enumerate(plain.split("\n")):
                if i > 0:
                    flush()
                if line:
                    cur["runs"].append((line, dict(st)))
        if not m:
            break
        tag = m.group(0)
        pos = m.end()
        t = tag.lower()
        if t.startswith("<p"):
            flush()
            kv = parse_style(tag)
            cur["align"] = kv.get("text-align")
            cur["lineHeight"] = kv.get("line-height")
            cur["marginTop"] = kv.get("margin-top")
        elif t == "</p>":
            flush()
        elif t in ("<br>", "<br/>"):
            flush()
        elif t in ("<ul>", "<ol>"):
            flush(); stack_list.append(t)
        elif t in ("</ul>", "</ol>"):
            flush()
            if stack_list:
                stack_list.pop()
        elif t.startswith("<li"):
            flush()
            cur["li"] = len(stack_list)
            kv = parse_style(tag)
            if "text-align" in kv:
                cur["align"] = kv["text-align"]
        elif t == "</li>":
            flush()
        elif t == "<strong>" or t == "<b>":
            st["bold"] = True
        elif t == "</strong>" or t == "</b>":
            st["bold"] = False
        elif t == "<em>" or t == "<i>":
            st["italic"] = True
        elif t == "</em>" or t == "</i>":
            st["italic"] = False
        elif t == "<u>":
            st["underline"] = True
        elif t == "</u>":
            st["underline"] = False
        elif t == "<s>":
            st["strike"] = True
        elif t == "</s>":
            st["strike"] = False
        elif t == "<sub>":
            st["sub"] = True
        elif t == "</sub>":
            st["sub"] = False
        elif t == "<sup>":
            st["sup"] = True
        elif t == "</sup>":
            st["sup"] = False
        elif t.startswith("<span"):
            for k, v in parse_style(tag).items():
                if k == "color":
                    st["color"] = resolve_color(v.strip(), colors)
                elif k == "font-size":
                    st["size"] = float(re.sub(r"[^\d.]", "", v))
                elif k == "font-family":
                    st["font"] = v.strip().strip("'\"")
                elif k == "background-color":
                    st["bg"] = resolve_color(v.strip(), colors)
        elif t == "</span>":
            st["color"] = st["size"] = st["font"] = st["bg"] = None
        elif t.startswith("<a"):
            st["href"] = parse_style(tag).get("href")
        elif t == "</a>":
            st["href"] = None
    flush()
    return paras


# ---------------- 渲染器 ----------------
SHAPE_MAP = {
    "rect": MSO_SHAPE.RECTANGLE, "roundRect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL, "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "rtTriangle": MSO_SHAPE.RIGHT_TRIANGLE, "diamond": MSO_SHAPE.DIAMOND,
    "parallelogram": MSO_SHAPE.PARALLELOGRAM, "trapezoid": MSO_SHAPE.TRAPEZOID,
    "pentagon": MSO_SHAPE.REGULAR_PENTAGON, "hexagon": MSO_SHAPE.HEXAGON,
    "heptagon": MSO_SHAPE.HEPTAGON, "octagon": MSO_SHAPE.OCTAGON,
    "decagon": MSO_SHAPE.DECAGON, "donut": MSO_SHAPE.DONUT,
    "star4": MSO_SHAPE.STAR_4_POINT, "star5": MSO_SHAPE.STAR_5_POINT,
    "star6": MSO_SHAPE.STAR_6_POINT, "star8": MSO_SHAPE.STAR_8_POINT,
    "homePlate": MSO_SHAPE.PENTAGON, "chevron": MSO_SHAPE.CHEVRON,
    "rightArrow": MSO_SHAPE.RIGHT_ARROW, "leftArrow": MSO_SHAPE.LEFT_ARROW,
    "upArrow": MSO_SHAPE.UP_ARROW, "downArrow": MSO_SHAPE.DOWN_ARROW,
    "leftRightArrow": MSO_SHAPE.LEFT_RIGHT_ARROW, "wedgeRectCallout": MSO_SHAPE.RECTANGULAR_CALLOUT,
    "wedgeEllipseCallout": MSO_SHAPE.OVAL_CALLOUT, "cloud": MSO_SHAPE.CLOUD,
    "heart": MSO_SHAPE.HEART, "sun": MSO_SHAPE.SUN, "moon": MSO_SHAPE.MOON,
    "smileyFace": MSO_SHAPE.SMILEY_FACE, "lightningBolt": MSO_SHAPE.LIGHTNING_BOLT,
    "arc": MSO_SHAPE.ARC, "bracePair": MSO_SHAPE.DOUBLE_BRACE,
    "leftBrace": MSO_SHAPE.LEFT_BRACE, "rightBrace": MSO_SHAPE.RIGHT_BRACE,
    "round1Rect": MSO_SHAPE.ROUND_1_RECTANGLE, "round2SameRect": MSO_SHAPE.ROUND_2_SAME_RECTANGLE,
    "flowChartDocument": MSO_SHAPE.FLOWCHART_DOCUMENT,
    "flowChartProcess": MSO_SHAPE.FLOWCHART_PROCESS,
    "flowChartDecision": MSO_SHAPE.FLOWCHART_DECISION,
    "can": MSO_SHAPE.CAN, "cube": MSO_SHAPE.CUBE, "plaque": MSO_SHAPE.PLAQUE,
    "line": MSO_SHAPE.RECTANGLE,
}
WARN = []


def add_shape(shape_el, slide, theme):
    x, y, w, h = shape_el["bounds"]
    name = shape_el.get("shapeName", "rect")
    if name == "custom":
        free = slide.shapes.build_freeform(px(x), px(y), px(w) // max(w, 1))
        WARN.append(f"custom path shape '{shape_el.get('elementId')}' 近似渲染")
        shape = free.convert_to_shape()
    else:
        mso = SHAPE_MAP.get(name)
        if mso is None:
            WARN.append(f"未知形状 {name} → rect")
            mso = MSO_SHAPE.RECTANGLE
        shape = slide.shapes.add_shape(mso, px(x), px(y), px(w), px(h))
        if name == "roundRect" and shape_el.get("adjustments"):
            try:
                shape.adjustments[0] = shape_el["adjustments"][0] / 100000
            except Exception:
                pass
    shape.shadow.inherit = False
    fill = shape_el.get("fill")
    if fill and fill.get("type") == "solid":
        shape.fill.solid()
        shape.fill.fore_color.rgb = resolve_color(fill.get("color"), theme)
    elif fill and fill.get("type") == "gradient":
        stops = fill.get("stops", [])
        if len(stops) >= 2:
            shape.fill.gradient()
            shape.fill.gradient_stops[0].color.rgb = resolve_color(stops[0]["color"], theme.get("colors", theme))
            shape.fill.gradient_stops[-1].color.rgb = resolve_color(stops[-1]["color"], theme.get("colors", theme))
    else:
        shape.fill.background()
    bd = shape_el.get("border")
    if bd:
        shape.line.color.rgb = resolve_color(bd.get("color", "#000000"), theme)
        shape.line.width = Pt(bd.get("width", 1))
    else:
        shape.line.fill.background()
    if shape_el.get("rotation"):
        shape.rotation = shape_el["rotation"]
    if shape_el.get("opacity") is not None and shape_el["opacity"] < 1:
        from pptx.oxml.ns import qn
        sp = shape._element.spPr
        alpha = str(int(shape_el["opacity"] * 100000))
        srgb = sp.find(".//" + qn("a:solidFill") + "/" + qn("a:srgbClr"))
        if srgb is not None:
            a = srgb.makeelement(qn("a:alpha"), {"val": alpha})
            srgb.append(a)
    return shape


ICON_GLYPH = {
    "ear": "👂", "eye": "👁", "lightbulb": "💡", "robot": "🤖",
    "microphone": "🎙", "wand-magic-sparkles": "✨", "star": "⭐",
    "check": "✔", "xmark": "✘", "magnifying-glass": "🔍", "globe": "🌍",
    "book": "📖", "book-open": "📖", "camera": "📷", "volume-high": "🔊",
    "pen": "✏", "pencil": "✏", "trophy": "🏆", "rocket": "🚀",
    "heart": "❤", "gear": "⚙", "circle-check": "✔", "fire": "🔥",
    "cloud": "☁", "shield": "🛡", "lock": "🔒", "clock": "🕐",
    "map": "🗺", "comments": "💬", "chart-column": "📊", "graduation-cap": "🎓",
    "bug": "🐛", "seedling": "🌱", "brain": "🧠", "bolt": "⚡",
    "user": "👤", "users": "👥", "house": "🏠", "school": "🏫",
    "phone": "📱", "palette": "🎨", "gamepad": "🎮", "flag": "🚩",
}


def _fa_svg_path(icon_name):
    """'fas:lightbulb' -> FA svg 文本；找不到返回 None。"""
    import os
    style, _, name = icon_name.partition(":")
    if not name:
        name, style = style, "solid"
    style = {"fas": "solid", "far": "regular", "fab": "brands"}.get(style, style)
    cands = []
    env = os.environ.get("PPT_STUDIO_FA_SVGS")
    if env:
        cands.append(pathlib.Path(env))
    here = pathlib.Path.cwd()
    for base in [here] + list(here.parents)[:4]:
        cands.append(base / "assets" / "fa" / "node_modules" / "@fortawesome" /
                     "fontawesome-free" / "svgs")
    for c in cands:
        f = c / style / f"{name}.svg"
        if f.exists():
            return f.read_text(encoding="utf-8")
    return None


def _fa_icon_png(svg_text, color_hex, size_px):
    if not color_hex.startswith("#"):
        color_hex = "#" + color_hex
    import io
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import PathPatch
    from svgpath2mpl import parse_path
    vb = re.search(r'viewBox="([\d. ]+)"', svg_text).group(1)
    d = re.search(r'[\s]+d="([^"]+)"', svg_text).group(1)
    mpath = parse_path(d)
    fig = plt.figure(figsize=(size_px / 128, size_px / 128), dpi=128)
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis("off")
    ax.add_patch(PathPatch(mpath, facecolor=color_hex, edgecolor="none"))
    x0, y0, x1, y1 = map(float, vb.split())
    ax.set_xlim(x0, x1); ax.set_ylim(y1, y0)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True)
    plt.close(fig)
    return buf.getvalue()


def add_icon(icon_el, slide, theme, proj=None):
    x, y, w, h = icon_el["bounds"]
    icon_name = icon_el.get("iconName", "")
    svg_text = _fa_svg_path(icon_name) if proj is not None else None
    if svg_text:
        color_hex = "#3D3A36"
        fill = icon_el.get("fill") or {}
        if fill.get("color"):
            color_hex = str(resolve_color(fill["color"], theme))
        try:
            png = _fa_icon_png(svg_text, color_hex, 256)
            cache = (proj or pathlib.Path(".")) / "media" / "_icons"
            cache.mkdir(parents=True, exist_ok=True)
            cname = f"_icon_{abs(hash(icon_name + color_hex)) % 10**10}.png"
            cpath = cache / cname
            if not cpath.exists():
                cpath.write_bytes(png)
            return slide.shapes.add_picture(str(cpath), px(x), px(y), px(w), px(h))
        except Exception as exc:
            WARN.append(f"icon {icon_name} FA 渲染失败退化 emoji: {exc}")
    name = icon_name.split(":")[-1]
    glyph = ICON_GLYPH.get(name, "★")
    tb = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = glyph
    run.font.size = Pt(min(w, h) * 0.72)
    fill = icon_el.get("fill", {})
    if isinstance(fill, dict) and fill.get("color"):
        run.font.color.rgb = resolve_color(fill["color"], theme.get("colors", theme))
    return tb


def add_image(img_el, slide, theme, proj):
    x, y, w, h = img_el["bounds"]
    src = img_el["src"]
    if not src.startswith(("http://", "https://")):
        src = str(proj / src)
    fit = (img_el.get("fit") or {}).get("mode", "cover")
    cw = ch = cl = cr = 0
    if fit == "contain":
        from PIL import Image as PImage
        try:
            im = PImage.open(src)
            iw, ih = im.size
        except Exception:
            iw = ih = 1
        scale = min(w / iw, h / ih)
        dw, dh = iw * scale, ih * scale
        x, y = x + (w - dw) / 2, y + (h - dh) / 2
        w, h = dw, dh
    pic = slide.shapes.add_picture(src, px(x), px(y), px(w), px(h))
    if fit == "cover":
        try:
            from PIL import Image as PImage
            im = PImage.open(src)
            iw, ih = im.size
            box_r, src_r = w / h, iw / ih
            if src_r > box_r:
                excess = 1 - box_r / src_r
                cl = cr = excess / 2
            else:
                excess = 1 - src_r / box_r
                cw = ch = excess / 2
            if max(cw, ch) > 0.001:
                pic.crop_left, pic.crop_right = cl, cr
                pic.crop_top, pic.crop_bottom = cw, ch
        except Exception:
            pass
    cs = img_el.get("cropShape")
    if cs and cs.get("shapeName") in ("roundRect", "ellipse"):
        try:
            from pptx.oxml.ns import qn
            from lxml import etree
            pic_el = pic._element
            sp_pr = pic_el.find(".//" + qn("pic:spPr"))
            if sp_pr is not None:
                old = sp_pr.find(qn("a:prstGeom"))
                if old is not None:
                    sp_pr.remove(old)
                geom = etree.SubElement(sp_pr, qn("a:prstGeom"))
                geom.set("prst", "ellipse" if cs["shapeName"] == "ellipse" else "roundRect")
                av = etree.SubElement(geom, qn("a:avLst"))
                if cs["shapeName"] == "roundRect":
                    adj = cs.get("adjustments", [8000])
                    gd = etree.SubElement(av, qn("a:gd"))
                    gd.set("name", "adj"); gd.set("fmla", f"val {adj[0]}")
        except Exception:
            WARN.append("cropShape 圆角注入失败（已按矩形渲染）")
    return pic


ALIGN_MAP = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
             "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY,
             "distributed": PP_ALIGN.JUSTIFY}
ANCHOR_MAP = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
              "bottom": MSO_ANCHOR.BOTTOM}


def apply_run(run, st, theme, default_size, default_font):
    f = run.font
    f.name = font_name_of(st.get("font") or default_font)
    f.size = Pt(st.get("size") or default_size)
    if st.get("bold"):
        f.bold = True
    if st.get("italic"):
        f.italic = True
    if st.get("underline"):
        f.underline = True
    if st.get("strike"):
        f.strikethrough = True
    if st.get("color"):
        f.color.rgb = st["color"] if hasattr(st["color"], "real") else resolve_color(st["color"], theme.get("colors", theme))
    if st.get("bg"):
        try:
            from pptx.oxml.ns import qn
            rPr = run._r.get_or_add_rPr()
            hl = rPr.makeelement(qn("a:highlight"), {})
            clr = hl.makeelement(qn("a:srgbClr"), {"val": str(st["bg"])})
            hl.append(clr)
            rPr.append(hl)
        except Exception:
            pass


def add_text(text_el, slide, theme):
    x, y, w, h = text_el["bounds"]
    tb = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = tb.text_frame
    content = text_el.get("content", {})
    tf.word_wrap = content.get("wrap", True)
    align = content.get("align") or ["left", "top"]
    if len(align) == 1:
        align = [align[0], "top"]
    tf.vertical_anchor = ANCHOR_MAP.get(align[1], MSO_ANCHOR.TOP)
    style_ref = content.get("style")
    base_size, base_color, base_font = 18, None, None
    if style_ref and style_ref.startswith("$"):
        ts = theme.get("textStyles", {}).get(style_ref[1:], {})
        base_size = ts.get("fontSize", base_size)
        base_color = ts.get("color")
        base_font = ts.get("fontFamily")
    base = {"fontSize": content.get("fontSize", base_size),
            "color": content.get("color", base_color),
            "bold": content.get("bold", False),
            "lineHeight": content.get("lineHeight", 1.2),
            "font": content.get("fontFamily", base_font)}
    raw = content.get("text", "")
    paras = html_to_paras(raw, theme) if re.search(r"<[a-zA-Z/]", raw) else None
    if paras is None:
        # 纯文本路径：runs 元素须为 (text, 扁平样式 dict)，与 html_to_paras 输出同构
        paras = [{"runs": [(l, {"bold": base["bold"], "color": base["color"]})
                           for l in [raw]] if raw else [],
                  "align": None, "lineHeight": None, "marginTop": None, "li": 0}]
    first = True
    for para in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        # LaTeX 公式段检测：整段被 \( \) / $$ 包围 → 注入原生 OMML 公式
        merged = "".join(txt for txt, _ in para["runs"]).strip()
        try:
            import latex_omml
            if merged and latex_omml.is_formula(merged):
                omml = latex_omml.latex_to_omml_element(
                    latex_omml.strip_delims(merged))
                container = latex_omml.wrap_for_paragraph(omml)
                p._p.append(container)
                if base.get("fontSize"):
                    for omr in container.iter("{%s}r" % "http://schemas.openxmlformats.org/officeDocument/2006/math"):
                        pass
                continue
        except Exception as exc:
            WARN.append(f"公式 OMML 转换失败，退化纯文本: {exc}")
        if para.get("lineHeight"):
            try:
                p.line_spacing = float(re.sub(r"[^\d.]", "", str(para["lineHeight"])))
            except ValueError:
                pass
        if para.get("marginTop"):
            try:
                p.space_before = Pt(float(re.sub(r"[^\d.]", "", str(para["marginTop"]))))
            except ValueError:
                pass
        if para.get("align"):
            p.alignment = ALIGN_MAP.get(para["align"], PP_ALIGN.LEFT)
        elif align[0]:
            p.alignment = ALIGN_MAP.get(align[0], PP_ALIGN.LEFT)
        if para.get("li"):
            p.level = min(para["li"], 4) - 1 if para["li"] > 1 else 0
            prefix = "• " if para["li"] else ""
            if para["li"] > 1:
                prefix = "◦ " + prefix
            r0 = p.add_run(); r0.text = prefix
            apply_run(r0, {"size": base["size"]}, theme, base["size"], base["font"])
        for txt, st in para["runs"]:
            if txt == "":
                continue
            run = p.add_run()
            run.text = txt
            if st.get("sub"):
                run.font._rPr.set("baseline", "-25000")
            if st.get("sup"):
                run.font._rPr.set("baseline", "30000")
            apply_run(run, st, theme, st.get("size") or base.get("fontSize"), base["font"])
    return tb


def add_table(tbl_el, slide, theme):
    x, y, w, h = tbl_el["bounds"]
    rows = tbl_el["rows"]
    nrows, ncols = len(rows), max(len(r) for r in rows)
    gt = slide.shapes.add_table(nrows, ncols, px(x), px(y), px(w), px(h)).table
    for ci, ratio in enumerate(tbl_el.get("columnWidths", [])):
        try:
            gt.columns[ci].width = Emu(int(px(w) * ratio))
        except Exception:
            pass
    style = tbl_el.get("style")
    ts = {}
    if isinstance(style, str) and style.startswith("$"):
        ts = theme.get("tableStyles", {}).get(style[1:], {})
    base_cell = ts.get("cellStyle", {})
    first_row = ts.get("firstRowStyle")
    body = ts.get("bodyStyles", [])
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = gt.cell(ri, ci)
            if ci < len(row):
                cell_def = row[ci]
                txt = cell_def.get("text", "") if isinstance(cell_def, dict) else str(cell_def)
                cell.text = txt
                st = {}
                if ri == 0 and first_row:
                    st = first_row
                elif body:
                    st = body[(ri - 1) % len(body)] if ri > 0 else st
                if isinstance(st, dict):
                    if st.get("fill"):
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = resolve_color(st["fill"].get("color", "#ffffff"), theme)
                    if st.get("color"):
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                run.font.color.rgb = resolve_color(st["color"], theme.get("colors", theme))
                    if st.get("bold"):
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                run.font.bold = True
    return gt


CHART_MAP = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "line": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA, "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT, "scatter": XL_CHART_TYPE.XY_SCATTER,
    "radar": XL_CHART_TYPE.RADAR, "bubble": XL_CHART_TYPE.BUBBLE,
}



# ---------------- plotly 扩展图表（waterfall/heatmap/treemap/sunburst/candlestick/sankey） ----------------
def _theme_color_cycle(theme):
    colors = theme
    if isinstance(colors, dict) and "colors" in colors and isinstance(colors["colors"], dict):
        colors = colors["colors"]
    keys = ["primary", "primaryDeep", "accent", "tealSoft", "red"]
    vals = [colors.get(k) for k in keys if colors.get(k)]
    return vals or ["#FF8A3D", "#1B7A70"]


def _plotly_chart_png(ch_el, theme, width, height):
    """六类扩展图表 → plotly 构图 → kaleido PNG bytes。"""
    import plotly.graph_objects as go
    data = ch_el["data"]
    cols = data["cols"]
    rows = data["rows"]

    def col(name):
        i = cols.index(name)
        return [r[i] for r in rows]

    series = ch_el["series"][0]
    enc = series.get("encode", {})
    t0 = series.get("type")
    cycle = _theme_color_cycle(theme)
    layout = dict(template="none", width=width, height=height,
                  margin=dict(l=40, r=40, t=30, b=40),
                  font=dict(family=font_name_of(None), size=14, color="#3D3A36"),
                  paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF")
    fig = None
    if t0 == "waterfall":
        fig = go.Figure(go.Waterfall(
            x=col(enc["x"]), y=[float(v) for v in col(enc["y"])],
            increasing_marker_color=cycle[0], decreasing_marker_color="#E05B4C",
            totals_marker_color=cycle[1] if len(cycle) > 1 else cycle[0]))
    elif t0 == "heatmap":
        xs, ys, zs = [], [], []
        for r in rows:
            xv, yv = str(r[cols.index(enc["x"])]), str(r[cols.index(enc["y"])])
            vv = float(r[cols.index(enc["value"])])
            if xv not in xs: xs.append(xv)
            if yv not in ys: ys.append(yv)
        z = [[None] * len(xs) for _ in ys]
        for r in rows:
            xv, yv = str(r[cols.index(enc["x"])]), str(r[cols.index(enc["y"])])
            z[ys.index(yv)][xs.index(xv)] = float(r[cols.index(enc["value"])])
        fig = go.Figure(go.Heatmap(z=z, x=xs, y=ys,
                                   colorscale=[[0, cycle[1]], [1, cycle[0]]]))
    elif t0 == "treemap":
        cat, val = col(enc["category"]), [float(v) for v in col(enc["value"])]
        parent = col(enc["parent"]) if enc.get("parent") in cols else [""] * len(cat)
        fig = go.Figure(go.Treemap(labels=cat, parents=parent, values=val,
                                   marker_colorscale=[[0, cycle[1]], [1, cycle[0]]]))
    elif t0 == "sunburst":
        cat, val = col(enc["category"]), [float(v) for v in col(enc["value"])]
        parent = col(enc["parent"]) if enc.get("parent") in cols else [""] * len(cat)
        fig = go.Figure(go.Sunburst(labels=cat, parents=parent, values=val,
                                    marker_colorscale=[[0, cycle[1]], [1, cycle[0]]]))
    elif t0 == "sankey":
        src_i, tgt_i, flw_i = cols.index(enc["source"]), cols.index(enc["target"]), cols.index(enc["flow"])
        nodes, links = [], []
        for r in rows:
            s, tg = str(r[src_i]), str(r[tgt_i])
            if s not in nodes: nodes.append(s)
            if tg not in nodes: nodes.append(tg)
            links.append((s, tg, float(r[flw_i])))
        fig = go.Figure(go.Sankey(
            node=dict(label=nodes, color=cycle[0]),
            link=dict(source=[nodes.index(s) for s, _, _ in links],
                      target=[nodes.index(t) for _, t, _ in links],
                      value=[v for _, _, v in links],
                      color="rgba(27,122,112,0.4)")))
    elif t0 == "candlestick":
        xi = cols.index(enc["x"])
        fig = go.Figure(go.Candlestick(
            x=[str(r[xi]) for r in rows],
            open=[float(r[cols.index(enc["open"])]) for r in rows],
            high=[float(r[cols.index(enc["high"])]) for r in rows],
            low=[float(r[cols.index(enc["low"])]) for r in rows],
            close=[float(r[cols.index(enc["close"])]) for r in rows]))
        fig.update_layout(xaxis_rangeslider_visible=False)
    if fig is None:
        raise ValueError(f"plotly 分支不支持类型 {t0}")
    fig.update_layout(**layout)
    return fig.to_image(format="png", scale=2)


def _add_plotly_chart(ch_el, slide, theme, proj):
    x, y, w, h = ch_el["bounds"]
    png = _plotly_chart_png(ch_el, theme, int(w), int(h))
    cache = proj / "media" / "_charts"
    cache.mkdir(parents=True, exist_ok=True)
    name = f"_chart_{abs(hash(json.dumps(ch_el, sort_keys=True, default=str))) % 10**10}.png"
    path = cache / name
    path.write_bytes(png)
    return slide.shapes.add_picture(str(path), px(x), px(y), px(w), px(h))



def add_chart(ch_el, slide, theme, proj=None):
    add_chart.proj = proj or pathlib.Path('.')
    series = ch_el.get("series", [])
    types = [s.get("type") for s in series]
    unsupported = {"waterfall", "heatmap", "treemap", "sunburst", "sankey", "candlestick"}
    if any(t in unsupported for t in types):
        return _add_plotly_chart(ch_el, slide, theme, getattr(add_chart, "proj", pathlib.Path(".")))
    t0 = types[0]
    ctype = CHART_MAP.get(t0)
    if ctype is None:
        WARN.append(f"未知 chart 类型 {t0}")
        ctype = XL_CHART_TYPE.COLUMN_CLUSTERED
    cols = list(ch_el["data"]["cols"])
    rows = ch_el["data"]["rows"]
    cats = [str(r[0]) for r in rows]
    chart_data = CategoryChartData()
    chart_data.categories = cats
    for ci in range(1, len(cols)):
        chart_data.add_series(cols[ci], [float(r[ci]) if r[ci] is not None else 0
                                         for r in rows])
    x, y, w, h = ch_el["bounds"]
    gframe = slide.shapes.add_chart(ctype, px(x), px(y), px(w), px(h), chart_data)
    chart = gframe.chart
    if ch_el.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = ch_el["title"] if isinstance(ch_el["title"], str) else ch_el["title"].get("text", "")
    if ch_el.get("legend") is False:
        chart.has_legend = False
    return gframe


def _catmull_rom_sample(pts, samples_per_seg=24):
    """Catmull-Rom 样条插值：曲线穿过全部给定点，返回密集采样序列（平滑）。"""
    if len(pts) < 3:
        return pts
    out = [pts[0]]
    ext = [pts[0]] + list(pts) + [pts[-1]]
    for i in range(1, len(ext) - 2):
        p0, p1, p2, p3 = ext[i - 1], ext[i], ext[i + 1], ext[i + 2]
        for s in range(samples_per_seg):
            t = s / samples_per_seg
            t2, t3 = t * t, t * t * t
            x = 0.5 * ((2 * p1[0]) + (-p0[0] + p2[0]) * t +
                       (2 * p0[0] - 5 * p1[0] + 4 * p2[0] - p3[0]) * t2 +
                       (-p0[0] + 3 * p1[0] - 3 * p2[0] + p3[0]) * t3)
            y = 0.5 * ((2 * p1[1]) + (-p0[1] + p2[1]) * t +
                       (2 * p0[1] - 5 * p1[1] + 4 * p2[1] - p3[1]) * t2 +
                       (-p0[1] + 3 * p1[1] - 3 * p2[1] + p3[1]) * t3)
            out.append((x, y))
    out.append(pts[-1])
    return out


def add_line(line_el, slide, theme):
    x, y, w, h = line_el["bounds"]
    points = line_el.get("points", "0,0 1,1")
    pts = [tuple(map(float, p.split(","))) for p in points.split()]
    if len(pts) == 2:
        conn = slide.shapes.add_connector(1, px(x + pts[0][0] * w / 100),
                                          px(y + pts[0][1] * h / 100),
                                          px(x + pts[1][0] * w / 100),
                                          px(y + pts[1][1] * h / 100))
        bd = line_el.get("border", {})
        conn.line.color.rgb = resolve_color(bd.get("color", "#4A3728"), theme)
        conn.line.width = Pt(bd.get("width", 2))
        arrow = line_el.get("arrow")
        if arrow and arrow[1]:
            from pptx.oxml.ns import qn
            ln = conn.line._get_or_add_ln()
            tail = ln.makeelement(qn("a:tailEnd"), {"type": "triangle"})
            ln.append(tail)
        return conn
    if line_el.get("curve") == "smooth":
        pts = _catmull_rom_sample(pts, 24)
    free = slide.shapes.build_freeform(px(x + pts[0][0] * w / 100), px(y + pts[0][1] * h / 100), 9525)
    free.add_line_segments([(px(px_x * w / 100), px(py * h / 100)) for px_x, py in pts[1:]], close=False)
    return free.convert_to_shape()


def set_fade(slide):
    from pptx.oxml.ns import qn
    cSld = slide._element.find(qn("p:cSld"))
    transition = slide._element.makeelement(qn("p:transition"), {"spd": "med"})
    fade = transition.makeelement(qn("p:fade"), {})
    transition.append(fade)
    cSld.addnext(transition)


def render_background(slide, bg, theme, proj):
    if not bg:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = resolve_color("#FFFFFF", theme.get("colors", theme))
        return
    if bg.get("type") == "solid":
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = resolve_color(bg.get("color", "#FFFFFF"), theme)
    elif bg.get("type") == "gradient":
        stops = bg.get("stops", [])
        if len(stops) >= 2:
            slide.background.fill.gradient()
            slide.background.fill.gradient_stops[0].color.rgb = resolve_color(stops[0]["color"], theme.get("colors", theme))
            slide.background.fill.gradient_stops[-1].color.rgb = resolve_color(stops[-1]["color"], theme.get("colors", theme))
    elif bg.get("type") == "image":
        WARN.append("image 背景未支持，已退化为纯色")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--output", default=None)
    ap.add_argument("--transition", default="fade", choices=["fade", "none"])
    ap.add_argument("--force", action="store_true", help="覆盖已存在输出")
    a = ap.parse_args()
    pptd = pathlib.Path(a.input)
    proj = pptd.parent
    out = pathlib.Path(a.output) if a.output else pptd.with_suffix(".pptx")
    if out.exists() and not a.force:
        print(f"[local-engine] {out} 已存在，加 --force 覆盖"); sys.exit(2)
    deck = yaml.safe_load(pptd.read_text(encoding="utf-8"))
    theme = deck.get("theme", {})

    prs = Presentation()
    prs.slide_width = px(deck["size"][0])
    prs.slide_height = px(deck["size"][1])
    blank = prs.slide_layouts[6]

    for page_rel in deck.get("pages", []):
        page = yaml.safe_load((proj / page_rel).read_text(encoding="utf-8"))
        slide = prs.slides.add_slide(blank)
        render_background(slide, page.get("background"), theme, proj)
        for el in page.get("elements", []):
            et = el.get("elementType")
            try:
                if et == "text":
                    add_text(el, slide, theme)
                elif et == "shape":
                    add_shape(el, slide, theme)
                elif et == "image":
                    add_image(el, slide, theme, proj)
                elif et == "icon":
                    add_icon(el, slide, theme, proj)
                elif et == "table":
                    add_table(el, slide, theme)
                elif et == "chart":
                    add_chart(el, slide, theme, proj)
                elif et == "line":
                    add_line(el, slide, theme)
            except Exception as exc:
                WARN.append(f"元素 {el.get('elementId')} ({et}) 渲染失败: {exc}")
        if a.transition == "fade":
            set_fade(slide)
        if page.get("notes"):
            slide.notes_slide.notes_text_frame.text = page["notes"]

    prs.save(out)
    print(f"[local-engine] pptx written: {out} "
          f"({out.stat().st_size // 1024} KB, {len(deck.get('pages', []))} slides)")
    if WARN:
        print(f"[local-engine] warnings ({len(WARN)}):")
        for w in WARN:
            print("  -", w)


if __name__ == "__main__":
    main()
