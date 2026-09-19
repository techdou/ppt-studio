#!/usr/bin/env python3
"""讲义/教材 MD → 课件 PPTX（批量）。

设计动机：讲义应当是图文交叉的正式文档（图片跟着内容走），
不需要为做 PPT 在文末另写"页面规划表"——本脚本直接从讲义结构生成课件：
H1 → 封面页、学习目标 → 目标页、每个 H2 节 → 一页内容页（配图就近取节内第一张图）、
术语卡 → 术语页，不足 8 页自动补"课堂要点回顾"。

排版要点（实测踩坑沉淀）：
- 行预算截断 fit_text：PowerPoint 实际换行比"字数/行宽"理论值密约 15%，
  参数已留余量；截断统一收在句末，杜绝半句。
- 有配图的页正文占左半区（430px），无图占全宽（840px）。
- 页码框宽度 ≥76px，否则两位数页码会被竖排拆开。
- 教师参考节（课前准备/流程表/兜底）与引用块（> 教师引导）默认不上课件。

用法:
  单文件: python handout_to_pptx.py --input 讲义.md --out 第1课.pptx
  批量:   python handout_to_pptx.py --input-dir <根目录> --out-dir <输出根>
          （根目录下每个含 第*课讲义.md / 第*课.md 的数字子目录算一课，
            输出镜像子目录结构，文件名 第NN课-课名.pptx）
  主题:   --theme theme.json（缺省为奶油浅底橙青配色）
"""
import argparse
import json
import pathlib
import re
import shutil
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

DEFAULT_THEME = {
    "primary": "#FF8A3D", "primaryDeep": "#D96A1F", "accent": "#1B7A70",
    "bg": "#FFF8EF", "brown": "#4A3728", "text": "#3D3A36",
    "muted": "#8A7E70", "soft": "#FFE8D6", "font": "微软雅黑",
}
DEFAULT_SKIP = ["课前准备", "流程表", "万一不行", "教师兜底", "补充图库", "页面规划"]
UNSAFE = re.compile(r'[\\/:*?"<>|\s]+')


def load_theme(path):
    theme = dict(DEFAULT_THEME)
    if path:
        theme.update(json.loads(pathlib.Path(path).read_text(encoding="utf-8")))
    return theme


class Deck:
    def __init__(self, theme):
        self.t = theme
        self.prs = Presentation()
        self.prs.slide_width = Emu(int(960 * 9525))
        self.prs.slide_height = Emu(int(540 * 9525))
        self.blank = self.prs.slide_layouts[6]

    def _c(self, h):
        h = str(h).lstrip("#")
        while len(h) < 6:
            h += "0"
        try:
            return RGBColor(int(h[:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        except ValueError:
            return RGBColor(0, 0, 0)

    def slide(self):
        s = self.prs.slides.add_slide(self.blank)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = self._c(self.t["bg"])
        return s

    def bar(self, slide, bounds, color):
        st = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Emu(int(bounds[0] * 9525)), Emu(int(bounds[1] * 9525)),
            Emu(int(bounds[2] * 9525)), Emu(int(bounds[3] * 9525)))
        st.fill.solid()
        st.fill.fore_color.rgb = self._c(color)
        st.line.fill.background()
        st.shadow.inherit = False
        return st

    def text(self, slide, bounds, content, size=16, color=None, bold=False,
             align=PP_ALIGN.LEFT):
        tb = slide.shapes.add_textbox(
            Emu(int(bounds[0] * 9525)), Emu(int(bounds[1] * 9525)),
            Emu(int(bounds[2] * 9525)), Emu(int(bounds[3] * 9525)))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        for i, line in enumerate(content.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = 1.5
            r = p.add_run()
            r.text = line
            r.font.size = Pt(size)
            r.font.name = self.t["font"]
            r.font.color.rgb = self._c(color or self.t["text"])
            r.font.bold = bold
        return tb

    def picture(self, slide, src, bounds):
        if pathlib.Path(src).exists():
            slide.shapes.add_picture(
                str(src), Emu(int(bounds[0] * 9525)), Emu(int(bounds[1] * 9525)),
                Emu(int(bounds[2] * 9525)), Emu(int(bounds[3] * 9525)))

    def header(self, slide, title, page_no):
        self.bar(slide, [40, 28, 620, 54], self.t["primary"])
        self.text(slide, [56, 32, 590, 46], title, 22, "#FFFFFF", bold=True)
        self.text(slide, [854, 30, 76, 40], f"{page_no:02d}", 26,
                  self.t["primaryDeep"], bold=True)

    def footer(self, slide, label):
        self.text(slide, [40, 512, 500, 22], label, 11,
                  self.t["muted"])

    def save(self, path, title, author="课件"):
        self.prs.core_properties.title = title
        self.prs.core_properties.author = author
        self.prs.save(str(path))


def fit_text(text, chars_per_line, max_lines):
    """按行预算截断：段落超出容量时在句末收尾，杜绝半句溢出。"""
    lines, out = 0, []
    for para in text.split("\n"):
        n = max(1, -(-len(para) // chars_per_line))
        if lines + n > max_lines:
            remain = max_lines - lines
            if remain > 0:
                take = para[:remain * chars_per_line]
                for i in range(len(take) - 1, max(0, len(take) - 80), -1):
                    if take[i] in "。！？；：，——":
                        take = take[:i + 1]
                        break
                out.append(take)
            break
        out.append(para)
        lines += n
    return "\n".join(out)


def parse_handout(md_path, skip_heads):
    text = md_path.read_text(encoding="utf-8")
    title_m = re.search(r"^#\s*(.+)$", text, re.M)
    title = title_m.group(1).strip() if title_m else md_path.stem
    goals = re.findall(r"^- (.+)$", text, re.M)[:3]
    # H1 之后、第一个 H2 之前的首图作封面题图
    before_h2 = text.split("## ", 1)[0]
    cover_m = re.search(r"!\[[^\]]*\]\((assets/images/[^)]+)\)", before_h2)
    cover_img = str(md_path.parent / cover_m.group(1)) if cover_m else None

    secs = []
    for chunk in re.split(r"^## ", text, flags=re.M)[1:]:
        lines = chunk.split("\n")
        head = lines[0].strip()
        if any(k in head for k in skip_heads):
            continue
        body_lines, img = [], None
        img_m = re.search(r"!\[[^\]]*\]\((assets/images/[^)]+)\)", chunk)
        if img_m:
            img = str(md_path.parent / img_m.group(1))
        for l in lines[1:]:
            l = l.strip()
            if not l or l.startswith(("![", "|", "---", ">", "*")):
                continue
            body_lines.append(re.sub(r"\*\*(.+?)\*\*", r"\1", l))
        body = "\n".join(body_lines[:16])
        if body.strip():
            secs.append({"head": re.sub(r"^[^\u4e00-\u9fa5A-Za-z0-9]+", "", head),
                         "text": body, "img": img})
    terms = []
    term_m = re.search(r"##[^\n]*术语卡(.*?)(?:\n## |\Z)", text, re.S)
    if term_m:
        for row in re.findall(r"^\|([^|]+)\|([^|]+)\|", term_m.group(1), re.M):
            t, d = row[0].strip(), row[1].strip()
            if t and t not in ("术语", ":---", "---") and not set(t) <= set("-: "):
                terms.append((t, d))
    return {"title": title, "goals": goals, "cover_img": cover_img,
            "secs": secs, "terms": terms}


def process(md_path, theme, out_dir, meta_label, author, min_pages=8):
    lesson = parse_handout(md_path, DEFAULT_SKIP)
    no_m = re.search(r"(\d+)", md_path.stem)
    no = int(no_m.group(1)) if no_m else 0
    name_m = re.match(r"第\s*\d+\s*课\s*(.+)", lesson["title"])
    lesson_name = name_m.group(1).strip() if name_m else lesson["title"]
    fname = f"第{no:02d}课-{UNSAFE.sub('', lesson_name)}.pptx"

    deck = Deck(theme)
    s = deck.slide()
    deck.bar(s, [0, 0, 960, 12], theme["primary"])
    deck.bar(s, [56, 96, 218, 40], theme["accent"])
    deck.text(s, [60, 100, 210, 32], meta_label, 15, "#FFFFFF", bold=True)
    deck.text(s, [52, 156, 430, 150], lesson["title"], 40, theme["brown"], bold=True)
    if lesson["cover_img"]:
        deck.picture(s, lesson["cover_img"], [508, 84, 412, 412])
    deck.footer(s, meta_label)
    s = deck.slide()
    deck.header(s, "这节课，我们要……", 2)
    for i, g in enumerate(lesson["goals"][:3]):
        y = 130 + i * 110
        deck.bar(s, [70, y, 44, 44], theme["accent"])
        deck.text(s, [74, y + 4, 36, 36], str(i + 1), 20, "#FFFFFF", bold=True)
        deck.text(s, [128, y, 700, 100], g, 17)
    page = 3
    for sec in lesson["secs"]:
        s = deck.slide()
        deck.header(s, sec["head"], page)
        has_img = bool(sec["img"])
        body = fit_text(sec["text"], 22 if has_img else 32, 15 if has_img else 12)
        if body:
            deck.text(s, [60, 100, 430 if has_img else 840, 380], body,
                      12 if has_img else 14)
        if has_img:
            deck.picture(s, sec["img"], [520, 130, 400, 300])
        page += 1
    if lesson["terms"]:
        s = deck.slide()
        deck.header(s, "本课术语卡", page)
        for i, (t, d) in enumerate(lesson["terms"][:6]):
            y = 110 + i * 55
            deck.bar(s, [60, y, 840, 46], theme["soft"])
            deck.text(s, [76, y + 4, 200, 38], t, 17, theme["brown"], bold=True)
            deck.text(s, [280, y + 4, 600, 38], d, 14, theme["muted"])
    while len(deck.prs.slides) < min_pages:
        s = deck.slide()
        deck.bar(s, [40, 28, 620, 54], theme["accent"])
        deck.text(s, [56, 32, 590, 46], "课堂要点回顾", 22, "#FFFFFF", bold=True)
        deck.text(s, [854, 30, 76, 40], f"{len(deck.prs.slides)+1:02d}", 26,
                  theme["primaryDeep"], bold=True)
        for i, (t, d) in enumerate(lesson["terms"][:4]):
            deck.text(s, [80, 110 + i * 80, 800, 70], f"• {t}: {d}", 16)
    out = out_dir / fname
    deck.save(out, f"第{no}课 {lesson_name}", author)
    return out, len(deck.prs.slides)


def main():
    ap = argparse.ArgumentParser(description="讲义 MD → 课件 PPTX")
    ap.add_argument("--input", help="单个讲义 md")
    ap.add_argument("--out", help="单文件输出 pptx 路径")
    ap.add_argument("--input-dir", help="批量模式：讲义根目录")
    ap.add_argument("--out-dir", help="批量模式：输出根目录")
    ap.add_argument("--theme", help="主题 JSON（缺省奶油浅底橙青）")
    ap.add_argument("--label", default="", help="封面标签（如 小学 · 40 分钟）")
    ap.add_argument("--author", default="课件", help="PPTX 作者元数据")
    a = ap.parse_args()
    theme = load_theme(a.theme)

    if a.input:
        md = pathlib.Path(a.input)
        out = pathlib.Path(a.out) if a.out else md.parent / (md.stem + ".pptx")
        out.parent.mkdir(parents=True, exist_ok=True)
        path, n = process(md, theme, out.parent, a.label, a.author)
        if path.resolve() != out.resolve():
            path.replace(out)
        print(f"[ok] {out} ({n} pages)")
        return
    if a.input_dir:
        root, out_root = pathlib.Path(a.input_dir), pathlib.Path(a.out_dir or (a.input_dir + "-pptx"))
        count = 0
        for sub in sorted(p for p in root.iterdir() if p.is_dir()):
            mds = list(sub.glob("第*课*讲义.md")) or list(sub.glob("第*课.md"))
            if not mds:
                continue
            dest = out_root / sub.name
            dest.mkdir(parents=True, exist_ok=True)
            # 讲义图片相对 assets/ 解析，批量模式在工作目录就近出图
            path, n = process(mds[0], theme, dest, a.label, a.author)
            count += 1
            print(f"[deck] {sub.name} -> {path.name} ({n} pages)")
        print(f"total {count} decks -> {out_root}")
        return
    ap.error("需要 --input 或 --input-dir")


if __name__ == "__main__":
    main()
