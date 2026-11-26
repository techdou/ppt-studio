#!/usr/bin/env python3
"""导出 PPT 每页图片 + 拼接总览图（本机渲染，无网络依赖）。

依赖本机已安装 Microsoft PowerPoint（COM 自动化）。输出结构：
  <output>/pages/1.png ... N.png     每页 PNG（1280x720）
  <output>/overview.jpg              全页拼接总览（供视觉 QA）

用法:
  python export_images.py <deck.pptx | deck.pptd> [--output DIR] [--force]
输入为 .pptd 时，先用本地引擎渲染出临时 pptx。
"""
import argparse
import pathlib
import re
import shutil
import subprocess
import sys

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))


def ensure_pptx(input_path: pathlib.Path, workdir: pathlib.Path) -> pathlib.Path:
    if input_path.suffix == ".pptx":
        return input_path
    import export_pptx_local
    out = workdir / (input_path.stem + ".pptx")
    sys.argv = ["export_pptx_local.py", str(input_path), "--output", str(out), "--force"]
    export_pptx_local.main()
    return out


def export_pngs_via_com(pptx: pathlib.Path, png_dir: pathlib.Path, width=1280, height=720):
    png_dir.mkdir(parents=True, exist_ok=True)
    ps = f"""
$ErrorActionPreference = 'Stop'
$pp = New-Object -ComObject PowerPoint.Application
# Open an untitled copy; do not quit PowerPoint because COM can reuse the user's instance.
$pres = $pp.Presentations.Open('{str(pptx)}', $true, $true, $false)
try {{
    $pres.Export('{str(png_dir)}', 'PNG', {width}, {height})
}} finally {{
    $pres.Close()
}}
Write-Output exported
"""
    r = subprocess.run(["powershell", "-Command", ps], capture_output=True, text=True,
                       timeout=300)
    if "exported" not in r.stdout:
        raise RuntimeError(f"COM 导出失败: {r.stderr[:300]}")
    # PowerPoint outputs SlideN.png or 幻灯片N.PNG. Sort by slide number, not name.
    files = []
    for path in png_dir.iterdir():
        match = re.fullmatch(r"(?:Slide|幻灯片)([0-9]+)\.png", path.name, re.IGNORECASE)
        if match and path.is_file():
            files.append((int(match[1]), path))
    renamed = []
    for i, (_, f) in enumerate(sorted(files), 1):
        dst = png_dir / f"{i}.png"
        shutil.move(str(f), str(dst))
        renamed.append(dst)
    return renamed


def stitch_overview(pngs, overview_path, cols=2, cell_w=960):
    from PIL import Image
    thumbs = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        ratio = cell_w / im.width
        thumbs.append(im.resize((cell_w, int(im.height * ratio)), Image.LANCZOS))
    cell_h = max(t.height for t in thumbs)
    rows = (len(thumbs) + cols - 1) // cols
    board = Image.new("RGB", (cols * cell_w + (cols + 1) * 16,
                              rows * cell_h + (rows + 1) * 16), "#FFFFFF")
    for i, t in enumerate(thumbs):
        x = 16 + (i % cols) * (cell_w + 16)
        y = 16 + (i // cols) * (cell_h + 16)
        board.paste(t, (x, y))
    board.save(overview_path, quality=85, optimize=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help=".pptx 或 .pptd 项目")
    ap.add_argument("--output", "-o", default=None)
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    input_path = pathlib.Path(args.input).resolve()
    out_dir = pathlib.Path(args.output).resolve() if args.output else \
        (input_path.parent if input_path.suffix == ".pptx" else input_path) / ".qa-images"
    if out_dir.exists() and not args.force:
        print(f"[local-engine] {out_dir} 已存在，加 --force 覆盖")
        sys.exit(2)

    workdir = out_dir / ".build"
    workdir.mkdir(parents=True, exist_ok=True)
    pptx = ensure_pptx(input_path, workdir)

    pages_dir = out_dir / "pages"
    if pages_dir.exists():
        shutil.rmtree(pages_dir)
    png_dir = workdir / "_png"
    pngs = export_pngs_via_com(pptx, png_dir)
    pages_dir.mkdir(parents=True, exist_ok=True)
    final = []
    for i, src in enumerate(pngs, 1):
        dst = pages_dir / f"{i}.png"
        shutil.copy2(src, dst)
        final.append(dst)

    overview = out_dir / "overview.jpg"
    stitch_overview(final, overview)
    shutil.rmtree(workdir, ignore_errors=True)
    print(f"[local-engine] exported {len(final)} pages -> {out_dir}")
    print(f"[local-engine] overview: {overview}")


if __name__ == "__main__":
    main()
