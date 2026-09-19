#!/usr/bin/env python3
"""PPTD 文本元素读写工具（供本地查看器 ui/server.mjs 调用，也可手动使用）。

用法:
  python page_text_tool.py texts <deck.pptd> [--page N]
      列出各页（或指定页）的 text 元素，JSON 输出:
      {"deck": "...", "size": [960,540], "title": "...",
       "pages": [{"page": 1, "file": "pages/1.page", "pageType": "cover",
                  "texts": [{"elementId": "t1", "bounds": [..], "preview": "...", "rich": false}]}]}

  python page_text_tool.py update <deck.pptd> --page N --element ID --text-file F
      将指定 text 元素的 content.text 替换为文件 F 的内容（保留元素其他字段）。
      F 内容原样写入，可含富文本标签（<p>/<span>/<strong> 等）。

依赖: PyYAML（与导出引擎同一环境）。
"""
import argparse
import html
import json
import pathlib
import re
import sys

import yaml


class _LiteralDumper(yaml.SafeDumper):
    pass


def _str_representer(dumper, data):
    if "\n" in data:
        # block scalar 风格更贴近 PPTD 规范示例，且 diff 友好
        return dumper.represent_scalar("tag:yaml.org,2002:str", data, style="|")
    return dumper.represent_scalar("tag:yaml.org,2002:str", data)


_LiteralDumper.add_representer(str, _str_representer)


def _load_deck(deck_path: pathlib.Path):
    manifest = yaml.safe_load(deck_path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict) or "pages" not in manifest:
        raise SystemExit(f"无效的 PPTD manifest: {deck_path}")
    return manifest


def _page_path(deck_path: pathlib.Path, rel: str) -> pathlib.Path:
    p = (deck_path.parent / rel).resolve()
    if not p.is_file():
        raise SystemExit(f"页面文件不存在: {rel}")
    return p


def _strip_tags(text: str) -> str:
    plain = re.sub(r"<[^>]+>", "", text or "")
    return re.sub(r"\s+", " ", html.unescape(plain)).strip()


def cmd_texts(deck_path: pathlib.Path, page_filter=None) -> None:
    manifest = _load_deck(deck_path)
    out = {
        "deck": str(deck_path),
        "title": manifest.get("title"),
        "size": manifest.get("size", [960, 540]),
        "pages": [],
    }
    for idx, rel in enumerate(manifest["pages"], 1):
        if page_filter is not None and idx != page_filter:
            continue
        data = yaml.safe_load(_page_path(deck_path, rel).read_text(encoding="utf-8"))
        texts = []
        for el in data.get("elements", []):
            if el.get("elementType") != "text":
                continue
            raw = str((el.get("content") or {}).get("text") or "")
            texts.append({
                "elementId": el.get("elementId"),
                "bounds": el.get("bounds"),
                "preview": _strip_tags(raw)[:120],
                "rawText": raw,
                "rich": "<" in raw,
            })
        out["pages"].append({
            "page": idx, "file": rel,
            "pageType": data.get("pageType"), "texts": texts,
        })
    print(json.dumps(out, ensure_ascii=False))


def cmd_update(deck_path: pathlib.Path, page_no: int, element_id: str, text_file: pathlib.Path) -> None:
    manifest = _load_deck(deck_path)
    if not (1 <= page_no <= len(manifest["pages"])):
        raise SystemExit(f"页码越界: {page_no}（共 {len(manifest['pages'])} 页）")
    rel = manifest["pages"][page_no - 1]
    page_path = _page_path(deck_path, rel)
    data = yaml.safe_load(page_path.read_text(encoding="utf-8"))
    hits = [el for el in data.get("elements", [])
            if el.get("elementType") == "text" and el.get("elementId") == element_id]
    if len(hits) != 1:
        raise SystemExit(f"第 {page_no} 页未找到唯一 text 元素 elementId={element_id!r}（匹配 {len(hits)} 个）")
    new_text = text_file.read_text(encoding="utf-8")
    hits[0]["content"]["text"] = new_text
    page_path.write_text(
        yaml.dump(data, Dumper=_LiteralDumper, allow_unicode=True,
                  sort_keys=False, width=4096),
        encoding="utf-8")
    print(json.dumps({"ok": True, "page": page_no, "elementId": element_id,
                      "file": str(page_path)}, ensure_ascii=False))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("texts", help="列出文本元素")
    p1.add_argument("deck", help=".pptd manifest 路径")
    p1.add_argument("--page", type=int, default=None)

    p2 = sub.add_parser("update", help="替换文本内容")
    p2.add_argument("deck", help=".pptd manifest 路径")
    p2.add_argument("--page", type=int, required=True)
    p2.add_argument("--element", required=True, help="elementId")
    p2.add_argument("--text-file", required=True, help="新文本内容文件（UTF-8）")

    args = ap.parse_args()
    deck_path = pathlib.Path(args.deck).resolve()
    if args.cmd == "texts":
        cmd_texts(deck_path, args.page)
    else:
        cmd_update(deck_path, args.page, args.element, pathlib.Path(args.text_file).resolve())


if __name__ == "__main__":
    main()
