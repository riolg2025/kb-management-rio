#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""inventory — 旧 PPTX 全量盘点（ppt-retrofit Phase 0）

提取：逐页文本（形状递归+表格+graphicFrame/SmartArt——SmartArt 文字不进普通文本
提取，判"空页"必须看 smartart 字段）、图片清单（格式/像素/sha1/字节数）、
主题色、字体、逐页形状与连接线统计。

用法：
  python3 -I inventory.py --pptx 原稿.pptx --out-json inventory.json \
      [--slides 1-55] [--deep]
"""
import argparse, hashlib, io, json, re, sys

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # GROUP
            yield from iter_shapes(sh.shapes)


def slide_texts(slide):
    """普通文本：文本框/形状/表格（组合递归）"""
    buf = []
    for sh in iter_shapes(slide.shapes):
        try:
            if sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        if c.text.strip():
                            buf.append(c.text)
                continue
        except Exception:
            pass
        try:
            t = sh.text_frame.text
            if t.strip():
                buf.append(t)
        except Exception:
            pass
    return buf


def diagram_texts(slide):
    """SmartArt/图表等 graphicFrame 关联部件的文字（diagramData/chart parts）
    —— 这是判"空页"的生死线：SmartArt 文字不进普通文本提取。"""
    out = []
    try:
        rels = slide.part.rels
    except Exception:
        return out
    for rid, rel in rels.items():
        rt = getattr(rel, "reltype", "") or ""
        if not ("diagramData" in rt or rt.endswith("/chart")):
            continue
        try:
            root = etree.fromstring(rel.target_part.blob)
            ts = [t.text for t in root.iter(qn("a:t")) if t.text and t.text.strip()]
            if ts:
                out.append(" ".join(ts))
        except Exception:
            continue
    return out


def slide_images(slide):
    out = []
    seen_in_slide = set()
    for sh in iter_shapes(slide.shapes):
        if sh.shape_type != 13:  # PICTURE
            continue
        try:
            img = sh.image
        except Exception:
            continue
        sha = hashlib.sha1(img.blob).hexdigest()
        px = None
        try:
            from PIL import Image
            px = Image.open(io.BytesIO(img.blob)).size
        except Exception:
            pass
        out.append({"name": sh.name, "ext": img.ext, "px": px,
                    "bytes": len(img.blob), "sha1": sha,
                    "dup_in_slide": sha in seen_in_slide})
        seen_in_slide.add(sha)
    return out


def theme_info(prs):
    colors, fonts = {}, {}
    try:
        master = prs.slide_masters[0]
        theme_part = None
        for rel in master.part.rels.values():
            if (getattr(rel, "reltype", "") or "").endswith("/theme"):
                theme_part = rel.target_part
                break
        if theme_part is not None:
            root = etree.fromstring(theme_part.blob)
            scheme = root.find(f".//{{{NS_A}}}clrScheme")
            if scheme is not None:
                for child in scheme:
                    tag = etree.QName(child).localname
                    c = child.find(f"{{{NS_A}}}srgbClr")
                    s = child.find(f"{{{NS_A}}}sysClr")
                    colors[tag] = c.get("val") if c is not None else (
                        s.get("lastClr") if s is not None else None)
            for kind in ("majorFont", "minorFont"):
                f = root.find(f".//{{{NS_A}}}{kind}")
                if f is not None:
                    ea = f.find(f"{{{NS_A}}}ea")
                    lat = f.find(f"{{{NS_A}}}latin")
                    fonts[kind] = {"latin": lat.get("typeface") if lat is not None else None,
                                   "ea": ea.get("typeface") if ea is not None else None}
    except Exception:
        pass
    return colors, fonts


def parse_range(spec, total):
    if not spec:
        return range(1, total + 1)
    m = re.match(r"^(\d+)-(\d+)$", spec)
    if m:
        return range(int(m.group(1)), int(m.group(2)) + 1)
    return [int(x) for x in spec.split(",")]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--out-json", default=None)
    ap.add_argument("--slides", default=None, help="如 1-55 或 3,7,46")
    ap.add_argument("--deep", action="store_true", help="逐形状清单（名称/类型/框/文本头）")
    a = ap.parse_args()

    prs = Presentation(a.pptx)
    total = len(prs.slides._sldIdLst)
    colors, fonts = theme_info(prs)
    W, H = prs.slide_width / 914400, prs.slide_height / 914400

    result = {"file": a.pptx, "slides_total": total,
              "canvas_inch": [round(W, 3), round(H, 3)],
              "theme": {"colors": colors, "fonts": fonts},
              "slides": [], "dedup_candidates": {}}
    sha_slides = {}

    for no in parse_range(a.slides, total):
        sl = prs.slides[no - 1]
        texts = slide_texts(sl)
        dgm = diagram_texts(sl)
        imgs = slide_images(sl)
        n_cxn = 0
        for el in sl.shapes._spTree.iter():
            if el.tag == qn("p:cxnSp"):
                n_cxn += 1
        entry = {"no": no,
                 "text_count": len(texts),
                 "texts": texts,
                 "smartart_or_chart": dgm,
                 "images": imgs,
                 "shape_count": len(sl.shapes),
                 "connectors": n_cxn,
                 "empty_suspect": not (texts or dgm or imgs)}
        if a.deep:
            entry["shapes"] = [
                {"name": sh.name, "type": str(sh.shape_type),
                 "box": [round(sh.left / 914400, 2), round(sh.top / 914400, 2),
                         round(sh.width / 914400, 2), round(sh.height / 914400, 2)],
                 "text": (sh.text_frame.text or "")[:40]
                         if hasattr(sh, "text_frame") else ""}
                for sh in sl.shapes]
        result["slides"].append(entry)
        for im in imgs:
            sha_slides.setdefault(im["sha1"], []).append(no)
    result["dedup_candidates"] = {k: v for k, v in sha_slides.items() if len(v) > 1}

    js = json.dumps(result, ensure_ascii=False, indent=1)
    if a.out_json:
        with open(a.out_json, "w", encoding="utf-8") as f:
            f.write(js)
        print(f"已写出 {a.out_json}")

    for e in result["slides"]:
        flag = " ⚠️疑似空页(查SmartArt后再判)" if e["empty_suspect"] else ""
        dgm = f" SmartArt/图表:{len(e['smartart_or_chart'])}段" if e["smartart_or_chart"] else ""
        print(f"P{e['no']:02d} 文本{e['text_count']:3d} 图{len(e['images']):2d}"
              f" 形状{e['shape_count']:3d} 连线{e['connectors']:2d}{dgm}{flag}")
    if result["dedup_candidates"]:
        print("⚠ 同文件多引用（去重候选，裁剪变体复用不重复入库）:",
              {k[:8]: v for k, v in result["dedup_candidates"].items()})


if __name__ == "__main__":
    main()
