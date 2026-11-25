#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""structure_diff — 照画页结构对账 / 原稿图 OCR 对账（ppt-retrofit 自查五件套 #2#3#4）

mode slide（结构对账）：
  python3 -I structure_diff.py slide --src 原稿.pptx --src-slide 23 \
      --dst 成品.pptx --dst-slide 23 [--gap 0.45] [--json]
  提取双方：形状总数、连接线数、y-带聚类（带=纵向间距>gap）、每带形状数、
  文本清单；报告：块数差异、每带差异、原稿文本在成品中未命中的清单。

mode ocr（原稿图 OCR 对账，五件套 #4）：
  python3 -I structure_diff.py ocr --image p031_1.png \
      --dst 成品.pptx --dst-slide 31 [--lang chi_sim+eng]
  原稿图 OCR 提词（len≥2），逐词查成品页规范化全文，报告未命中词。
  需 pytesseract + tesseract（含 chi_sim），缺失时给出安装提示后退出。
"""
import argparse
import hashlib
import re
import sys

from pptx import Presentation
from pptx.oxml.ns import qn


def norm(t):
    return re.sub(r"[\s\*｜→–—·+×&><\/、，。：:；;()（）\"“”\"'‘’'【】\[\]\-_#%.]+",
                  "", t).lower()


def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:
            yield from iter_shapes(sh.shapes)


def is_bg(sh, CW, CH):
    try:
        w, h = sh.width / 914400, sh.height / 914400
        return w * h > 0.8 * CW * CH
    except Exception:
        return False


def slide_profile(slide):
    buf_texts, tops, n_blocks = [], [], 0
    for sh in iter_shapes(slide.shapes):
        if sh.name and sh.name.startswith("decor"):
            continue
        n_blocks += 1
        try:
            t = sh.text_frame.text.strip()
            if t:
                buf_texts.append(t)
        except Exception:
            pass
        try:
            if not is_bg(sh, 13.34, 7.51):
                tops.append(round(sh.top / 914400, 2))
        except Exception:
            pass
    n_cxn = sum(1 for el in slide.shapes._spTree.iter() if el.tag == qn("p:cxnSp"))
    return {"blocks": n_blocks, "connectors": n_cxn,
            "texts": buf_texts, "tops": sorted(tops)}


def cluster_bands(tops, gap):
    if not tops:
        return []
    bands, cur = [], [tops[0]]
    for t in tops[1:]:
        if t - cur[-1] > gap:
            bands.append(cur)
            cur = [t]
        else:
            cur.append(t)
    bands.append(cur)
    return bands


def mode_slide(a):
    src = Presentation(a.src).slides[a.src_slide - 1]
    dst = Presentation(a.dst).slides[a.dst_slide - 1]
    ps, pd = slide_profile(src), slide_profile(dst)
    bs = cluster_bands(ps["tops"], a.gap)
    bd = cluster_bands(pd["tops"], a.gap)

    print(f"== 结构对账 原稿s{a.src_slide} → 成品s{a.dst_slide} ==")
    print(f"块数: 原稿 {ps['blocks']} / 成品 {pd['blocks']}"
          + ("" if ps['blocks'] == pd['blocks'] else "  ⚠ 不一致"))
    print(f"连接线: 原稿 {ps['connectors']} / 成品 {pd['connectors']}"
          + ("" if ps['connectors'] == pd['connectors'] else "  ⚠ 不一致"))
    print(f"分带: 原稿 {len(bs)} {['%d块' % len(b) for b in bs]}")
    print(f"      成品 {len(bd)} {['%d块' % len(b) for b in bd]}")
    if len(bs) != len(bd):
        print("  ⚠ 分带数不一致（布局重排可解释则忽略，照画页应一致）")
    else:
        for i, (x, y) in enumerate(zip(bs, bd)):
            if len(x) != len(y):
                print(f"  ⚠ 带{i+1}: 原稿 {len(x)} 块 / 成品 {len(y)} 块")

    dst_all = norm("".join(pd["texts"]))
    src_all = norm("".join(ps["texts"]))
    miss = []
    for t in ps["texts"]:
        for seg in re.split(r"[\n｜，。：:；;！？!?]+", t):
            n = norm(seg)
            if len(n) >= 2 and n not in dst_all:
                miss.append(seg.strip()[:38])
    # 去重保序
    seen, uniq = set(), []
    for m_ in miss:
        if m_ not in seen:
            seen.add(m_)
            uniq.append(m_)
    print(f"原稿文本未命中成品: {len(uniq)}")
    for m_ in uniq[:40]:
        print(f"    - {m_}")
    extra = 0
    for t in pd["texts"]:
        for seg in re.split(r"[\n｜，。：:；;！？!?]+", t):
            n = norm(seg)
            if len(n) >= 4 and n not in src_all:
                extra += 1
    print(f"成品新增文本(导语/图注类，信息性): {extra}")
    if a.json:
        import json
        print(json.dumps({"src": {k: v for k, v in ps.items() if k != 'texts'},
                          "dst": {k: v for k, v in pd.items() if k != 'texts'},
                          "bands_src": [len(b) for b in bs],
                          "bands_dst": [len(b) for b in bd],
                          "missing_texts": uniq},
                         ensure_ascii=False))


def mode_ocr(a):
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        sys.exit("需要 pytesseract + Pillow（apt 装 tesseract-ocr 与中文包 chi-sim）")
    words = pytesseract.image_to_string(Image.open(a.image), lang=a.lang)
    prs = Presentation(a.dst)
    buf = []

    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                walk(sh.shapes)
                continue
            try:
                if sh.has_table:
                    for r in sh.table.rows:
                        for c in r.cells:
                            buf.append(c.text)
                    continue
            except Exception:
                pass
            try:
                if sh.text_frame.text:
                    buf.append(sh.text_frame.text)
            except Exception:
                pass

    walk(prs.slides[a.dst_slide - 1].shapes)
    dst_all = norm("".join(buf))
    seen, miss, total = set(), [], 0
    for w in re.split(r"[\s\n]+", words):
        w = w.strip()
        n = norm(w)
        if len(n) < 2:
            continue
        total += 1
        if n in seen:
            continue
        seen.add(n)
        if n not in dst_all:
            miss.append(w[:30])
    print(f"== OCR 对账 {a.image} → 成品s{a.dst_slide} ==")
    print(f"OCR 词数: {total}  去重: {len(seen)}  未命中: {len(miss)}")
    for m_ in miss[:50]:
        print(f"    - {m_}")
    if miss:
        print("⚠ 未命中逐条人工核：真缺失(要补) vs 图内装饰词/编号(进 STOP)")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="mode", required=True)

    p1 = sub.add_parser("slide", help="照画页结构对账")
    p1.add_argument("--src", required=True, help="原稿 pptx")
    p1.add_argument("--src-slide", type=int, required=True)
    p1.add_argument("--dst", required=True, help="成品 pptx")
    p1.add_argument("--dst-slide", type=int, required=True)
    p1.add_argument("--gap", type=float, default=0.45, help="分带纵向间距阈值 in")
    p1.add_argument("--json", action="store_true")
    p1.set_defaults(func=mode_slide)

    p2 = sub.add_parser("ocr", help="原稿图 OCR 对账")
    p2.add_argument("--image", required=True)
    p2.add_argument("--dst", required=True)
    p2.add_argument("--dst-slide", type=int, required=True)
    p2.add_argument("--lang", default="chi_sim+eng")
    p2.set_defaults(func=mode_ocr)

    a = ap.parse_args()
    a.func(a)


if __name__ == "__main__":
    main()
