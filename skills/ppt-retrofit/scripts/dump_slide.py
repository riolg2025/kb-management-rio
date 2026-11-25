#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dump_slide — 照画页结构规格提取（原稿 → 重建规格书）

  python3 -I dump_slide.py --pptx 原稿.pptx --slide 11 [--pics] [--brief]

- shapes: 序号/几何(prst)/bbox(in)/fill/线条/文本(≤60字)
- cxns:   连接线 bbox + flip + 颜色 + st/end 挂接的 shape id
- --pics: 附加图片 bbox + sha1 前 8 位（可对回 assets.json）
- --brief: 只输出文本框 (x,y) + 文本，供配对/清点用
坐标单位 in，画布 13.333×7.5。
"""
import argparse
import hashlib
import re

from pptx import Presentation
from pptx.oxml.ns import qn

EMU = 914400.0


def bbox(el):
    xfrm = el.find(qn("p:spPr") + "/" + qn("a:xfrm"))
    if xfrm is None:
        xfrm = el.find(qn("p:grpSpPr") + "/" + qn("a:xfrm"))
    if xfrm is None:
        return None
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    x, y = int(off.get("x")) / EMU, int(off.get("y")) / EMU
    w, h = int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU
    flips = (xfrm.get("flipH"), xfrm.get("flipV"))
    return x, y, w, h, flips


def fill_of(el):
    for tag in ("a:solidFill",):
        sf = el.find(qn("p:spPr") + "/" + qn(tag))
        if sf is not None:
            c = sf.find(qn("a:srgbClr"))
            if c is not None:
                return "#" + c.get("val")
            if sf.find(qn("a:schemeClr")) is not None:
                return "scheme:" + sf.find(qn("a:schemeClr")).get("val")
        sf2 = el.find(qn("p:style"))
        if sf2 is not None:
            r = sf2.find(qn("a:fillRef"))
            if r is not None:
                return "fillRef" + r.get("idx", "")
    ln = el.find(qn("p:spPr") + "/" + qn("a:ln") + "/" + qn("a:solidFill") + "/" + qn("a:srgbClr"))
    if ln is not None:
        return "line#" + ln.get("val")
    return ""


def texts_of(sp):
    out = []
    for t in sp.iter(qn("a:t")):
        out.append(t.text or "")
    return " ".join(out).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--slide", type=int, required=True)
    ap.add_argument("--pics", action="store_true")
    ap.add_argument("--brief", action="store_true")
    a = ap.parse_args()
    prs = Presentation(a.pptx)
    slide = prs.slides[a.slide - 1]
    tree = slide.shapes._spTree
    id2shape = {}
    counter = [0]

    def walk(tree, depth, tf=(0.0, 0.0, 1.0, 1.0)):
        """tf=(dx, dy, sx, sy)：父组合的坐标变换（平移+缩放）"""
        for el in tree:
            tag = el.tag.split("}")[1]
            if tag == "grpSp":
                counter[0] += 1
                xfrm = el.find(qn("p:grpSpPr") + "/" + qn("a:xfrm"))
                ndx, ndy, nsx, nsy = tf
                if xfrm is not None and not a.brief:
                    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
                    cho, che = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
                    if all(v is not None for v in (off, ext, cho, che)):
                        ox, oy = int(off.get("x")) / EMU, int(off.get("y")) / EMU
                        cx, cy = int(cho.get("x")) / EMU, int(cho.get("y")) / EMU
                        sx = int(ext.get("cx")) / EMU / (int(che.get("cx")) / EMU or 1)
                        sy = int(ext.get("cy")) / EMU / (int(che.get("cy")) / EMU or 1)
                        ndx, ndy = ox - cx * sx + ndx * sx, oy - cy * sy + ndy * sy
                        nsx, nsy = sx * nsx, sy * nsy
                        print("%sgrp%-2d (%5.2f,%5.2f %5.2fx%5.2f scale=%.2f,%.2f)"
                              % ("  " * depth, counter[0], ox, oy,
                                 int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU, sx, sy))
                walk(el, depth + 1, (ndx, ndy, nsx, nsy))
                continue
            if tag not in ("sp", "cxnSp", "pic"):
                continue
            counter[0] += 1
            i = counter[0]
            bb = bbox(el)
            if bb is None:
                continue
            X, Y = bb[0] * tf[2] + tf[0], bb[1] * tf[3] + tf[1]
            W, H = bb[2] * tf[2], bb[3] * tf[3]
            t = texts_of(el)
            pre = "  " * depth
            if tag == "sp":
                if a.brief:
                    if t:
                        print("%ssp%-3d (%5.2f,%5.2f) %s" % (pre, i, X, Y, t.replace(chr(10), "⏎")[:60]))
                    continue
                prst = ""
                g = el.find(qn("p:spPr") + "/" + qn("a:prstGeom"))
                if g is not None:
                    prst = g.get("prst")
                if el.find(qn("p:spPr") + "/" + qn("a:custGeom")) is not None:
                    prst = "custGeom"
                sid = el.find(qn("p:nvSpPr") + "/" + qn("p:cNvPr")).get("id")
                id2shape[sid] = i
                print("%ssp%-3d %-12s (%5.2f,%5.2f %5.2fx%5.2f) %-16s %s"
                      % (pre, i, prst or "rect", X, Y, W, H, fill_of(el),
                         t.replace(chr(10), "⏎")[:60]))
            elif tag == "cxnSp":
                nv = el.find(qn("p:nvCxnSpPr") + "/" + qn("p:cNvPr")).get("id")
                st = el.find(qn("p:nvCxnSpPr") + "/" + qn("p:cNvCxnSpPr") + "/" + qn("a:stCxn"))
                en = el.find(qn("p:nvCxnSpPr") + "/" + qn("p:cNvCxnSpPr") + "/" + qn("a:endCxn"))
                spPr = el.find(qn("p:spPr"))
                g = spPr.find(qn("a:prstGeom"))
                prst = g.get("prst") if g is not None else "?"
                ln = spPr.find(qn("a:ln"))
                dash = hd = tl = ""
                if ln is not None:
                    d = ln.find(qn("a:prstDash"))
                    dash = d.get("val") if d is not None else ""
                    h_ = ln.find(qn("a:headEnd"))
                    t_ = ln.find(qn("a:tailEnd"))
                    hd = h_.get("type") if h_ is not None else ""
                    tl = t_.get("type") if t_ is not None else ""
                arw = ("←" if hd == "arrow" else "") + ("→" if tl == "arrow" else "")
                extras = " ".join(x for x in (dash, arw) if x)
                print("%scxn %-14s (%5.2f,%5.2f %5.2fx%5.2f) flip=%s %-10s %-8s st=%s end=%s"
                      % (pre, prst, X, Y, W, H, bb[4], fill_of(el), extras,
                         id2shape.get(st.get("id"), st.get("id")) if st is not None else "-",
                         id2shape.get(en.get("id"), en.get("id")) if en is not None else "-"))
            elif tag == "pic" and a.pics:
                blip = el.find(qn("p:blipFill") + "/" + qn("a:blip"))
                rid = blip.get(qn("r:embed")) if blip is not None else ""
                sha = ""
                if rid:
                    part = slide.part.rels[rid].target_part
                    sha = hashlib.sha1(part.blob).hexdigest()[:8]
                print("%spic%-3d (%5.2f,%5.2f %5.2fx%5.2f) %s" % (pre, i, X, Y, W, H, sha))

    walk(tree, 0)


if __name__ == "__main__":
    main()
