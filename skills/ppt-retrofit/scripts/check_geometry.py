#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_geometry — 越界 + 文本框交叠（ppt-retrofit 校验三件套之一）

用法：
  python3 -I check_geometry.py --pptx 成品.pptx [--canvas 13.333x7.5] [--overlap 0.05]

全绿标准：越界 0 / 交叠 0。
"""
import argparse
from pptx import Presentation
from pptx.oxml.ns import qn


def overlaps(b1, b2):
    x1, y1, w1, h1 = b1
    x2, y2, w2, h2 = b2
    ox = max(0.0, min(x1 + w1, x2 + w2) - max(x1, x2))
    oy = max(0.0, min(y1 + h1, y2 + h2) - max(y1, y2))
    return ox * oy


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--canvas", default="13.333x7.5")
    ap.add_argument("--overlap", type=float, default=0.05, help="交叠面积阈值 in²")
    a = ap.parse_args()
    CW, CH = (float(x) for x in a.canvas.split("x"))
    TOL = 0.02

    prs = Presentation(a.pptx)
    oob = 0
    for si, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if sh.name and sh.name.startswith("decor"):
                continue
            try:
                l, t = sh.left / 914400, sh.top / 914400
                w, h = sh.width / 914400, sh.height / 914400
            except Exception:
                continue
            # rot=90°/270° 的形状按旋转后的视觉包围盒判（原稿连接线常用
            # "竖包围盒+旋转"落盘，未旋转包围盒会误报越界/占位）
            rot = sh.rotation
            if rot in (90.0, 270.0):
                cx, cy = l + w / 2, t + h / 2
                l, t, w, h = cx - h / 2, cy - w / 2, h, w
            r, b = l + w, t + h
            if l < -TOL or t < -TOL or r > CW + TOL or b > CH + TOL:
                print(f"s{si:02d} 越界: {sh.name} ({l:.2f},{t:.2f})-({r:.2f},{b:.2f})")
                oob += 1
    print("越界:", oob)

    tot = 0
    for si, sl in enumerate(prs.slides, 1):
        boxes = []
        for sh in sl.shapes:
            if sh.shape_type != 17:  # TEXT_BOX
                continue
            if sh.name and sh.name.startswith("decor"):
                continue
            boxes.append((sh.name, sh.left / 914400, sh.top / 914400,
                          sh.width / 914400, sh.height / 914400))
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                n1, *b1 = boxes[i]
                n2, *b2 = boxes[j]
                area = overlaps(b1, b2)
                if area > a.overlap:
                    print(f"s{si:02d} 交叠 [{n1}]×[{n2}] {area:.3f} in²")
                    tot += 1
    print("交叠:", tot)


if __name__ == "__main__":
    main()
