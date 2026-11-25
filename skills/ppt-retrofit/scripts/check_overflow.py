#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_overflow — 文字溢出估算（ppt-retrofit 校验三件套之一）

启发式估宽：CJK≈1.02em、大写/数字≈0.60em、小写≈0.52em、空格≈0.30em。
逐框估算行数×行高 vs 框高（wrap 时）、行宽 vs 框宽（nowrap 时）。
组合形状（继承图）与竖排窄框跳过。旗标是"疑似"非"确诊"，逐条人工过。

用法：
  python3 -I check_overflow.py --pptx 成品.pptx

全绿标准：旗标总数 0。
"""
import argparse
import math

from pptx import Presentation
from pptx.oxml.ns import qn


def run_w(text, pt):
    w = 0.0
    for ch in text:
        o = ord(ch)
        if o >= 0x2E80:
            w += 1.02
        elif ch == " ":
            w += 0.30
        elif ch.isupper() or ch.isdigit():
            w += 0.60
        else:
            w += 0.52
    return w * pt / 72.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--h-tol", type=float, default=1.12)
    ap.add_argument("--w-tol", type=float, default=0.05)
    a = ap.parse_args()

    prs = Presentation(a.pptx)
    flags = 0
    for i, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if sh.shape_type == 6:  # 组合（继承图）跳过
                continue
            try:
                if sh.has_table:
                    continue
            except Exception:
                pass
            try:
                tf = sh.text_frame
            except Exception:
                continue
            if tf is None or not tf.text.strip():
                continue
            w = (sh.width or 0) / 914400
            h = (sh.height or 0) / 914400
            if w <= 0 or h <= 0 or w < 0.5:
                continue
            try:
                bp = tf._txBody.find(qn("a:bodyPr"))
                if bp is not None and bp.get("vert"):
                    continue  # eaVert 竖排框
            except Exception:
                pass
            wrap = tf.word_wrap
            est_h = 0.0
            over_w = 0.0
            for p in tf.paragraphs:
                rs = [(r.text or "", r.font.size.pt if r.font.size else 12.0)
                      for r in p.runs]
                if not rs:
                    continue
                lw = sum(run_w(t, s) for t, s in rs)
                mx = max(s for _, s in rs)
                ls = p.line_spacing
                if ls is None:
                    lh = mx / 72.0
                elif isinstance(ls, (float, int)):
                    lh = mx / 72.0 * ls
                else:
                    lh = ls.pt / 72.0
                nlines = (1 if wrap is False
                          else max(1, math.ceil(lw / max(w, 0.05) - 0.02)))
                est_h += nlines * lh
                if p.space_after:
                    est_h += p.space_after.pt / 72.0
                if p.space_before:
                    est_h += p.space_before.pt / 72.0
                if wrap is False:
                    over_w = max(over_w, lw / w)
            t0 = tf.text.replace("\n", "⏎")[:26]
            ratio = est_h / h
            if ratio > a.h_tol and est_h - h > a.w_tol:
                print(f"s{i:02d} 高溢出 ×{ratio:.2f} box={w:.2f}x{h:.2f} "
                      f"est={est_h:.2f} «{t0}»")
                flags += 1
            if over_w > 1.03:
                print(f"s{i:02d} 宽溢出 ×{over_w:.2f} box_w={w:.2f} «{t0}»")
                flags += 1
    print("旗标总数:", flags)


if __name__ == "__main__":
    main()
