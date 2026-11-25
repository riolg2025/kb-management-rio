#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""graft_deck — 复杂图形整体沿用：原稿页形状 XML 原样移植进成品页（保真第一手段）

  python3 -I scripts/graft_deck.py --src 原稿.pptx --dst 成品.pptx \
      --src-slide 15 [--dst-slide 15] \
      [--exclude-text 词1,词2] [--only-groups] \
      [--font 微软雅黑] [--no-font] [--swap-theme]

做什么：
- 原稿页顶层形状（sp/pic/cxnSp/grpSp）deepcopy 进成品页：几何/样式/文本字节级保真
- 内嵌图片按源文件字节重挂（python-pptx 按 sha1 去重复用，继承图哈希校验不受影响）
- 默认强制 --font 字体（latin+ea 写进 rPr/defRPr/endParaRPr，防目标文档主题字体污染）
- --swap-theme：成品 theme1.xml 换成原稿的（沿用图形带 schemeClr/fillRef 时用）

何时用（决策树"xml"档位，复杂图形默认动作）：
- 嵌套组带缩放变换的拓扑/架构图（存储坐标≠视觉坐标，重画必错位）
- 几十个小图拼贴的界面页、严格逻辑图——重画错漏率远高于整体搬运
注意：原稿形状用显式 srgbClr 时跨文档安全；带 schemeClr/主题引用就加 --swap-theme。
"""
import argparse
import hashlib
import io
import os
import zipfile

from copy import deepcopy
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn


def _norm(s):
    return "".join(s.split())


def _force_font(el, name):
    for e in el.iter():
        if not isinstance(e.tag, str):
            continue
        if etree.QName(e).localname in ("rPr", "defRPr", "endParaRPr"):
            for tag in ("latin", "ea"):
                sub = e.find(qn("a:" + tag))
                if sub is None:
                    sub = etree.SubElement(e, qn("a:" + tag))
                sub.set("typeface", name)


def _remap_images(el, dst_slide, src_part):
    """源图片字节原样挂进目标包（sha1 不变）"""
    for blip in el.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid is None:
            continue
        blob = src_part.rels[rid].target_part.blob
        _part, rid2 = dst_slide.part.get_or_add_image_part(io.BytesIO(blob))
        blip.set(qn("r:embed"), rid2)


def graft(dst_slide, src_slide, font=None, excludes=(), only_groups=False):
    """把 src 页顶层形状并入 dst 页，返回搬运块数"""
    src_part = src_slide.part
    tree = dst_slide.shapes._spTree
    n = 0
    for el in src_slide.shapes._spTree:
        if not isinstance(el.tag, str):
            continue
        ln = etree.QName(el).localname
        if ln not in ("sp", "pic", "cxnSp", "grpSp"):
            continue
        if only_groups and ln != "grpSp":
            continue
        text = _norm("".join(t.text or "" for t in el.iter(qn("a:t"))))
        if any(x in text for x in excludes):
            continue
        new = deepcopy(el)
        _remap_images(new, dst_slide, src_part)
        if font:
            _force_font(new, font)
        tree.append(new)
        n += 1
    return n


def swap_theme(dst_path, src_path):
    """目标文档 theme1.xml ← 源文档（仅影响 schemeClr/fillRef 的解析）"""
    orig = zipfile.ZipFile(src_path).read("ppt/theme/theme1.xml")
    tmp = dst_path + ".tmp"
    with zipfile.ZipFile(dst_path) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = orig if item.filename == "ppt/theme/theme1.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    os.replace(tmp, dst_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="原稿 pptx")
    ap.add_argument("--dst", required=True, help="成品 pptx（就地修改，注意备份）")
    ap.add_argument("--src-slide", type=int, required=True)
    ap.add_argument("--dst-slide", type=int, help="默认与 --src-slide 相同")
    ap.add_argument("--exclude-text", default="", help="规整文本含任一子串即跳过（逗号分隔，如原页题/页脚）")
    ap.add_argument("--only-groups", action="store_true", help="只搬 grpSp 组")
    ap.add_argument("--font", default="微软雅黑", help="强制字体（默认微软雅黑）")
    ap.add_argument("--no-font", action="store_true", help="不改字体，保持原稿字体声明")
    ap.add_argument("--swap-theme", action="store_true", help="成品 theme1.xml 换成原稿的")
    a = ap.parse_args()

    excludes = [x for x in a.exclude_text.split(",") if x.strip()]
    dst_slide_no = a.dst_slide or a.src_slide
    before = len(Presentation(a.dst).slides[dst_slide_no - 1].shapes._spTree)

    src_sl = Presentation(a.src).slides[a.src_slide - 1]
    dst_prs = Presentation(a.dst)
    dst_sl = dst_prs.slides[dst_slide_no - 1]
    n = graft(dst_sl, src_sl,
              font=None if a.no_font else a.font,
              excludes=excludes, only_groups=a.only_groups)
    dst_prs.save(a.dst)
    if a.swap_theme:
        swap_theme(a.dst, a.src)

    # 保存后重开自检：XML 写坏时 soffice 会静默输出空 PDF，这里先拦一道
    chk = Presentation(a.dst)
    shapes = list(chk.slides[dst_slide_no - 1].shapes)
    print(f"p{dst_slide_no} 沿用 {n} 块；自检重开 OK，页内形状 {len(shapes)} 个（改动前 spTree 子元素 {before}）")
    if a.swap_theme:
        print("theme1.xml 已换成原稿主题")


if __name__ == "__main__":
    main()
