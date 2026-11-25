#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""extract_assets — Phase 3 资产管线：提取原稿全部图片 → 去重入库（ppt-retrofit）

  python3 -I extract_assets.py --pptx 原稿.pptx --out-dir assets [--manifest assets.json]

规则：
- 命名 p{页号:03d}_{序号:02d}.{ext}，序号=该页内出现顺序（含组合内递归）
- sha1 去重：跨页/同页重复引用只落一个文件，manifest 记 slides 列表
- TIFF → PNG 转换（PPTX 内嵌 TIFF 兼容性差）
- <40px 小图标 flag=tiny（多为装饰点，重建时通常不继承）
- manifest: {"files": {文件名: 元数据}, "slides": {页号: [文件名,...]}}
"""
import argparse
import hashlib
import io
import json
import os
import re

from PIL import Image
from pptx import Presentation

TINY_PX = 40


def iter_pictures(shapes):
    for sh in shapes:
        if sh.shape_type == 6:
            yield from iter_pictures(sh.shapes)
            continue
        try:
            if sh.shape_type == 13:
                yield sh
        except Exception:
            pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--out-dir", default="assets")
    ap.add_argument("--manifest", default="")
    a = ap.parse_args()
    man_path = a.manifest or os.path.join(a.out_dir, "assets.json")
    os.makedirs(a.out_dir, exist_ok=True)

    prs = Presentation(a.pptx)
    files, slides_map = {}, {}
    seen = {}          # sha1 -> filename
    counter = {}       # slide no -> per-slide index

    for si, slide in enumerate(prs.slides, 1):
        lst = []
        for pic in iter_pictures(slide.shapes):
            counter[si] = counter.get(si, 0) + 1
            try:
                blob = pic.image.blob
            except Exception as e:
                print(f"  ! P{si:02d} 图 {counter[si]} 提取失败: {e}")
                continue
            ext = pic.image.ext.lower()
            # 注意：sha1 必须在 TIFF→PNG 转换之后计算——manifest 登记的
            # sha1 是"落盘文件的字节指纹"（继承校验用成品内嵌图 sha1 对表）。
            # 若取转换前 TIFF 的 sha，登记值与文件不符，校验必挂。
            if ext in ("tiff", "tif"):
                im = Image.open(io.BytesIO(blob))
                if im.mode not in ("RGB", "RGBA"):
                    im = im.convert("RGBA")
                buf = io.BytesIO()
                im.save(buf, "PNG")
                blob, ext = buf.getvalue(), "png"
            sha1 = hashlib.sha1(blob).hexdigest()
            if sha1 in seen:
                lst.append(seen[sha1])
                files[seen[sha1]]["slides"].append(si)
                continue
            px = ""
            try:
                px = list(Image.open(io.BytesIO(blob)).size)
            except Exception:
                pass
            fname = f"p{si:03d}_{counter[si]:02d}.{ext}"
            with open(os.path.join(a.out_dir, fname), "wb") as f:
                f.write(blob)
            tiny = bool(px) and min(px) < TINY_PX
            files[fname] = {"sha1": sha1, "px": px, "bytes": len(blob),
                            "slides": [si], "tiny": tiny}
            seen[sha1] = fname
            lst.append(fname)
        if lst:
            slides_map[si] = lst

    man = {"files": files, "slides": slides_map}
    with open(man_path, "w", encoding="utf-8") as f:
        json.dump(man, f, ensure_ascii=False, indent=1)
    dup_groups = {}
    for fn, m in files.items():
        if len(m["slides"]) > 1:
            dup_groups[fn] = m["slides"]
    print(f"提取图片: 唯一文件 {len(files)} 个 → {a.out_dir}/")
    print(f"重复引用组: {len(dup_groups)}")
    for fn, ss in sorted(dup_groups.items(), key=lambda x: -len(x[1]))[:10]:
        print(f"    {fn} ← 页 {ss}")
    tiny = [fn for fn, m in files.items() if m["tiny"]]
    print(f"小图标(tiny): {len(tiny)}")
    print(f"manifest → {man_path}")


if __name__ == "__main__":
    main()
