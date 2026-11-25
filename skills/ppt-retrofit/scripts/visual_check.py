#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""visual_check — 渲染自查三件套（LibreOffice 实渲，交付前最后一道闸）

子命令：
  render  <pptx> [--pages 1-99] [--dpi 110]      渲染每页为 <outdir>/<文件名>/pNN.png
  diff    <原稿pptx> <成品pptx> --pages 15,18    沿用页像素比对（灰度阈值 diff%），出热区图
  sbs     <原稿pptx> <成品pptx> --pages 21,27    重画页并排大图（原稿左|成品右），供肉眼核验

用法约定（self-check.md 渲染自查节）：
- 整页沿用页 → diff 阈值 6%（字体替换文字回流的正常波动由阈值吸收）
- 严格重画/重排页 → diff 必然超标，直接 sbs 并排肉眼核验
- SmartArt 在 LO 是回退渲染（乱叠），文本可对账、布局不可比——这类页走文字对账

环境准备：无 soffice 时先装 LibreOffice（apt install libreoffice-impress；
有本地离线 deb 缓存优先用缓存）。soffice 全程 headless +
独立 UserInstallation，可并发不锁 profile。
"""
import argparse
import os
import shutil
import subprocess
import sys

LO_PROFILE = "file:///tmp/lo_profile_visual_check"


def sh(cmd):
    r = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stdout[-1500:]); print(r.stderr[-1500:], file=sys.stderr)
        raise SystemExit(f"命令失败: {cmd}")
    return r.stdout


def pages_arg(s):
    out = set()
    for part in s.split(","):
        if "-" in part:
            a, b = part.split("-"); out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    return sorted(out)


def deck_dir(outdir, pptx):
    name = os.path.splitext(os.path.basename(pptx))[0][:40]
    return os.path.join(outdir, name)


def render(pptx, pages, dpi, outdir):
    import fitz  # PyMuPDF：PDF→PNG
    out = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0][:40])
    os.makedirs(out, exist_ok=True)
    pdf = os.path.join(out, "deck.pdf")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise SystemExit("soffice 不存在——先安装 libreoffice-impress")
    if os.path.exists(pdf):
        os.remove(pdf)
    sh(f'{soffice} --headless --norestore --nolockcheck '
       f'-env:UserInstallation={LO_PROFILE} --convert-to pdf '
       f'--outdir "{out}" "{pptx}"')
    src_pdf = os.path.join(out, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if not os.path.exists(src_pdf):
        raise SystemExit("PDF 未生成")
    if os.path.getsize(src_pdf) < 3000:
        raise SystemExit("PDF 仅 %d 字节——疑似空输出：先用 python-pptx 重开文件自检形状数（XML 写坏时 soffice 会静默出空 PDF）" % os.path.getsize(src_pdf))
    if src_pdf != pdf:
        os.replace(src_pdf, pdf)
    doc = fitz.open(pdf)
    n = doc.page_count
    for p in pages:
        if p > n:
            print(f"  超出页数 {p}>{n}，跳过"); continue
        pg = doc[p - 1]
        zoom = dpi / 72.0
        pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        png = os.path.join(out, f"p{p:02d}.png")
        pix.save(png)
        print(f"  渲染 p{p:02d} → {png}  ({pix.width}x{pix.height})")
    return out


def _load(outdir, pptx, p):
    from PIL import Image
    path = os.path.join(deck_dir(outdir, pptx), f"p{p:02d}.png")
    if not os.path.exists(path):
        raise SystemExit(f"缺渲染图 {path}——先对两份文件各跑一次 render")
    return Image.open(path).convert("L")


def diff(orig, new, pages, thresh, outdir):
    import numpy as np
    from PIL import Image
    print(f"{'页':>4} {'差异像素%':>9}  判定   （阈值 {thresh}%，热区图同名 _diff.png）")
    for p in pages:
        a, b = _load(outdir, orig, p), _load(outdir, new, p)
        w = min(a.width, b.width); h = min(a.height, b.height)
        a = a.resize((w, h)); b = b.resize((w, h))
        A, B = np.asarray(a, dtype=np.int16), np.asarray(b, dtype=np.int16)
        D = np.abs(A - B)
        mask = D > 28
        pct = mask.mean() * 100
        verdict = "OK" if pct <= thresh else "!!差异"
        print(f"p{p:02d}  {pct:8.2f}%  {verdict}")
        heat = Image.merge("RGB", (
            Image.fromarray(np.where(mask, 255, np.asarray(a) // 3).astype("uint8")),
            Image.fromarray((np.asarray(a) // 3).astype("uint8")),
            Image.fromarray((np.asarray(a) // 3).astype("uint8"))))
        heat.save(os.path.join(deck_dir(outdir, new), f"p{p:02d}_diff.png"))


def sbs(orig, new, pages, outdir):
    from PIL import Image, ImageDraw
    for p in pages:
        a = Image.open(os.path.join(deck_dir(outdir, orig), f"p{p:02d}.png"))
        b = Image.open(os.path.join(deck_dir(outdir, new), f"p{p:02d}.png"))
        h = 780
        a = a.resize((int(a.width * h / a.height), h))
        b = b.resize((int(b.width * h / b.height), h))
        sheet = Image.new("RGB", (a.width + b.width + 24, h + 34), "#222222")
        sheet.paste(a, (0, 34)); sheet.paste(b, (a.width + 24, 34))
        d = ImageDraw.Draw(sheet)
        d.text((8, 8), f"ORIGINAL p{p:02d}", fill="#ffdd55")
        d.text((a.width + 32, 8), f"REBUILT p{p:02d}", fill="#66ff66")
        out = os.path.join(deck_dir(outdir, new), f"p{p:02d}_sbs.png")
        sheet.save(out)
        print(f"  并排 p{p:02d} → {out}")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("render")
    r.add_argument("pptx")
    r.add_argument("--pages", default="1-99")
    r.add_argument("--dpi", type=int, default=110)
    d = sub.add_parser("diff"); d.add_argument("orig"); d.add_argument("new")
    d.add_argument("--pages", required=True); d.add_argument("--thresh", type=float, default=6.0)
    s = sub.add_parser("sbs"); s.add_argument("orig"); s.add_argument("new")
    s.add_argument("--pages", required=True)
    for p in (r, d, s):
        p.add_argument("--outdir", default="previews", help="渲染输出根目录（默认 ./previews）")
    a = ap.parse_args()
    pages = pages_arg(a.pages) if hasattr(a, "pages") else []
    if a.cmd == "render":
        render(a.pptx, pages, a.dpi, a.outdir)
    elif a.cmd == "diff":
        diff(a.orig, a.new, pages, a.thresh, a.outdir)
    elif a.cmd == "sbs":
        sbs(a.orig, a.new, pages, a.outdir)


if __name__ == "__main__":
    main()
