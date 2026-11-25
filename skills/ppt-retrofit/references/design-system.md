# 设计体系 — tokens、字阶、页型库、helper 模板

## 画布与边距

- 画布 13.333 × 7.5 in（16:9）；全片边距体系 x=0.55，内容宽 12.23
- 页标题 20pt，y 0.20–0.82 带 logo 位（12.28, 0.24）；模块标题统一 14pt（mod_title）
- 章节页/深色页复用 section_page（深色体系见下）

## 字阶（写死为默认，禁止新增档位）

| 渲染档 | 用途 |
|---|---|
| 6.5 | 名词格（窄格多行，下限） |
| 10 | chip / 辅助文字 / 图注 |
| 12 | 正文 |
| 14 | 小标题 / 徽章 / 编号 / 模块标题（与模块标题同高，靠色块样式区分） |
| 16 | 强调数字 |
| 20 | 页标题 |
| 26+ | 显示档（幽灵编号、THANK YOU 等）不动 |

**SIZE_MAP 反推坑（本项目踩过最疼的坑）**：SIZE_MAP 把入参映射到渲染档——**要渲染 12pt 必须传 10；要渲染 10pt 传 7.0–8.5；传 12 会渲染成 14**。14/20/26+ 原样。写页面时先想渲染目标再反推入参。

## 色彩 bootstrap（每项目从原稿提取，不用本表硬编码）

- BRAND = 原稿 theme accent1；辅助色从 logo 取色（参考项目绿 #00C800 取自 logo 绿点）
- 三色轮换 `TRI = [BRAND, 辅助色, 深色]` 用于编号/类目/序号交替，避免单色疲劳
- `TINT_OF`：每色浅底派生（卡片浅染）；`LIGHT_OF`：深底上的亮色变体
- 中性色固定：TEXT #1F2937 / GRAY #6B7280 / LINE #E3E8F0 / BG #F8FAFC
- 深色页体系（用户要求"降低阅读疲劳"时启用）：DEEP→DARK 渐变底 + 幽灵大编号 + 亮色变体文字 + 进度点；卡片 DKCARD / 描边 DKLINE / 正文 DKTXT

## helper 源码模板

拷到项目 `<out>/scripts_gen/deck_lib.py`，只改顶部 tokens 区（标 ⚠ 的行）。字体名按客户环境（微软雅黑在用户 Windows/WPS 渲染，构建机没有也能出文件）。

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deck_lib — 设计 tokens 与助手（ppt-retrofit 模板）"""
import os
from lxml import etree
from PIL import Image
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

OUT_PPTX = "成品.pptx"            # ⚠ 每项目改
IMG = "images"                     # ⚠ 资产目录（相对 cwd 或绝对路径）
CW, CH = 13.333, 7.5

# ---------- tokens ⚠ 每项目改 ----------
BRAND, AUX, DARK = "2A6BE5", "00C800", "1E3A5F"
DEEP = "0F1C2E"
TINT, BG = "EDF3FE", "F8FAFC"
AUX_L, AUX_TINT = "4DE167", "E6F8EC"
DARK_TINT = "E8ECF4"
TEXT, GRAY, LINE = "1F2937", "6B7280", "E3E8F0"
TRI = [BRAND, AUX, DARK]
TINT_OF = {BRAND: TINT, AUX: AUX_TINT, DARK: DARK_TINT}
LIGHT_OF = {BRAND: "4F8DF9", AUX: AUX_L, DARK: "8FB4F2"}
DKCARD, DKLINE, DKTXT = "1B3050", "2E4A73", "C9D8F0"
FONT = "微软雅黑"                  # ⚠ 每项目改

# ---------- 字阶收敛（渲染档见 design-system.md；注意反推坑） ----------
SIZE_MAP = {
    5.5: 6.5, 6.0: 6.5,
    7.0: 10.0, 7.2: 10.0, 7.5: 10.0, 8.0: 10.0, 8.5: 10.0,
    9.0: 12.0, 9.5: 12.0, 10.0: 12.0,
    10.5: 14.0, 11.0: 14.0, 11.5: 14.0, 12.0: 14.0,
    12.5: 14.0, 13.0: 14.0, 13.5: 14.0,
    15.0: 16.0, 18.0: 16.0,
}

def RGB(s): return RGBColor.from_string(s)

def _alpha(shape, pct):
    sf = shape._element.spPr.find(qn("a:solidFill"))
    if sf is None: return
    c = sf.find(qn("a:srgbClr"))
    if c is None: return
    a = etree.SubElement(c, qn("a:alpha")); a.set("val", str(int(pct * 1000)))

def _shadow(shape, blur=90000, dist=30000, alpha=20):
    spPr = shape._element.spPr
    el = etree.SubElement(spPr, qn("a:effectLst"))
    sh = etree.SubElement(el, qn("a:outerShdw"))
    sh.set("blurRad", str(blur)); sh.set("dist", str(dist))
    sh.set("dir", "5400000"); sh.set("rotWithShape", "0")
    c = etree.SubElement(sh, qn("a:srgbClr")); c.set("val", DARK)
    a = etree.SubElement(c, qn("a:alpha")); a.set("val", str(alpha * 1000))

def shp(slide, st, x, y, w, h, fill=None, line=None, lw=0.75, shadow=False,
        radius=None, alpha=None, name=None):
    s = slide.shapes.add_shape(st, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = RGB(fill)
    else: s.fill.background()
    if line: s.line.color.rgb = RGB(line); s.line.width = Pt(lw)
    else: s.line.fill.background()
    if radius is not None:
        try: s.adjustments[0] = radius
        except Exception: pass
    if alpha: _alpha(s, alpha)
    if shadow: _shadow(s)
    if name: s.name = name
    s.text_frame.word_wrap = True
    return s

def _style_run(r, text, size, bold, color, name=FONT, italic=False):
    size = SIZE_MAP.get(size, size)          # 字阶收敛单点入口
    r.text = text
    f = r.font
    f.size = Pt(size); f.bold = bold; f.italic = italic; f.name = name
    f.color.rgb = RGB(color)
    rPr = r._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None: ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", name)

def P(text, size=10, bold=False, color=TEXT, align=None, sa=0, sb=0, ls=None, runs=None):
    return dict(text=text, size=size, bold=bold, color=color, align=align,
                sa=sa, sb=sb, ls=ls, runs=runs)

def txt(slide, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap; tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    if isinstance(paras, dict): paras = [paras]
    first = True
    for pa in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = pa.get("align") or align
        if pa.get("sa"): p.space_after = Pt(pa["sa"])
        if pa.get("sb"): p.space_before = Pt(pa["sb"])
        if pa.get("ls"): p.line_spacing = pa["ls"]
        if pa.get("runs"):
            for rr in pa["runs"]:
                _style_run(p.add_run(), rr[0], rr[1], rr[2], rr[3],
                           rr[4] if len(rr) > 4 else FONT)
        else:
            for i, seg in enumerate(pa["text"].split("\n")):
                if i > 0:
                    p = tf.add_paragraph(); p.alignment = pa.get("align") or align
                    if pa.get("sa"): p.space_after = Pt(pa["sa"])
                    if pa.get("ls"): p.line_spacing = pa["ls"]
                _style_run(p.add_run(), seg, pa["size"], pa["bold"], pa["color"])
    return tb

def pic(slide, fname, x, y, w, h, border=None, shadow=False):
    path = fname if os.path.isabs(fname) else os.path.join(IMG, fname)
    iw, ih = Image.open(path).size
    ar = iw / ih
    if w / h > ar: pw, ph = h * ar, h
    else: pw, ph = w, w / ar
    px, py = x + (w - pw) / 2, y + (h - ph) / 2
    p = slide.shapes.add_picture(path, Inches(px), Inches(py), Inches(pw), Inches(ph))
    if border: p.line.color.rgb = RGB(border); p.line.width = Pt(0.75)
    if shadow: _shadow(p)
    return p

def framed_pic(slide, fname, x, y, w, h, caption=None, cap_size=8):
    """统一截图相框：白底细边卡包图（w/h=图片显示区，按真实宽高比定框），图注并入框底"""
    pad = 0.08
    cap_h = 0.26 if caption else 0.0
    fw, fh = w + 2 * pad, h + 2 * pad + cap_h
    card(slide, x - pad, y - pad, fw, fh, fill="FFFFFF", line=LINE, radius=0.05, shadow=False)
    pic(slide, fname, x, y, w, h, border=LINE)
    if caption:
        txt(slide, x - pad, y + h + pad + 0.02, fw, cap_h - 0.02,
            P(caption, cap_size, False, GRAY, PP_ALIGN.CENTER))
    return fw, fh

def gradient(shape, c1, c2, angle=115.0):
    shape.fill.gradient()
    st = shape.fill.gradient_stops
    st[0].color.rgb = RGB(c1); st[0].position = 0.0
    st[1].color.rgb = RGB(c2); st[1].position = 1.0
    try: shape.fill.gradient_angle = angle
    except Exception: pass

def chip(slide, x, y, w, h, text, fill=BRAND, tcolor="FFFFFF", size=9, bold=True, radius=0.5):
    s = shp(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, radius=radius)
    tf = s.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _style_run(p.add_run(), text, size, bold, tcolor)
    return s

def card(slide, x, y, w, h, fill="FFFFFF", line=LINE, radius=0.07, shadow=True):
    return shp(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
               fill=fill, line=line, lw=0.75, shadow=shadow, radius=radius)

def mod_title(slide, x, y, text, accent=BRAND, w=4.5, h=0.36, size=14):
    """模块标题统一：小色块 + 粗体深色文本"""
    shp(slide, MSO_SHAPE.RECTANGLE, x, y + (h - 0.26) / 2, 0.1, 0.26, fill=accent)
    txt(slide, x + 0.2, y, w, h, P(text, size, True, DARK), anchor=MSO_ANCHOR.MIDDLE)

def vshape(slide, x, y, w, h, text, fill, tcolor="FFFFFF", size=10, bold=True):
    """竖排（eaVert）色块文字，用于泳道/分区纵条标签"""
    s = shp(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill=fill)
    tf = s.text_frame; tf.word_wrap = False; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    tf._txBody.find(qn("a:bodyPr")).set("vert", "eaVert")
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _style_run(p.add_run(), text, size, bold, tcolor)
    return s

def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def section_page(prs, num, title, eng, accent):
    """章节页：深色底 + 幽灵大编号 + 进度点（深色体系样板）"""
    s = new_slide(prs)
    bg = shp(s, MSO_SHAPE.RECTANGLE, 0, 0, CW, CH, name="decor-bg")
    gradient(bg, DEEP, DARK, 115)
    acc = LIGHT_OF.get(accent, accent)
    g = txt(s, 6.9, 1.15, 5.95, 4.4, P(num, 200, True, DKCARD), align=PP_ALIGN.RIGHT)
    g.name = "decor-ghost"
    shp(s, MSO_SHAPE.OVAL, 10.75, 0.72, 1.85, 1.85, line=DKLINE, lw=1.5, name="decor-ring")
    shp(s, MSO_SHAPE.OVAL, 12.18, 2.32, 0.22, 0.22, fill=AUX, name="decor-dot")
    shp(s, MSO_SHAPE.RECTANGLE, 0.92, 3.02, 0.6, 0.09, fill=acc)
    txt(s, 0.92, 3.26, 8.5, 0.35, P(f"SECTION {num} · {eng}", 12, True, acc))
    txt(s, 0.9, 3.66, 10.6, 1.2, P(title, 40, True, "FFFFFF"))
    idx = int(num) - 1
    for i in range(6):
        shp(s, MSO_SHAPE.OVAL, 0.92 + i * 0.42, 6.35, 0.16, 0.16,
            fill=(acc if i == idx else DKLINE))
    return s

def clean_empty_runs(tb):
    """表格必调：删除无文本 run（<a:t/>）。空单元格写空串会产生此 run，
    LibreOffice 会把所在行撑高约 2 倍致表格溢出（gotcha #26）。
    判断必须用 not t.text（自闭合 <a:t/> 在 lxml 里是 None，=="" 匹配不到）"""
    for t_el in list(tb._tbl.iter(qn("a:t"))):
        if not t_el.text:
            run_el = t_el.getparent()
            run_el.getparent().remove(run_el)

def verify(path):
    """越界检查（三件套之一，check_geometry.py 参数化版另备）"""
    from pptx import Presentation
    prs2 = Presentation(path)
    issues = []
    for si, sl in enumerate(prs2.slides, 1):
        for sh in sl.shapes:
            if sh.name and sh.name.startswith("decor"): continue
            l, t = sh.left / 914400, sh.top / 914400
            r, b = l + sh.width / 914400, t + sh.height / 914400
            if l < -0.02 or t < -0.02 or r > CW + 0.02 or b > CH + 0.02:
                issues.append(f"slide{si} 越界: {sh.name}")
    return issues
```

## 页型库（论点类型 → 页型，几何骨架要点）

| 论点 | 页型 | 几何骨架要点 |
|---|---|---|
| 因果/递进 | 要求语带 + 结果行卡 | 语带 TINT 底 + 左色条；行卡等高、行距均匀（可断言） |
| 分层依赖 | 层带 + 下箭头 | 行卡按 TINT_OF 层色浅染，层间 DOWN_ARROW |
| 总分结构 | 中心放射 | 中心实底圆垫+描边（连线画完再垫圆遮线头），连线从圆缘出、端点亮色圆点 |
| 并列枚举 | 卡片网格 | 列宽一致对齐边距体系；有类目加色带头（浅底带 + 14pt 类目名 + 左色条） |
| 供给/支撑 | 供数节点 + 分配扇形 | 顶节点居中对齐中列；横杠 + 支路 DOWN_ARROW；列卡行首色块标题 + 角标大数字 |
| 流程 | 步骤链 / 泳道 | 泳道边界先定，节点落带内做断言；序号徽章不与文字交叠 |
| 截图组 | 相框网格 | framed_pic；**框按图片真实宽高比定**（盒子比例与图不符 = 大片不对称留白，丑页主因之一） |
| 深色呼吸页 | 章节页 / 总结页 | section_page 复用；总结页放射卡阵两行底边对齐 |
