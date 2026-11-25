#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_text — 大纲 vs 成品 PPTX 文本完整性审计（ppt-retrofit 校验三件套之一）

规则：大纲机器层（```machine 围栏，无围栏退回旧行为）内容行清洗 → 按分隔符/关系
动词拆原子段 → norm 后在成品页全文中做子串匹配。版式描述词（STOP）与资产操作
元语言（PROC）不计入缺失。

⚠ 词表是生长型的：审计 → 人工核对（真缺失 vs 版式词/继承图内文字）→ 扩词表。
每项目都要养词表，种子表必漏。

用法：
  python3 -I check_text.py --pptx 成品.pptx --outline outline.md \
      [--from 1 --to 55] [--gone 37,46]
"""
import argparse
import re

from pptx import Presentation
from pptx.oxml.ns import qn


def norm(t):
    return re.sub(r"[\s\x0b\*｜→–—·+×&><\/、，。：:；;()（）\"“”\"'‘’'【】\[\]\-_#%.]+",
                  "", t).lower()


# ---- 种子词表（项目内生长） ----
STOP = {
    # 版式/结构描述
    "内容", "资产", "动作", "导语", "标题", "原标题", "左半页", "右半页",
    "左列", "右列", "左栏", "右栏", "左侧", "右侧", "左上", "左下", "右上", "右下",
    "顶部", "底部", "中央", "中段", "顶栏", "底部结论", "顶部叙述", "图注",
    "标签", "标语", "竖条", "横条", "分隔带", "泳道", "四象限", "网格", "象限配图",
    "文本", "要点", "其余要点", "全文", "清单",
    "三层", "四层", "五层", "分层", "自上而下", "两翼", "底座", "层带",
    "说明", "点说明", "条", "张", "页签", "截图", "张配图", "截图网格", "相框",
    "流程图", "架构图", "示意图", "环图", "表格", "图表", "大屏", "banner",
    # 常见继承图内高频词（按项目扩充）
    "logo", "cmdb", "easyops", "api", "agent", "itil", "oa", "sso", "vs",
}
PROC = [
    # 资产操作元语言
    "继承", "重绘", "重排", "重画", "包装", "原生", "对照", "转录", "风格统一",
    "模板", "章节", "自动布局", "结构化", "隐藏引用", "有效项", "重设计",
    "沿用", "原图", "手绘", "原为", "修复", "重做", "截自", "去重", "重复",
    "保留", "可选", "压缩", "误作", "裁剪", "放大", "变体", "同一文件",
    "一条不减", "一字不改", "按原稿", "user确认", "用户确认", "自动带", "逐条核对",
]

DELIM = re.compile(
    r"[｜→/、，。：:；;()（）\"“”\"'‘’'【】\[\]\n＋+\-—–\s·&*]"
    r"|属于|包含|所属|连接|创建|部署|关联|支撑|覆盖")


def segments(line):
    line = re.sub(r"（[^）]*）\s*$", "", line)   # 仅行尾括号=元语言注释，剥离（行内括号是正文）
    line = re.sub(r"^\s*[-*]+\s*", "", line)
    line = re.sub(r"^\s*\d+\s*[.、]\s*", "", line)
    line = re.sub(r"^\s*\*{0,2}\d{1,2}\s+", "", line)
    line = re.sub(r"[①②③④⑤⑥⑦⑧⑨⑩⑪⑫]", "", line)
    line = re.sub(r"[（(]图[）)]", "｜", line)
    line = re.sub(r"(左|右)下特性\s*\d+\s*条\s*[：:]", "｜", line)
    line = re.sub(r"特性\s*\d+\s*条\s*[：:]", "", line)
    out = []
    for seg in DELIM.split(line):
        n = norm(seg)
        if len(n) < 2 or n in STOP:
            continue
        if any(p in n for p in PROC):
            continue
        if re.search(r"img\d+", n) or re.search(r"p\d{3}_\d", n):
            continue
        if re.fullmatch(r"\d+([x×]\d+)?", n) or re.fullmatch(r"p?\d{1,3}", n):
            continue
        out.append((seg.strip()[:38], n))
    return out


def parse_outline(path):
    """#### PNN 开页；##/### 关页（防尾部附录混入）；```machine 围栏优先。"""
    text = open(path, encoding="utf-8").read()
    pages, cur, infence = {}, None, False
    fenced, fence_lines, plain = set(), {}, {}
    for ln in text.splitlines():
        m = re.match(r"^#{4}\s*P(\d{2})", ln)
        if m:
            cur = int(m.group(1))
            pages[cur] = []
            fenced.discard(cur)
            fence_lines[cur], plain[cur] = [], []
            infence = False
            continue
        if re.match(r"^#{1,3}\s", ln):
            cur = None
            infence = False
            continue
        if cur is None:
            continue
        s = ln.strip()
        if s.startswith("```"):
            infence = not infence
            if infence:
                fenced.add(cur)
            continue
        if infence:
            if re.match(r"^\s*-\s*资产：", ln):
                continue          # 资产行=机器指令，不做文本审计（继承图文字走 OCR 对账）
            fence_lines[cur].append(ln)
        else:
            if re.match(r"^\s*[-*]\s*\*\*(资产|动作|原标题|这页在说|怎么处理|原稿的图|需要你拍板)\*\*", ln):
                continue
            if "slide0" in ln or "连线标注" in ln:
                continue
            plain[cur].append(ln)
    for no in list(pages):
        pages[no] = fence_lines[no] if no in fenced else plain[no]
    return pages


def page_text(slide):
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
                t = sh.text_frame.text
            except Exception:
                continue
            if t:
                buf.append(t)

    walk(slide.shapes)
    return norm("".join(buf))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--outline", required=True)
    ap.add_argument("--from", dest="lo", type=int, default=1)
    ap.add_argument("--to", dest="hi", type=int, default=999)
    ap.add_argument("--gone", default="", help="已删除的大纲页号，逗号分隔")
    a = ap.parse_args()
    gone = {int(x) for x in re.split(r"[,\s]+", a.gone.strip()) if x}

    prs = Presentation(a.pptx)
    pages = parse_outline(a.outline)
    hi = min(a.hi, max(pages) if pages else a.hi)
    tot = 0
    for no in range(a.lo, hi + 1):
        if no in gone:
            print(f"P{no:02d} 已删除，跳过")
            continue
        idx = no - len([p for p in gone if p < no])
        if idx > len(prs.slides._sldIdLst):
            print(f"P{no:02d} 超出成品页数，跳过")
            continue
        ptext = page_text(prs.slides[idx - 1])
        miss = []
        for ln in pages.get(no, []):
            for raw, n in segments(ln):
                if n not in ptext:
                    miss.append(raw)
        tot += len(miss)
        print(f"P{no:02d} 缺失 {len(miss)}" + (":" if miss else ""))
        for m_ in miss:
            print(f"    - {m_}")
    print("总缺失:", tot)


if __name__ == "__main__":
    main()
