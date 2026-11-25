---
name: ppt-retrofit
description: 改造/重做已有旧 PPTX 为原生可编辑新 PPTX（保真重做）。触发词：改造旧 PPT、重做这份 PPT、把方案 PPT 重排、原生可编辑重做、丑页重排、旧 PPT 翻新。反指示：从文档新建 PPT 走 pagecraft（html-slide-gen / ppt-generator）；只改几个字/换张图的小编辑直接 python-pptx 改，不必走本流程。
---

# ppt-retrofit — 旧 PPT 原生改造

把人写的旧 PPTX 改造成**原生可编辑**的新 PPTX：内容保真（文字一字不减、图按类别处理）、结构保真（架构/流程图严格照原稿）、版式升级（自由发挥页重新设计）。技术栈 python-pptx + Pillow，自包含，不依赖任何其他插件。

## 两个入口

| 入口 | 何时用 | 流程 |
|---|---|---|
| **A 全片重建** | 整套 PPT 重做 | Phase 0→5 全走，**G1 大纲 gate 和 G2 样张 gate 强制** |
| **B 点名修页** | 用户点名某几页丑/错 | 跳过 G1/G2：定位页（见铁律 5）→ 按决策树重排 → 校验三件套回归 → 交付 |

## 铁律（MUST，违反任何一条都构成返工风险）

1. **大纲 = 唯一事实源**。全片重建必须经用户确认大纲后才动工；确认结果写进大纲"确认记录"节。
2. **分类决策必读 `references/decision-tree.md`**；拿不准某页/某图归哪类 → 停下来问用户，不猜。
3. **重绘前必读原稿图**：提取出的原稿图直接 Read 看图写规格，禁止只凭文字提取重建。
4. **自查五件套 + 校验三件套全绿才许交付**（`references/self-check.md`）；**有渲染器（LibreOffice）时渲染自查是第四道闸，必须做**（沿用页 diff ≤6% + 重画页 sbs 并排目检）；无渲染器先实测 `sudo -n whoami` 可装则装（`apt install libreoffice-impress`，有本地离线 deb 缓存优先用缓存），实在装不了才降级：交付说明声明"本轮无渲染目检，采用结构审计"，并给高风险页 Top-5。
5. **用户报的页码 = 成品物理页序**，与原稿页号可能错位（删除页后差 1）。对齐方式：从成品逐页提取 ≥20pt 标题对照大纲，**禁止推算**。
6. **全程说用户语言**：大纲、确认问题、交付说明一律用 `references/outline-template.md` 词汇表左列的说法，内部术语（tokens/SIZE_MAP/strict 等）不出现在用户可见文本里。
7. **执行中发现与大纲标记不符**（如标"重绘"实为截图）→ 回大纲改标记、重新经用户确认，不悄悄改。
8. 分批构建，每批全绿再交用户翻看；改一批回归一批。

## Phase 0 · 盘点（gate G0）

```bash
python3 -I scripts/inventory.py --pptx <原稿.pptx> --out-json <out>/inventory.json [--slides 1-55] [--deep]
```

- 提取：逐页文本（含**表格、组合递归、graphicFrame/SmartArt**）、图片清单（格式/像素/sha1）、主题色、字体、形状统计
- ⚠️ 判"空页"必须先看 inventory 的 smartart 字段——SmartArt 文字不进普通文本提取
- 人工补：逐页粗读原稿，给每页标分类初判（见决策树）
- **G0**：inventory.json 存在，每页有文本/图/SmartArt 之一，或已确认真空页

## Phase 1 · 逐页大纲（gate G1 用户确认）

按 `references/outline-template.md` 写 outline.md：

- 每页必含意图句"**这页在说什么**"（一句大白话，自由发挥页的命根子）
- 用户层（怎么处理/原稿的图/需要你拍板的）+ 机器层（```machine 围栏，脚本消费）
- **分组确认**：照画页、存疑图、删除页逐条封闭式确认；自由发挥页批量带过
- **G1**：用户明确确认分类标记、删除页、意图句后才进 Phase 2

## Phase 2 · 风格样张（gate G2 用户确认）

- 色彩 bootstrap：从原稿 theme 提品牌色（`references/design-system.md`）；字阶用技能默认档位，不加档
- 产出 **3–4 页样张 mini PPTX**（封面/章节页/内容页/继承图页各一）+ 一页纸设计规范
- 风格问卷一次问完：**字体遵用户约定（必问，不从原稿主题推断，问一次记死）**？深色页要不要？页脚要不要？图标风格？页眉 logo 位置？
- **G2**：样张确认后**锁定 tokens + helper 库**；铺量阶段只填内容不改体系，要改回 Phase 2 改样张

## Phase 3 · 资产管线

- 图片提取到 `<out>/images/`，命名 `pNNN_k.ext`（NNN=大纲页号，k=页内序号，用页号映射表防错位）
- TIFF→PNG、大图压缩、按 sha1 去重（同文件多引用 → 裁剪/放大变体，不重复入库）

## Phase 4 · 分批重建（gate G4 校验全绿）

- 拷 `references/design-system.md` 的 deck_lib 模板，改 tokens 区；每批一个 deck_pNN_MM.py + build_all 总装
- 照画页：先跑 `dump_slide.py` 拿原稿结构清单当规格书（含连接线箭头/线型/折线、组变换）；重建后 `structure_diff.py --mode slide` 对账
- **复杂图形（嵌套组拓扑/界面拼贴/严格架构图）默认 `graft_deck.py` 整体沿用**，重画是例外（判定见 decision-tree.md 图级档位）
- 表格构建后必调 `clean_empty_runs()`（deck_lib 模板；空单元格无文本 run 会把 LibreOffice 行高撑到 2 倍，gotcha #26）
- 页型从页型库选（布局编码论点，见 decision-tree.md"自由发挥"节）
- 每批：build → 三件套全绿 → 渲染自查（diff/sbs）→ 交用户翻看

## Phase 5 · 交付

- 全量回归：三件套 + structure_diff 对账 + 渲染自查（沿用页 diff、重画页 sbs 逐页过目）
- 产出验收对照表（照画页逐块 ✓/✗）+ **每页处理方式声明（沿用/重画/重排）** + 高风险页 Top-5 + 密度档位声明 + 衍生文件登记清单（见 self-check.md"交付包里必须有"）
- **命名检查**：文件名不得出现"AI"等机器感字样
- 交付必须附可点击链接（相对空间路径）；若入知识库，落盘挂载点即入库

## 校验脚本速查

```bash
python3 -I scripts/check_geometry.py  --pptx 成品.pptx                 # 越界+文本框交叠
python3 -I scripts/check_overflow.py  --pptx 成品.pptx                 # CJK 溢出启发式
python3 -I scripts/check_text.py      --pptx 成品.pptx --outline outline.md \
                                      --from 1 --to 55 --gone 37       # 文本保真审计
python3 -I scripts/structure_diff.py  slide --src 原稿.pptx --src-slide 23 \
                                      --dst 成品.pptx --dst-slide 23   # 照画页结构对账
python3 -I scripts/structure_diff.py  ocr   --image p031_1.png \
                                      --dst 成品.pptx --dst-slide 31   # 原稿图 OCR 对账
python3 -I scripts/dump_slide.py      --pptx 原稿.pptx --slide 11 [--pics] [--brief]
                                      # 照画页规格提取（形状/连接线样式/组变换/图片sha）
python3 -I scripts/graft_deck.py      --src 原稿.pptx --dst 成品.pptx \
                                      --src-slide 15 [--dst-slide 15] [--exclude-text 页题,页脚] \
                                      [--swap-theme]                   # 复杂图形整体沿用（XML 移植）
python3 -I scripts/visual_check.py    render 成品.pptx --pages 1-99    # 渲染出图（需 LibreOffice）
python3 -I scripts/visual_check.py    diff   原稿.pptx 成品.pptx --pages 15,18   # 沿用页像素比对（阈值 6%）
python3 -I scripts/visual_check.py    sbs    原稿.pptx 成品.pptx --pages 21,27   # 重画页并排目检
```

词表生长循环：check_text 报缺失 → 人工核对是"真缺失"还是"版式描述词/继承图内文字" → 后者进 STOP/PROC。每项目都要养词表，别指望种子词表一步到位。

## 环境注意

- `python3 -I` 必加；脚本全部参数化传路径，不依赖状态文件
- 渲染自查需要 LibreOffice：先实测 `sudo -n whoami`（别信上一轮"无渲染器"的旧结论），可装则 `apt install libreoffice-impress`（有本地离线 deb 缓存优先用缓存）；实在装不了才降级为结构审计并按铁律 4 声明
- 审计脚本读用户上传的文件时保持 `-I`
