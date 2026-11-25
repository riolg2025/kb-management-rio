# kb-management-rio

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](#requirements)
[![Claude Code](https://img.shields.io/badge/Claude%20Code-plugin-8B5CF6.svg)](#installation)
[![Agents](https://img.shields.io/badge/AGENTS.md-compatible-10B981.svg)](AGENTS.md)

> A knowledge-base maintenance skill collection for AI coding agents.
> Its first skill, **ppt-retrofit**, retrofits legacy, hand-made PowerPoint decks into **fully editable, native PPTX** — content- and structure-faithful, layout upgraded.

**中文文档见下文 ⬇️**

---

## English

### Why ppt-retrofit

Old decks usually die in one of three ways when they get rebuilt:

1. **Pasted as images** — nothing is editable anymore.
2. **Rewritten by AI** — the original argument and meaning quietly drift.
3. **"Redrawn" architecture diagrams** — dozens of shapes, connectors and nested groups get mistranslated into something that looks similar but means something else.

ppt-retrofit takes fidelity as the first constraint, not the afterthought:

- **Text is carried over verbatim** — machine-audited page by page, zero missing runs allowed.
- **Complex graphics are transplanted, not redrawn** — XML-level graft preserves geometry, styles and semantics byte-for-byte.
- **Only genuinely free-form pages get redesigned** — and even then the argument must survive (every page has a confirmed intent sentence).
- **Machine audits + pixel-level render checks stand guard before delivery** — text reconciliation, geometry, overflow, structural reconciliation, plus real LibreOffice rendering with pixel diff and side-by-side review.

Battle-tested: a 41-page enterprise solution deck was rebuilt end-to-end with this pipeline — every one of the 41 pages machine-audited, rendered, and visually reviewed before handoff.

### What's inside

| Component | What it does |
|---|---|
| **SKILL.md** | The skill itself: two entries, 8 iron rules, Phase 0–5 workflow with user-confirmation gates |
| `references/decision-tree.md` | Per-page / per-graphic decision system: use original image / **graft as-is** / redraw |
| `references/outline-template.md` | Two-layer outline (human-readable + machine block) as the single source of truth |
| `references/self-check.md` | Four-gate quality system: machine audits + render self-check |
| `references/design-system.md` | Design tokens, font scale, page-type library, `deck_lib` helper template |
| `references/gotchas.md` | **29 documented gotchas** from real projects — symptom → root cause → prevention |
| `scripts/` | **9 parameterized Python scripts** covering the whole pipeline |
| `docs/optimization-backlog.md` | 6 distilled experience rules + work-item backlog |

### The decision system (core idea)

Every page and every graphic falls into exactly one of three tiers — and the default is the safest one:

| Tier | When | Mechanism |
|---|---|---|
| **bitmap** | UI screenshots, certificates, photo-rich complex diagrams | Use the original image directly |
| **graft (XML)** | Complex logic diagrams, nested-group topologies, UI collages — **the default for complex graphics** | `graft_deck.py` deep-copies shapes at XML level, re-links images by SHA-1, enforces fonts, optionally swaps the theme |
| **redraw** | Simple, countable structures (resource trees, short chains, rings) | Rebuild natively — only with a verified spec sheet |

Redrawing is the exception and needs a reason. Grafting is immune to "drawn wrong" by construction.

### Pipeline at a glance

| Phase | What happens | Gate |
|---|---|---|
| 0 | Inventory the source deck (texts incl. tables/groups/SmartArt, images with SHA-1, theme) | G0 |
| 1 | Page-by-page outline — intent sentence per page, user confirms handling | **G1 (user)** |
| 2 | Style sample (3–4 pages) + design questionnaire | **G2 (user)** |
| 3 | Asset pipeline (extract, dedupe by SHA-1, convert) | — |
| 4 | Batch rebuild → machine audits → render self-check | G4 (all green) |
| 5 | Delivery: acceptance table, per-page handling statement, risk notes | — |

### Installation

**Claude Code** (one-liner):

```
/plugin marketplace add riolg2025/kb-management-rio
/plugin install kb-management-rio@kb-management-rio
```

**Other agents (Codex, Cursor, Jules, …)**: clone this repo and point your agent at [AGENTS.md](AGENTS.md) — it explains the skill layout and script usage in tool-agnostic terms. There is no plugin-market concept outside Claude Code, but the skill is plain markdown + Python and works with any agent that can read files and run commands.

### Quick start

```bash
git clone https://github.com/riolg2025/kb-management-rio.git
cd kb-management-rio

# 1. Inventory the source deck
python3 -I skills/ppt-retrofit/scripts/inventory.py --pptx /path/to/legacy.pptx --out-json inventory.json --deep

# 2. Dump one slide as a rebuild spec sheet (geometry, fills, connector styles, group transforms)
python3 -I skills/ppt-retrofit/scripts/dump_slide.py --pptx /path/to/legacy.pptx --slide 11 --pics

# 3. Graft a complex diagram page into your new deck, as-is
python3 -I skills/ppt-retrofit/scripts/graft_deck.py \
    --src /path/to/legacy.pptx --dst new.pptx --src-slide 15 --dst-slide 15 --swap-theme

# 4. Audit the result
python3 -I skills/ppt-retrofit/scripts/check_geometry.py --pptx new.pptx
python3 -I skills/ppt-retrofit/scripts/check_text.py --pptx new.pptx --outline outline.md --from 1 --to 41

# 5. Render self-check (needs LibreOffice)
python3 -I skills/ppt-retrofit/scripts/visual_check.py render new.pptx --pages 1-41
python3 -I skills/ppt-retrofit/scripts/visual_check.py diff /path/to/legacy.pptx new.pptx --pages 15,18   # grafted pages
python3 -I skills/ppt-retrofit/scripts/visual_check.py sbs   /path/to/legacy.pptx new.pptx --pages 21,27   # redrawn pages
```

For full-deck rebuilds, read `skills/ppt-retrofit/SKILL.md` first — the outline (Phase 1) and style sample (Phase 2) **require user confirmation** by design.

### Scripts reference

| Script | Purpose |
|---|---|
| `inventory.py` | Per-slide inventory: texts (incl. tables, groups, SmartArt), images with SHA-1, theme colors, shape stats |
| `extract_assets.py` | Extract images to an asset dir, dedupe by SHA-1, TIFF→PNG (hash computed after conversion) |
| `dump_slide.py` | Dump one slide as a rebuild spec: shapes with geometry/fill, connectors with arrows/dash/bend types, nested-group transforms, pictures with SHA-1 |
| `graft_deck.py` | Transplant original slide shapes verbatim (XML deepcopy + SHA-1 image re-link + font enforcement + optional theme swap) |
| `check_geometry.py` | Out-of-bounds + textbox overlap (rotation-aware) |
| `check_overflow.py` | CJK text-overflow heuristic |
| `check_text.py` | Outline vs deck text-completeness audit (per-page substring matching) |
| `structure_diff.py` | Structural reconciliation: `slide` mode (blocks/bands/connectors) or `ocr` mode (source bitmap OCR vs deck text) |
| `visual_check.py` | Render QA via LibreOffice: `render` / `diff` (pixel threshold) / `sbs` (side-by-side review) |

### Requirements

- **Required**: Python 3.9+ with `python-pptx`, `Pillow`, `lxml`, `numpy`
- **Optional** (render self-check): `PyMuPDF` + LibreOffice Impress (`apt install libreoffice-impress`)
- CJK fonts (e.g. Microsoft YaHei) don't need to exist on the build machine — font names are written into the XML and rendered on the end user's machine

### Known limitations

- LibreOffice renders SmartArt as a fallback layout (overlapping boxes) — layout is not comparable, but the extracted text is; audit those pages by text reconciliation instead
- GitHub-hosted environments without LibreOffice: the render gate is skipped; the delivery statement then explicitly says "no render check this round, structural audit only"

### Contributing

Issues and PRs welcome. When you hit a new pitfall, the convention is: document it in `references/gotchas.md` (symptom → root cause → prevention), distill it into a rule in `docs/optimization-backlog.md`, then attach a work item.

### License

[MIT](LICENSE)

---

## 中文

### 为什么做 ppt-retrofit

旧 PPT 重建通常死于三种方式：

1. **贴成图片**——完全不可编辑；
2. **AI 重新创作**——原稿论点和语义悄悄漂移；
3. **"重画"架构图**——几十个形状、连接线、嵌套组被翻译成"看起来像、意思变了"的东西。

ppt-retrofit 把保真当作第一约束，而不是事后补救：

- **文字一字不减**——逐页机器对账，缺失必须为零；
- **复杂图形移植而非重画**——XML 级搬运，几何、样式、语义原样保真；
- **只有真正自由的页面才重新设计**——且每页必须有经用户确认的"这页在说什么"意图句，论点不能丢；
- **交付前机器对账 + 像素级渲染自查双保险**——文字完整性、几何、溢出、结构对账，外加真实 LibreOffice 渲染的像素比对和并排目检。

实战验证：一套 41 页企业方案 PPT 用本流程全片重建——41 页逐页机器审计、渲染出图、目检后交付。

### 仓库内容

| 组成 | 说明 |
|---|---|
| **SKILL.md** | 技能主文档：双入口、8 条铁律、Phase 0–5 工作流与用户确认关卡 |
| `references/decision-tree.md` | 逐页逐图决策体系：直接用原图 / **整体沿用（XML 移植）** / 重画 |
| `references/outline-template.md` | 双层大纲（用户层 + 机器层），唯一事实源 |
| `references/self-check.md` | 四道闸质量体系：机器对账 + 渲染自查 |
| `references/design-system.md` | 设计 tokens、字阶、页型库、`deck_lib` 助手模板 |
| `references/gotchas.md` | **29 条实战坑清单**——症状 → 根因 → 规避 |
| `scripts/` | **9 个参数化 Python 脚本**，覆盖全流程 |
| `docs/optimization-backlog.md` | 6 条经验规律 + 待优化工作项 |

### 决策体系（核心思想）

每页每图必落三类之一，默认选最安全的一档：

| 档位 | 适用 | 机制 |
|---|---|---|
| **bitmap（原图）** | UI 截图、证书、含照片的复杂图 | 直接使用原稿图片 |
| **graft（XML 移植）** | 复杂逻辑图、嵌套组拓扑、界面拼贴——**复杂图形的默认动作** | `graft_deck.py` XML 级深拷贝、图片按 SHA-1 重挂、字体强制、可选主题替换 |
| **redraw（重画）** | 结构清晰可数的简单图（资源树、短流程链、环图） | 原生重建——必须有逐块核对过的规格书 |

重画是例外、要给理由；移植从构造上就不会"画歪"。

### 流程一览

| 阶段 | 做什么 | 关卡 |
|---|---|---|
| 0 | 原稿盘点（文本含表格/组合/SmartArt、图片带 SHA-1、主题） | G0 |
| 1 | 逐页大纲——每页意图句，处理方式经用户确认 | **G1（用户）** |
| 2 | 风格样张（3–4 页）+ 设计问卷 | **G2（用户）** |
| 3 | 资产管线（提取、按 SHA-1 去重、格式转换） | — |
| 4 | 分批重建 → 机器对账 → 渲染自查 | G4（全绿） |
| 5 | 交付：验收对照表、逐页处理方式声明、风险提示 | — |

### 安装

**Claude Code**（一键）：

```
/plugin marketplace add riolg2025/kb-management-rio
/plugin install kb-management-rio@kb-management-rio
```

**其他 Agent 工具（Codex / Cursor / Jules 等）**：clone 本仓库，让 Agent 读 [AGENTS.md](AGENTS.md)（跨工具通用说明）。Claude Code 之外没有插件市场概念，但本技能就是纯 markdown + Python，任何能读写文件、执行命令的 Agent 都能用。

### 快速上手

```bash
git clone https://github.com/riolg2025/kb-management-rio.git
cd kb-management-rio

# 1. 盘点原稿
python3 -I skills/ppt-retrofit/scripts/inventory.py --pptx 旧稿.pptx --out-json inventory.json --deep

# 2. 把某一页导出为重建规格书（几何/填充/连接线样式/组变换）
python3 -I skills/ppt-retrofit/scripts/dump_slide.py --pptx 旧稿.pptx --slide 11 --pics

# 3. 把复杂图形页原样移植进新稿
python3 -I skills/ppt-retrofit/scripts/graft_deck.py --src 旧稿.pptx --dst 新稿.pptx --src-slide 15 --dst-slide 15 --swap-theme

# 4. 审计成品
python3 -I skills/ppt-retrofit/scripts/check_geometry.py --pptx 新稿.pptx
python3 -I skills/ppt-retrofit/scripts/check_text.py --pptx 新稿.pptx --outline outline.md --from 1 --to 41

# 5. 渲染自查（需 LibreOffice）
python3 -I skills/ppt-retrofit/scripts/visual_check.py render 新稿.pptx --pages 1-41
python3 -I skills/ppt-retrofit/scripts/visual_check.py diff 旧稿.pptx 新稿.pptx --pages 15,18   # 移植页
python3 -I skills/ppt-retrofit/scripts/visual_check.py sbs   旧稿.pptx 新稿.pptx --pages 21,27   # 重画页
```

整片重建请先读 `skills/ppt-retrofit/SKILL.md`——按设计，大纲（Phase 1）和风格样张（Phase 2）**必须经用户确认**才能继续。

### 脚本速查

| 脚本 | 用途 |
|---|---|
| `inventory.py` | 逐页盘点：文本（含表格/组合/SmartArt）、图片带 SHA-1、主题色、形状统计 |
| `extract_assets.py` | 图片提取入库：按 SHA-1 去重、TIFF→PNG（哈希取转换后字节） |
| `dump_slide.py` | 单页导出重建规格：形状几何/填充、连接线箭头/线型/折线、嵌套组变换、图片 SHA-1 |
| `graft_deck.py` | 原稿页形状原样移植（XML 深拷贝 + SHA-1 图片重挂 + 字体强制 + 可选主题替换） |
| `check_geometry.py` | 越界 + 文本框交叠（旋转感知） |
| `check_overflow.py` | 中文溢出启发式检查 |
| `check_text.py` | 大纲 vs 成品文本完整性审计（逐页子串匹配） |
| `structure_diff.py` | 结构对账：`slide` 模式（块/带/连线）或 `ocr` 模式（原稿位图 OCR 对成品文本） |
| `visual_check.py` | LibreOffice 渲染质检：`render` / `diff`（像素阈值）/ `sbs`（并排目检） |

### 环境要求

- **必需**：Python 3.9+，及 `python-pptx`、`Pillow`、`lxml`、`numpy`
- **可选**（渲染自查）：`PyMuPDF` + LibreOffice Impress（`apt install libreoffice-impress`）
- 中文字体（如微软雅黑）构建机可不装——字体名写入 XML，由最终用户的机器渲染

### 已知限制

- LibreOffice 对 SmartArt 是回退渲染（盒子乱叠）——布局不可比，但提取文本可对账，这类页请走文字逐条核对
- 无 LibreOffice 的环境会跳过渲染闸，交付声明会明写"本轮无渲染目检，采用结构审计"

### 参与贡献

欢迎 Issue 和 PR。踩到新坑的沉淀约定：先写进 `references/gotchas.md`（症状 → 根因 → 规避），再提炼为 `docs/optimization-backlog.md` 里的规律并挂工作项。

### 许可证

[MIT](LICENSE)
