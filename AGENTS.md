# AGENTS.md — 给 AI 编码助手的仓库使用说明

本仓库是一个 **Agent Skills 技能包**（含一个 Claude Code 插件壳）。任何能读写文件、执行 shell 命令的 AI 编码助手（Codex、Cursor、Jules 等）都可以按本说明使用其中的技能。

## 这个仓库是什么

`skills/ppt-retrofit/` 是一个完整的技能：把人写的旧 PPTX 改造成**原生可编辑**的新 PPTX——内容保真（文字一字不减）、结构保真（架构/流程图按原稿处理）、版式升级。技术栈 python-pptx + Pillow，纯文件操作，不依赖任何特定 Agent 平台。

## 如何使用这个技能（通用流程）

1. **先读技能主文档**：`skills/ppt-retrofit/SKILL.md`——里面有完整工作流（双入口：全片重建 / 点名修页）、8 条铁律、分阶段执行与检查点。
2. **按需读参考文档**（都在 `skills/ppt-retrofit/references/`）：
   - 接到任务先读 `decision-tree.md`——决定每页每图"直接用原图 / 整体沿用原稿图形 / 重画"
   - 重排自由发挥页读它的"找分组信号"与"布局编码论点"节
   - 交付前自查读 `self-check.md`（机器对账 + 渲染自查）
   - 踩到奇怪现象先查 `gotchas.md`（29 条实战坑，症状→根因→规避）
3. **脚本用法**（`skills/ppt-retrofit/scripts/`，全部参数化）：

```bash
python3 -I skills/ppt-retrofit/scripts/inventory.py     --pptx 原稿.pptx --out-json inventory.json --deep
python3 -I skills/ppt-retrofit/scripts/dump_slide.py    --pptx 原稿.pptx --slide 11 --pics
python3 -I skills/ppt-retrofit/scripts/graft_deck.py    --src 原稿.pptx --dst 成品.pptx --src-slide 15 [--swap-theme]
python3 -I skills/ppt-retrofit/scripts/check_geometry.py --pptx 成品.pptx
python3 -I skills/ppt-retrofit/scripts/check_text.py    --pptx 成品.pptx --outline outline.md --from 1 --to 41
python3 -I skills/ppt-retrofit/scripts/structure_diff.py slide --src 原稿.pptx --src-slide 23 --dst 成品.pptx --dst-slide 23
python3 -I skills/ppt-retrofit/scripts/visual_check.py  render 成品.pptx --pages 1-99   # 需 LibreOffice
```

4. **重要执行约束**：python 调用一律加 `-I`（隔离模式）；脚本全部参数传路径，不依赖状态文件。

## 环境依赖

- 必需：python3 + python-pptx、Pillow、lxml、numpy
- 可选：PyMuPDF + LibreOffice（渲染自查用；缺失则按 SKILL.md 铁律 4 降级为结构审计并声明）
- CJK 字体（如微软雅黑）在构建机可不装——字体名写入 XML，由最终用户的机器渲染

## 给 Claude Code 用户的快捷安装

```bash
/plugin marketplace add riolg2025/kb-management-rio
/plugin install kb-management-rio@kb-management-rio
```

其他工具（Codex 等）没有插件市场概念：clone 本仓库后，把 `skills/ppt-retrofit/SKILL.md` 当作任务指令的权威来源即可（等价于把上面的通用流程读给你的 Agent 听）。
