# kb-management-rio

知识库维护技能集插件。把知识库运营中的重复性工作沉淀为可复用的 Agent 技能，首个技能是 **ppt-retrofit（旧 PPT 原生改造）**。

## ppt-retrofit — 旧 PPT 原生改造

把人写的旧 PPTX 改造成**原生可编辑**的新 PPTX：内容保真（文字一字不减、图片按类别处理）、结构保真（架构/流程图严格按原稿或整体沿用）、版式升级（自由发挥页重新设计）。技术栈 python-pptx + Pillow，自包含，不依赖其他插件。

**两个入口**：

- **A 全片重建**：整套 PPT 重做。流程 = 盘点（inventory）→ 逐页大纲（用户确认）→ 风格样张（用户确认）→ 资产管线 → 分批重建 → 渲染自查 → 交付
- **B 点名修页**：用户点名某几页丑/错。定位页 → 按决策树重排 → 回归校验 → 交付

**核心机制**：

- 分类决策树：每页每图在"直接用原图 / 整体沿用原稿图形（XML 移植）/ 重画"三档里选，复杂图形默认整体沿用，重画是例外
- 自查体系：文字对账 + 几何检查 + 溢出启发式 + 结构对账四道机器闸，外加 LibreOffice 渲染自查（沿用页像素比对、重画页并排目检），把图形错误挡在交付前
- 坑清单 29 条 + 经验规律 6 条 + 待优化清单，全部来自真实项目实战（41 页企业方案 PPT 全片重建）

**脚本 9 个**（全部参数化，`python3 -I scripts/<脚本>.py --pptx <文件>`）：inventory / extract_assets / dump_slide / graft_deck / check_geometry / check_overflow / check_text / structure_diff / visual_check

## 环境要求

- python3 及 python-pptx、Pillow、lxml、numpy（文本/几何/结构类脚本必需）
- PyMuPDF + LibreOffice Impress：仅渲染自查需要；缺失时技能自动降级为结构审计并在交付说明中声明
- 微软雅黑等客户字体在构建机可不装（字体名写进 XML，在用户机器上渲染）

## 目录结构

```
.claude-plugin/plugin.json    # 插件清单
skills/ppt-retrofit/          # 技能本体（SKILL.md + references 5 篇 + scripts 9 个）
docs/optimization-backlog.md  # 经验规律与待优化清单（测试期持续沉淀）
```
