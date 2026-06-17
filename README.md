# Codex 科研论文绘图 Skill

当前版本：`v1.1.1`

`Codex 科研论文绘图 Skill` 是一个面向科研论文配图工作流的 Codex Skill。它用于把参考图片、截图或生成图还原为 **Microsoft Visio `.vsdx` 原生可编辑图形**，并从同一个 Visio 源文件导出 PNG、SVG、PDF、PPTX 等交付格式。

核心目标不是把图片贴进 Visio，而是让 Codex 通过 Visio 原生形状、文本、连线、分组和样式重建论文配图、模型框架图、流程图和多面板科学图。

## 适用场景

适合：

- 根据 PNG/JPG/截图重建 Visio 图。
- 将 AI 生成的论文模型图转成可编辑 `.vsdx`。
- 按参考图修改已有 Visio 文件的布局、配色、字体或模块结构。
- 对复杂多面板科学图进行结构化复刻。
- 检查 `.vsdx` 是否误用了整张参考图嵌入。
- 给 Visio 图统一论文风格字体、配色和线条规范。
- 从保存后的 `.vsdx` 导出 SVG、PDF、PPTX 或 PNG。
- 对复杂多面板图先做面板四角/边界标定，减少子模块移位、串区和重叠。

不适合：

- 只需要把图片插入 Visio 页面。
- 只需要普通图片编辑、抠图或美化。
- 不要求 Visio 原生可编辑性的纯位图复刻。

## 核心原则

最终交付的 `.vsdx` 应尽量由以下对象构成：

- Visio 原生矩形、圆形、线条、箭头、连接线。
- 可编辑文本。
- 可编辑分组。
- 原生近似绘制的小图表、热图、节点图、立方体、堆叠图。

禁止用整张参考图作为最终页面内容来冒充还原。参考图只能作为临时描摹依据；最终文件中不应留下完整的大尺寸参考 PNG/JPG。

`.vsdx` 是可编辑母版。SVG/PDF/PPTX 是从这个母版导出的交付物，不应该单独重画出彼此不一致的版本。

## 仓库结构

```text
.
├── README.md
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
│   └── rebuild-guidelines.md
└── scripts/
    ├── visio_export_formats.ps1
    ├── visio_page_tools.ps1
    └── visio_rebuild_scaffold.ps1
```

文件说明：

- `SKILL.md`：Codex Skill 的主入口，包含触发描述、工作流、验收标准和安全规则。
- `agents/openai.yaml`：Codex UI 元数据。
- `references/rebuild-guidelines.md`：复杂科学图还原准则，包括面板拆解、绘图顺序、样式参数、导出策略和验证 rubric。
- `scripts/visio_export_formats.ps1`：可复用导出函数，支持 PNG、SVG、PDF、PPTX。
- `scripts/visio_page_tools.ps1`：辅助检查脚本，用于备份、导出、检查 `.vsdx` 包结构。
- `scripts/visio_rebuild_scaffold.ps1`：Visio 原生绘图脚手架，用于快速编写一比一重建脚本，并内置全局坐标和面板局部坐标 helper。

## 环境要求

推荐环境：

- Windows。
- Microsoft Visio。
- PowerShell。
- Microsoft PowerPoint，用于 PPTX 导出。
- Git。
- Codex Desktop 或支持本地文件与工具调用的 Codex 环境。

说明：

- 完整 Visio 自动绘图依赖 Visio COM Automation，因此主要面向 Windows + Microsoft Visio。
- SVG 和 PNG 由 Visio 页面导出。
- PDF 由 Visio 固定格式导出。
- PPTX 默认由 PowerPoint COM 创建单页演示文稿，并插入 Visio 导出的 SVG 页面渲染。
- 不安装 Visio 时，仍可做 `.vsdx` 包结构检查或有限 XML 修改，但不适合完整一比一重建。

## 安装方式

将本仓库克隆或复制到 Codex skills 目录。

Windows 示例：

```powershell
git clone https://github.com/pengjunchi0/codex-visio-paper-figure-skill.git "$env:USERPROFILE\.codex\skills\visio-image-rebuilder"
```

安装后重启 Codex 或开启新会话，使 skill 被重新发现。

## 推荐使用方式

示例请求：

```text
使用 visio-image-rebuilder，根据这张参考图片重建 C:\path\model.vsdx，要求最终是 Visio 原生可编辑形状，不要整图嵌入，并导出 SVG、PDF、PPTX。
```

```text
把这个 .vsdx 按参考图更换配色，保持布局不变，最终仍然可编辑，并给我一个 PDF 预览和 PPTX。
```

```text
检查这个 .vsdx 是否只是嵌入了整张 PNG，如果是，请改成原生 Visio 形状重建，再导出 SVG。
```

## 面板标定与防重叠

v1.1.1 增加了面向复杂多面板图的坐标约束建议和脚手架 helper。推荐流程是：

1. 先标定整张参考图尺寸和 Visio 页面尺寸。
2. 再标定每个主要 panel 的左上角、宽高，必要时记录四角点。
3. panel 内部元素使用 0-1 局部坐标绘制，而不是直接手写全图坐标。
4. 导出预览后检查子模块是否越出父 panel、相邻 panel 是否重叠、箭头和文字是否穿过无关模块。

`visio_rebuild_scaffold.ps1` 中已提供：

- `RectRel`
- `TextRel`
- `OvalRel`
- `LineRel`
- `Assert-RelBox`
- `Assert-RelPoint`

这些 helper 会把局部坐标映射回全局参考坐标，并在局部元素越出 panel 边界时直接报错，避免复杂图后半部分出现整体移位或重叠。

## 多格式导出

导出已有 `.vsdx`：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_page_tools.ps1 `
  -VsdxPath "C:\path\model.vsdx" `
  -ExportFormats svg,pdf,pptx `
  -OutputDir "C:\path\exports" `
  -InspectPackage
```

重建并导出：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_rebuild_scaffold.ps1 `
  -VsdxPath "C:\path\model.vsdx" `
  -PageW 16 `
  -PageH 9 `
  -RefW 1600 `
  -RefH 900 `
  -PreviewPath "C:\path\exports\model.png" `
  -ExportFormats svg,pdf,pptx `
  -OutputDir "C:\path\exports"
```

## 验收标准

一个合格的 Visio 还原结果应满足：

- 主体布局和参考图一致。
- 主要模块、标题、编号、箭头和说明文字齐全。
- 文字可编辑。
- 图形对象可单独选中和修改。
- 没有整张参考图作为最终底图。
- 配色、字体和线条风格统一。
- 复杂多面板图的内部元素不应明显移位、跨 panel 串区或互相重叠。
- 有原文件备份。
- 请求的 PNG/SVG/PDF/PPTX 从同一个保存后的 `.vsdx` 导出，并且文件非空。

## Version History

### v1.1.1 - Panel calibration and anti-overlap

- 在 `SKILL.md` 中加入复杂多面板图的面板标定流程。
- 在 `references/rebuild-guidelines.md` 中新增 panel-local 坐标、四角标定和防重叠检查准则。
- 在 `visio_rebuild_scaffold.ps1` 中新增 `RectRel`、`TextRel`、`OvalRel`、`LineRel` 等局部坐标 helper。
- 新增 `Assert-RelBox` 和 `Assert-RelPoint`，当局部元素越出父 panel 边界时提前报错。
- 更新 `agents/openai.yaml` 和 README，说明 v1.1.1 的防移位、防重叠能力。

### v1.1 - Multi-format outputs

- 新增 `scripts/visio_export_formats.ps1`。
- `visio_page_tools.ps1` 支持 `-ExportFormats png,svg,pdf,pptx`、`-OutputDir`、`-OutputBaseName`。
- `visio_rebuild_scaffold.ps1` 支持重建后直接导出 PNG/SVG/PDF/PPTX。
- `SKILL.md` 和 `references/rebuild-guidelines.md` 增加多格式导出策略和验收标准。

### v1.0 - Initial usable version

- 建立 Codex Skill 基础结构。
- 明确核心规则：最终 `.vsdx` 应由 Visio 原生可编辑形状构成，不能用整张参考图嵌入冒充还原。
- 提供 Visio COM 绘图脚手架和 `.vsdx` 检查工具。
- 提供复杂多面板论文图的绘图流程设计。

### Planned v1.2

- 增加更多可复用 motif helper，例如 `DrawCube`、`DrawHeatmap`、`DrawGraph`、`DrawStackedSequence`、`DrawMiniChart`。
- 增加参考图尺寸读取脚本。
- 增加 shape inventory 导出脚本，用于分析已有 `.vsdx` 的文本、颜色、位置和分组。
