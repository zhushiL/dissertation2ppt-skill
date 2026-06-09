# dissertation2ppt

`dissertation2ppt` 是一个面向 Codex 的论文答辩 PPT 生成 skill，用于把中文学位论文材料转换为可编辑、证据充分、结构紧凑的答辩 PowerPoint。它优先处理 Word `.docx` 论文，因为 `.docx` 更容易保留章节、图题、表格、公式和原始插图。

该 skill 主要服务于本科、硕士、博士论文答辩，尤其适合控制科学、自动化、机器人、计算机视觉、人工智能、传感、三维重建、检测、分割等工程技术方向。

## 核心目标

- 生成 16:9 宽屏答辩 PPTX。
- 默认输出 18-25 页紧凑型答辩稿，不为凑页数扩充内容。
- 使用中文答辩结构，包含正式 `目录` 页。
- 以论文原始图表、实验结果、方法框架和结论作为主要证据。
- 尽量保持文本、表格、图示、图表等对象可编辑，避免把整套 PPT 做成截图。
- 使用纯白画布和学校色前景结构，除非用户提供模板并明确要求遵循模板。
- 保持 `总结与展望` 为独立且有内容的章节，最后单独保留 `敬请各位老师批评指正` 结束页。

## 仓库结构

```text
.
├── README.md
└── dissertation2ppt/
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    └── scripts/
        └── validate_dissertation_ppt.py
```

- `dissertation2ppt/SKILL.md`：skill 的主说明文件，也是本仓库的核心内容。
- `dissertation2ppt/agents/openai.yaml`：Codex 中展示和默认提示相关的元数据。
- `dissertation2ppt/scripts/validate_dissertation_ppt.py`：PPTX 结构检查脚本，主要用于显式审查已有文稿或开发调试。

## 安装

把仓库克隆到 Codex skills 目录：

```bash
git clone https://github.com/zhushiL/dissertation2ppt-skill.git ~/.codex/skills/dissertation2ppt
```

如果已经把仓库下载到其他位置，也可以复制 skill 目录：

```bash
mkdir -p ~/.codex/skills
cp -R dissertation2ppt ~/.codex/skills/dissertation2ppt
```

安装后，在 Codex 中用 `$dissertation2ppt` 触发。

## 依赖

- Codex 本地 skills 功能。
- Python 3.10 或更高版本。
- `python-pptx`，用于生成和编辑 PPTX：

```bash
python -m pip install python-pptx
```

仓库内的结构检查脚本只使用 Python 标准库。

## 使用示例

从 Word 论文生成答辩 PPT：

```text
使用 $dissertation2ppt，根据 /path/to/thesis.docx 生成中文硕士论文答辩 PPT。
```

只生成计划并等待确认：

```text
使用 $dissertation2ppt，根据 /path/to/thesis.docx 先生成 output/ppt_plan.md，然后暂停等我确认。
```

按模板风格生成：

```text
使用 $dissertation2ppt，根据 /path/to/thesis.docx 制作答辩 PPT，并参考 /path/to/template.pptx 的视觉模板。
```

## 默认交付物

一次正常运行会把新生成的文件集中放在 `output/` 或用户指定的交付目录中：

```text
output/
├── ppt_plan.md
├── dissertation_defense.pptx
├── asset_manifest.md
└── assets/
    └── figures/
```

- `output/ppt_plan.md` 必须先于 PPTX 生成，用于锁定内容主线、页数、证据层级、视觉系统和构建时质量门。
- `output/dissertation_defense.pptx` 是最终可编辑答辩 PPTX。
- `output/assets/figures/` 存放提取或选用的论文图、辅助图和相关素材。
- `output/asset_manifest.md` 只在使用图像、Logo、截图、论文成果、下载资产或辅助视觉素材时生成，用于记录来源。

skill 不会在正常交付后额外生成例行质量报告，也不会默认启动单独的交付后审查流程。质量控制应在计划和逐页构建阶段完成；只有用户明确要求审查已有 PPT 时，才使用额外校验或审查流程。

## 工作流程

1. 先检查论文、模板和相邻素材，寻找学校 Logo 和可用身份素材。
2. 如果找不到可验证的官方 Logo，暂停并要求用户提供 Logo，除非用户明确同意无 Logo 方案。
3. 提取题目、作者、导师、学院、专业、学位层级、答辩日期、研究问题、技术路线、核心贡献、实验数据、指标、图表、结论和科研成果。
4. 在生成 PPT 前写入 `output/ppt_plan.md`，锁定页数、章节节奏、证据主线、版式系统和构建质量门。
5. 按计划逐页构建，并在每页构建时检查论点、证据、字体、可读性、图表比例、留白和页面标记。
6. 所有构建时质量门满足后，保存最终可编辑 PPTX。

## 版式与内容原则

- 使用 `目录`，不要把目录页写成 `答辩提纲`。
- 内容页使用有意义的编号章节标题，例如 `4 基于条件引导的点云补全方法`。
- 页码统一放在右下角，除非用户模板已有一致且合理的位置规则。
- 不使用持续出现的章节导航条、底部进度导航或重复 `01 02 03 04` 轨道。
- 普通页可编辑文字统一使用 Microsoft YaHei / `微软雅黑`。
- 科研成果页中文使用 SimSun / `宋体`，英文使用 Times New Roman。
- 观众需要阅读的正文字号不低于 16 pt，内容页章节标题不低于 24 pt。
- 方法图、对比图、重建图、分割图、检测图、实验结果和关键表格要足够大，不能只作为装饰。
- 表格应保留关键行列、数据集、指标、基线方法和学生方法，并突出核心结果。
- 禁止伪造学校 Logo、实验数据、论文成果、专利、软著、数据集、合作单位标识或产品标识。

## 开发与维护

- 以 `dissertation2ppt/SKILL.md` 作为 skill 行为规则的源文件。
- 辅助脚本放在 `dissertation2ppt/scripts/`。
- 不要提交真实论文、未公开数据、生成的 `output/`、提取图像、预览图、个人 Codex 运行目录或其他私有材料。
- 修改规则时，应同步检查 README 是否仍然准确描述交付物、流程和限制。

## 检查已有 PPTX

仓库提供的脚本可用于显式检查已有 PPTX 的结构风险：

```bash
python dissertation2ppt/scripts/validate_dissertation_ppt.py output/dissertation_defense.pptx
```

输出 JSON：

```bash
python dissertation2ppt/scripts/validate_dissertation_ppt.py --json output/dissertation_defense.pptx
```

该脚本适合审查或开发调试，不是 `$dissertation2ppt` 正常生成流程中的例行交付后步骤。

## 许可

当前仓库尚未包含 `LICENSE` 文件。正式分发或接受外部贡献前，建议补充明确的开源许可证。
