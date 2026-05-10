# dissertation2ppt

`dissertation2ppt` 是一个面向中文学位论文答辩场景的 Codex Skill，用于把本科、硕士、博士论文材料转换为结构完整、证据充分、可编辑的中文答辩 PowerPoint。

它不是通用 PPT 模板，也不是简单的论文摘要工具。它把论文内容重新组织成适合答辩委员会阅读的叙事：研究问题、技术路线、核心工作、实验依据、成果产出、总结与展望，并尽量保留论文中的原始图表和关键数据。

## Features

- 面向中文高校答辩场景，默认生成简体中文内容。
- 默认输出 18-25 页、16:9、纯白背景的正式答辩 PPTX。
- 优先使用论文原始图表、实验结果、对比表和消融分析，而不是装饰性素材。
- 使用可编辑 PPT 对象承载文本、形状、表格、图示和备注，避免整页截图式输出。
- 在生成 PPT 前先产出 `output/ppt_plan.md`，明确答辩逻辑、页码规划、素材取舍和 QA 风险。
- 支持输出素材清单 `output/asset_manifest.md` 和质量检查报告 `output/qa_report.md`。
- 提供结构化校验脚本，检查页数、比例、白底规则、文本溢出风险、整页截图风险等问题。

## Best For

- 控制科学与工程、自动化、机器人、计算机视觉、人工智能、传感检测、三维重建、图像分割、目标检测等工程技术类论文。
- 需要从 `.docx` 学位论文中提取章节结构、图、表、实验结果和科研成果，并整理为答辩汇报。
- 需要一份可以继续在 PowerPoint / WPS 中人工修改的 `.pptx` 文件。

## Repository Structure

```text
dissertation2ppt/
├── SKILL.md
├── README.md
├── agents/
│   └── openai.yaml
└── scripts/
    └── validate_dissertation_ppt.py
```

- `SKILL.md`: Codex Skill 主说明，定义角色、输出契约、排版规则、内容组织方式和 QA 要求。
- `agents/openai.yaml`: Skill 在 Agent 列表中的展示信息和默认提示词。
- `scripts/validate_dissertation_ppt.py`: 只依赖 Python 标准库的 PPTX 结构校验脚本。

## Installation

将仓库克隆到 Codex skills 目录：

```bash
git clone https://github.com/<your-name>/dissertation2ppt.git ~/.codex/skills/dissertation2ppt
```

如果已经把项目下载到本地，也可以复制目录：

```bash
mkdir -p ~/.codex/skills
cp -R dissertation2ppt ~/.codex/skills/dissertation2ppt
```

然后在 Codex 中使用 `$dissertation2ppt` 触发该 skill。

## Requirements

- Codex，且支持本地 Skills。
- Python 3.10+。
- 用于生成 PPTX 的 `python-pptx`：

```bash
python -m pip install python-pptx
```

校验脚本 `scripts/validate_dissertation_ppt.py` 只使用 Python 标准库，不需要额外安装依赖。

## Usage

推荐提供 Word 版论文，因为 `.docx` 更容易保留标题层级、图表、公式和题注。

示例提示词：

```text
Use $dissertation2ppt to turn /path/to/thesis.docx into an 18-25 slide editable Chinese defense PPTX.
```

如果需要先审阅规划再生成 PPT：

```text
Use $dissertation2ppt to create output/ppt_plan.md from /path/to/thesis.docx first. Stop after the plan and wait for my approval.
```

默认输出：

```text
output/
├── ppt_plan.md
├── dissertation_defense.pptx
├── asset_manifest.md
├── qa_report.md
└── assets/
    └── figures/
```

## Workflow

1. 读取论文材料，提取题目、作者、导师、学院、专业、摘要、目录、章节结构、图表、实验结果和科研成果。
2. 生成 `output/ppt_plan.md`，明确答辩假设、专家诊断、品牌配色、页码大纲、图表计划和 QA 风险。
3. 根据论文章节重组答辩叙事，而不是机械复制论文目录。
4. 生成可编辑 `.pptx`，并尽量用原始论文图表作为证据。
5. 记录使用过的图、表、logo、截图或外部素材。
6. 运行结构化校验，并在可用时进行渲染预览检查。

## Validation

生成 PPTX 后运行：

```bash
python scripts/validate_dissertation_ppt.py output/dissertation_defense.pptx
```

输出 JSON 报告：

```bash
python scripts/validate_dissertation_ppt.py output/dissertation_defense.pptx --json
```

校验脚本会检查：

- 是否为 16:9。
- 页数是否在默认范围内。
- 是否存在非白色整页背景风险。
- 是否存在疑似整页截图。
- 是否缺少原生文本。
- 是否存在文本溢出候选。
- 是否包含不应出现在答辩 PPT 中的内部标记。
- 是否存在明显的底部大面积空白。

## Design Principles

- 一页一个答辩点。
- 证据优先，装饰克制。
- 图表尽量来自论文原文，不伪造实验结果。
- 表格数据尽量完整保留，过密时拆页而不是缩到不可读。
- 封面正式居中，内容页使用学校标识、编号标题和细分割线。
- 默认纯白背景，用校色线条、标签、图示和轻量结构增强层次。
- 详细解释放进 speaker notes，不把生成说明、素材说明或审计信息放到答辩页上。

## Limitations

- 该项目依赖 AI Agent 理解论文内容，最终 PPT 需要人工审阅，尤其是数值、公式、论文成果和作者信息。
- PDF 输入通常不如 `.docx` 稳定，可能需要额外裁剪和人工确认图表内容。
- 学校 logo、官方配色和答辩日期等信息建议由用户明确提供，避免误用。
- 未公开论文、涉密材料或包含敏感数据的论文，应只在你有权限的环境中处理。

## Roadmap

- 增加更多中文高校答辩模板风格示例。
- 增强 `.docx` 图表提取辅助脚本。
- 增加 slide preview 渲染检查工作流。
- 增加常见论文类型的页面规划示例。
- 增加 GitHub Actions 示例，用于校验 skill 文件和脚本。

## Contributing

欢迎提交 issue 和 pull request，尤其是：

- 不同学科论文的答辩结构改进。
- 更严格的 PPTX 结构校验规则。
- 更好的图表提取、裁剪和素材清单流程。
- 中文答辩场景下的版式规范、措辞规范和 QA 清单。

请避免提交真实论文原文、未授权图片、个人隐私信息或涉密材料。

## License

开源前请补充 `LICENSE` 文件，并在此处明确许可证类型。

如果没有特殊限制，建议在 MIT License 或 Apache License 2.0 中选择一个。
