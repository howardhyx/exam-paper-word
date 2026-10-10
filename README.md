# 多学科试卷编写 · exam-paper-word

一个用于 Codex 的中文 Skill：根据学科、学段、考查范围、题型分值和难度要求组卷，输出可编辑的 Word 试题卷及答案与评分标准。

适用于单元检测、阶段考试、综合复习和模拟考试，也可以先讨论命题方案，或调整已有试卷。它不限定为高中、数学或某一种题型。

## 能做什么

- 先建立命题蓝图，再逐题生成题干、解答与采分点，核对考点覆盖及各级分值。
- 分别生成学生试题卷和教师答案卷，保持题号、图表与分值一致。
- 数学计算、推导和证明默认留连续空白；文字回答可用横线，中文作文可用方格，作图可留独立区域。
- 横线按实际栏宽生成，不使用重复下划线；支持 A4 单栏、A3 双栏和用户模板。
- 根据完整解答的预计手写量估算答题空间，渲染后再核对实际容量。

## 安装

把下面这句话发给 Codex：

```text
使用 $skill-installer，安装 https://github.com/howardhyx/exam-paper-word/tree/main/skills/exam-paper-word 中的 Skill。
```

也可以下载本仓库的 ZIP，将 `skills/exam-paper-word` 整个文件夹放到个人技能目录 `~/.agents/skills/`，或项目内的 `.agents/skills/`。最终结构应为 `.agents/skills/exam-paper-word/SKILL.md`，不要只复制 SKILL.md。

某些已有安装环境仍使用 `~/.codex/skills/`；使用该环境的实际技能目录即可，避免重复安装到多个位置。未识别时重新打开 Codex。安装位置以 [Codex 官方 Skill 文档](https://developers.openai.com/codex/skills) 为准。

## 使用示例

```text
使用 $exam-paper-word，制作一份高中数学阶段检测，范围为函数与导数，满分100分，考试90分钟，预计难度中等。
```

```text
使用 $exam-paper-word，按附件教材制作小学英语单元检测。
```

请提供学科、范围、学段、总分、时长及特殊版式要求。题型结构与难度会影响命题结果；有样卷或教材时可一并提供。

## 运行环境

Skill 的文字规则由 Codex 执行，并不是一条命令就能独立完成组卷的程序。

附带的 Word 排版组件需要 Python 3.10 或以上，以及 `python-docx`。两个核查脚本只需要 Python 标准库。如果环境已经提供依赖，可直接使用；否则在自己的 Python 环境中运行：

```text
python -m pip install -r requirements.txt
```

`requirements.txt` 只用于排版组件、测试和示例，不是安装 Skill 的前提。Word 公式转换及逐页渲染需要当前环境另行提供工具；这些工具没有打包进仓库。环境有 documents 文档技能时优先使用其流程。没有可靠渲染器时，应说明未完成页面核验。

## 文件说明

| 路径 | 用途 |
|---|---|
| `skills/exam-paper-word/SKILL.md` | Skill 入口与组卷流程 |
| `skills/exam-paper-word/references/` | 命题蓝图、Word 版式、答题空间规则 |
| `skills/exam-paper-word/scripts/check_blueprint.py` | 核对分值与评分记录 |
| `skills/exam-paper-word/scripts/answer_space.py` | 估算空间并核对容量记录 |
| `skills/exam-paper-word/scripts/word_components.py` | 原生 Word 留白、横线、方格、绘图区和选项组件 |
| `examples/component-demo.docx` | 排版组件样页，包含 A4、跨页和 A3 双栏示例 |
| `examples/build_component_demo.py` | 可复现样页的作者程序 |
| `tests/test_components.py` | 空间与排版结构的功能测试 |

样页用于展示组件，不是一份完整试卷。生成过程中的容量记录属于配置值，不代表已经测得页面上的可写容量。

## 开发与检查

在仓库根目录运行：

```text
python -m unittest discover -s tests -v
python examples/build_component_demo.py
python skills/exam-paper-word/scripts/check_blueprint.py --help
python skills/exam-paper-word/scripts/answer_space.py --help
```

示例会在 `examples/generated/` 生成 Word 和配置记录；渲染这些新生成文件后仍需查看全部页面。修改 Skill 时同时检查入口与参考文件的链接，保持组件接口、文档示例和测试一致。

空间组件会依据已填写的手写量计算尺寸，但不会自动理解一道数学题并推断所有推导步骤。分值和容量脚本核对的是结构与所填数据，不能代替学科内容审核或页面检查。正式使用前应复核题目、答案、评分标准和打印效果；预估难度不等于实测难度。

## 反馈与许可证

欢迎通过 Issues 提交可复现的问题或通过 Pull Request 改进。排版问题请说明纸张、栏数、所用字体和渲染工具，并提供不含学生个人信息的最小示例。

本项目使用 [MIT License](LICENSE)。许可证适用于本仓库内容；用户另行提供的教材、试题、图片和引用材料的权利归原权利人所有。
