# 答题空间记录与 Word 排版组件

## 先选择作答方式

根据实际书写任务而不是“简答题”等题型名称选用组件。计算、推导、证明默认用 `blank` 连续空白；文字回答按学科习惯用 `blank` 或 `lined`；中文作文可用 `grid`；作图用 `drawing`，题目需要时再插入所需坐标框。外语写作不默认采用中文作文格。用户模板和明确要求优先。

组件只生成答题区和选项，不替代整卷模板、公式转换、题号管理或页面核验。它们不往学生卷写入答案、容量估计或其他工作标签。

## 逐小问记录空间

在同一题目记录的 `answer_spaces` 数组中，按实际作答任务记录独立空间，`id` 对应题号/小问。同一道作文的多个评分维度共用一条空间记录；一个小问若有分开的文字和作图区，可有两条不同 `id` 的记录。作答在答题卡、题干填空或其他位置完成且不需独立区域时，显式用空数组，不擅自增加区域。记录不依附于分值单位，避免重复统计。

计划阶段可以先估计，完整解答完成后必须更新。独立作答区的记录包含：

| 字段 | 含义 |
|---|---|
| `id`、`mode` | 对应小问或独立区域；`blank`、`lined`、`grid`、`drawing`，无需区域可用 `none` |
| `basis` | 书写量依据，例如“列式1行、变形3行、结论1行；宽分式需加高” |
| `width_mm` | 当前栏和缩进确定后的区域宽度；作文格为其最大可用宽度 |
| `estimated_lines`、`line_height_mm` | 空白/横线区域的预计手写行数和行高；行高默认9毫米，可按学段和公式调整 |
| `reserve_lines` | 改写、推导等额外余量，不机械规定每题相同 |
| `drawing_height_mm` | 绘图区高度；纯作图区必填，文字区可另加绘图高度 |
| `estimated_chars`、`reserve_chars`、`cell_size_mm` | 作文格的预计字数、额外格数和方格边长；标点等占格也计入容量 |
| `allocated_height_mm` | 排版分配的高度；省略时按最低容量计算，可增加，不能低于需求 |
| `observed` | 渲染后的实际容量记录，见下文；不能直接把组件返回值当作实际测量 |

空白/横线区最低高度 =（预计手写行数 + 余量行数）× 行高 + 绘图高度。依据实际宽度估计行数，不使用教师答案的电脑排版行数，也不按分值乘固定系数。大分式、分类讨论和宽表达式分别考虑行高及换行；改变栏宽或缩进后重新估计书写量。

作文格每行格数 = 向下取整（可用宽度 ÷ 格子边长）；行数向上取整以满足预计字数及余量。格子保持正方形，右侧剩余不足一格的空间可留空，不套用横线铺满宽度的规则。纯绘图区按作图步骤和实际图形尺寸估计，不能被文字行数替代。

例子：数学推导预计5行，行高10毫米，另留1行改写余量，总高度至少60毫米；采用连续空白，不加横线。

```json
{
  "id": "12(1)",
  "mode": "blank",
  "basis": "列式1行、推导3行、结论1行，另留1行改写空间",
  "width_mm": 170,
  "estimated_lines": 5,
  "line_height_mm": 10,
  "reserve_lines": 1,
  "allocated_height_mm": 60
}
```

最终页面检查后，在该记录中补 `observed`：实际可写总高度、宽度，以及空白/横线区的实际手写容量或作文格总格数；含绘图区时另记绘图区高度。跨页时可用 `page_segments` 记录各页段的位置、高度及行/格数，总高度不包含页眉页脚和不能作答的页间空间。空白区的行数是按手写行高估算的容量，不是隐藏表格的行数。例：

```json
{
  "height_mm": 60,
  "width_mm": 170,
  "writing_lines": 6,
  "drawing_height_mm": 0,
  "page_segments": [{"page": 2, "height_mm": 60, "writing_lines": 6}]
}
```

有线区数可用横线，作文区数可用格子；空白和绘图区结合页面标尺或渲染尺寸估计实际容量。数字需来自页面检查，不能为通过脚本照抄目标值。脚本只核对算术与所填观测记录，不能判断留白是否足够完成真实任务。

在完整试卷工作记录上分别运行：

```text
python scripts/answer_space.py work/blueprint.json --stage plan --report work/space-plan.json
python scripts/answer_space.py work/blueprint.json --stage final --report work/space-final.json
```

计划阶段允许缺少空间记录但会提示；最终阶段每题需有 `answer_spaces`，非空且非 `none` 区域需有页面观测。分配不足、实际容量不足或宽度不符返回失败。此检查与 `check_blueprint.py` 的分值检查分别执行。命题蓝图、未排版草案无需提前提供页面观测。

## 调用组件

`scripts/word_components.py` 使用环境提供的 `python-docx`；`scripts/answer_space.py` 只需 Python 标准库。使用依赖加载工具获取解释器，勿固定本机路径或自动安装办公软件。将技能的 `scripts` 目录加入导入路径，在本次任务的工作目录中运行作者程序。

```python
import sys
from pathlib import Path
from docx import Document

skill_dir = Path("实际技能目录")
sys.path.insert(0, str(skill_dir / "scripts"))
from word_components import add_answer_area, add_choices

doc = Document()  # 页面、正文样式等由本次试卷模板设置
question = doc.add_paragraph("12．（1）写出计算过程。")
record = {
    "id": "12(1)", "mode": "blank", "basis": "五行推导，另留一行余量",
    "estimated_lines": 5, "line_height_mm": 10, "reserve_lines": 1
}
area = add_answer_area(doc, record, question_paragraph=question)
record["width_mm"] = area["configured"]["width_mm"]
record["allocated_height_mm"] = area["configured"]["height_mm"]

question = doc.add_paragraph("13．选择正确选项。")
add_choices(doc, ["选项内容一", "选项内容二", "选项内容三", "选项内容四"],
            columns=2, question_paragraph=question)
doc.save("work/试题卷.docx")
```

- `add_answer_area(doc, spec, ...)` 返回原生表格对象和 `configured` 容量。空白区仅有无边框的高度结构，横线区只有横向细线，作文格为方格；行高为最小值，不用固定行高裁切文字。普通书写区按行自然跨页，单个绘图区保持连续。
- `add_choices(doc, choices, columns=1|2|4, ...)` 生成无边框选项表格，字符串自动加选项字母；复杂公式/图片可传入接收单元格的函数，自行插入原生公式并加选项字母。它不会把 LaTeX 字符串当成公式。先按内容选择列数，渲染后放不下就减少列数，不缩小正文硬挤。
- 通用参数为 `section`、`column_index`、`width_mm`、`left_mm`、`right_mm` 和字体设置。默认按文档末节的当前栏计算，支持等宽或明确指定的不等宽栏；生成当前正文时传正确节和栏。容器内或复杂版式需要另行落实，组件只用于正文，不能把整页宽度直接传给窄栏。
- `question_paragraph` 只关联题干末段与首个答题/选项行，不锁住整个长题或整个跨页答题区。页末若只容得下无意义的小片作答区，仍需联合题干调整分页。绘图区过高不能在一页内放下时，组件报告错误，改任务安排或页面设置。
- 按用户模板调用这些组件，或使用符合相同规则的原生实现；不强制改造现有模板。组件的宽度和容量是配置值，最终 DOCX 必须重新渲染，核查边框、字体、跨页与实际可写容量。
