---
name: thesis-zh
description: 按《本科生论文（设计）撰写要求》撰写、排版、检查或修改中文学位论文（Word/.docx）。当任务是写本科/研究生毕业论文、毕业设计说明书，或需要检查论文的字号字体、标题级次、页眉页码、参考文献著录格式、三线表、插图规范时使用。生成或检查论文前必须先加载本 skill，并使用其中的 scripts/ 脚本，不要凭记忆猜字号边距。
metadata:
  source: 本科生论文（设计）撰写要求.docx（用户提供的本校规范）
  applies-to: 中文 Word 学位论文（本科毕业设计/论文，可迁移至硕士）
  companion: office-docx
---

# 中文本科论文（设计）撰写与排版

本 skill 封装了《本科生论文（设计）撰写要求》的全部硬性格式规定。**所有字号、字体、边距、页码、著录格式都以此为准，不要凭印象填写。**

配套资源（相对本 skill 目录）：

| 文件 | 用途 |
|---|---|
| `references/thesis-format.md` | 规范全文的可执行版（逐条精确数值 + 落地方式 + 检查项编号） |
| `references/checklist.md` | 交付前逐条核对清单（标注哪些能脚本判定、哪些需人工） |
| `scripts/thesisdocx.py` | 排版助手库：页面、页码、字体、标题、三线表、图表题注、参考文献 |
| `scripts/audit_thesis.py` | 规范校验器：11 项机械检查，输出错误/警告/提示 |
| `scripts/example_thesis.py` | 端到端示例：生成一份合规论文骨架，可直接照抄结构 |

## 先加载 office-docx

本 skill 只管“论文规范”；DOCX 的读写、渲染、交付流程遵循 `office-docx` skill。两者同时使用：

1. 加载 `office-docx`，按它调用 `load_workspace_dependencies` 取得 Python 路径。
2. 用该 Python 执行本 skill 的 `scripts/thesisdocx.py` 与 `scripts/audit_thesis.py`。
3. 交付前按下文「渲染检查」用捆绑 LibreOffice Kit 做版式抽查（**先 convert 成 PDF，再 render 该 PDF**），最后用 `present` 交付。

脚本只依赖 `python-docx`（捆绑环境已含），不要安装任何包。

## 工作流程

1. **确认范围**：论文题目、学科、层次（本科/硕士）、是否已有封面、开题报告表、原创性声明模板。用户没给就先问一次，不要自己编造学校固定的封面文字。
2. **先写内容，再排版**。先产出章节文本与图表，最后统一套格式——这样不会在改稿时反复破坏格式。
3. **照 `example_thesis.py` 搭骨架**：它已按“封面节（无页眉页码）→ 前置节（罗马数字页码）→ 正文节（阿拉伯数字页码）”分好节。分节是页码规范的前提，不要把所有内容放在一个节里。
4. **用 `audit_thesis.py` 校验**，按报告逐条修正，直到 0 错误。
5. **渲染抽查**：按「渲染检查」先 convert 成 PDF 再渲染该 PDF，抽查正文 1—2 页，确认标题、三线表宽度、页码、页眉没有跑版。
6. **交付**：`present` 最终 `.docx`。

## 关键 API（`thesisdocx`）

```python
import thesisdocx as T

doc = T.new_thesis_document()          # A4 + 规范边距 + 1.5 倍行距 + ThesisH1—H4 样式

# —— 分节与页码（页码规范的前提）——
fm   = T.start_front_matter_section(doc)  # 前置节：大写罗马数字页码
body = T.start_body_section(doc)          # 正文节：阿拉伯数字从 1 开始
T.setup_front_matter_page_numbers(doc)    # 只需给首个前置节编罗马页码时用

# —— 页眉（内容必须等于本页一级标题，换章节要重设）——
T.set_header_text(body, "一、引言")        # 5 号宋体、居中
T.set_header_text(body, ["很长的论文题目", "第二行"])   # 需要分两行时

# —— 标题：编号自动补全，一级“一、”二级“（一）”三级“1”四级“（1）” ——
T.add_heading(doc, 1, "引言")              # -> 一、引言
T.add_heading(doc, 2, "研究背景")          # -> （一）研究背景
T.add_heading(doc, 3, "数据来源")          # -> 1数据来源（原文第 3 级不加点号）
T.add_heading(doc, 1, "一、已有编号")      # -> 原样保留，不会变成“已有编号、”

# —— 正文与题头 ——
T.add_body(doc, "正文段落……")             # 小4宋体 + Times New Roman 小4，首行缩进
T.add_front_matter_head(doc, "摘要")       # -> “摘　要”，小2加粗宋体，其后隔一行
T.add_keywords(doc, ["甲", "乙"])          # “关键词”4号黑体 + 全角空格 + 内容小4宋体
T.add_english_abstract_head(doc, "Abstract")
T.add_body_mixed(doc, [("中文部分", False), (" English part", True)])

# —— 公式：居中、末尾不加标点 ——
T.add_formula(doc, "E = mc²")
T.add_formula(doc, "y = ax + b", number=1)  # 编号右对齐

# —— 三线表：题注必须先写（表序表名在表上）——
T.add_table_caption(doc, 1, "实验数据统计")
table = doc.add_table(rows=4, cols=4)
T.apply_three_line_table(table)             # 上下 1.5 磅、表头下 0.5 磅、无左右边线
T.style_table_text(table)                   # 表内文字小5宋体
T.empty_cell_dash(table.cell(3, 3))         # 空缺格填“——”
T.add_table_note(doc, "注：……")

# —— 插图：先插图片，题注写在图下 ——
T.add_figure_caption(doc, 1, "系统结构", cite=1)   # “图1 系统结构”+ 上标引用号

# —— 参考文献 ——
T.add_reference(doc, 1, "张三.数据结构[M].北京：高等教育出版社，2020.15-20.")
```

`add_heading` 的编号是**按同级别调用顺序递增**的。跨文档或重新开始时先调用 `T.reset_heading_numbers()`（`new_thesis_document()` 已自动重置）。

## 铁律速查（写之前先记住这 10 条）

| 项目 | 规格 |
|---|---|
| 页面 | A4；上下边距 2.54cm，左右 3.17cm，页眉 1.5cm，页脚 1.75cm |
| 行距 | 全文统一 1.5 倍；页码居中 |
| 论文题目 | 小3号黑体；副标题 4号黑体 |
| 正文 | 中文小4号宋体，英文 Times New Roman 小4号 |
| 正文中标题 | 小4号加粗宋体，一、二级标题**上下均空一行** |
| 标题级次 | 1—4 级依次为 “一、”→“（一）”→“1”→“（1）”（第 3 级不加点号） |
| 页眉 | 除封面、扉页外每页都要；居中；内容为**本页一级标题**；5号宋体 |
| 页码 | 声明/开题/目录/摘要用**大写罗马数字**，正文起用阿拉伯数字 |
| 图表 | 表：三线表（上下 1.5 磅、中间 0.5 磅、无左右边线），表序表名在**表上** |
| 引用标注 | 参考文献用 `[序号]`；注释用 ①②③ 且**按页重新编号** |

字号换算（OOXML 半点值）：小3=15pt(30)、4号=14pt(28)、小4=12pt(24)、5号=10.5pt(21)、小5=9pt(18)。
中文字体名：宋体 `SimSun`、黑体 `SimHei`、楷体 `KaiTi`、仿宋 `FangSong`；西文 `Times New Roman`。
设置字体时必须同时写 `w:ascii`/`w:hAnsi`（西文）和 `w:eastAsia`（中文），否则中文不会真正变成宋体/黑体——`T.set_run_font()` 已处理。

## 高频扣分点

- 摘要页**不写论文题目**；“摘要”/“目录”两字之间要有**一个全角空格**，且题头后**隔一行**再写正文。
- 正文中标题的“上下均空一行”是**真的空一整行**，不是段前段后 6 磅——`add_heading` 会自动插入空行。
- 页眉必须跟着**一级标题变**，不能全文用同一个页眉文字。
- 表格**不加左、右边线**；数据空缺格填 “——”（两个字符宽），不能留空。
- 图号图名在**图下**，表序表名在**表上**——两者位置相反，最容易搞错。
- 公式居中、**末尾不加标点**；换行优先在等号 “＝” 处断。
- 注释序号按页重编，不能全篇连续编号。
- 封面、扉页**不设页眉页码**；正文必须另起一节，否则罗马数字/阿拉伯数字页码分不开。
- 参考文献 9 种文献类型各有固定著录格式，见 `references/thesis-format.md`。
- 本 skill 的自定义标题样式基于 Word 内置 `Heading 1—4`，因此**目录域（TOC）能正常收录**；目录生成后需在 Word 中按 F9 更新。

## 渲染检查（版式抽查）

**必须两步：先 `convert` 成 PDF，再 `render` 这个 PDF。不要直接 `render` `.docx`。**

直接渲染 `.docx` 在当前捆绑 LibreOffice Kit 上会把内容横向放大约 1.28 倍后按请求的像素框裁切，表现为：居中内容右偏约 2.6cm、右对齐的公式编号被切在页缘、表格右边界顶到页面边缘——**全是假跑版**。已用 3 行最小文档复现，与论文内容无关；同一文档导出 PDF 后文字坐标（左边距 89.95pt = 3.17cm）分毫不差。

```sh
# 正确：两步走
<node> <cli> convert --input thesis.docx --output thesis.pdf
<node> <cli> render  --input thesis.pdf  --output-dir preview --dpi 120 --pages 1,6
```

判定基准（A4 + 规范边距）：正文区 x ∈ [3.17, 17.83] cm，A4 宽 21cm 时居中内容的墨迹中心应在 **10.50cm**。只有越出 17.83cm 才算真跑版；三线表的左右边线落在 2.98—17.62cm 属于单元格边距的正常现象。

若渲染报 `spawn EPERM`，是沙箱不允许 Node 起 `soffice` 子进程，放宽沙箱后重试即可，不是文档问题。

## 命令

```sh
# 生成合规骨架（可照抄结构）
<python> "<skill>/scripts/example_thesis.py" 论文初稿.docx

# 校验（每次排版后运行；0 错误才算通过）
<python> "<skill>/scripts/audit_thesis.py" thesis.docx
<python> "<skill>/scripts/audit_thesis.py" thesis.docx --json report.json
<python> "<skill>/scripts/audit_thesis.py" thesis.docx --only page,fonts,headings,tables
```

检查项：`page` `fonts` `headings` `header` `abstract` `toc` `tables` `figures` `refs` `formula` `notes`。
退出码：0 无错误；1 存在错误；2 无法读取文档。校验器只能判机械项，**页眉是否等于本页一级标题、图内文字字号、坐标单位、注释按页重编、内容准确性必须人工核对**——见 `references/checklist.md`。
