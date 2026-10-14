# thesis-zh

面向 [DeepSeek Harness](https://github.com/deepseek-ai) 的 Agent Skill：按《本科生论文（设计）撰写要求》**撰写、排版、检查中文学位论文**（Word / `.docx`）。

所有字号、字体、边距、页码、页眉、图表题注、参考文献著录格式都以规范原文为准，**不靠记忆猜数值**。

## 能力

| 资源 | 用途 |
|---|---|
| `SKILL.md` | 规范要点 + 关键 API + 铁律速查 + 高频扣分点 |
| `references/thesis-format.md` | 规范全文的可执行版（逐条精确数值 + 落地方式 + 检查项编号） |
| `references/checklist.md` | 交付前逐条核对清单（标注哪些能脚本判定、哪些需人工） |
| `scripts/thesisdocx.py` | 排版助手库：页面、页码、字体、标题级次、三线表、图表题注、参考文献 |
| `scripts/audit_thesis.py` | 规范校验器：11 项机械检查，输出错误/警告/提示，退出码 0/1/2 |
| `scripts/example_thesis.py` | 端到端示例：生成一份合规论文骨架，可直接照抄结构 |

## 安装

复制到 DSH 的用户级 skill 目录（与 `tabbit` 等同级），DSH 会实时发现，无需重启：

```powershell
git clone https://github.com/<you>/thesis-zh "$env:USERPROFILE\.agents\skills\thesis-zh"
```

其他 DSH 识别的 skill 根目录：`<项目>/.dsh/skills`、`<项目>/.agents/skills`、`~/.dsh/skills`。

## 依赖

只用 `python-docx`（DSH 捆绑 Python 已含），**不要安装任何包**。DOCX 的读写、渲染、交付流程遵循配套的 `office-docx` skill。

## 用法

```sh
# 生成合规骨架（可照抄结构）
<python> "<skill>/scripts/example_thesis.py" 论文初稿.docx

# 校验（每次排版后运行；0 错误才算通过）
<python> "<skill>/scripts/audit_thesis.py" thesis.docx
<python> "<skill>/scripts/audit_thesis.py" thesis.docx --json report.json
<python> "<skill>/scripts/audit_thesis.py" thesis.docx --only page,fonts,headings,tables
```

检查项：`page` `fonts` `headings` `header` `abstract` `toc` `tables` `figures` `refs` `formula` `notes`。

校验器只能判机械项——**页眉是否等于本页一级标题、图内文字字号、坐标单位、注释是否按页重编、内容准确性必须人工核对**。

## ⚠️ 渲染陷阱（版式抽查必读）

做版式抽查时**必须两步：先 `convert` 成 PDF，再 `render` 这个 PDF。不要直接 `render` `.docx`。**

直接渲染 `.docx` 在捆绑 LibreOffice Kit 上会把内容横向放大约 1.28 倍后按请求的像素框裁切，表现为：居中内容右偏约 2.6cm、右对齐的公式编号被切在页缘、表格右边界顶到页面边缘——**全是假跑版**。已用 3 行最小文档复现，与文档内容无关；同一文档导出 PDF 后文字坐标（左边距 89.95pt = 3.17cm）分毫不差。

```sh
# 正确：两步走
<node> <cli> convert --input thesis.docx --output thesis.pdf
<node> <cli> render  --input thesis.pdf  --output-dir preview --dpi 120 --pages 1,6
```

判定基准（A4 + 规范边距）：正文区 x ∈ [3.17, 17.83] cm，居中内容的墨迹中心应在 **10.50cm**。只有越出 17.83cm 才算真跑版；三线表的左右边线落在 2.98—17.62cm 属于单元格边距的正常现象。

## 主要格式规格（速查）

| 项目 | 规格 |
|---|---|
| 页面 | A4；上下 2.54cm，左右 3.17cm，页眉 1.5cm，页脚 1.75cm |
| 行距 | 全文 1.5 倍；页码居中 |
| 正文 | 中文小 4 号宋体，西文 Times New Roman 小 4 号 |
| 标题级次 | `一、` → `（一）` → `1` → `（1）`（第 3 级不加点号） |
| 页眉 | 除封面/扉页外每页；居中；内容为**本页一级标题**；5 号宋体 |
| 页码 | 声明/开题/目录/摘要用大写罗马数字，正文起用阿拉伯数字 |
| 表格 | 三线表（上下 1.5 磅、表头下 0.5 磅、无左右边线），**表序表名在表上** |
| 插图 | **图号图名在图下**（与表相反） |

字号换算（OOXML 半点值）：小3=15pt(30)、4号=14pt(28)、小4=12pt(24)、5号=10.5pt(21)、小5=9pt(18)。

## 说明

`references/thesis-format.md` 是某校《本科生论文（设计）撰写要求》的可执行化整理，供同规范院校参考；使用前请与你所在学校的现行规定核对。
