# ks-kernel-changelog-txt2docx

把纯文本格式的内核发布日志整理成排版好的 Word 文档。

| 文件 | 说明 |
|---|---|
| `txt2docx.py` | 转换脚本 |
| `README.md` | 本文档 |
| `LICENSE` | MIT 许可证 |

## 依赖

- `python3`
- `python-docx`

安装依赖：

```bash
pip install python-docx
```

检查依赖：

```bash
python3 -c "import docx; print('python-docx OK')"
```

## 快速开始

```bash
python3 txt2docx.py <输入.txt> <输出.docx>
```

例：

```bash
python3 txt2docx.py SP3-2403.txt SP3-2403.docx
```

## 参数

| 参数 | 说明 |
|---|---|
| `src` | 源文本文件路径（位置参数，必填） |
| `dst` | 输出 docx 路径（位置参数，必填） |
| `--skip <小节名>` | 整段删除该小节，可重复指定。默认值就是 `测试建议`，即不写也会删掉「测试建议」 |
| `--title <标题>` | 指定文档标题。不写则自动取源文件首行；若首行是版本号，用默认标题「内核更新日志」 |

```bash
# 查看帮助
python3 txt2docx.py --help

# 指定标题
python3 txt2docx.py in.txt out.docx --title "V10SP4 内核发布日志"

# 额外再删一个小节
python3 txt2docx.py in.txt out.docx --skip 测试建议 --skip 已知问题
```

## 批量处理同系列文件

```bash
for pair in "sp1:sp1" "sp2.txt:sp2" "sp3-2303.txt:sp3-2303" "SP3-2403.txt:SP3-2403"; do
  src="${pair%%:*}"; out="${pair##*:}"
  python3 txt2docx.py "$src" "$out.docx"
done
```

## 输出说明

脚本运行后会打印：

```
文档标题：银河麒麟服务器操作系统 V10SP4 4.19 内核发布日志
版本标题 30 个，二级标题 63 个，正文段落 1960 个
已删除「测试建议」小节正文 221 行
输出：SP3-2403.docx
```

| 输出项 | 含义 |
|---|---|
| 版本标题 | 一级标题数量，即内核版本数 |
| 二级标题 | 各小节合计（特性增强 / 问题修复 / CVE修复 / 配置参数变更 等） |
| 正文段落 | 条目数。挤在一行的多条 CVE 会按 CVE 拆开，所以通常多于源文件行数 |
| 已删除…正文 | 被整段删除的小节内容行数 |

## 文档结构

| 层级 | 内容 | 样式 |
|---|---|---|
| 文档标题 | 源文件首行，居中 | Title，宋体 22pt |
| 一级标题 | `2026-9-23  Version 4.19.90-89.47.v2401` | Heading 1，16pt |
| 二级标题 | 特性增强 / 特性更改 / 问题修复 / CVE修复 / 配置参数变更 | Heading 2，13pt |
| 正文 | 每个条目一段 | Normal，10.5pt |

字体统一为**中文宋体 + 英文数字 Times New Roman**。标题里的中文也是宋体（不是黑体）。

> **字体实现说明**：python-docx 默认模板的标题样式是通过 Word 主题字体引用的（`w:asciiTheme="majorHAnsi"` / `w:eastAsiaTheme="majorEastAsia"`）指定 Calibri Light 的，而按 OOXML 规范**主题属性的优先级高于显式字体属性**——只设置 `w:ascii` 会被 Word 忽略。所以脚本做了三步：删除所有样式的 theme 属性后再写死字体、同步设置 `docDefaults`、最后在 zip 层扫一遍 `styles.xml` 清掉表格样式条件格式块里的残留（那些 `python-docx` 的 API 碰不到）。

## 解析规则

### 小节头识别

只认下面三种写法，避免把 `CVE-2021-47523 task#225598` 这类**正文行**误判成标题：

1. **整行就是小节名** — `问题修复`、`CVE`、`功能增强`
2. **小节名 + 冒号 + 同行内容** — `CVE修复: CVE-2021-47523 task#225598 ...`
3. **带 `#` 前缀的 markdown 写法** — `### CVE修复 - CVE-2024-26935 task#217556 - ...`

第 3 种只在有 `#` 前缀时才认，所以和正文不可能混淆。这一条是必须的：如果放宽成「小节名 + 任意内容」，仅 `CVE` 这一个词就会把四个源文件里 3000 多行 `CVE-xxxx task#...` 正文全部误判。

### 小节名归一

源日志里混用的异名会统一成规范名：

| 源文件写法 | 文档中的标题 |
|---|---|
| `功能特性`、`功能增强` | 特性增强 |
| `CVE` | CVE修复 |

### CVE 条目拆分

- 挤在一行的多条 CVE 拆成一 CVE 一段
- **逗号分隔的多个 CVE 整行保留**（如 `CVE-2022-49160, CVE-2025-71238 task#...`），因为逗号分隔意味着它们共享后面的 task 引用，拆开会丢失对应关系
- 拆分产生的纯符号残渣（如分隔用的 `-`）直接丢弃

### 版本顺序

保持源文件原序（通常是最新在前），不做重排。

## 校验方法

转换后建议核对一次，确认没丢内容。下面这段脚本会比对 CVE 数量、查找未收录的正文行、检查字体和残留：

```python
import re, zipfile
from collections import Counter
from docx import Document
from docx.oxml.ns import qn

SRC, DST = 'SP3-2403.txt', 'SP3-2403.docx'

src = open(SRC, encoding='utf-8').read()
d = Document(DST)
dtxt = '\n'.join(p.text for p in d.paragraphs)
flat = re.sub(r'\s+', '', dtxt)

# CVE 数量是否一致（差异应只来自被删的「测试建议」）
s = Counter(re.findall(r'CVE-\d{4}-\d+', src))
t = Counter(re.findall(r'CVE-\d{4}-\d+', dtxt))
print('源里多出（被删的）:', dict(s - t))
print('docx 里多出:', dict(t - s))

# 无意义符号段落 / markdown 残留
print('纯符号段落:', [p.text for p in d.paragraphs
                  if p.text.strip() and not re.search(r'[\w一-鿿]', p.text)] or '无')
print('markdown 残留:', [p.text[:30] for p in d.paragraphs
                     if p.text.lstrip().startswith('#')] or '无')

# 字体
for n in ('Normal', 'Title', 'Heading 1', 'Heading 2'):
    st = d.styles[n]
    print(n, st.element.rPr.rFonts.get(qn('w:ascii')),
          st.element.rPr.rFonts.get(qn('w:eastAsia')))
xml = zipfile.ZipFile(DST).read('word/styles.xml').decode()
print('主题字体残留:', re.findall(r'w:(?:ascii|hAnsi|eastAsia|cs)Theme="[^"]*"', xml) or '无')
```

`源里多出（被删的）` 这一项应该只列出出现在「测试建议」小节里的 CVE 号（如 `CVE-2026-64564poc测试` 这种），其它项目都应为「无」。

## 常见改动

改脚本顶部的常量即可：

| 想改什么 | 位置 |
|---|---|
| 字体 | `CJK_FONT` / `LATIN_FONT` |
| 默认删除的小节 | `DEFAULT_SKIP` |
| 新增异名小节 | `SECTION_ALIASES`，如加 `'新叫法': '特性增强'` |
| 新增小节类型 | `SECTIONS` 元组 |
| 各级标题字号 | `build()` 里 `set_fonts(doc.styles['Heading 1'], size=16)` 等三行 |
| 正文默认字号 | `build()` 里 `set_fonts(doc.styles['Normal'], size=10.5)` |
