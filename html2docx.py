#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把内核发布日志整理成 docx。支持纯文本(.txt)与网页(.html/.htm)两种输入。

用法：
    python3 html2docx.py 输入.html 输出.docx [--skip 测试建议] [--title 标题]
    python3 html2docx.py 输入.txt  输出.docx [--skip 测试建议] [--title 标题]

结构：版本 = 一级标题，小节 = 二级标题（特性增强/特性更改/问题修复/CVE修复/配置参数变更）。
字体：中文宋体，英文数字 Times New Roman。
"""

import argparse
import re
import shutil
import tempfile
import zipfile
from html.parser import HTMLParser
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

CJK_FONT = '宋体'
LATIN_FONT = 'Times New Roman'

DEFAULT_TITLE = '内核更新日志'
DEFAULT_SKIP = ('测试建议',)

VERSION_RE = re.compile(r'^\s*(\d{4}-\d{1,2}-\d{1,2})\s*\[(Version\s+[^\]]+)\]\s*$')

SECTIONS = ('特性增强', '特性更改', '问题修复', 'CVE修复', '配置参数变更', '测试建议')

# 源日志里混用的异名小节，统一归到规范名
SECTION_ALIASES = {'功能特性': '特性增强', '功能增强': '特性增强', 'CVE': 'CVE修复'}

# 小节头只认两种写法，避免正文被误判（"CVE-2021-47523 task#..." 绝不能匹配）：
#   1) 整行就是小节名，可带异名，如 "CVE" / "功能增强"
#   2) 小节名 + 冒号 + 同行内容，如 "CVE修复: CVE-2021-47523 ..."
SECTION_LINE_RE = re.compile(r'^\s*(%s)\s*$'
                             % '|'.join(SECTIONS + tuple(SECTION_ALIASES)))
SECTION_INLINE_RE = re.compile(r'^\s*(%s)\s*[:：]\s*(\S.*)$' % '|'.join(SECTIONS))

# 误加的 markdown 标题写法："### CVE修复 - CVE-xxx - CVE-yyy"。只有带 # 前缀时才认，
# 所以不会和正文的 "CVE-2021-47523 task#..." 混淆。
MD_SECTION_RE = re.compile(r'^#{1,6}\s*(%s)\s*[:：]?\s*[-–—]?\s*(\S.*)$'
                           % '|'.join(SECTIONS + tuple(SECTION_ALIASES)))


def match_section(line):
    """返回 (规范小节名, 同行内容)；不是小节头则返回 None。"""
    m = SECTION_LINE_RE.match(line)
    if m:
        return SECTION_ALIASES.get(m.group(1), m.group(1)), ''
    m = SECTION_INLINE_RE.match(line)
    if m:
        return m.group(1), m.group(2).strip()
    return None

# 一条 CVE 记录：CVE 号 + 其后的 task#/bug#/[...] 引用
CVE_RE = re.compile(r'CVE-\d{4}-\d+(?:\s+(?:\[[^\]]*\]|task#\S+|bug#\S+))*')

# 偶尔混进来的 markdown 标题前缀，如 "### CVE修复 - CVE-xxx"
MD_HEAD_RE = re.compile(r'^#{1,6}\s*')

# 网页另存为常见的导航/装饰文本，提取标题时应直接跳过（不当作标题或正文）
NAV_LINES = ('返回首页', '返回顶部', '回到顶部', '上一篇', '下一篇', '首页',
             '打印', '打印本页', '收藏', '分享', '上一页', '下一页')

THEME_ATTRS = ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme')

# 会被整体丢弃的标签（脚本、样式、注释等），不产生任何文本
SKIP_TAGS = {'script', 'style', 'noscript', 'template', 'head', 'title'}
# 块级标签：进入/退出时都会起一行，保证标题、段落、表格等不粘连
BLOCK_TAGS = {'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'tr',
              'br', 'pre', 'blockquote', 'section', 'article', 'header',
              'footer', 'table', 'ul', 'ol', 'dl', 'dt', 'dd', 'hr'}


class _HtmlToText(HTMLParser):
    """把 HTML 转成纯文本行，结构尽量贴近原 txt：标题、段落、表格各占一行。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._lines = []        # 已落定的行
        self._buf = []          # 当前行累积的文本片段
        self._pre = 0           # <pre> 深度，内部保留空白与换行
        self._skip = 0          # 丢弃区深度（script/style 等）

    def _flush(self):
        line = ''.join(self._buf).strip()
        if line:
            self._lines.append(line)
        self._buf = []

    def _newline(self):
        # pre 内部换行单独成行（原样保留空白）
        if self._pre:
            self._flush()
        else:
            self._flush()

    def _cell_gap(self):
        # 相邻单元格之间补一个空格，避免内容粘连
        if self._buf and not self._pre:
            self._buf.append(' ')

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in SKIP_TAGS:
            self._skip += 1
            return
        if self._skip:
            return
        if tag == 'pre':
            self._pre += 1
            self._flush()
            return
        if tag in ('td', 'th'):
            self._cell_gap()
            return
        if tag in BLOCK_TAGS:
            self._flush()

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in SKIP_TAGS:
            if self._skip:
                self._skip -= 1
            return
        if self._skip:
            return
        if tag == 'pre':
            if self._pre:
                self._pre -= 1
            self._flush()
            return
        if tag in BLOCK_TAGS:
            self._flush()

    def handle_data(self, data):
        if self._skip:
            return
        if self._pre:
            # pre 内的换行拆成多行；纯空白行会被 flush 时的 strip 清掉
            parts = data.split('\n')
            for i, part in enumerate(parts):
                if i:
                    self._flush()
                self._buf.append(part)
            return
        # 普通文本压缩连续空白，避免 HTML 里的缩进/换行把段落打散
        self._buf.append(re.sub(r'\s+', ' ', data))
    def text(self):
        self._flush()
        return self._lines


def html_to_text(data):
    """把 HTML 文档转成纯文本行列表。"""
    parser = _HtmlToText()
    parser.feed(data)
    parser.close()
    return parser.text()


def read_source(src):
    """读入源文件：html/htm 转文本，txt 直接按行读。"""
    p = Path(src)
    data = p.read_text(encoding='utf-8')
    if p.suffix.lower() in ('.html', '.htm'):
        return html_to_text(data)
    return data.splitlines()


def set_fonts(style, size=None, bold=None, color=None):
    """在样式上写死中西文字体。必须先清掉 theme 属性，否则 Word 会优先用主题字体。"""
    if not hasattr(style, 'font'):    # 编号样式没有字体
        return
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    for attr in THEME_ATTRS:
        key = qn('w:' + attr)
        if rFonts.get(key) is not None:
            del rFonts.attrib[key]
    rFonts.set(qn('w:ascii'), LATIN_FONT)
    rFonts.set(qn('w:hAnsi'), LATIN_FONT)
    rFonts.set(qn('w:eastAsia'), CJK_FONT)
    rFonts.set(qn('w:cs'), LATIN_FONT)

    style.font.name = LATIN_FONT
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color


def set_doc_default_fonts(doc):
    """兜底：文档默认字体也统一，避免未显式设置的地方走主题字体。"""
    for rPr in doc.styles.element.xpath('./w:docDefaults/w:rPrDefault/w:rPr'):
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = rPr.makeelement(qn('w:rFonts'), {})
            rPr.insert(0, rFonts)
        for attr in THEME_ATTRS:
            key = qn('w:' + attr)
            if rFonts.get(key) is not None:
                del rFonts.attrib[key]
        rFonts.set(qn('w:ascii'), LATIN_FONT)
        rFonts.set(qn('w:hAnsi'), LATIN_FONT)
        rFonts.set(qn('w:eastAsia'), CJK_FONT)
        rFonts.set(qn('w:cs'), LATIN_FONT)


def strip_theme_fonts(path):
    """最后扫一遍 styles.xml：表格样式等藏在条件格式块里的主题字体引用也一并清掉。"""
    with zipfile.ZipFile(path) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        tmp_path = tmp.name

    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == 'word/styles.xml':
                data = re.sub(rb'\s+w:(?:ascii|hAnsi|eastAsia|cs)Theme="[^"]*"',
                              b'', data)
            zout.writestr(info, data)

    shutil.move(tmp_path, path)


def split_cve_line(line):
    """把挤在一行里的多条 CVE 拆成一 CVE 一段；拆不动就原样保留。"""
    matches = list(CVE_RE.finditer(line))
    if len(matches) < 2:
        return [line.strip()]
    for a, b in zip(matches, matches[1:]):
        # 逗号分隔的多个 CVE 属于同一条记录，拆开会丢引用关系，整行保留
        if re.search(r'[,，、]', line[a.end():b.start()]):
            return [line.strip()]
    residue = CVE_RE.sub('', line).strip()
    parts = [m.group(0).strip() for m in matches]
    # 行首行尾的非 CVE 残留内容，纯符号（如分隔用的 "-"）没有信息量，丢掉
    if residue and re.search(r'[\w一-鿿]', residue):
        parts.append(residue)
    return parts


def build(src, dst, skip=(), title=None):
    lines = read_source(src)

    doc = Document()

    set_doc_default_fonts(doc)
    for style in doc.styles:          # 所有样式统一字体，清掉主题字体残留
        set_fonts(style)
    set_fonts(doc.styles['Normal'], size=10.5)
    set_fonts(doc.styles['Title'], size=22, bold=True,
              color=RGBColor(0x00, 0x00, 0x00))
    for name, size in (('Heading 1', 16), ('Heading 2', 13)):
        set_fonts(doc.styles[name], size=size, bold=True,
                  color=RGBColor(0x00, 0x00, 0x00))

    stats = {'ver': 0, 'sec': 0, 'para': 0, 'skipped': 0}

    # 提取标题：跳过开头的导航/装饰行与空行，把第一行有实际内容的
    # 非版本号/非小节名文本当作标题
    body = []
    for raw in lines:
        s = raw.strip()
        if not s:
            continue
        if not body and not VERSION_RE.match(s) and match_section(s) is None:
            if s in NAV_LINES:
                continue          # "返回首页" 之类的导航文本直接丢弃
            if title is None:
                title = s
            continue
        body.append(s)

    doc_title = doc.add_heading(title or DEFAULT_TITLE, level=0)
    doc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    skipping = False
    for line in body:
        # "### CVE修复 - ..." 这类误加的 markdown 前缀，只在去掉后能识别成
        # 版本/小节时才剥掉，避免误伤正文
        norm = MD_HEAD_RE.sub('', line)
        if norm != line and VERSION_RE.match(norm):
            line = norm               # "### 2026-9-5 [Version ...]"

        m = VERSION_RE.match(line)
        if m:
            skipping = False          # 版本边界，跳过段结束
            doc.add_heading('%s  %s' % (m.group(1), m.group(2)), level=1)
            stats['ver'] += 1
            continue

        md = MD_SECTION_RE.match(line)
        if md:
            sec = (SECTION_ALIASES.get(md.group(1), md.group(1)), md.group(2).strip())
        else:
            sec = match_section(line)
        if sec:
            name, rest = sec
            if name in skip:
                skipping = True
                continue
            skipping = False
            doc.add_heading(name, level=2)
            stats['sec'] += 1
            if rest:                  # "CVE修复: CVE-xxx ..." 这种同行写法
                for part in split_cve_line(rest):
                    doc.add_paragraph(part)
                    stats['para'] += 1
            continue

        if skipping:
            stats['skipped'] += 1
            continue

        for part in split_cve_line(line):
            doc.add_paragraph(part)
            stats['para'] += 1

    doc.save(dst)
    strip_theme_fonts(dst)

    print('文档标题：%s' % (title or DEFAULT_TITLE))
    print('版本标题 %(ver)d 个，二级标题 %(sec)d 个，正文段落 %(para)d 个' % stats)
    if skip:
        print('已删除「%s」小节正文 %d 行' % ('/'.join(skip), stats['skipped']))
    else:
        print('未删除任何小节')
    print('输出：%s' % dst)


def main():
    ap = argparse.ArgumentParser(
        description='把内核发布日志整理成 docx，支持 .txt / .html / .htm 输入。')
    ap.add_argument('src')
    ap.add_argument('dst')
    ap.add_argument('--skip', action='append', default=None,
                    help='要整段删除的小节名，可重复')
    ap.add_argument('--title', default=None, help='文档标题，默认取源文件首行')
    args = ap.parse_args()
    build(args.src, args.dst,
          skip=tuple(args.skip) if args.skip is not None else DEFAULT_SKIP,
          title=args.title)


if __name__ == '__main__':
    main()
