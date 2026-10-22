#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
md2docx.py — 将 Markdown 面试题库转换为 Word (.docx)

设计目标
--------
1. 优先使用 pandoc：对表格 / 代码块 / 多级标题 / 可点击目录支持最完整，质量最好。
2. 无 pandoc 时自动回退到内置的 python-docx 轻量转换器（标题/段落/表格/列表/代码/引用/加粗/行内代码/图片）。
3. 全程统一字体与字色：中文「微软雅黑」、英文/数字「Times New Roman」、字色黑色；
   代码块保留等宽字体 Consolas（含其中的中文注释仍走微软雅黑）。
   代码块还会加「浅灰底纹 + 细边框」构成代码框，行内代码也加浅灰底，观感接近 GitHub。
   标题统一为「黑色 + 微软雅黑」（覆盖 pandoc 默认蓝绿主题色/主题字体），二级标题加粗。
4. 生成 Word 原生目录字段（打开时自动更新页码），并保留题库原有的「目录总览」表。
5. 页码与分节：目录与正文「下一页」分节；目录用罗马数字 I/II/III、正文用阿拉伯数字从 1 起，
   页脚居中、距底 0.5cm（--no-page-numbers 关闭，--footer-distance-cm 调距离）。
6. Mermaid 流程图渲染成图片：```mermaid 块用本地 mmdc + 系统 Chrome/Edge 离线渲染为 PNG 嵌入
   （缺依赖或语法错自动降级为代码文本；--no-mermaid 关闭）。
7. 图片：Markdown 图片 ![alt](path) 嵌入 Word，过宽自动等比缩到正文宽度。
8. 文档最前面自动生成一张【封面页】：整页垂直居中显示标题（单文件=文档一级标题/文件名，
   整目录=目录名），封面不显示页码（可用 --no-blank-page 关闭）。
9. LaTeX 数学公式：$...$ / $$...$$ / \(...\) / \[...\] 都能渲染——
   pandoc 引擎转成 Word【原生公式】(OMML，可在 Word 里直接编辑)；
   内置引擎用 matplotlib mathtext 离线渲染成图片嵌入。
   代码块/行内代码里的 $ 不受影响（不会误伤 ${name}、$100 等）；--no-math 关闭。

用法
----
    python md2docx.py                          # 转换默认题库文件
    python md2docx.py 输入.md                   # 指定输入
    python md2docx.py 输入.md -o 输出.docx       # 指定输出
    python md2docx.py --toc-depth 3            # 目录深入到每一道题(默认 2:模块+小节)
    python md2docx.py --no-toc                 # 不生成自动目录
    python md2docx.py --font 宋体               # 更换中文字体
    python md2docx.py --engine python-docx     # 强制使用内置引擎(不调用 pandoc)
    python md2docx.py --table-style gray       # 表格配色: blue(默认)/gray/green/teal
    python md2docx.py --table-font-size 9      # 表格列太多挤时调小字号
    python md2docx.py --latin-font Arial       # 英文/数字字体(默认 Times New Roman)
    python md2docx.py --text-color 222222      # 正文字色 16 进制(默认 000000 黑)
    python md2docx.py --no-blank-page          # 不在最前插入封面页(默认插入带标题封面)
    python md2docx.py --page-size Letter       # 纸张 A4(默认)/Letter
    python md2docx.py --code-bg EEF1F5         # 代码块/行内代码底纹色(默认 F6F8FA)
    python md2docx.py --code-border C8CDD3      # 代码块边框色(默认 D0D7DE)
    python md2docx.py --code-font-size 9       # 代码字号(默认 9.5)
    python md2docx.py --no-code-box            # 代码块只用等宽字体，不加底纹/边框
    python md2docx.py --no-page-numbers        # 不加页码/分节
    python md2docx.py --footer-distance-cm 0.8 # 页码距页面底部的距离(默认 0.5cm)
    python md2docx.py --no-mermaid             # 不渲染 mermaid 流程图(保持代码文本)
    python md2docx.py --keep-heading-together  # 保留标题“与下段同页/段中不分页”(会显示左侧小黑点)
    python md2docx.py --no-math               # 不渲染 LaTeX 数学公式(保持 $...$ 原文)

提示：Word 打开后按「视图 → 导航窗格」可看到全部标题的可点击大纲，
     与自动目录深度无关（因为每个标题都是真正的 Word 标题样式）。
"""

import argparse
import os
import re
import shutil
import subprocess
import sys

DEFAULT_INPUT = "LangChain_LangGraph_面试题库规范版.md"
DEFAULT_FONT = "微软雅黑"                 # 中文（eastAsia）
DEFAULT_LATIN_FONT = "Times New Roman"    # 英文/数字（ascii/hAnsi）
DEFAULT_TEXT_COLOR = "000000"             # 正文字色（黑色）
DEFAULT_PAGE_SIZE = "A4"                   # 纸张：A4 / Letter（显式设定，避免默认 US Letter）
PAGE_SIZES_CM = {"A4": (21.0, 29.7), "Letter": (21.59, 27.94)}  # (宽, 高) 厘米

# 代码块外观（可被 CLI 覆盖）：GitHub 风格浅灰底 + 细边框 + 等宽字号
DEFAULT_CODE_BG = "F6F8FA"                # 代码块 / 行内代码底纹
DEFAULT_CODE_BORDER = "D0D7DE"            # 代码块边框色
DEFAULT_CODE_FONT_SIZE = 9.5             # 代码字号（pt）
DEFAULT_QUOTE_BAR = "5B8DB8"             # 引用块（> blockquote）左侧竖条色
DEFAULT_QUOTE_BG = "EEF3F9"              # 引用块底纹色（很浅的蓝）
CODE_FONT = "Consolas"                    # 等宽字体（与 _inject_fonts 保持一致）
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"  # 保留空白缩进用

# 代码/等宽样式名（小写子串）：命中则保留等宽字体，不改成 Times New Roman
_CODE_STYLE_HINTS = ("source code", "verbatim", "code", "html", "macro")

# 表格配色方案：表头底色 / 表头字色 / 斑马纹底色 / 边框色
TABLE_SCHEMES = {
    "blue":  {"header": "305496", "hfont": "FFFFFF", "zebra": "D9E1F2", "border": "8EAADB"},
    "gray":  {"header": "595959", "hfont": "FFFFFF", "zebra": "EDEDED", "border": "BFBFBF"},
    "green": {"header": "538135", "hfont": "FFFFFF", "zebra": "E2EFDA", "border": "A9D08E"},
    "teal":  {"header": "2F6C74", "hfont": "FFFFFF", "zebra": "DAEEF0", "border": "9CC3C8"},
}

# OOXML 子元素的 schema 顺序（插入时据此定位，避免 Word 打开报错/修复）
_TBLPR_ORDER = ["w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
                "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
                "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
                "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription"]
_TRPR_ORDER = ["w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore",
               "w:wAfter", "w:cantSplit", "w:trHeight", "w:tblHeader",
               "w:tblCellSpacing", "w:jc", "w:hidden"]
_TCPR_ORDER = ["w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge",
               "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
               "w:tcFit", "w:vAlign", "w:hideMark"]
# 运行属性 rPr 子元素顺序（注入 rFonts / color 时据此定位）
_RPR_ORDER = ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
              "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
              "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
              "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
              "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
              "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
              "w:eastAsianLayout", "w:specVanish", "w:oMath"]
# 段落属性 pPr 子元素顺序（注入 pBdr / shd / keepLines 时据此定位）
_PPR_ORDER = ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore",
              "w:framePr", "w:widowControl", "w:numPr", "w:suppressLineNumbers",
              "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
              "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
              "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
              "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
              "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
              "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
              "w:rPr", "w:sectPr", "w:pPrChange"]
# 分节属性 sectPr 子元素顺序（注入 pgNumType 时据此定位）
_SECTPR_ORDER = ["w:headerReference", "w:footerReference", "w:footnotePr",
                 "w:endnotePr", "w:type", "w:pgSz", "w:pgMar", "w:paperSrc",
                 "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols",
                 "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
                 "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
                 "w:printerSettings", "w:sectPrChange"]


# --------------------------------------------------------------------------- #
# 通用工具
# --------------------------------------------------------------------------- #
def find_pandoc():
    """返回 pandoc 可执行路径，找不到返回 None。"""
    return shutil.which("pandoc")


def human_size(num_bytes):
    for unit in ("B", "KB", "MB", "GB"):
        if num_bytes < 1024:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} TB"


def _ensure_parent_dir(path):
    """确保输出文件的父目录存在，不存在则创建（避免 -o 指到不存在的目录时失败）。"""
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)


def _derive_cover_title(md_path):
    """取封面标题：优先用文档首个一级标题(# )，取不到则用文件名(去扩展名)。"""
    try:
        with open(md_path, encoding="utf-8") as fh:
            for _ in range(300):
                line = fh.readline()
                if not line:
                    break
                m = re.match(r"^#\s+(.+?)\s*$", line.strip())
                if m:
                    return m.group(1).strip()
    except OSError:
        pass
    return os.path.splitext(os.path.basename(md_path))[0]


# --------------------------------------------------------------------------- #
# 目录批量导出：自然排序收集 + 合并（--folder / -fd）
# --------------------------------------------------------------------------- #
def _natural_key(name):
    """自然排序键：把字符串里的连续数字段按数值比较，使 2<10、67<100、1.9<1.10。"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def collect_markdown_files(folder):
    """收集目录【顶层】的 .md/.markdown 文件并自然排序；跳过隐藏/临时(.、~$ 开头)文件。
    返回绝对/相对（随传入 folder）路径列表。"""
    exts = (".md", ".markdown")
    names = []
    for name in os.listdir(folder):
        if name.startswith(".") or name.startswith("~$"):
            continue
        if name.lower().endswith(exts) and os.path.isfile(os.path.join(folder, name)):
            names.append(name)
    names.sort(key=_natural_key)
    return [os.path.join(folder, n) for n in names]


# 合并多个 md 时，文件之间插入的“下一页”分页符（按引擎选不同写法）：
#   pandoc      → openxml 原生分页块（需 raw_attribute 扩展，已在 --from 里开启）
#   python-docx → 哨兵注释，由内置解析器识别后插入 WD_BREAK.PAGE
_PAGEBREAK_OOXML = ('\n\n```{=openxml}\n'
                    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n\n')
_PAGEBREAK_SENTINEL = "MD2DOCX_PAGEBREAK"


def merge_markdown_files(files, engine, page_break=True):
    """把多个 md 文件按给定顺序拼成一个大 md 字符串；文件之间可插“下一页”分页符。
    每个文件末尾补一个换行，避免相邻文件的首尾行粘连成一段。"""
    if page_break:
        brk = _PAGEBREAK_OOXML if engine == "pandoc" else f"\n\n{_PAGEBREAK_SENTINEL}\n\n"
    else:
        brk = "\n\n"
    parts = []
    for idx, fp in enumerate(files):
        with open(fp, encoding="utf-8") as fh:
            txt = fh.read()
        if idx:
            parts.append(brk)
        parts.append(txt.rstrip() + "\n")
    return "".join(parts)


# --------------------------------------------------------------------------- #
# 中文字体注入（同时服务于 pandoc 参考文档 与 python-docx 兜底）
# --------------------------------------------------------------------------- #
def _inject_fonts(doc, cjk_font, latin_font, text_color):
    """给 Document 注入字体与字色（作用于文档默认与全部具名样式）：

        中文 (eastAsia)        -> cjk_font   （默认 微软雅黑）
        英文/数字 (ascii/hAnsi) -> latin_font（默认 Times New Roman）
        文字颜色 (color)        -> text_color（默认 000000 黑色）

    代码/等宽样式（Source Code、Verbatim 等）保留等宽字体 Consolas，只同步中文与黑色，
    以免整段代码被改成比例字体而难以阅读；表头白字由表格美化阶段以直接格式覆盖。

    标题（Heading/Title/Subtitle）：pandoc 参考文档默认给标题主题色(蓝绿)+主题字体，
    这里统一改成【黑色 + 微软雅黑】（连英文/数字也用微软雅黑），二级标题额外加粗。
    关键：必须删掉 rFonts 的 *Theme 与 color 的 themeColor/themeShade/themeTint 属性，
    否则主题属性会盖过我们设的显式字体/颜色，标题仍是蓝绿色、英文仍是主题字体。
    """
    from docx.oxml.ns import qn

    def _strip_theme(el, attrs):
        for a in attrs:
            key = qn(a)
            if el.get(key) is not None:
                del el.attrib[key]

    def _apply(rpr, mono=False, heading=False):
        rfonts = _oxml_child(rpr, "w:rFonts", _RPR_ORDER)
        if heading:
            latin = cjk_font          # 标题英文/数字也走微软雅黑
        elif mono:
            latin = "Consolas"
        else:
            latin = latin_font
        rfonts.set(qn("w:eastAsia"), cjk_font)
        rfonts.set(qn("w:ascii"), latin)
        rfonts.set(qn("w:hAnsi"), latin)
        _strip_theme(rfonts, ("w:asciiTheme", "w:eastAsiaTheme",
                              "w:hAnsiTheme", "w:cstheme"))
        color = _oxml_child(rpr, "w:color", _RPR_ORDER)
        color.set(qn("w:val"), text_color)
        _strip_theme(color, ("w:themeColor", "w:themeShade", "w:themeTint"))

    def _set_bold(rpr):
        from docx.oxml.ns import qn as _qn
        _oxml_child(rpr, "w:b", _RPR_ORDER).set(_qn("w:val"), "true")
        _oxml_child(rpr, "w:bCs", _RPR_ORDER).set(_qn("w:val"), "true")

    # 1) 文档默认 docDefaults/rPrDefault/rPr
    try:
        styles_el = doc.styles.element
        doc_defaults = styles_el.find(qn("w:docDefaults"))
        if doc_defaults is not None:
            rpr_default = doc_defaults.find(qn("w:rPrDefault"))
            if rpr_default is not None:
                rpr = rpr_default.find(qn("w:rPr"))
                if rpr is None:
                    rpr = rpr_default.makeelement(qn("w:rPr"), {})
                    rpr_default.append(rpr)
                _apply(rpr, mono=False)
    except Exception:
        pass

    # 2) 遍历所有具名样式（代码类保留等宽字体；标题走黑色+微软雅黑）
    for style in doc.styles:
        try:
            name = (style.name or "").lower()
            is_heading = name.startswith("heading") or name in ("title", "subtitle")
            mono = any(h in name for h in _CODE_STYLE_HINTS)
            rpr = style.element.get_or_add_rPr()
            _apply(rpr, mono=mono, heading=is_heading)
            if name == "heading 2":       # 二级标题加粗
                _set_bold(rpr)
        except Exception:
            continue


def _set_page_size(doc, size=DEFAULT_PAGE_SIZE):
    """把文档所有节的纸张显式设为 size（A4/Letter），避免 python-docx/pandoc 默认的 US Letter。"""
    from docx.shared import Cm
    w_cm, h_cm = PAGE_SIZES_CM.get(size, PAGE_SIZES_CM[DEFAULT_PAGE_SIZE])
    for sec in doc.sections:
        sec.page_width = Cm(w_cm)
        sec.page_height = Cm(h_cm)


def build_reference_doc(pandoc, cjk_font, latin_font, text_color, ref_path,
                        page_size=DEFAULT_PAGE_SIZE):
    """导出 pandoc 默认参考文档并注入字体/字色，返回参考文档路径。

    若未安装 python-docx，则直接返回未改字体的原始参考文档（pandoc 仍能用）。
    """
    with open(ref_path, "wb") as fh:
        subprocess.run(
            [pandoc, "--print-default-data-file", "reference.docx"],
            stdout=fh,
            check=True,
        )
    try:
        from docx import Document
    except ImportError:
        print("  · 未安装 python-docx，跳过字体注入（可 pip install python-docx 获得更佳效果）")
        return ref_path

    doc = Document(ref_path)
    _inject_fonts(doc, cjk_font, latin_font, text_color)
    _set_page_size(doc, page_size)
    doc.save(ref_path)
    return ref_path


# --------------------------------------------------------------------------- #
# 引擎一：pandoc
# --------------------------------------------------------------------------- #
def _mark_toc_dirty(docx_path):
    """把目录(TOC)字段标记为 dirty，让 Word 打开时【静默】刷新目录页码。

    刻意【不】在 settings.xml 写 <w:updateFields w:val="true"/>——那会让 Word 每次打开都弹
    “该文档包含的域可能引用了其他文件，是否更新该文档中的这些域” 的提示（该提示只由这个
    全局设置触发）。改用字段级 <w:fldChar ... w:dirty="true"> 后，Word 会自动更新目录且
    不再弹此提示；页脚 PAGE 域本就由 Word 分页时实时计算，无需该设置。
    顺带移除历史遗留的 updateFields（老文档重转时也能一并去掉提示）。返回是否改动。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError:
        return False
    doc = Document(docx_path)
    body = doc.element.body
    changed = False
    for p_el in body.iter(qn("w:p")):
        # 该段是否是 TOC 域（其 instrText 里含 "TOC"）
        if not any("TOC" in (it.text or "") for it in p_el.iter(qn("w:instrText"))):
            continue
        for fc in p_el.iter(qn("w:fldChar")):          # 把该域的 begin 标 dirty
            if fc.get(qn("w:fldCharType")) == "begin":
                fc.set(qn("w:dirty"), "true")
                changed = True
                break
    # 无论如何都确保 settings.xml 里没有 updateFields（它才是弹框根源）
    settings = doc.settings.element
    for el in settings.findall(qn("w:updateFields")):
        settings.remove(el)
        changed = True
    if changed:
        doc.save(docx_path)
    return changed


def _prepend_cover_page(docx_path, title, cjk_font=DEFAULT_FONT):
    """在文档最前插入一张【封面页】：单独成节（下一页分节）、纸张/页边距沿用正文，
    整页【垂直居中】显示标题文字；该节页脚留空（封面不显示页码，由 apply_page_numbering
    保证）。原有内容顺延到封面之后另起一页。返回 True/False。
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        print("  · 未安装 python-docx，跳过封面页插入")
        return False
    import copy
    doc = Document(docx_path)
    body = doc.element.body
    final_sect = body.find(qn("w:sectPr"))

    # 封面节 sectPr：克隆纸张/页边距，去掉页眉页脚与页码格式，设「下一页 + 垂直居中」
    cover_sect = (copy.deepcopy(final_sect) if final_sect is not None
                  else OxmlElement("w:sectPr"))
    for tag in ("w:headerReference", "w:footerReference", "w:pgNumType"):
        for el in cover_sect.findall(qn(tag)):
            cover_sect.remove(el)
    _oxml_child(cover_sect, "w:type", _SECTPR_ORDER).set(qn("w:val"), "nextPage")
    _oxml_child(cover_sect, "w:vAlign", _SECTPR_ORDER).set(qn("w:val"), "center")

    # 标题段（微软雅黑、黑色、加粗、居中、大字号）
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = tp.add_run(title or "")
    run.bold = True
    run.font.size = Pt(28)
    rpr = run._element.get_or_add_rPr()
    rf = _oxml_child(rpr, "w:rFonts", _RPR_ORDER)
    rf.set(qn("w:ascii"), cjk_font)
    rf.set(qn("w:hAnsi"), cjk_font)
    rf.set(qn("w:eastAsia"), cjk_font)
    _oxml_child(rpr, "w:color", _RPR_ORDER).set(qn("w:val"), "000000")

    # 承载封面节 sectPr 的空段落（该节最后一段）
    bp = OxmlElement("w:p")
    bppr = OxmlElement("w:pPr")
    bppr.append(cover_sect)
    bp.append(bppr)

    # 移动到正文最前：标题段 → 分节段
    body.insert(0, tp._p)
    tp._p.addnext(bp)
    doc.save(docx_path)
    return True


def _apply_pagebreak_markers(docx_path):
    """把正文里独占一段的分页哨兵(_PAGEBREAK_SENTINEL)转成真正的硬分页符。

    pandoc 会把单文件里的哨兵行原样渲染成一个文本段落，这里把它改成分页
    （封面独占一页即靠此实现）。python-docx 引擎在解析阶段已直接处理哨兵，
    不会有残留，故此函数对其为空操作。返回处理的分页数。
    """
    try:
        from docx import Document
        from docx.enum.text import WD_BREAK
    except ImportError:
        return 0
    doc = Document(docx_path)
    n = 0
    for p in doc.paragraphs:
        if p.text.strip() == _PAGEBREAK_SENTINEL:
            for r in list(p.runs):                       # 清掉哨兵文本
                r._element.getparent().remove(r._element)
            p.add_run().add_break(WD_BREAK.PAGE)         # 换成硬分页
            n += 1
    if n:
        doc.save(docx_path)
    return n


# --------------------------------------------------------------------------- #
# 表格美化 + 跨页处理（pandoc 与 python-docx 两个引擎共用的后处理）
# --------------------------------------------------------------------------- #
def _oxml_child(parent, tag, order):
    """按 schema 顺序获取或创建 parent 下的子元素，返回该元素。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tq = qn(tag)
    exist = parent.find(tq)
    if exist is not None:
        return exist
    el = OxmlElement(tag)
    full = [qn(t) for t in order]
    my = full.index(tq)
    for child in parent:
        if child.tag in full and full.index(child.tag) > my:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def _style_table_layout(table):
    """表格占满页宽、自动列宽、整体居中。"""
    from docx.oxml.ns import qn
    from docx.enum.table import WD_TABLE_ALIGNMENT
    pr = table._tbl.tblPr
    w = _oxml_child(pr, "w:tblW", _TBLPR_ORDER)
    w.set(qn("w:type"), "pct")
    w.set(qn("w:w"), "5000")  # 5000 = 100% 页宽
    _oxml_child(pr, "w:tblLayout", _TBLPR_ORDER).set(qn("w:type"), "autofit")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER


def _style_table_borders(table, color, sz=4):
    """全边框 + 内部横竖线，柔和配色。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    b = _oxml_child(table._tbl.tblPr, "w:tblBorders", _TBLPR_ORDER)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = b.find(qn("w:" + edge))
        if e is None:
            e = OxmlElement("w:" + edge)
            b.append(e)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), color)


def _style_cell_margins(table, top=38, left=110, bottom=38, right=110):
    """加大单元格内边距，文字不贴边。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    m = _oxml_child(table._tbl.tblPr, "w:tblCellMar", _TBLPR_ORDER)
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = m.find(qn("w:" + side))
        if e is None:
            e = OxmlElement("w:" + side)
            m.append(e)
        e.set(qn("w:w"), str(val))
        e.set(qn("w:type"), "dxa")


def _row_flag(row, tag):
    """给行属性打布尔标记（cantSplit=禁止跨页断裂 / tblHeader=跨页重复表头）。"""
    from docx.oxml.ns import qn
    trPr = row._tr.get_or_add_trPr()
    _oxml_child(trPr, tag, _TRPR_ORDER).set(qn("w:val"), "true")


def _cell_shading(cell, fill):
    from docx.oxml.ns import qn
    shd = _oxml_child(cell._tc.get_or_add_tcPr(), "w:shd", _TCPR_ORDER)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def _cell_vcenter(cell):
    from docx.oxml.ns import qn
    _oxml_child(cell._tc.get_or_add_tcPr(), "w:vAlign", _TCPR_ORDER).set(qn("w:val"), "center")


def _fmt_cell_paras(cell, size_pt, bold=False, color=None, center=False):
    """统一单元格内段落：紧凑行距、字号，可选加粗/字色/居中。"""
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    for p in cell.paragraphs:
        pf = p.paragraph_format
        pf.space_before = Pt(1)
        pf.space_after = Pt(1)
        pf.line_spacing = 1.0
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.size = Pt(size_pt)
            if bold:
                r.bold = True
            if color:
                r.font.color.rgb = RGBColor.from_string(color)


def beautify_tables(docx_path, scheme="blue", font_size=9.5):
    """统一美化 docx 内所有表格，并处理跨页。返回处理的表格数。

    - 视觉：占满页宽、柔和边框、表头深色底+白字加粗居中、数据区隔行斑马纹、垂直居中。
    - 跨页：表头行 tblHeader 自动重复；每行 cantSplit 禁止被切成两页。
    """
    try:
        from docx import Document
    except ImportError:
        print("  · 未安装 python-docx，跳过表格美化（pip install python-docx 后可启用）")
        return 0
    sc = TABLE_SCHEMES.get(scheme, TABLE_SCHEMES["blue"])
    doc = Document(docx_path)
    count = 0
    for table in doc.tables:
        _style_table_layout(table)
        _style_table_borders(table, sc["border"], sz=4)
        _style_cell_margins(table)
        for ri, row in enumerate(table.rows):
            _row_flag(row, "w:cantSplit")            # 该行禁止跨页断裂
            if ri == 0:
                _row_flag(row, "w:tblHeader")        # 表头行跨页自动重复
            for cell in row.cells:
                _cell_vcenter(cell)
                if ri == 0:
                    _cell_shading(cell, sc["header"])
                    _fmt_cell_paras(cell, font_size, bold=True, color=sc["hfont"], center=True)
                else:
                    if ri % 2 == 0:                  # 隔行斑马纹
                        _cell_shading(cell, sc["zebra"])
                    _fmt_cell_paras(cell, font_size)
        count += 1
    doc.save(docx_path)
    return count


# --------------------------------------------------------------------------- #
# 代码块美化（浅灰底纹 + 细边框 + 等宽字号；pandoc 与 python-docx 两个引擎共用）
# --------------------------------------------------------------------------- #
# 原来只把代码字体设成 Consolas，导出后代码块和正文一样是白底裸文字，缺少"代码框"的
# 视觉区隔。这里补上：给整段代码加浅灰底纹 + 四周细边框，行内代码也加浅灰底。
# 关键点：pandoc 把「一个围栏代码块」整体渲染成"单个段落"（行间用 <w:br/> 断行），
# 所以对这一个段落加底纹/边框，天然就是一个连续的代码框，不会每行一个框。
def _style_code_ppr(pPr, bg, border, sz=6, space=4):
    """给段落属性 pPr 加：四周细边框 + 浅灰底纹（构成一个代码框）。

    sz 为边框粗细（1/8 磅，6≈0.75 磅）；space 为文字到边框的间距（磅）。
    若相邻两段边框定义相同，Word 会自动把它们并成一个连续边框（不画内部横线），
    因此即便某引擎把代码拆成多段，视觉上仍是一个整框。
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    if border:
        pbdr = _oxml_child(pPr, "w:pBdr", _PPR_ORDER)
        for edge in ("top", "left", "bottom", "right"):  # 不设 between，避免行间出现横线
            e = pbdr.find(qn("w:" + edge))
            if e is None:
                e = OxmlElement("w:" + edge)
                pbdr.append(e)
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(sz))
            e.set(qn("w:space"), str(space))
            e.set(qn("w:color"), border)
    if bg:
        shd = _oxml_child(pPr, "w:shd", _PPR_ORDER)
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), bg)


def _rpr_shading(rpr, fill):
    """给运行属性 rPr 加底纹（行内代码用）。"""
    from docx.oxml.ns import qn
    shd = _oxml_child(rpr, "w:shd", _RPR_ORDER)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def _run_shading(run, fill):
    """给 python-docx 的 Run 对象加底纹（行内代码用）。"""
    _rpr_shading(run._element.get_or_add_rPr(), fill)


def _set_run_size(r_el, pt):
    """给某个 <w:r> 设定字号（磅），用于统一代码块内字号。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    rpr = r_el.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r_el.insert(0, rpr)
    half = str(int(round(pt * 2)))  # OOXML 字号单位是半磅
    _oxml_child(rpr, "w:sz", _RPR_ORDER).set(qn("w:val"), half)
    _oxml_child(rpr, "w:szCs", _RPR_ORDER).set(qn("w:val"), half)


def beautify_code_blocks(docx_path, code_bg=DEFAULT_CODE_BG, code_border=DEFAULT_CODE_BORDER,
                         code_font_size=DEFAULT_CODE_FONT_SIZE, code_box=True):
    """后处理：给 docx 里所有代码块段落加底纹+边框、统一等宽字号，并给行内代码加底纹。

    识别方式走 OOXML 样式 id（pandoc 输出里代码块段落 = SourceCode，行内代码 = VerbatimChar），
    用 body.iter 遍历（连表格单元格里的代码也覆盖到）。返回处理的代码块段落数。
    注意：不改各语法高亮记号（KeywordTok/CommentTok…）的颜色，保留 pandoc 的着色。
    """
    if not code_box:
        return 0
    try:
        from docx import Document
    except ImportError:
        print("  · 未安装 python-docx，跳过代码块美化（pip install python-docx 后可启用）")
        return 0
    from docx.oxml.ns import qn
    doc = Document(docx_path)
    body = doc.element.body
    nblk = 0
    for p_el in body.iter(qn("w:p")):
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            continue
        ps = pPr.find(qn("w:pStyle"))
        if ps is None or ps.get(qn("w:val")) != "SourceCode":
            continue
        _style_code_ppr(pPr, code_bg, code_border)
        for r_el in p_el.iter(qn("w:r")):      # 统一块内字号
            _set_run_size(r_el, code_font_size)
        nblk += 1
    for r_el in body.iter(qn("w:r")):          # 行内代码底纹
        rpr = r_el.find(qn("w:rPr"))
        if rpr is None:
            continue
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None and rs.get(qn("w:val")) == "VerbatimChar":
            _rpr_shading(rpr, code_bg)
    doc.save(docx_path)
    return nblk


# --------------------------------------------------------------------------- #
# 引用块 > blockquote 美化（左竖条 + 浅底纹，两个引擎共用）
# --------------------------------------------------------------------------- #
# pandoc 把 `>` 转成 BlockText 样式段落（仅斜体+缩进，无边框底纹→看不出是引用）；
# python-docx 兜底引擎则直接在建段时套用。识别用到的样式 id 全集：
_QUOTE_STYLE_IDS = {"BlockText", "Quote", "IntenseQuote", "BlockQuote"}


def _style_quote_ppr(pPr, bar=DEFAULT_QUOTE_BAR, bg=DEFAULT_QUOTE_BG,
                     sz=24, space=10, indent=170):
    """给段落属性 pPr 加：左侧竖条(bar) + 浅底纹(bg) + 左缩进，构成引用块观感。

    只画【左】边框 → 相邻引用段之间不出现横线，多段引用视觉上连成一条竖条。
    sz=竖条粗细(1/8 磅，24≈3 磅)；space=竖条到文字的间距(磅)；indent=左缩进(twips)。
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    if bar:
        pbdr = _oxml_child(pPr, "w:pBdr", _PPR_ORDER)
        left = pbdr.find(qn("w:left"))
        if left is None:
            left = OxmlElement("w:left")
            pbdr.append(left)
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), str(sz))
        left.set(qn("w:space"), str(space))
        left.set(qn("w:color"), bar)
    if bg:
        shd = _oxml_child(pPr, "w:shd", _PPR_ORDER)
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), bg)
    if indent:
        ind = _oxml_child(pPr, "w:ind", _PPR_ORDER)
        ind.set(qn("w:left"), str(indent))


def beautify_blockquotes(docx_path, bar=DEFAULT_QUOTE_BAR, bg=DEFAULT_QUOTE_BG,
                         quote_box=True):
    """后处理：给 docx 里所有引用块段落（pandoc=BlockText 等）加左竖条 + 浅底纹，
    让 `>` 引用在视觉上清晰可辨。用 body.iter 遍历（表格单元格里的引用也覆盖）。
    返回处理的段落数。"""
    if not quote_box:
        return 0
    try:
        from docx import Document
    except ImportError:
        return 0
    from docx.oxml.ns import qn
    doc = Document(docx_path)
    body = doc.element.body
    n = 0
    for p_el in body.iter(qn("w:p")):
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            continue
        ps = pPr.find(qn("w:pStyle"))
        if ps is None or ps.get(qn("w:val")) not in _QUOTE_STYLE_IDS:
            continue
        _style_quote_ppr(pPr, bar, bg)
        n += 1
    if n:
        doc.save(docx_path)
    return n


# --------------------------------------------------------------------------- #
# 去掉标题左侧的“小黑点”：取消「与下段同页」keepNext 与「段中不分页」keepLines
# --------------------------------------------------------------------------- #
# Word 会在开了 keepNext / keepLines（或 pageBreakBefore / suppressLineNumbers）的段落
# 左侧画一个【非打印】的小黑方块作为标记。pandoc 参考文档与 Word 内置标题样式默认给各级
# 标题都开了 keepNext+keepLines，于是每个标题左边都有个小黑点。等价的手动操作是在
# 「段落 → 换行和分页」里取消“与下段同页/段中不分页”这两项。
def strip_keep_flags(docx_path):
    """删除段落的 w:keepNext（与下段同页）与 w:keepLines（段中不分页），消除标题左侧小黑点。

    从【样式表】styles.xml（标题样式在此定义，是小黑点的主要来源）和【正文】document.xml
    （直接格式）两处一并删除。显式关闭（w:val=false/0/off）的本就不显示标记，予以保留。
    返回删除的属性个数。
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError:
        return 0
    doc = Document(docx_path)
    removed = 0
    roots = [doc.element]                 # 正文 document.xml
    try:
        roots.append(doc.styles.element)  # 样式表 styles.xml（标题样式在此）
    except Exception:
        pass
    for root in roots:
        for pPr in root.iter(qn("w:pPr")):
            for tag in ("w:keepNext", "w:keepLines"):
                for el in pPr.findall(qn(tag)):
                    if el.get(qn("w:val")) in ("false", "0", "off"):
                        continue          # 显式关闭：不显示标记，保留其“显式关闭”语义
                    pPr.remove(el)
                    removed += 1
    if removed:
        doc.save(docx_path)
    return removed


# --------------------------------------------------------------------------- #
# 页码与分节（目录页与正文分节：目录用罗马数字 I/II/III，正文用阿拉伯数字从 1 起）
# --------------------------------------------------------------------------- #
# pandoc 的目录被包在一个 <w:sdt> 里（body 的直接子节点），其后紧跟正文首个标题，
# 文末只有一个 <w:sectPr>。做法：在 <w:sdt> 之后插一个"下一页"分节符，把文档切成
#   段1（目录，含可能的封面空白页）→ 罗马数字页码
#   段2（正文）           → 阿拉伯数字页码、从 1 重新计数
# 页脚放一个居中的 PAGE 域（两段共用，各自按本段的 pgNumType 渲染），距底部 0.5cm。
def _add_page_field(paragraph, cjk_font=DEFAULT_FONT):
    """在段落里插入一个居中的 PAGE 域（begin→instr→separate→缓存"1"→end）。
    带 separate + 缓存结果，未按 F9 前也能显示页码；字体随中文字体、黑色。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    def _run(children):
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rf = OxmlElement("w:rFonts")
        rf.set(qn("w:ascii"), cjk_font)
        rf.set(qn("w:hAnsi"), cjk_font)
        rf.set(qn("w:eastAsia"), cjk_font)
        rpr.append(rf)
        col = OxmlElement("w:color")
        col.set(qn("w:val"), "000000")
        rpr.append(col)
        r.append(rpr)
        for ch in children:
            r.append(ch)
        return r

    def _fld(t):
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), t)
        return el

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    cached = OxmlElement("w:t")
    cached.text = "1"
    paragraph._p.append(_run([_fld("begin")]))
    paragraph._p.append(_run([instr]))
    paragraph._p.append(_run([_fld("separate")]))
    paragraph._p.append(_run([cached]))
    paragraph._p.append(_run([_fld("end")]))


def _set_pgnumtype(sect_pr, fmt, start=None):
    """在 sectPr 里设置 pgNumType（页码格式/起始），按 schema 顺序插入。"""
    from docx.oxml.ns import qn
    for el in sect_pr.findall(qn("w:pgNumType")):
        sect_pr.remove(el)
    pnt = _oxml_child(sect_pr, "w:pgNumType", _SECTPR_ORDER)
    pnt.set(qn("w:fmt"), fmt)
    if start is not None:
        pnt.set(qn("w:start"), str(start))
    else:
        if qn("w:start") in pnt.attrib:
            del pnt.attrib[qn("w:start")]


def apply_page_numbering(docx_path, footer_cm=0.5, cjk_font=DEFAULT_FONT, has_cover=False):
    """给导出的 docx 加「分节 + 多页码体系」。返回 True/False。

    节的构成（存在才有，顺序固定）：封面 → 目录 → 正文
      · 封面(若有)：页脚留空，不显示页码；
      · 目录(若有)：罗马数字 I/II/III，从 I 起（封面不计入，故显式 start=1）；
      · 正文       ：阿拉伯数字，从 1 起。
    各节之间「下一页」分节；页脚 PAGE 域居中、距底部 footer_cm 厘米。
    has_cover 表示最前是否已由 _prepend_cover_page 生成封面节。
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        from docx.shared import Cm
        from docx.enum.section import WD_SECTION
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except ImportError:
        print("  · 未安装 python-docx，跳过页码/分节")
        return False
    import copy

    # ---- 第一趟：若有目录(sdt)，在其后插入「下一页」分节符，把目录与正文切成两节 ----
    doc = Document(docx_path)
    body = doc.element.body
    sdt = body.find(qn("w:sdt"))               # pandoc 目录容器
    final_sect = body.find(qn("w:sectPr"))     # 文末（正文段）分节属性
    if final_sect is None:
        return False
    two_section = sdt is not None
    if two_section:
        toc_sect = copy.deepcopy(final_sect)   # 克隆以继承同样的纸张/页边距
        for tag in ("w:headerReference", "w:footerReference", "w:pgNumType", "w:type"):
            for el in toc_sect.findall(qn(tag)):
                toc_sect.remove(el)
        p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        pPr.append(toc_sect)                    # 段落 pPr 内含 sectPr = 该段最后一段
        p.append(pPr)
        sdt.addnext(p)                          # 放到目录之后、正文之前
        doc.save(docx_path)
        doc = Document(docx_path)               # 重新打开，识别新增 section

    secs = doc.sections
    # ---- 正文节（永远是最后一节）：阿拉伯数字从 1 ----
    body_sec = secs[-1]
    body_sec.footer_distance = Cm(footer_cm)
    body_sec.start_type = WD_SECTION.NEW_PAGE
    _set_pgnumtype(body_sec._sectPr, "decimal", start=1)

    if two_section:
        # ---- 目录节（正文前一节）：罗马数字，从 I 起 ----
        toc_sec = secs[-2]
        toc_sec.footer_distance = Cm(footer_cm)
        _set_pgnumtype(toc_sec._sectPr, "upperRoman", start=1)
        toc_sec.footer.is_linked_to_previous = False
        fp = toc_sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_page_field(fp, cjk_font)
        body_sec.footer.is_linked_to_previous = True   # 正文继承目录页脚（按本节阿拉伯渲染）
    else:
        body_sec.footer.is_linked_to_previous = False
        fp = body_sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_page_field(fp, cjk_font)

    # ---- 封面节（若有，永远是第一节）：页脚留空、不显示页码 ----
    if has_cover and len(secs) >= 2:
        cover_sec = secs[0]
        cover_sec.footer_distance = Cm(footer_cm)
        cover_sec.footer.is_linked_to_previous = False   # 独立、空页脚
        for cp in cover_sec.footer.paragraphs:           # 清掉任何残留内容
            for cr in list(cp.runs):
                cr._element.getparent().remove(cr._element)

    doc.save(docx_path)
    return True


# --------------------------------------------------------------------------- #
# Mermaid 流程图渲染（```mermaid → PNG，用本地 mmdc + 系统 Chrome/Edge，离线）
# 以及图片自适应页宽（mermaid 图与用户自带图片，超出正文宽度就等比缩小）
# --------------------------------------------------------------------------- #
_MERMAID_RE = re.compile(r"(?ms)^[ \t]*```[ \t]*mermaid[ \t]*\r?\n(.*?)\r?\n[ \t]*```[ \t]*$")


def _find_chromium():
    """找一个可用的 Chromium 内核浏览器给 mmdc 用（Chrome 优先，其次 Edge）。"""
    envp = os.environ.get("PUPPETEER_EXECUTABLE_PATH")
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pfx = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    cands = [envp,
             os.path.join(pf, "Google", "Chrome", "Application", "chrome.exe"),
             os.path.join(pfx, "Google", "Chrome", "Application", "chrome.exe"),
             os.path.join(pfx, "Microsoft", "Edge", "Application", "msedge.exe"),
             os.path.join(pf, "Microsoft", "Edge", "Application", "msedge.exe")]
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def render_mermaid_blocks(md_text, tmpdir, cjk_font=DEFAULT_FONT, scale=2, width=1400):
    """把 md 里的 ```mermaid 块用 mmdc 渲染成 PNG，替换成 ![](绝对路径.png)。

    返回 (新md文本, 生成png列表, 成功数, 总块数)。多块【并行】渲染（每块独立 mmdc+浏览器
    进程，天然可并发）。缺 mmdc/浏览器或某块渲染失败时，该块原样保留为代码文本（降级不报错）。
    CJK 通过 mermaid 配置里的 fontFamily 走中文字体。
    """
    import json
    from concurrent.futures import ThreadPoolExecutor
    matches = list(_MERMAID_RE.finditer(md_text))
    total = len(matches)
    if total == 0:
        return md_text, [], 0, 0
    exe = shutil.which("mmdc")
    browser = _find_chromium()
    if not exe or not browser:
        return md_text, [], 0, total       # 缺依赖 → 全部保留为文本
    # mmdc CLI 代际兼容：v12 起 -w/--width 被 --size（最大边长）取代，旧版只有 -w
    try:
        _help = subprocess.run([exe, "--help"], capture_output=True, text=True,
                               timeout=30, encoding="utf-8", errors="replace").stdout or ""
    except Exception:
        _help = ""
    size_opt = ["--size"] if ("--size" in _help and "-w," not in _help) else ["-w"]
    os.makedirs(tmpdir, exist_ok=True)
    pptr = os.path.join(tmpdir, "_pptr.json")
    with open(pptr, "w", encoding="utf-8") as fh:
        json.dump({"executablePath": browser, "args": ["--no-sandbox"]}, fh)
    mmcfg = os.path.join(tmpdir, "_mmcfg.json")
    with open(mmcfg, "w", encoding="utf-8") as fh:
        json.dump({"fontFamily": f'{cjk_font}, "Microsoft YaHei", sans-serif'}, fh)

    def _render(idx, code):
        mmd = os.path.join(tmpdir, f"_mmd_{idx}.mmd")
        png = os.path.join(tmpdir, f"_mmd_{idx}.png")
        with open(mmd, "w", encoding="utf-8") as fh:
            fh.write(code)
        cmd = [exe, "-i", mmd, "-o", png, "-b", "white",
               *size_opt, str(width), "-s", str(scale), "-p", pptr, "-c", mmcfg]
        if os.name == "nt" and exe.lower().endswith((".cmd", ".bat")):
            cmd = ["cmd", "/c"] + cmd
        try:
            subprocess.run(cmd, capture_output=True, text=True, timeout=120,
                           encoding="utf-8", errors="replace")
        except Exception:
            pass
        return png if os.path.isfile(png) else None

    # 并行渲染（上限 4，避免同时开太多浏览器）；结果按块序保存
    with ThreadPoolExecutor(max_workers=min(4, total)) as ex:
        results = list(ex.map(lambda p: _render(p[0], p[1].group(1)),
                              list(enumerate(matches))))

    # 按原顺序回填：成功→图片引用，失败→原代码块
    pngs = [r for r in results if r]
    out = []
    last = 0
    for m, res in zip(matches, results):
        out.append(md_text[last:m.start()])
        if res:
            out.append(f"\n\n![]({res.replace(chr(92), '/')})\n\n")  # 绝对路径，免 resource-path 歧义
        else:
            out.append(m.group(0))                                    # 渲染失败 → 保留原代码块
        last = m.end()
    out.append(md_text[last:])
    return "".join(out), pngs, len(pngs), total


# --------------------------------------------------------------------------- #
# LaTeX 数学公式（$…$ / $$…$$ / \(…\) / \[…\]）
#   pandoc 引擎：+tex_math_dollars 扩展 → Word 原生公式（OMML），无需出图；
#   内置引擎  ：matplotlib mathtext 离线渲染成 PNG —— 行内公式→内嵌小图标记，
#               块级公式→独立图片引用。两引擎共用本段的「分隔符归一化 + 代码保护」。
# --------------------------------------------------------------------------- #
# 行内 $ 的边界 guards 与 pandoc 规则一致：开 $ 右侧不能是空白、闭 $ 左侧不能是空白
# 且右侧不能紧跟数字 —— 因此 $100 和 $200 不会被误判成公式。
_LATEX_PAREN_RE = re.compile(r"\\\((.+?)\\\)")
_LATEX_BRACKET_RE = re.compile(r"(?s)\\\[(.+?)\\\]")
_MATH_DISPLAY_RE = re.compile(r"(?s)\$\$\s*(.+?)\s*\$\$")
_MATH_INLINE_RE = re.compile(r"(?<![\w$])\$(?!\s)([^$\n]+?)(?<!\s)\$(?!\d)")
_MATH_IMG_RE = re.compile(r"MD2DOCXMATH<([^<>]+)>")


def _math_to_png(tex, out_png, font_size=12):
    """离线把一个 LaTeX 片段渲染成 PNG（matplotlib mathtext，白底黑字）。

    以 dpi=300 输出并保留 dpi 元数据 → python-docx 按原生物理尺寸嵌入，
    行内公式高度正好与正文字号对齐。语法不支持时抛异常，由调用方降级为原文。
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    fig = plt.figure(figsize=(8, 3))
    fig.patch.set_facecolor("white")
    fig.text(0.01, 0.5, f"${tex.strip()}$", fontsize=font_size,
             va="center", ha="left", color="black")
    fig.savefig(out_png, dpi=300, bbox_inches="tight", pad_inches=0.02)
    plt.close(fig)


def preprocess_math(md_text, tmpdir, to_images=True, inline_fs=11, display_fs=12):
    """识别 LaTeX 数学公式并转换。代码围栏与行内代码 `…` 里的内容一律不动。

    1) 分隔符归一化（两引擎共用）：\\(…\\) → $…$，\\[…\\] → $$…$$；
    2) to_images=True（内置引擎）：再把 $…$/$$…$$ 渲染成 PNG ——
       行内 → 内嵌标记 MD2DOCXMATH<路径>（由 _add_inline 转成小图）；
       块级 → 独立图片引用 ![](绝对路径)。
       渲染失败的表达式原样保留（降级不报错）。
    返回 (新文本, 生成png列表, 转换/识别数)。
    """
    if "$" not in md_text and "\\(" not in md_text and "\\[" not in md_text:
        return md_text, [], 0
    if to_images:
        try:
            import matplotlib  # noqa: F401
        except ImportError:
            to_images = False
    os.makedirs(tmpdir, exist_ok=True)
    pngs, count = [], 0

    def _render(tex, display):
        nonlocal count
        png = os.path.join(tmpdir, f"_math_{len(pngs) + 1}.png")
        try:
            _math_to_png(tex, png, display_fs if display else inline_fs)
        except Exception:
            return None
        pngs.append(png)
        count += 1
        return png.replace(chr(92), "/")

    def _proc(t):
        # 注意 strip：\( x \) 内侧空格若不剥掉，$ x $ 会被行内公式规则(开$右侧非空白)拒收
        t = _LATEX_BRACKET_RE.sub(lambda m: f"$${m.group(1).strip()}$$", t)
        t = _LATEX_PAREN_RE.sub(lambda m: f"${m.group(1).strip()}$", t)
        if not to_images:                       # pandoc：归一化后交 tex_math_dollars
            return t
        def _d(m):                              # 块级公式 → 独立图片
            p = _render(m.group(1), True)
            return f"\n\n![]({p})\n\n" if p else m.group(0)
        t = _MATH_DISPLAY_RE.sub(_d, t)
        def _i(m):                              # 行内公式 → 内嵌标记
            p = _render(m.group(1), False)
            return f"MD2DOCXMATH<{p}>" if p else m.group(0)
        parts = re.split(r"(`[^`]*`)", t)       # 行内代码里的 $ 不动
        for k in range(0, len(parts), 2):
            parts[k] = _MATH_INLINE_RE.sub(_i, parts[k])
        return "".join(parts)

    # 按代码围栏切段：围栏内（含 mermaid 源码）原样保留
    segs, buf, in_code = [], [], False
    for line in md_text.split("\n"):
        if line.strip().startswith("```"):
            if in_code:
                buf.append(line)
                segs.append((True, "\n".join(buf)))
                buf, in_code = [], False
            else:
                segs.append((False, "\n".join(buf)))
                buf, in_code = [line], True
        else:
            buf.append(line)
    segs.append((in_code, "\n".join(buf)))
    # 关键：段与段之间必须补回 "\n"。split("\n") 拆出的换行属于“行分隔符”，
    # 若直接 "".join 会把相邻两段粘成一行——围栏标记 ``` 不再位于行首，
    # 代码块就不被识别，块内以 # 开头的注释行会被当成标题（实测踩过的坑）。
    parts = [_proc(t) if not code else t for code, t in segs]
    result = "\n".join(parts)
    if not to_images:                           # 仅归一化：估算识别数（打日志用）
        count = sum(len(_MATH_DISPLAY_RE.findall(t)) + len(_MATH_INLINE_RE.findall(t))
                    for code, t in segs if not code)
    return result, pngs, count


def fit_images_to_page(docx_path):
    """把超过"正文可用宽度"的行内图片等比缩小到刚好铺满正文宽度。返回处理张数。
    同时惠及 mermaid 生成图与用户自带的大图，避免图片溢出页面。"""
    try:
        from docx import Document
    except ImportError:
        return 0
    doc = Document(docx_path)
    shapes = doc.inline_shapes
    if not shapes:
        return 0
    from docx.shared import Cm
    # pandoc/自插节的 sectPr 可能缺显式 pgSz/pgMar：按 Word 的继承规则，
    # 缺失项沿用「前一节」已知值，全缺再用 A4 + 3.18cm 边距兜底；
    # 取所有节里「最窄的正文可用宽度」做上限，保证图片在任何一节都不溢出。
    last_pw, last_lm, last_rm = Cm(21.0), Cm(3.18), Cm(3.18)
    widths = []
    for s in doc.sections:
        if s.page_width is not None:
            last_pw = s.page_width
        if s.left_margin is not None:
            last_lm = s.left_margin
        if s.right_margin is not None:
            last_rm = s.right_margin
        widths.append(last_pw - last_lm - last_rm)
    content_w = min(widths) if widths else (Cm(21.0) - Cm(3.18) * 2)
    n = 0
    for shp in shapes:
        try:
            w, h = shp.width, shp.height
            if w and h and content_w and w > content_w:
                shp.height = int(h * (content_w / w))
                shp.width = int(content_w)
                n += 1
        except Exception:
            continue
    if n:
        doc.save(docx_path)
    return n


def convert_with_pandoc(pandoc, src, out, cjk_font, toc, toc_depth,
                        table_scheme, table_font, latin_font, text_color,
                        blank_first_page=True,
                        code_bg=DEFAULT_CODE_BG, code_border=DEFAULT_CODE_BORDER,
                        code_font_size=DEFAULT_CODE_FONT_SIZE, code_box=True,
                        page_numbers=True, footer_cm=0.5,
                        mermaid=True, image_attrs=False, math=True,
                        quote_bar=DEFAULT_QUOTE_BAR, quote_bg=DEFAULT_QUOTE_BG,
                        quote_box=True, cover_title=None,
                        page_size=DEFAULT_PAGE_SIZE):
    workdir = os.path.dirname(os.path.abspath(out)) or "."
    ref_path = os.path.join(workdir, ".ref_cjk.docx")
    tmpdir = os.path.join(workdir, ".md2docx_tmp")
    src_dir = os.path.dirname(os.path.abspath(src)) or "."
    tmp_files = []
    print(f"  · 生成参考文档(字体注入): {ref_path}")
    build_reference_doc(pandoc, cjk_font, latin_font, text_color, ref_path, page_size)

    # 预处理：LaTeX 数学公式（\(…\)/\[…\] 归一化为 $ 形式，配合 tex_math_dollars
    # 由 pandoc 转成 Word 原生公式 OMML）+ ```mermaid 渲染成图片（失败自动降级为代码文本）
    src_in = src
    if math or mermaid:
        try:
            with open(src, encoding="utf-8") as fh:
                original = fh.read()
            new_md = original
            if math:
                new_md, _, n_math = preprocess_math(new_md, tmpdir, to_images=False)
                if n_math:
                    print(f"  · 识别到 {n_math} 处 LaTeX 数学公式（转 Word 原生公式）")
            if mermaid:
                new_md, pngs, ok, total = render_mermaid_blocks(new_md, tmpdir, cjk_font)
                if total:
                    if ok:
                        print(f"  · 已渲染 {ok}/{total} 个 mermaid 流程图为图片")
                    if ok < total:
                        print(f"  · 注意：{total - ok} 个 mermaid 块未渲染（缺 mmdc/浏览器或语法错误），按代码文本导出")
                    tmp_files += pngs
            if new_md != original:
                src_in = os.path.join(tmpdir, "_src.md")
                with open(src_in, "w", encoding="utf-8") as fh:
                    fh.write(new_md)
                tmp_files.append(src_in)
        except Exception as e:
            print(f"  · 预处理跳过（{type(e).__name__}）")

    cmd = [
        pandoc,
        src_in,
        "-o", out,
        "--from=gfm+raw_attribute"
        + ("+tex_math_dollars" if math else "")   # $…$/$$…$$ → Word 原生公式(OMML)
        + ("+attributes" if image_attrs else ""),  # +attributes 才认 {width=} 图片尺寸（封面 logo）
        "--standalone",
        f"--reference-doc={ref_path}",
        "-M", "toc-title=目录",
        # 让 md 里的相对路径图片能被找到：源目录 + 临时目录（mermaid 图为绝对路径不受影响）
        "--resource-path", os.pathsep.join([src_dir, workdir, tmpdir]),
    ]
    if toc:
        cmd += ["--toc", f"--toc-depth={toc_depth}"]

    print("  · 调用 pandoc: " + " ".join(
        (f'"{c}"' if (" " in c or "题" in c) else c) for c in cmd))
    try:
        try:
            subprocess.run(cmd, check=True, timeout=600)
        except subprocess.TimeoutExpired:
            print("[错误] pandoc 转换超时（>600s），已中止。", file=sys.stderr)
            raise
        except subprocess.CalledProcessError:
            if os.path.isfile(out):
                print(f"[提示] 目标文件可能正被 Word/WPS 打开、无法写入：{out}", file=sys.stderr)
                print("       请关闭该文档后重试，或用 -o 指定新的输出文件名。", file=sys.stderr)
            raise

        if toc:
            _mark_toc_dirty(out)  # 让 Word 打开时静默刷新目录（不再弹“更新域”提示）

        n = beautify_tables(out, table_scheme, table_font)
        if n:
            print(f"  · 已美化 {n} 张表格（配色={table_scheme}｜表头跨页重复 + 行禁止断裂）")

        m = beautify_code_blocks(out, code_bg, code_border, code_font_size, code_box)
        if m:
            print(f"  · 已美化 {m} 个代码块（浅灰底纹 + 细边框 + 等宽 {code_font_size}pt）")

        q = beautify_blockquotes(out, quote_bar, quote_bg, quote_box)
        if q:
            print(f"  · 已美化 {q} 个引用块（左竖条 + 浅底纹）")

        pb = _apply_pagebreak_markers(out)
        if pb:
            print(f"  · 已应用 {pb} 处分页标记（如封面独占一页）")

        fi = fit_images_to_page(out)
        if fi:
            print(f"  · 已把 {fi} 张过宽图片等比缩放到正文宽度")

        cover_ok = False
        if blank_first_page:
            cover_ok = _prepend_cover_page(out, cover_title, cjk_font)
            if cover_ok:
                print(f"  · 已插入封面页（标题：{cover_title or '(空)'}；封面不显示页码）")

        if page_numbers and apply_page_numbering(out, footer_cm, cjk_font,
                                                 has_cover=cover_ok):
            print(f"  · 已加页码：封面不计页码、目录=罗马数字 I/II/III、正文=阿拉伯数字从 1 起"
                  f"（下一页分节，页脚距底 {footer_cm}cm）")
    finally:
        # 无论成功或失败，都清理临时参考文档与 mermaid 临时文件，避免残留
        try:
            os.remove(ref_path)
        except OSError:
            pass
        for fp in tmp_files:
            try:
                os.remove(fp)
            except OSError:
                pass
        for junk in ("_pptr.json", "_mmcfg.json"):
            try:
                os.remove(os.path.join(tmpdir, junk))
            except OSError:
                pass
        try:
            os.rmdir(tmpdir)          # 仅当空目录时才会成功，非空则保留
        except OSError:
            pass


# --------------------------------------------------------------------------- #
# 引擎二：python-docx 兜底转换器（轻量 Markdown 解析）
# --------------------------------------------------------------------------- #
_INLINE = re.compile(
    r"\[(?P<ltext>[^\]]+)\]\((?P<lurl>[^)\s]+)\)"      # [文字](url) 链接
    r"|\*\*(?P<bold>.+?)\*\*"                            # **加粗**
    r"|`(?P<code>[^`]+)`"                                # `行内代码`
    r"|~~(?P<strike>.+?)~~"                              # ~~删除线~~
    r"|(?<!\*)\*(?!\*)(?P<italic>[^*\n]+?)\*(?!\*)"      # *斜体*（避开 **）
)


def _add_hyperlink(paragraph, url, text):
    """在段落里插入一个真正的超链接（蓝色 + 下划线）。建立关系失败则退化为普通文字。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    except Exception:
        paragraph.add_run(text)
        return
    hyper = OxmlElement("w:hyperlink")
    hyper.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color"); c.set(qn("w:val"), "0563C1"); rpr.append(c)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rpr.append(u)
    r.append(rpr)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
    r.append(t)
    hyper.append(r)
    paragraph._p.append(hyper)


def _add_inline(paragraph, text, code_font="Consolas", code_bg=None):
    """把一段文本按 [链接](url) / **加粗** / `行内代码` / ~~删除线~~ / *斜体* 解析后加入段落。
    传入 code_bg 时给行内代码补一层浅灰底纹（与代码块观感一致）。
    含 MD2DOCXMATH<路径> 标记（行内数学公式图片）时，把标记转成内嵌小图。"""
    if "MD2DOCXMATH<" in text:                       # 行内公式 → 内嵌小图
        pos = 0
        for m in _MATH_IMG_RE.finditer(text):
            if m.start() > pos:
                _add_inline(paragraph, text[pos:m.start()], code_font, code_bg)
            run = paragraph.add_run()
            try:
                run.add_picture(m.group(1))
            except Exception:
                paragraph.add_run("[公式渲染失败]")
            pos = m.end()
        if pos < len(text):
            _add_inline(paragraph, text[pos:], code_font, code_bg)
        return
    pos = 0
    for m in _INLINE.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        if m.group("ltext") is not None:
            _add_hyperlink(paragraph, m.group("lurl"), m.group("ltext"))
        elif m.group("bold") is not None:
            paragraph.add_run(m.group("bold")).bold = True
        elif m.group("code") is not None:
            run = paragraph.add_run(m.group("code"))
            run.font.name = code_font
            if code_bg:
                _run_shading(run, code_bg)
        elif m.group("strike") is not None:
            paragraph.add_run(m.group("strike")).font.strike = True
        elif m.group("italic") is not None:
            paragraph.add_run(m.group("italic")).italic = True
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def _split_table_row(line):
    """拆分一行管道表格：|a|b|c| -> ['a','b','c']，处理首尾空单元。"""
    cells = line.strip().split("|")
    if cells and cells[0] == "":
        cells = cells[1:]
    if cells and cells[-1] == "":
        cells = cells[:-1]
    return [c.strip() for c in cells]


def _is_table_separator(line):
    """判断是否为表头分隔行，如 |---|:--:|---|。"""
    s = line.strip()
    if "|" not in s or "-" not in s:
        return False
    body = s.strip("|")
    for cell in body.split("|"):
        cell = cell.strip()
        if not cell or set(cell) - set(":- "):
            return False
    return True


def convert_with_python_docx(src, out, cjk_font, table_scheme, table_font,
                             latin_font, text_color, blank_first_page=True,
                             code_bg=DEFAULT_CODE_BG, code_border=DEFAULT_CODE_BORDER,
                             code_font_size=DEFAULT_CODE_FONT_SIZE, code_box=True,
                             page_numbers=True, footer_cm=0.5, mermaid=True, math=True,
                             quote_bar=DEFAULT_QUOTE_BAR, quote_bg=DEFAULT_QUOTE_BG,
                             quote_box=True, cover_title=None,
                             page_size=DEFAULT_PAGE_SIZE):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.shared import Pt, RGBColor, Cm, Emu, Inches
    from docx.oxml.ns import qn

    inline_bg = code_bg if code_box else None  # 行内代码底纹（关代码框时不加）
    src_dir = os.path.dirname(os.path.abspath(src)) or "."
    workdir = os.path.dirname(os.path.abspath(out)) or "."
    tmpdir = os.path.join(workdir, ".md2docx_tmp")
    tmp_files = []

    with open(src, "r", encoding="utf-8") as fh:
        md_text = fh.read()
    if mermaid:                                     # 先把 ```mermaid 渲染成图片引用
        try:
            md_text, pngs, ok, total = render_mermaid_blocks(md_text, tmpdir, cjk_font)
            if ok:
                tmp_files += pngs
                print(f"  · 已渲染 {ok}/{total} 个 mermaid 流程图为图片")
            elif total:
                print(f"  · 注意：{total} 个 mermaid 块未渲染（缺 mmdc/浏览器），按代码文本导出")
        except Exception as e:
            print(f"  · mermaid 渲染跳过（{type(e).__name__}）")
    if math:                                        # LaTeX 公式 → PNG（matplotlib 离线）
        try:
            md_text, mpngs, n_math = preprocess_math(md_text, tmpdir, to_images=True)
            if mpngs:
                tmp_files += mpngs
                print(f"  · 已把 {n_math} 处 LaTeX 数学公式渲染为图片")
            elif n_math:
                print(f"  · 注意：{n_math} 处数学公式未渲染（缺 matplotlib 或语法不支持），按原文保留")
        except Exception as e:
            print(f"  · 数学公式渲染跳过（{type(e).__name__}）")
    lines = md_text.split("\n")

    doc = Document()
    _inject_fonts(doc, cjk_font, latin_font, text_color)
    _set_page_size(doc, page_size)

    # 正文可用宽度（用于图片自适应），无显式页面设置时用 A4-2.54cm 兜底
    _sec0 = doc.sections[0]
    _cw = (_sec0.page_width or Cm(21.0)) - (_sec0.left_margin if _sec0.left_margin is not None else Cm(2.54)) \
          - (_sec0.right_margin if _sec0.right_margin is not None else Cm(2.54))

    def _add_image(path, attr=None):
        """嵌入一张图片。attr 里带 {width=Xcm|Xin} 则按指定宽度(等比缩高)；
        否则仅在超出正文宽度时等比缩到正文宽度。path 可为绝对或相对源目录。"""
        cand = path if os.path.isabs(path) else os.path.join(src_dir, path)
        if not os.path.isfile(cand):
            cand2 = os.path.join(tmpdir, path)
            cand = cand2 if os.path.isfile(cand2) else cand
        if not os.path.isfile(cand):
            doc.add_paragraph(f"[缺图片: {path}]")
            return
        wm = re.search(r"width\s*=\s*([\d.]+)\s*(cm|in)?", attr or "")
        try:
            if wm:                                        # 显式定尺（如封面 logo）
                val = float(wm.group(1))
                doc.add_picture(cand, width=(Inches(val) if wm.group(2) == "in" else Cm(val)))
            else:
                doc.add_picture(cand)
                shp = doc.inline_shapes[-1]
                if shp.width and _cw and shp.width > _cw:
                    shp.height = int(shp.height * (_cw / shp.width))
                    shp.width = int(_cw)
        except Exception:
            doc.add_paragraph(f"[图片无法嵌入: {path}]")

    i, n = 0, len(lines)
    in_code = False
    code_buf = []

    def flush_code():
        if not code_buf:
            return
        p = doc.add_paragraph()
        run = p.add_run("\n".join(code_buf))
        run.font.name = CODE_FONT
        run.font.size = Pt(code_font_size)
        run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        for t in run._element.findall(qn("w:t")):   # 保留每行前导空白（缩进）
            t.set(XML_SPACE, "preserve")
        pf = p.paragraph_format
        pf.left_indent = Pt(6)
        pf.right_indent = Pt(6)
        pf.space_before = Pt(4)
        pf.space_after = Pt(4)
        pf.line_spacing = 1.0
        if code_box:                                # 浅灰底纹 + 细边框，构成代码框
            _style_code_ppr(p._p.get_or_add_pPr(), code_bg, code_border)
        code_buf.clear()

    while i < n:
        raw = lines[i]
        line = raw.rstrip("\n")

        # 代码围栏
        if line.strip().startswith("```"):
            if in_code:
                flush_code()
                in_code = False
            else:
                in_code = True
            i += 1
            continue
        if in_code:
            code_buf.append(line)
            i += 1
            continue

        stripped = line.strip()

        # 空行
        if not stripped:
            i += 1
            continue

        # 合并分页哨兵（--folder 合并模式在文件之间插入）→ 下一页
        if stripped == _PAGEBREAK_SENTINEL:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            i += 1
            continue

        # 独立成行的图片：![alt](path) 或 ![alt](path){width=Xcm}
        mi = re.match(r"^!\[[^\]]*\]\(\s*([^)\s]+)(?:\s+[^)]*)?\)(?:\{([^}]*)\})?$", stripped)
        if mi:
            _add_image(mi.group(1), mi.group(2))
            i += 1
            continue

        # 分隔线（非表格）
        if re.fullmatch(r"-{3,}|\*{3,}|_{3,}", stripped):
            doc.add_paragraph().add_run("").add_break()
            i += 1
            continue

        # 标题
        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            level = len(m.group(1))
            title = m.group(2).strip()
            doc.add_heading(title, level=min(level, 6))
            i += 1
            continue

        # 表格：当前行含 | 且下一行是分隔行
        if "|" in stripped and i + 1 < n and _is_table_separator(lines[i + 1]):
            header = _split_table_row(stripped)
            rows = []
            j = i + 2
            while j < n and "|" in lines[j] and lines[j].strip():
                if _is_table_separator(lines[j]):
                    j += 1
                    continue
                rows.append(_split_table_row(lines[j]))
                j += 1
            cols = len(header)
            table = doc.add_table(rows=1, cols=cols)
            # 边框 / 底色 / 斑马纹 / 跨页统一交给 beautify_tables 处理
            for c, txt in enumerate(header):
                cell = table.rows[0].cells[c]
                cell.paragraphs[0].text = ""
                _add_inline(cell.paragraphs[0], txt, code_bg=inline_bg)
                for r in cell.paragraphs[0].runs:
                    r.bold = True
            for row in rows:
                cells = table.add_row().cells
                for c in range(cols):
                    val = row[c] if c < len(row) else ""
                    cells[c].paragraphs[0].text = ""
                    _add_inline(cells[c].paragraphs[0], val, code_bg=inline_bg)
            i = j
            continue

        # 引用块 > blockquote：普通段 + 左竖条 + 浅底纹（不依赖模板是否有 Intense Quote）
        if stripped.startswith(">"):
            content = stripped.lstrip(">").strip()
            p = doc.add_paragraph()
            _add_inline(p, content, code_bg=inline_bg)
            if quote_box:
                _style_quote_ppr(p._p.get_or_add_pPr(), quote_bar, quote_bg)
            i += 1
            continue

        # 无序列表
        m = re.match(r"^(\s*)[-*+]\s+(.*)$", raw)
        if m:
            indent = len(m.group(1))
            p = doc.add_paragraph(style="List Bullet")
            if indent >= 2:
                p.paragraph_format.left_indent = Pt(18 * (indent // 2 + 1))
            _add_inline(p, m.group(2), code_bg=inline_bg)
            i += 1
            continue

        # 有序列表
        m = re.match(r"^(\s*)\d+\.\s+(.*)$", raw)
        if m:
            p = doc.add_paragraph(style="List Number")
            _add_inline(p, m.group(2), code_bg=inline_bg)
            i += 1
            continue

        # 普通段落
        p = doc.add_paragraph()
        _add_inline(p, stripped, code_bg=inline_bg)
        i += 1

    if in_code:
        flush_code()

    doc.save(out)
    n = beautify_tables(out, table_scheme, table_font)
    if n:
        print(f"  · 已美化 {n} 张表格（配色={table_scheme}｜表头跨页重复 + 行禁止断裂）")
    cover_ok = False
    if blank_first_page:
        cover_ok = _prepend_cover_page(out, cover_title, cjk_font)
        if cover_ok:
            print(f"  · 已插入封面页（标题：{cover_title or '(空)'}；封面不显示页码）")
    if page_numbers and apply_page_numbering(out, footer_cm, cjk_font, has_cover=cover_ok):
        print(f"  · 已加页码：封面不计页码、正文阿拉伯数字从 1 起（页脚距底 {footer_cm}cm）"
              f"；内置引擎无自动目录，故不分罗马数字节")
    print("  · 提示：内置引擎不生成自动目录，请在 Word 中「引用→目录」一键插入；"
          "或安装 pandoc 获得完整目录。")
    for fp in tmp_files:                       # 清理 mermaid 临时图片
        try:
            os.remove(fp)
        except OSError:
            pass
    for junk in ("_pptr.json", "_mmcfg.json"):
        try:
            os.remove(os.path.join(tmpdir, junk))
        except OSError:
            pass
    try:
        os.rmdir(tmpdir)
    except OSError:
        pass


# --------------------------------------------------------------------------- #
# 主入口
# --------------------------------------------------------------------------- #
def _convert_one(engine, pandoc, src, out, args, page_numbers, code_box,
                 quote_box, toc, cover_title=None):
    """按引擎把单个 md(src) 转成 docx(out)。单文件、目录合并、目录批量三处共用，
    避免重复那一长串参数。toc 单独传：合并模式让整本共用一个目录。"""
    if engine == "pandoc":
        convert_with_pandoc(pandoc, src, out, args.font,
                            toc=toc, toc_depth=args.toc_depth,
                            table_scheme=args.table_style, table_font=args.table_font_size,
                            latin_font=args.latin_font, text_color=args.text_color,
                            blank_first_page=not args.no_blank_page,
                            code_bg=args.code_bg, code_border=args.code_border,
                            code_font_size=args.code_font_size, code_box=code_box,
                            page_numbers=page_numbers, footer_cm=args.footer_distance_cm,
                            mermaid=not args.no_mermaid, image_attrs=args.image_attrs,
                            math=not args.no_math,
                            quote_bar=args.quote_bar, quote_bg=args.quote_bg,
                            quote_box=quote_box,
                            cover_title=cover_title, page_size=args.page_size)
    else:
        convert_with_python_docx(src, out, args.font,
                                 table_scheme=args.table_style, table_font=args.table_font_size,
                                 latin_font=args.latin_font, text_color=args.text_color,
                                 blank_first_page=not args.no_blank_page,
                                 code_bg=args.code_bg, code_border=args.code_border,
                                 code_font_size=args.code_font_size, code_box=code_box,
                                 page_numbers=page_numbers, footer_cm=args.footer_distance_cm,
                                 mermaid=not args.no_mermaid,
                                 math=not args.no_math,
                                 quote_bar=args.quote_bar, quote_bg=args.quote_bg,
                                 quote_box=quote_box,
                                 cover_title=cover_title, page_size=args.page_size)
    # 去掉标题左侧的小黑点（取消「与下段同页/段中不分页」），除非用户显式要求保留
    if not args.keep_heading_together and os.path.isfile(out):
        k = strip_keep_flags(out)
        if k:
            print(f"  · 已取消 {k} 处「与下段同页/段中不分页」（标题左侧小黑点消失）")


def main():
    parser = argparse.ArgumentParser(
        description="将 Markdown 转换为 Word (.docx)；单文件或整目录(--folder)批量/合并")
    parser.add_argument("input", nargs="?", default=DEFAULT_INPUT,
                        help=f"输入 Markdown 文件（默认: {DEFAULT_INPUT}）；给了 --folder 时忽略")
    parser.add_argument("-fd", "--folder", default=None,
                        help="导出整个目录：把其中顶层 .md 按自然排序【合并成一个】docx；"
                             "配合 --split 改为每个文件各导一个")
    parser.add_argument("--split", action="store_true",
                        help="配合 --folder：不合并，目录下每个 .md 各导出一个 docx")
    parser.add_argument("-o", "--output", default=None,
                        help="输出路径：单文件/合并=输出 .docx；--split 时=输出【目录】")
    parser.add_argument("--font", default=DEFAULT_FONT,
                        help=f"中文字体（默认: {DEFAULT_FONT}）")
    parser.add_argument("--toc-depth", type=int, default=2,
                        help="自动目录深度：1=仅模块，2=模块+小节(默认)，3=含每道题")
    parser.add_argument("--no-toc", action="store_true", help="不生成自动目录")
    parser.add_argument("--engine", choices=["auto", "pandoc", "python-docx"],
                        default="auto", help="转换引擎（默认 auto：优先 pandoc）")
    parser.add_argument("--table-style", default="blue",
                        choices=list(TABLE_SCHEMES.keys()),
                        help="表格配色：blue(默认)/gray/green/teal")
    parser.add_argument("--table-font-size", type=float, default=9.5,
                        help="表格内文字字号(pt)，列多挤时可调小，默认 9.5")
    parser.add_argument("--latin-font", default=DEFAULT_LATIN_FONT,
                        help=f"英文/数字字体（默认: {DEFAULT_LATIN_FONT}）")
    parser.add_argument("--text-color", default=DEFAULT_TEXT_COLOR,
                        help=f"正文字色 6 位十六进制（默认: {DEFAULT_TEXT_COLOR}，即黑色）")
    parser.add_argument("--no-blank-page", action="store_true",
                        help="不在最前面插入封面页（默认插入一张带标题、垂直居中、无页码的封面）")
    parser.add_argument("--code-bg", default=DEFAULT_CODE_BG,
                        help=f"代码块/行内代码底纹色 6 位十六进制（默认: {DEFAULT_CODE_BG}）")
    parser.add_argument("--code-border", default=DEFAULT_CODE_BORDER,
                        help=f"代码块边框色 6 位十六进制（默认: {DEFAULT_CODE_BORDER}）")
    parser.add_argument("--code-font-size", type=float, default=DEFAULT_CODE_FONT_SIZE,
                        help=f"代码字号(pt)（默认: {DEFAULT_CODE_FONT_SIZE}）")
    parser.add_argument("--no-code-box", action="store_true",
                        help="代码块不加底纹/边框，仅保留等宽字体（默认加代码框）")
    parser.add_argument("--quote-bar", default=DEFAULT_QUOTE_BAR,
                        help=f"引用块(> blockquote)左侧竖条色 6 位十六进制（默认: {DEFAULT_QUOTE_BAR}）")
    parser.add_argument("--quote-bg", default=DEFAULT_QUOTE_BG,
                        help=f"引用块底纹色 6 位十六进制（默认: {DEFAULT_QUOTE_BG}）")
    parser.add_argument("--no-quote-box", action="store_true",
                        help="引用块不加左竖条/底纹（默认加，让 > 引用清晰可辨）")
    parser.add_argument("--no-page-numbers", action="store_true",
                        help="不加页码/分节（默认：目录罗马数字、正文阿拉伯数字从 1 起）")
    parser.add_argument("--footer-distance-cm", type=float, default=0.5,
                        help="页码距页面底部的距离(厘米)，默认 0.5")
    parser.add_argument("--no-mermaid", action="store_true",
                        help="不渲染 ```mermaid 流程图（默认：用本地 mmdc+浏览器渲染成图片，缺依赖则降级为代码文本）")
    parser.add_argument("--no-math", action="store_true",
                        help="不渲染 LaTeX 数学公式（默认：$…$/$$…$$/\\(…\\)/\\[…\\] → pandoc 引擎转 Word 原生公式、"
                             "内置引擎转 matplotlib 图片）")
    parser.add_argument("--keep-heading-together", action="store_true",
                        help="保留标题的「与下段同页/段中不分页」属性（Word 会在标题左侧显示小黑点）；"
                             "默认去掉这两项以消除小黑点")
    parser.add_argument("--image-attrs", action="store_true",
                        help="允许图片尺寸属性 {width=…}（封面 logo 自适应用；默认关，避免影响其他文档解析）")
    parser.add_argument("--page-size", default=DEFAULT_PAGE_SIZE,
                        choices=list(PAGE_SIZES_CM.keys()),
                        help=f"纸张尺寸：A4(默认)/Letter（默认 {DEFAULT_PAGE_SIZE}）")
    args = parser.parse_args()

    pandoc = find_pandoc()
    engine = args.engine
    if engine == "auto":
        engine = "pandoc" if pandoc else "python-docx"
    if engine == "pandoc" and not pandoc:
        print("[警告] 未找到 pandoc，自动改用内置 python-docx 引擎。")
        engine = "python-docx"
    if engine == "python-docx":
        try:
            import docx  # noqa: F401
        except ImportError:
            print("[错误] 内置引擎需要 python-docx：pip install python-docx", file=sys.stderr)
            sys.exit(1)

    code_box = not args.no_code_box
    quote_box = not args.no_quote_box
    page_numbers = not args.no_page_numbers
    toc = not args.no_toc
    print(f"[开始] 引擎={engine}  中文={args.font}  英文={args.latin_font}  "
          f"字色=#{args.text_color}  表格配色={args.table_style}  "
          f"代码框={'开' if code_box else '关'}  引用框={'开' if quote_box else '关'}  "
          f"页码={'开' if page_numbers else '关'}")

    # ---------- 分支一：整目录导出（--folder / -fd） ----------
    if args.folder:
        folder = args.folder
        if not os.path.isdir(folder):
            print(f"[错误] 目录不存在: {folder}", file=sys.stderr)
            sys.exit(1)
        files = collect_markdown_files(folder)
        if not files:
            print(f"[错误] 目录下没有 .md/.markdown 文件: {folder}", file=sys.stderr)
            sys.exit(1)
        folder_abs = os.path.abspath(folder)
        print(f"       目录: {folder_abs}  共 {len(files)} 个 md（自然排序，"
              f"{'每个各导一个' if args.split else '合并成一个'}）")
        for k, f in enumerate(files, 1):
            print(f"         {k:>2}. {os.path.basename(f)}")

        if args.split:
            # 批量：每个 md 各导一个 docx，输出到 -o 指定目录（默认=源目录）
            outdir = args.output or folder
            os.makedirs(outdir, exist_ok=True)
            ok = 0
            for f in files:
                out_f = os.path.join(
                    outdir, os.path.splitext(os.path.basename(f))[0] + ".docx")
                print(f"\n[转换] {os.path.basename(f)} → {out_f}")
                try:
                    _convert_one(engine, pandoc, f, out_f, args, page_numbers,
                                 code_box, quote_box, toc, _derive_cover_title(f))
                    if os.path.isfile(out_f):
                        ok += 1
                except Exception as e:
                    print(f"  · 跳过（{type(e).__name__}: {e}）", file=sys.stderr)
            print(f"\n[完成] 批量导出 {ok}/{len(files)} 个 docx → {os.path.abspath(outdir)}")
            if ok < len(files):
                sys.exit(1)
        else:
            # 合并：拼成一个大 md（文件间下一页），复用整条流水线导成一个 docx
            out = args.output or os.path.join(
                os.path.dirname(folder_abs) or ".",
                os.path.basename(folder_abs) + ".docx")
            _ensure_parent_dir(out)
            merged = merge_markdown_files(files, engine, page_break=True)
            merged_src = os.path.join(folder, ".md2docx_merged.md")
            with open(merged_src, "w", encoding="utf-8") as fh:
                fh.write(merged)
            print(f"       合并输出: {out}")
            try:
                _convert_one(engine, pandoc, merged_src, out, args, page_numbers,
                             code_box, quote_box, toc, os.path.basename(folder_abs))
            finally:
                try:
                    os.remove(merged_src)
                except OSError:
                    pass
            if os.path.isfile(out):
                print(f"[完成] 已合并 {len(files)} 个 md → {out}"
                      f"（{human_size(os.path.getsize(out))}）")
            else:
                print("[失败] 未生成输出文件。", file=sys.stderr)
                sys.exit(1)
        return

    # ---------- 分支二：单文件 ----------
    src = args.input
    if not os.path.isfile(src):
        print(f"[错误] 找不到输入文件: {src}", file=sys.stderr)
        sys.exit(1)
    out = args.output or (os.path.splitext(src)[0] + ".docx")
    _ensure_parent_dir(out)
    print(f"       输入: {src}")
    print(f"       输出: {out}")
    _convert_one(engine, pandoc, src, out, args, page_numbers, code_box, quote_box, toc,
                 _derive_cover_title(src))

    if os.path.isfile(out):
        print(f"[完成] 已生成 {out}（{human_size(os.path.getsize(out))}）")
    else:
        print("[失败] 未生成输出文件。", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
