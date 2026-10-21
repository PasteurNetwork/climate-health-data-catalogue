#!/usr/bin/env python3
# convert2docx.py — Convert a Markdown report into a styled, paginated DOCX.
#
# What it does: reads a Markdown file, renders it to HTML (with tables, fenced
# code blocks, footnotes and "sane" lists) exactly as convert2pdf.py does, then
# walks that HTML and rebuilds it as native Word content — headings, tables,
# lists, block quotes, code blocks, hyperlinks and figures — formatted from a
# style table that mirrors convert2pdf.py's CSS (same navy palette, type sizes,
# zebra rows and "Page X of Y" footer). LaTeX between dollar signs becomes a
# native Word equation rather than a picture of one, so it stays selectable and
# editable in Word's equation editor. Relative image paths are resolved
# against the Markdown file's own directory, so figures are embedded in the
# DOCX. Before saving it reports how many paragraphs, tables and images were
# written, and the output file size.
#
# Usage:
#     python convert2docx.py <input.md> [output.docx]
# If no output path is given, the DOCX is written next to the input file using
# the same name with a .docx extension.
#
# Equations need `pandoc` on PATH; a Markdown file with no dollar signs in it
# never calls it.

import argparse, copy, html.parser, markdown, os, pathlib, re, shutil, subprocess
import tempfile, urllib.parse, zipfile
import docx
from lxml import etree
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor, Twips
from docx.text.paragraph import Paragraph

# The DOCX counterpart of convert2pdf.py's `css` block: same colours, same type
# scale, expressed as Word properties. Border widths are in eighths of a point
# (CSS 1px = 0.75pt = 6 eighths); spacing is in points.
NAVY, INK, GRID, LINK = "0b3d66", "1a1a1a", "cfd8e0", "12609e"
FONT, MONO = "Arial", "DejaVu Sans Mono"   # DOCX has no font stacks: one name each

style = {
    "page":       dict(width=21.0, height=29.7, margin_y=1.4, margin_x=1.2),        # A4 portrait, cm
    "body":       dict(font=FONT, size=8.6, color=INK, line=1.4, space_after=4),
    "h1":         dict(size=17,   color=NAVY, space_before=0,  space_after=5, border=("bottom", 12, NAVY)),
    "h2":         dict(size=11.5, color=NAVY, space_before=15, space_after=4, border=("left", 24, NAVY), indent=6),
    "h3":         dict(size=9.6,  color=NAVY, space_before=10, space_after=3),
    "h4":         dict(size=8.8,  color=NAVY, space_before=8,  space_after=2),
    "th":         dict(size=8.3, fill=NAVY, color="ffffff", bold=True),
    "td":         dict(size=8.2, border=GRID, zebra="f4f7fa"),
    "table":      dict(margin_top=6, margin_bottom=12, cell_pad_y=60, cell_pad_x=90,   # padding in twips
                       char_width=0.52, max_chars=40),                               # column sizing
    "cols":       [12, 26, 12, 8, 42],   # % widths for 5-column tables, as in the CSS; [] sizes every table
    "a":          dict(color=LINK),
    "blockquote": dict(size=8.4, fill="fff8e6", border=("left", 24, "e0a800"), indent=9),
    "code":       dict(font=MONO, fill="eeeeff"),
    "pre":        dict(size=8.0, font=MONO, fill="f4f7fa", border=("left", 24, GRID), indent=6),
    "strong":     dict(color=NAVY),
    "figcaption": dict(size=8, color="666666"),
    "footer":     dict(size=8, color="888888"),
    "list":       dict(indent=312, hanging=255, bullets=("\u2022", "\u25e6", "\u25aa")),  # twips
    "hr":         dict(border=("bottom", 6, GRID)),
}

VOID = {"br", "hr", "img", "meta", "link", "input", "col", "area", "base", "source"}
BLOCKS = {"p", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "table", "thead",
          "tbody", "tfoot", "tr", "th", "td", "blockquote", "pre", "hr", "div", "figure"}


# --------------------------------------------------------------------------- equations

# The equation counterpart of the style table above: convert2pdf.py scans the
# Markdown for dollar signs with exactly this scanner, so the two scripts agree
# on what is an equation and what is prose. Where they part company is the
# target. convert2pdf.py compiles each equation with pdflatex and embeds a
# picture of it, because a PDF has no equation format; Word does have one
# (OMML), and a real equation beats a picture of one -- it stays selectable,
# searchable and editable in Word's equation editor, and it follows the body
# font size instead of pixelating when the page is zoomed. So here the
# equations go to pandoc, which writes OMML directly.
MATH_TOKEN = "qqmatheq{}qq"                      # inert in Markdown: plain letters and digits
MATH_TOKEN_RE = re.compile(r"qqmatheq(\d+)qq")

# One pass over the Markdown source. The first three alternatives match places
# where a dollar sign is not math -- a fenced block, an inline code span, an
# escaped \$ -- and matching them first is what leaves ```echo $PATH``` alone.
MATH_SCANNER = re.compile(r"""
      (?P<fence>^(?P<f>`{3,}|~{3,})[^\n]*\n.*?(?:^(?P=f)[^\n]*$|\Z))
    | (?P<code>(?P<t>`+)[^\n]*?(?P=t))
    | (?P<esc>\\\$)
    | (?P<display>\$\$(?P<deq>.+?)\$\$)
    | (?P<inline>\$(?P<ieq>[^\s$](?:[^$\n]|\n(?!\s*\n))*?)(?<!\s)\$(?!\d))
""", re.MULTILINE | re.DOTALL | re.VERBOSE)


def extract_math(md_text):
    """Swap every equation out of the Markdown for an inert token.

    Returns the rewritten Markdown plus a list of (kind, latex) pairs, where
    `kind` is "inline" or "display". A lone dollar never opens an equation,
    because an opening $ must be followed by a non-space and the closing $ must
    be preceded by a non-space and not followed by a digit: that is what keeps
    prose like "costs $5 and $6" out of the equation list. Math inside indented
    (four-space) code blocks is not protected -- use a fenced block for that.
    """
    equations = []

    def take(match):
        if match.group("esc") is not None:
            return "$"                           # Markdown does not unescape \$ itself
        for kind in ("display", "inline"):
            if match.group(kind) is None:
                continue
            latex = match.group("deq" if kind == "display" else "ieq").strip()
            equations.append((kind, re.sub(r"\n\s*\n", "\n", latex)))   # a blank line ends a TeX paragraph
            return MATH_TOKEN.format(len(equations) - 1)
        return match.group(0)                    # a fenced block or an inline code span

    return MATH_SCANNER.sub(take, md_text), equations


def render_math_omml(equations):
    """Convert every equation to Word's own OMML markup in one pandoc run.

    Pandoc writes inline math as a bare <m:oMath> and display math as an
    <m:oMathPara> carrying its own centring, which are exactly the two shapes
    Word expects inside a paragraph, so each is spliced in as it stands. Feeding
    pandoc one equation per paragraph is what keeps its output in step with the
    input: an equation pandoc cannot read comes back as plain text in its own
    paragraph, and that is how the loop below catches and names it.
    """
    pandoc = shutil.which("pandoc")
    if pandoc is None:
        raise SystemExit(f"pandoc not found: it is needed to turn the {len(equations)} "
                         "$...$ equations into Word equations")
    source = "\n\n".join(f"$${latex}$$" if kind == "display" else f"${latex}$"
                         for kind, latex in equations)
    with tempfile.TemporaryDirectory() as workdir:
        built = pathlib.Path(workdir) / "equations.docx"
        subprocess.run([pandoc, "--from=markdown", "--to=docx", "--output", str(built)],
                       input=source, text=True, check=True, capture_output=True)
        with zipfile.ZipFile(built) as archive:
            document = etree.fromstring(archive.read("word/document.xml"))

    paragraphs = list(document.iter(qn("w:p")))
    if len(paragraphs) != len(equations):
        raise SystemExit(f"pandoc returned {len(paragraphs)} paragraphs for {len(equations)} equations")
    math = []
    for index, ((kind, latex), par) in enumerate(zip(equations, paragraphs)):
        found = [child for child in par if child.tag in (qn("m:oMath"), qn("m:oMathPara"))]
        if not found:
            raise SystemExit(f"pandoc could not read equation {index + 1} of "
                             f"{len(equations)} as math:\n    {latex}")
        math.append(found[0])
    return math


# --------------------------------------------------------------------------- HTML tree

class Node:
    """One HTML element; `children` holds Nodes and plain strings."""
    __slots__ = ("tag", "attrs", "children")

    def __init__(self, tag, attrs=None):
        self.tag, self.attrs, self.children = tag, dict(attrs or {}), []


class TreeBuilder(html.parser.HTMLParser):
    """Parse python-markdown's HTML output into a small tag/text tree."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("root")
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs)
        self.stack[-1].children.append(node)
        if tag not in VOID:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(Node(tag, attrs))

    def handle_endtag(self, tag):
        for i in range(len(self.stack) - 1, 0, -1):   # tolerate unclosed inline tags
            if self.stack[i].tag == tag:
                del self.stack[i:]
                return

    def handle_data(self, data):
        self.stack[-1].children.append(data)


def parse_html(html_body):
    parser = TreeBuilder()
    parser.feed(html_body)
    parser.close()
    return parser.root


def is_block(node):
    return isinstance(node, Node) and node.tag in BLOCKS


# --------------------------------------------------------------------------- low-level XML helpers

# Child order inside <w:pPr> and <w:tblPr>: Word rejects a document that ignores it.
PPR_ORDER = ("w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:numPr",
             "w:pBdr", "w:shd", "w:tabs", "w:spacing", "w:ind", "w:contextualSpacing",
             "w:jc", "w:outlineLvl", "w:rPr", "w:sectPr")
TBLPR_ORDER = ("w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
               "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd",
               "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")


def insert_ordered(parent, el, order):
    """Insert `el` at its schema position, replacing any element of the same tag."""
    tag = "w:" + el.tag.split("}")[1]
    existing = parent.find(qn(tag))
    if existing is not None:
        parent.remove(existing)
    parent.insert_element_before(el, *order[order.index(tag) + 1:])
    return el


def pPr_insert(pPr, el):
    """Insert `el` into a <w:pPr> at its schema position."""
    return insert_ordered(pPr, el, PPR_ORDER)


def set_border(pPr, edge, size, color):
    """Add one paragraph border edge ('bottom'/'left'/...), size in eighths of a point."""
    bdr = pPr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = OxmlElement("w:pBdr")
        pPr_insert(pPr, bdr)
    el = OxmlElement("w:" + edge)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(size))
    el.set(qn("w:space"), "3")
    el.set(qn("w:color"), color)
    bdr.append(el)


def set_shading(pPr_or_tcPr, fill, insert=False):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr_insert(pPr_or_tcPr, shd) if insert else pPr_or_tcPr.append(shd)


def set_run_font(run, name):
    """Pin a run's typeface for every script slot, not just the Latin one."""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    fonts = rPr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rPr.insert(0, fonts)
    for slot in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(slot), name)


def set_mark_size(par, size_pt):
    """Shrink an empty spacer paragraph to (almost) nothing but its spacing."""
    pPr = par._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr_insert(pPr, rPr)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(int(size_pt * 2)))
        rPr.append(el)


def add_field(par, instr):
    """Append a Word field (PAGE, NUMPAGES, ...) as three runs."""
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    text = OxmlElement("w:instrText"); text.set(qn("xml:space"), "preserve"); text.text = f" {instr} "
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    for child in (begin, text, end):
        par.add_run()._r.append(child)


def numbering_level(level, ordered):
    """One <w:lvl> of a list definition: marker, indent and hanging indent."""
    spec = style["list"]
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(level))
    start = OxmlElement("w:start"); start.set(qn("w:val"), "1")
    fmt = OxmlElement("w:numFmt"); fmt.set(qn("w:val"), "decimal" if ordered else "bullet")
    text = OxmlElement("w:lvlText")
    text.set(qn("w:val"), f"%{level + 1}." if ordered else spec["bullets"][level % len(spec["bullets"])])
    jc = OxmlElement("w:lvlJc"); jc.set(qn("w:val"), "left")
    pPr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(spec["indent"] * (level + 1)))
    ind.set(qn("w:hanging"), str(spec["hanging"]))
    pPr.append(ind)
    rPr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for slot in ("w:ascii", "w:hAnsi", "w:cs"):
        fonts.set(qn(slot), style["body"]["font"])
    rPr.append(fonts)
    for child in (start, fmt, text, jc, pPr, rPr):
        lvl.append(child)
    return lvl


def add_numbering(doc, ordered):
    """Register a fresh 9-level list definition and return its numId.

    One definition per list is what makes every list restart at 1: Word shares
    counters between lists that share a definition.
    """
    numbering = doc.part.numbering_part.element
    abstract_id = 1 + max((int(a.get(qn("w:abstractNumId")))
                           for a in numbering.findall(qn("w:abstractNum"))), default=-1)
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType"); multi.set(qn("w:val"), "hybridMultilevel")
    abstract.append(multi)
    for level in range(9):
        abstract.append(numbering_level(level, ordered))
    first_num = numbering.find(qn("w:num"))            # all abstractNum must precede all num
    first_num.addprevious(abstract) if first_num is not None else numbering.append(abstract)

    num_id = 1 + max((int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))), default=0)
    num = OxmlElement("w:num"); num.set(qn("w:numId"), str(num_id))
    ref = OxmlElement("w:abstractNumId"); ref.set(qn("w:val"), str(abstract_id))
    num.append(ref)
    numbering.append(num)
    return num_id


def set_numbering(par, num_id, level):
    """Attach a paragraph to list definition `num_id` at indent level `level`."""
    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl"); ilvl.set(qn("w:val"), str(level))
    num = OxmlElement("w:numId"); num.set(qn("w:val"), str(num_id))
    numPr.append(ilvl)
    numPr.append(num)
    pPr_insert(par._p.get_or_add_pPr(), numPr)


# --------------------------------------------------------------------------- document setup

def configure_styles(doc):
    body = style["body"]
    normal = doc.styles["Normal"]
    normal.font.size = Pt(body["size"])
    normal.font.color.rgb = RGBColor.from_string(body["color"].upper())
    normal.paragraph_format.line_spacing = body["line"]
    normal.paragraph_format.space_after = Pt(body["space_after"])
    set_run_font(_StyleRun(normal), body["font"])

    for level in (1, 2, 3, 4):
        spec = style[f"h{level}"]
        heading = doc.styles[f"Heading {level}"]
        heading.base_style = normal
        heading.font.size = Pt(spec["size"])
        heading.font.bold = True
        heading.font.italic = False
        heading.font.color.rgb = RGBColor.from_string(spec["color"].upper())
        set_run_font(_StyleRun(heading), body["font"])
        fmt = heading.paragraph_format
        fmt.space_before = Pt(spec["space_before"])
        fmt.space_after = Pt(spec["space_after"])
        fmt.keep_with_next = True
        fmt.line_spacing = 1.15
        pPr = heading.element.get_or_add_pPr()
        if "border" in spec:
            set_border(pPr, *spec["border"])
        if "indent" in spec:
            fmt.left_indent = Pt(spec["indent"])


class _StyleRun:
    """Adapter so set_run_font() can also target a style's run properties."""

    def __init__(self, style_obj):
        self.font = style_obj.font
        self._r = style_obj.element


def add_page_footer(section):
    spec = style["footer"]
    par = section.footer.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    par.add_run("Page ")
    add_field(par, "PAGE")
    par.add_run(" of ")
    add_field(par, "NUMPAGES")
    for run in par.runs:
        run.font.size = Pt(spec["size"])
        run.font.color.rgb = RGBColor.from_string(spec["color"].upper())
        set_run_font(run, style["body"]["font"])


def new_document():
    doc = docx.Document()
    page = style["page"]
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(page["width"]), Cm(page["height"])
    section.top_margin = section.bottom_margin = Cm(page["margin_y"])
    section.left_margin = section.right_margin = Cm(page["margin_x"])
    configure_styles(doc)
    add_page_footer(section)
    return doc


# --------------------------------------------------------------------------- renderer

class Renderer:
    """Walk the HTML tree and emit Word content into `doc`."""

    def __init__(self, doc, base_dir, math):
        self.doc = doc
        self.base_dir = base_dir
        self.math = math
        self.images = 0
        self.equations = 0
        self.missing = []
        section = doc.sections[0]
        self.content_width = section.page_width - section.left_margin - section.right_margin

    # ---- blocks ----------------------------------------------------------
    def render(self, nodes):
        for node in nodes:
            if isinstance(node, str):
                if node.strip():
                    self.paragraph([node])
            else:
                self.block(node)

    def block(self, node):
        tag = node.tag
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self.heading(node)
        elif tag == "p":
            self.paragraph(node.children)
        elif tag in ("ul", "ol"):
            self.list(node, ordered=(tag == "ol"))
        elif tag == "table":
            self.table(node)
        elif tag == "blockquote":
            self.blockquote(node)
        elif tag == "pre":
            self.code_block(node)
        elif tag == "hr":
            self.rule()
        elif tag == "img":
            self.figure(node)
        elif tag in ("div", "section", "article", "figure", "details", "main", "root"):
            self.render(node.children)
        else:                                        # unknown wrapper: treat as a paragraph
            self.paragraph(node.children)

    def heading(self, node):
        level = min(int(node.tag[1]), 4)
        par = self.doc.add_paragraph(style=f"Heading {level}")
        self.inlines(par, node.children, {"size": style[f"h{level}"]["size"],
                                          "color": style[f"h{level}"]["color"], "bold": True})
        self.trim(par)

    def paragraph(self, children, fmt=None):
        images = [c for c in children if isinstance(c, Node) and c.tag == "img"]
        rest = [c for c in children if not (isinstance(c, Node) and c.tag == "img")]
        if images and not any(isinstance(c, str) and c.strip() for c in rest):
            for img in images:                       # a standalone image becomes a figure
                self.figure(img)
            return
        par = self.doc.add_paragraph()
        self.inlines(par, children, fmt or {})
        return self.trim(par)

    @staticmethod
    def trim(par):
        """Drop the trailing whitespace HTML leaves between a close tag and its parent."""
        while par.runs:
            run = par.runs[-1]
            if run._r is not par._p[-1]:         # an equation follows it, so that space is real
                break
            run.text = run.text.rstrip()
            if run.text:
                break
            run._r.getparent().remove(run._r)
        return par

    def list(self, node, ordered, level=0):
        num_id = add_numbering(self.doc, ordered)      # a definition per list, so it restarts at 1
        indent = Twips(style["list"]["indent"] * (min(level, 8) + 1))
        for item in node.children:
            if not (isinstance(item, Node) and item.tag == "li"):
                continue
            inline, blocks = self.split_item(item)
            par = self.doc.add_paragraph()
            par.paragraph_format.space_after = Pt(2)
            set_numbering(par, num_id, min(level, 8))
            self.inlines(par, inline, {})
            self.trim(par)
            for child in blocks:                        # nested lists and trailing blocks
                if child.tag in ("ul", "ol"):
                    self.list(child, ordered=(child.tag == "ol"), level=level + 1)
                    continue
                start = len(self.doc.paragraphs)
                self.paragraph(child.children) if child.tag == "p" else self.block(child)
                for extra in self.doc.paragraphs[start:]:
                    extra.paragraph_format.left_indent = indent

    @staticmethod
    def split_item(item):
        """Split an <li> into the content shown next to its marker and any trailing blocks."""
        inline, blocks, merged = [], [], False
        nested = any(is_block(child) for child in item.children)
        for child in item.children:
            if is_block(child):
                if child.tag == "p" and not merged and not blocks:
                    inline.extend(child.children)    # loose list: first <p> rides the marker
                    merged = True
                else:
                    blocks.append(child)
            elif blocks or (nested and not child.strip()):
                continue                             # layout whitespace around block tags
            else:
                inline.append(child)
        return inline, blocks

    def blockquote(self, node):
        spec = style["blockquote"]
        first = len(self.doc.paragraphs)
        self.render(node.children)
        for par in self.doc.paragraphs[first:]:
            par.paragraph_format.left_indent = Pt(spec["indent"])
            par.paragraph_format.right_indent = Pt(spec["indent"])
            par.paragraph_format.space_before = Pt(3)
            par.paragraph_format.space_after = Pt(3)
            pPr = par._p.get_or_add_pPr()
            set_border(pPr, *spec["border"])
            set_shading(pPr, spec["fill"], insert=True)
            for run in par.runs:
                run.font.size = Pt(spec["size"])

    def code_block(self, node):
        spec = style["pre"]
        text = self.plain_text(node).rstrip("\n")
        par = self.doc.add_paragraph()
        par.paragraph_format.left_indent = Pt(spec["indent"])
        par.paragraph_format.space_before = Pt(4)
        par.paragraph_format.space_after = Pt(8)
        par.paragraph_format.line_spacing = 1.0
        par.paragraph_format.keep_together = True
        pPr = par._p.get_or_add_pPr()
        set_border(pPr, *spec["border"])
        set_shading(pPr, spec["fill"], insert=True)
        for i, line in enumerate(text.split("\n")):
            run = par.add_run()
            if i:
                run.add_break()
            run.add_text(line)
            run.font.size = Pt(spec["size"])
            set_run_font(run, spec["font"])

    def rule(self):
        par = self.doc.add_paragraph()
        par.paragraph_format.space_before = Pt(6)
        par.paragraph_format.space_after = Pt(6)
        set_border(par._p.get_or_add_pPr(), *style["hr"]["border"])

    # ---- tables ----------------------------------------------------------
    def table(self, node):
        rows = self.table_rows(node)
        if not rows:
            return
        ncols = max(len(cells) for _, cells in rows)
        widths = self.col_widths(rows, ncols)
        table = self.doc.add_table(rows=len(rows), cols=ncols)
        self.set_table_layout(table, widths)

        body_index = 0
        for r, (kind, cells) in enumerate(rows):
            header = kind == "th"
            spec = style["th"] if header else style["td"]
            if header:
                repeat = OxmlElement("w:tblHeader")  # reprint the header after a page break
                table.rows[r]._tr.get_or_add_trPr().append(repeat)
            else:
                body_index += 1
            zebra = (not header) and body_index % 2 == 0        # CSS: tr:nth-child(even)
            for c in range(ncols):
                cell = table.cell(r, c)
                cell.width = widths[c]
                tcPr = cell._tc.get_or_add_tcPr()
                self.set_cell_borders(tcPr)
                if header:
                    set_shading(tcPr, style["th"]["fill"])
                elif zebra:
                    set_shading(tcPr, style["td"]["zebra"])
                if c < len(cells):
                    self.fill_cell(cell, cells[c], spec, header)
        self.space_table(table)

    @staticmethod
    def table_rows(node):
        """Flatten <thead>/<tbody>/<tr> into [(kind, [cell nodes]), ...]."""
        rows = []
        for child in node.children:
            groups = child.children if (isinstance(child, Node) and child.tag in
                                        ("thead", "tbody", "tfoot")) else [child]
            for tr in groups:
                if not (isinstance(tr, Node) and tr.tag == "tr"):
                    continue
                cells = [c for c in tr.children if isinstance(c, Node) and c.tag in ("th", "td")]
                if cells:
                    rows.append(("th" if cells[0].tag == "th" else "td", cells))
        return rows

    def col_widths(self, rows, ncols):
        """The CSS percentages for a 5-column table; otherwise share by content length.

        `table-layout:fixed` needs explicit widths, and the stylesheet only
        describes the five-column report table, so other shapes are sized from
        how much text each column actually holds.
        """
        if style["cols"] and ncols == len(style["cols"]):
            return [Emu(int(self.content_width * pct / 100)) for pct in style["cols"]]
        spec = style["table"]
        char = Pt(spec["char_width"] * style["td"]["size"])   # rough average advance width
        padding = Twips(2 * spec["cell_pad_x"])
        floors, weights = [], []
        for c in range(ncols):
            texts = [self.plain_text(cells[c]).strip() for _, cells in rows if c < len(cells)]
            longest_word = max((len(word) for text in texts for word in text.split()), default=1)
            floors.append(longest_word * char + padding)      # never narrower than one word
            weights.append(min(max((len(text) for text in texts), default=1), spec["max_chars"]))
        room = self.content_width - sum(floors)
        if room < 0:                                          # too many columns: share the floors
            return [Emu(int(self.content_width * floor / sum(floors))) for floor in floors]
        return [Emu(int(floor + room * weight / sum(weights)))
                for floor, weight in zip(floors, weights)]

    def set_table_layout(self, table, widths):
        tblPr = table._tbl.tblPr
        width = OxmlElement("w:tblW")                # CSS `table { width:100% }`
        width.set(qn("w:w"), "5000")
        width.set(qn("w:type"), "pct")
        insert_ordered(tblPr, width, TBLPR_ORDER)
        layout = OxmlElement("w:tblLayout")          # CSS `table-layout:fixed`
        layout.set(qn("w:type"), "fixed")
        insert_ordered(tblPr, layout, TBLPR_ORDER)
        margins = OxmlElement("w:tblCellMar")
        for edge, value in (("top", style["table"]["cell_pad_y"]), ("bottom", style["table"]["cell_pad_y"]),
                            ("left", style["table"]["cell_pad_x"]), ("right", style["table"]["cell_pad_x"])):
            el = OxmlElement("w:" + edge)
            el.set(qn("w:w"), str(value))
            el.set(qn("w:type"), "dxa")
            margins.append(el)
        insert_ordered(tblPr, margins, TBLPR_ORDER)
        for gridCol, column_width in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths):
            gridCol.set(qn("w:w"), str(int(column_width.twips)))
        for column, column_width in zip(table.columns, widths):
            column.width = column_width

    @staticmethod
    def set_cell_borders(tcPr):
        borders = OxmlElement("w:tcBorders")
        for edge in ("top", "left", "bottom", "right"):
            el = OxmlElement("w:" + edge)
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "6")                   # CSS 1px
            el.set(qn("w:color"), style["td"]["border"])
            borders.append(el)
        tcPr.append(borders)

    def fill_cell(self, cell, node, spec, header):
        par = cell.paragraphs[0]
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(0)
        par.paragraph_format.line_spacing = 1.15
        align = self.cell_alignment(node)
        if align is not None:
            par.alignment = align
        fmt = {"size": spec["size"]}
        if header:
            fmt.update(bold=spec["bold"], color=spec["color"], strong_color=spec["color"])
        self.inlines(par, node.children, fmt)
        self.trim(par)

    @staticmethod
    def cell_alignment(node):
        css_align = node.attrs.get("style", "")
        if "right" in css_align:
            return WD_ALIGN_PARAGRAPH.RIGHT
        if "center" in css_align:
            return WD_ALIGN_PARAGRAPH.CENTER
        return None

    def space_table(self, table):
        """CSS `table { margin: 8px 0 16px }` — Word expresses it as spacer paragraphs.

        A paragraph after the table is also what keeps two adjacent tables from
        being merged into one by Word, so it is always emitted.
        """
        previous = table._tbl.getprevious()
        if previous is not None and previous.tag == qn("w:p"):
            Paragraph(previous, self.doc).paragraph_format.space_after = Pt(style["table"]["margin_top"])
        else:                                        # nothing above (or another table): spacer
            before = self.doc.add_paragraph()
            before.paragraph_format.space_after = Pt(style["table"]["margin_top"])
            before.paragraph_format.line_spacing = 1.0
            set_mark_size(before, 2)
            table._tbl.addprevious(before._p)
        after = self.doc.add_paragraph()
        after.paragraph_format.space_after = Pt(style["table"]["margin_bottom"])
        after.paragraph_format.line_spacing = 1.0
        set_mark_size(after, 2)

    # ---- figures ---------------------------------------------------------
    def add_picture(self, par, node):
        """Embed an <img> into `par`, or record it as missing and return None."""
        src = urllib.parse.unquote(node.attrs.get("src", ""))
        path = pathlib.Path(src)
        if not path.is_absolute():
            path = self.base_dir / path
        if re.match(r"^[a-z]+://", src) or not path.is_file():
            self.missing.append(src)                 # remote and unresolved sources alike
            return None
        try:
            picture = par.add_run().add_picture(str(path))
        except Exception as exc:                     # unreadable or unsupported format
            self.missing.append(f"{src} ({exc.__class__.__name__})")
            return None
        if picture.width > self.content_width:       # CSS `img { max-width:100% }`
            picture.height = int(picture.height * self.content_width / picture.width)
            picture.width = self.content_width
        self.images += 1
        return picture

    def figure(self, node):
        src = urllib.parse.unquote(node.attrs.get("src", ""))
        caption = node.attrs.get("title") or node.attrs.get("alt") or ""
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(9)
        par.paragraph_format.space_after = Pt(3)
        par.paragraph_format.keep_with_next = True
        picture = self.add_picture(par, node)
        if picture is None:
            self.add_text(par, f"[missing image: {src}]", {"italic": True, "color": "888888"})
        if caption and picture is not None:
            spec = style["figcaption"]
            cap = self.doc.add_paragraph()
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.paragraph_format.space_after = Pt(9)
            self.add_text(cap, caption, {"size": spec["size"], "color": spec["color"]})

    # ---- inline ----------------------------------------------------------
    def inlines(self, par, nodes, fmt):
        for node in nodes:
            if isinstance(node, str):
                self.add_text(par, node, fmt)
            elif node.tag in ("strong", "b"):
                color = fmt.get("strong_color", style["strong"]["color"])
                self.inlines(par, node.children, {**fmt, "bold": True, "color": color})
            elif node.tag in ("em", "i", "cite"):
                self.inlines(par, node.children, {**fmt, "italic": True})
            elif node.tag in ("code", "kbd", "samp", "tt"):
                self.inlines(par, node.children, {**fmt, "code": True})
            elif node.tag in ("del", "s", "strike"):
                self.inlines(par, node.children, {**fmt, "strike": True})
            elif node.tag == "sup":
                self.inlines(par, node.children, {**fmt, "superscript": True})
            elif node.tag == "sub":
                self.inlines(par, node.children, {**fmt, "subscript": True})
            elif node.tag == "a":
                self.hyperlink(par, node, fmt)
            elif node.tag == "br":
                par.add_run().add_break()
            elif node.tag == "img":                  # an image sharing a line with text
                if self.add_picture(par, node) is None:
                    self.add_text(par, node.attrs.get("alt", ""), {**fmt, "italic": True})
            else:
                self.inlines(par, node.children, fmt)

    def hyperlink(self, par, node, fmt):
        url = node.attrs.get("href", "")
        link_fmt = {**fmt, "color": style["a"]["color"]}
        if not url or url.startswith("#"):           # anchors have no target in a DOCX
            self.inlines(par, node.children, link_fmt)
            return
        existing = list(par._p)
        self.inlines(par, node.children or [url], link_fmt)
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), par.part.relate_to(url, RT.HYPERLINK, is_external=True))
        for element in [e for e in par._p if e not in existing]:
            hyperlink.append(element)                # append() moves the runs
        par._p.append(hyperlink)

    def add_text(self, par, text, fmt):
        parts = MATH_TOKEN_RE.split(text)
        if len(parts) > 1:                       # text, index, text, index, ..., text
            for position, part in enumerate(parts):
                if position % 2:
                    self.add_math(par, int(part))
                elif part:
                    self.add_text(par, part, fmt)
            return
        if not fmt.get("raw"):
            text = re.sub(r"\s+", " ", text)         # HTML collapses whitespace
            if self.at_start(par) and not text.strip():
                return
            if self.at_start(par):
                text = text.lstrip()
        if not text:
            return
        run = par.add_run(text)
        run.bold = fmt.get("bold", False)
        run.italic = fmt.get("italic", False)
        run.underline = False                        # CSS `a { text-decoration:none }`
        run.font.strike = fmt.get("strike", False)
        if fmt.get("superscript"):                   # both share w:vertAlign: set only one
            run.font.superscript = True
        elif fmt.get("subscript"):
            run.font.subscript = True
        run.font.size = Pt(fmt.get("size", style["body"]["size"]))
        run.font.color.rgb = RGBColor.from_string(fmt.get("color", style["body"]["color"]).upper())
        if fmt.get("code"):
            set_run_font(run, style["code"]["font"])
            set_shading(run._r.get_or_add_rPr(), style["code"]["fill"])
        else:
            set_run_font(run, style["body"]["font"])
        return run

    @staticmethod
    def at_start(par):
        """True while the paragraph is still empty, counting an equation as content.

        Leading whitespace is dropped only at the very start of a paragraph, and
        an equation is not a run, so without this an equation opening a
        paragraph would swallow the space that follows it.
        """
        return (not par.runs and par._p.find(qn("m:oMath")) is None
                and par._p.find(qn("m:oMathPara")) is None)

    def add_math(self, par, index):
        """Drop pandoc's OMML for one equation into the paragraph, in reading order."""
        par._p.append(copy.deepcopy(self.math[index]))
        self.equations += 1

    @staticmethod
    def plain_text(node):
        out = []
        stack = [node]
        while stack:
            current = stack.pop(0)
            if isinstance(current, str):
                out.append(current)
            else:
                stack = list(current.children) + stack
        return "".join(out)


def markdown_to_docx(md_path, docx_path):
    md_text = pathlib.Path(md_path).read_text()
    md_text, equations = extract_math(md_text)

    html_body = markdown.markdown(md_text, extensions=["tables", "sane_lists", "fenced_code", "footnotes"])

    doc = new_document()
    # base_dir mirrors WeasyPrint's base_url: relative image paths (e.g. ![](./fig.png))
    # resolve against the Markdown file's own directory, so figures are embedded.
    base_dir = pathlib.Path(md_path).resolve().parent
    renderer = Renderer(doc, base_dir, render_math_omml(equations) if equations else [])
    renderer.render(parse_html(html_body).children)
    doc.save(docx_path)

    if renderer.missing:
        print("missing images:", ", ".join(sorted(set(renderer.missing))))
    print("paragraphs:", len(doc.paragraphs), "\ntables:", len(doc.tables),
          "\nimages:", renderer.images, "\nequations:", renderer.equations,
          "\nbytes:", os.path.getsize(docx_path))


def parse_args():
    """Parse command-line arguments and return (md_path, docx_path) as strings."""
    parser = argparse.ArgumentParser(description="Convert a Markdown report to a styled DOCX.")
    parser.add_argument("input", help="path to the input Markdown file")            # required input
    parser.add_argument("output", nargs="?", default=None, help="path to the output DOCX")  # optional output

    args = parser.parse_args()

    if not pathlib.Path(args.input).exists():
        parser.error(f"input file not found: {args.input}")

    md_path = args.input
    docx_path = args.output if args.output is not None else pathlib.Path(args.input).with_suffix(".docx")
    return str(md_path), str(docx_path)


if __name__ == "__main__":
    md_path, docx_path = parse_args()
    markdown_to_docx(md_path, docx_path)
