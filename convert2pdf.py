#!/usr/bin/env python3
# convert2pdf.py — Convert a Markdown report into a styled, paginated PDF.
#
# What it does: reads a Markdown file, renders it to HTML (with tables, fenced
# code blocks, footnotes and "sane" lists), wraps that HTML in an embedded CSS
# stylesheet for A4-portrait reports that sizes table columns to content, then writes to PDF via
# WeasyPrint. LaTeX between dollar signs is typeset by a real LaTeX run and
# embedded as vector SVG. Before saving it walks each page's box tree to measure the worst
# horizontal overflow (a probe for tables wider than the page) and prints the
# page count, that overflow, and the output file size.
#
# Usage:
#     python convert2pdf.py <input.md> [output.pdf]
# If no output path is given, the PDF is written next to the input file using
# the same name with a .pdf extension.
#
# Equations need `pdflatex` (with amsmath/amssymb) and `pdftocairo` on PATH; a
# Markdown file with no dollar signs in it calls neither.

import argparse, base64, markdown, os, pathlib, re, shutil, subprocess, tempfile, weasyprint

css = """
@page { size: A4 portrait; margin: 1.4cm 1.2cm;
  @bottom-right { content: "Page " counter(page) " of " counter(pages); font-size: 8pt; color:#888; } }
body { margin:0; font-family: 'DejaVu Sans', Arial, sans-serif; font-size: 8.6pt; color:#1a1a1a; line-height:1.4; }
h1 { font-size: 17pt; color:#0b3d66; border-bottom:2px solid #0b3d66; padding-bottom:5px; }
h2 { font-size: 11.5pt; color:#0b3d66; margin-top:20px; border-left:4px solid #0b3d66; padding-left:8px; page-break-after:avoid; }
table { border-collapse:collapse; width:100%; margin:8px 0 16px; table-layout:auto; }
th { background:#0b3d66; color:#fff; text-align:left; padding:5px 6px; font-size:8.3pt; }
td { border:1px solid #cfd8e0; padding:4px 6px; vertical-align:top; font-size:8.2pt; }
th, td { overflow-wrap:anywhere; word-break:break-word; }
tr:nth-child(even) td { background:#f4f7fa; }
a { color:#12609e; text-decoration:none; }
blockquote { background:#fff8e6; border-left:4px solid #e0a800; margin:10px 0; padding:6px 12px; font-size:8.4pt; }
pre { white-space:pre-wrap; overflow-wrap:anywhere; }
code { background:#eef; padding:1px 3px; border-radius:3px; }
strong { color:#0b3d66; }
img { max-width:100%; height:auto; display:block; margin:12px auto; }
img.mathinline { display:inline; margin:0; max-width:none; }   /* undo the block rule above */
img.mathdisplay { display:block; margin:10px auto; }
figure { margin:12px 0; text-align:center; page-break-inside:avoid; }
figcaption { font-size:8pt; color:#666; text-align:center; margin-top:4px; }
"""


# --------------------------------------------------------------------------- equations

# Everything between single dollars ($...$) is set inline with the surrounding
# text; everything between double dollars ($$...$$) is set as a centred display.
# Both are compiled by one real LaTeX run, so whatever this preamble supports
# will render. Display math is compiled inside $\displaystyle ...$, which is an
# hbox, so an outer environment such as `align` has to be written as its inner
# form (`aligned`, `gathered`, `cases`); snippet numbering would mean nothing.
MATH_PREAMBLE = r"""\documentclass[10pt]{article}
\usepackage[T1]{fontenc}
\usepackage{amsmath,amssymb}
"""
MATH_BASE = 10.0      # the \documentclass font size, in pt: one em of body text
MATH_BORDER = 1.0     # padding around each equation, in pt: keeps ink off the crop edge

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


def render_math(equations, workdir):
    """Compile every equation in one LaTeX run; return (svg, width, height, depth) each.

    Each snippet is measured into a box, the page is resized to exactly that box
    and the box's TeX width/height/depth are echoed to the log. The depth is
    what later puts the equation on the surrounding text's baseline rather than
    on the bottom of its own image.
    """
    for tool in ("pdflatex", "pdftocairo"):
        if shutil.which(tool) is None:
            raise SystemExit(f"{tool} not found: it is needed to typeset the {len(equations)} $...$ equations")

    snippets = "\n".join(r"\snippet{$" + (r"\displaystyle " if kind == "display" else "") + latex + "$}"
                         for kind, latex in equations)
    source = MATH_PREAMBLE + r"""
\newsavebox{\snip}
\newdimen\snipborder \snipborder=%.4fpt
\hoffset=-1in \voffset=-1in
\newcommand{\snippet}[1]{%%
  \sbox{\snip}{#1}%%
  \typeout{MATHBOX|\the\wd\snip|\the\ht\snip|\the\dp\snip}%%
  \pdfpagewidth=\dimexpr\wd\snip+2\snipborder\relax
  \pdfpageheight=\dimexpr\ht\snip+\dp\snip+2\snipborder\relax
  \shipout\vbox{\offinterlineskip\vbox to \snipborder{}\hbox{\kern\snipborder\copy\snip}}%%
}
\begin{document}
""" % MATH_BORDER + snippets + "\n\\end{document}\n"

    (workdir / "equations.tex").write_text(source)
    run = subprocess.run(["pdflatex", "-interaction=nonstopmode", "-halt-on-error",
                          "-no-shell-escape", "equations.tex"],
                         cwd=workdir, capture_output=True, text=True)
    log = (workdir / "equations.log").read_text(errors="replace")
    boxes = re.findall(r"^MATHBOX\|(.+?)pt\|(.+?)pt\|(.+?)pt$", log, re.MULTILINE)
    if run.returncode or len(boxes) != len(equations):
        raise SystemExit(math_error(equations, boxes, log))

    rendered = []
    for page, (width, height, depth) in enumerate(boxes, start=1):
        svg = workdir / f"equation{page}.svg"
        subprocess.run(["pdftocairo", "-svg", "-f", str(page), "-l", str(page),
                        "equations.pdf", svg.name], cwd=workdir, check=True, capture_output=True)
        rendered.append((svg.read_bytes(), float(width), float(height), float(depth)))
    return rendered


def math_error(equations, boxes, log):
    """Name the equation LaTeX stopped on: it is the first that reported no box."""
    index = len(boxes)
    source = equations[index][1] if index < len(equations) else "(unknown)"
    reason = "\n".join(line for line in log.splitlines() if line.startswith("!"))
    return (f"LaTeX failed on equation {index + 1} of {len(equations)}:\n    {source}\n"
            + (reason or log[-800:]))


def math_images(equations, rendered):
    """Build one <img> per equation, sized in em so it tracks the font around it.

    An inline equation also carries a negative vertical-align of its own depth,
    which drops the part below the baseline (the stroke of a fraction, a
    subscript) to where the text baseline already is. A display equation is left
    free to shrink: it gets a width but no height, so `max-width:100%` in the
    stylesheet can scale one that is wider than the text column.
    """
    tags = []
    for (kind, _), (svg, width, height, depth) in zip(equations, rendered):
        uri = "data:image/svg+xml;base64," + base64.b64encode(svg).decode()
        style = f"width:{(width + 2 * MATH_BORDER) / MATH_BASE:.4f}em"
        if kind == "inline":
            style += (f";height:{(height + depth + 2 * MATH_BORDER) / MATH_BASE:.4f}em"
                      f";vertical-align:{-(depth + MATH_BORDER) / MATH_BASE:.4f}em")
        tags.append(f'<img class="math{kind}" src="{uri}" style="{style}">')
    return tags


def markdown_to_pdf(md_path, pdf_path):
    md_text = pathlib.Path(md_path).read_text()
    md_text, equations = extract_math(md_text)

    html_body = markdown.markdown(md_text, extensions=["tables", "sane_lists", "fenced_code", "footnotes"])

    if equations:
        with tempfile.TemporaryDirectory() as workdir:
            tags = math_images(equations, render_math(equations, pathlib.Path(workdir)))
        html_body = MATH_TOKEN_RE.sub(lambda m: tags[int(m.group(1))], html_body)

    html_doc = f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style></head><body>{html_body}</body></html>"

    # base_url lets WeasyPrint resolve relative image paths (e.g. ![](./fig.png))
    # against the Markdown file's own directory, so figures are embedded in the PDF.
    base_url = str(pathlib.Path(md_path).resolve().parent)
    doc = weasyprint.HTML(string=html_doc, base_url=base_url).render()
    worst = 0
    for pg in doc.pages:
        def maxx(box):
            m = box.position_x + box.width
            for c in getattr(box, "children", []) or []: m = max(m, maxx(c))
            return m
        worst = max(worst, maxx(pg._page_box) - pg.width)
    doc.write_pdf(pdf_path)
    print("pages:", len(doc.pages), "\noverflow:", round(worst, 2),
          "\nequations:", len(equations), "\nbytes:", os.path.getsize(pdf_path))


def parse_args():
    """Parse command-line arguments and return (md_path, pdf_path) as strings."""
    parser = argparse.ArgumentParser(description="Convert a Markdown report to a styled PDF.")
    parser.add_argument("input", help="path to the input Markdown file") # required input
    parser.add_argument("output", nargs="?", default=None, help="path to the output PDF") # optional output
    args = parser.parse_args()

    # (optional) validate args.input exists
    if not pathlib.Path(args.input).exists():
        parser.error(f"input file not found: {args.input}")
    
    md_path = args.input
    pdf_path = args.output if args.output is not None else pathlib.Path(args.input).with_suffix(".pdf")   # derive default when None
    return str(md_path), str(pdf_path)

if __name__ == "__main__":
    md_path, pdf_path = parse_args()
    markdown_to_pdf(md_path, pdf_path)
