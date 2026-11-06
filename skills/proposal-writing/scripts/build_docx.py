#!/usr/bin/env python3
"""Build an editable Chinese proposal DOCX from a complete structured JSON draft."""
import argparse
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from proposal_contract import contract_errors
from word_limits import length_report
from formal_style import scan_docx


def set_font(target, east_asia, size, bold=False):
    font = target.font
    font.name = "Times New Roman"
    font.size = Pt(size)
    font.bold = bold
    font.italic = False
    font.color.rgb = RGBColor(0, 0, 0)
    rpr = target._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for key in ("ascii", "hAnsi", "cs"):
        fonts.set(qn("w:" + key), "Times New Roman")
    fonts.set(qn("w:eastAsia"), east_asia)
    for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn("w:" + key), None)


def element(tag, **attrs):
    node = OxmlElement("w:" + tag)
    for key, value in attrs.items():
        node.set(qn("w:" + key), str(value))
    return node


def bookmark(paragraph, ident):
    paragraph._p.insert(0, element("bookmarkStart", id=ident, name=f"proposal_{ident}"))
    paragraph._p.append(element("bookmarkEnd", id=ident))


def add_link(paragraph, text, anchor, font):
    link = element("hyperlink", anchor=anchor, history="1")
    run = element("r")
    props = element("rPr")
    props.append(element("rFonts", ascii="Times New Roman", hAnsi="Times New Roman", eastAsia=font))
    props.append(element("color", val="000000"))
    props.append(element("sz", val="22"))
    run.append(props)
    node = element("t")
    node.text = text
    run.append(node)
    link.append(run)
    paragraph._p.append(link)


def validate(data):
    if not isinstance(data, dict) or not isinstance(data.get("title"), str) or not data["title"].strip():
        raise ValueError("A nonempty project title is required")
    for key in ("subtitle", "status", "font_east_asia", "font_heading"):
        if key in data and (not isinstance(data[key], str) or not data[key].strip()):
            raise ValueError(f"{key} must be a nonempty string when supplied")
    for key in ("create_cover", "create_toc", "page_numbers"):
        if key in data and not isinstance(data[key], bool):
            raise ValueError(f"{key} must be a boolean when supplied")
    sections = data.get("sections")
    if not isinstance(sections, list) or not sections:
        raise ValueError("sections must contain the complete proposal draft")
    appendices = data.get("appendices", [])
    if not isinstance(appendices, list):
        raise ValueError("appendices must be a list")
    for section in sections + appendices:
        if not isinstance(section, dict) or not isinstance(section.get("heading"), str) or not section["heading"].strip():
            raise ValueError("Each section needs a heading")
        if not isinstance(section.get("blocks"), list) or not section["blocks"]:
            raise ValueError("Each section needs substantive content blocks")
        for block in section["blocks"]:
            if not isinstance(block, dict):
                raise ValueError("Each block must be an object")
            kind = block.get("type")
            if kind in ("paragraph", "heading", "bullet"):
                if not isinstance(block.get("text"), str) or not block["text"].strip():
                    raise ValueError("Text blocks must contain nonempty text")
                if kind == "heading" and (isinstance(block.get("level", 2), bool) or block.get("level", 2) not in range(2, 7)):
                    raise ValueError("Nested headings must use level 2 to 6")
            elif kind == "table":
                headers = block.get("headers")
                if not isinstance(headers, list) or not 1 <= len(headers) <= 6:
                    raise ValueError("Tables need 1 to 6 columns; split wide matrices using shared IDs")
                if not all(isinstance(x, str) and x.strip() for x in headers):
                    raise ValueError("Table headers must be nonempty strings")
                rows = block.get("rows")
                if not isinstance(rows, list) or not rows:
                    raise ValueError("Tables need substantive rows")
                if any(not isinstance(row, list) or len(row) != len(headers) or not all(isinstance(x, str) for x in row) for row in rows):
                    raise ValueError("Every table row must match the header count and use strings")
                widths = block.get("widths", [1] * len(headers))
                if not isinstance(widths, list) or len(widths) != len(headers) or any(isinstance(x, bool) or not isinstance(x, (int, float)) or not 0 < x < 100000 for x in widths):
                    raise ValueError("widths must provide a positive finite weight for each column")
            else:
                raise ValueError(f"Unsupported content block: {kind}")
    errors = contract_errors(data)
    if errors:
        raise ValueError("\n".join(errors))
    report = length_report(data)
    if report['errors']:
        raise ValueError('正文篇幅检查未通过：\n' + '\n'.join(report['errors']))


def add_table(doc, block, font):
    caption = block.get("caption")
    if caption:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.keep_with_next = True
        set_font(p.add_run(caption), font, 11)
    headers = block["headers"]
    weights = block.get("widths", [1] * len(headers))
    section = doc.sections[-1]
    usable = section.page_width - section.left_margin - section.right_margin
    widths = [int(usable * w / sum(weights)) for w in weights]
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for col, width in zip(table.columns, widths):
        col.width = width
    props = table._tbl.tblPr
    borders = element("tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(element(side, val="single", sz="4", color="D9D9D9"))
    props.append(borders)
    margins = element("tblCellMar")
    for side in ("top", "bottom", "left", "right"):
        margins.append(element(side, w="100", type="dxa"))
    props.append(margins)
    table.rows[0]._tr.get_or_add_trPr().append(element("tblHeader", val="true"))
    for row_index, values in enumerate([headers] + block["rows"]):
        row = table.rows[0] if row_index == 0 else table.add_row()
        if sum(len(v) for v in values) < 1400:
            row._tr.get_or_add_trPr().append(element("cantSplit"))
        for cell, text, width in zip(row.cells, values, widths):
            cell.width = width
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if row_index == 0:
                cell._tc.get_or_add_tcPr().append(element("shd", fill="E8EEF2", val="clear"))
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.first_line_indent = Pt(0)
            paragraph.paragraph_format.space_before = Pt(3)
            paragraph.paragraph_format.space_after = Pt(3)
            paragraph.paragraph_format.line_spacing = 1.15
            paragraph.paragraph_format.keep_with_next = False
            if len(block["rows"]) <= 5 and row_index < len(block["rows"]):
                paragraph.paragraph_format.keep_with_next = True
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_index == 0 or len(text) < 12 else WD_ALIGN_PARAGRAPH.LEFT
            set_font(paragraph.add_run(text), font, 10.5, row_index == 0)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.first_line_indent = Pt(0)
    spacer.paragraph_format.space_after = Pt(3)
    spacer.paragraph_format.line_spacing = 1
    spacer.paragraph_format.space_before = Pt(0)
    set_font(spacer.add_run(""), font, 3)


def default_fonts():
    """Prefer installed cross-platform CJK fonts on headless Unix renderers."""
    matcher = shutil.which("fc-match")
    if matcher:
        chosen = []
        for name in ("Noto Serif CJK SC", "Noto Sans CJK SC"):
            try:
                match = subprocess.run([matcher, "-f", "%{family}", name], check=True, capture_output=True, text=True, timeout=5)
                if name in match.stdout:
                    chosen.append(name)
            except (OSError, subprocess.SubprocessError):
                break
        if len(chosen) == 2:
            return tuple(chosen)
    if sys.platform == "darwin":
        return "Songti SC", "Heiti SC"
    return "SimSun", "SimHei"


def build(data, output):
    validate(data)
    aliases = {"宋体": "SimSun", "黑体": "SimHei"}
    default_body, default_heading = default_fonts()
    font = data.get("font_east_asia", default_body)
    heading_font = data.get("font_heading", default_heading)
    font = aliases.get(font, font)
    heading_font = aliases.get(heading_font, heading_font)
    doc = Document()
    page = doc.sections[0]
    page.page_width, page.page_height = Mm(210), Mm(297)
    page.top_margin = page.bottom_margin = page.left_margin = page.right_margin = Mm(25.4)
    page.footer_distance = Mm(12)
    normal = doc.styles["Normal"]
    set_font(normal, font, 12)
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.first_line_indent = Pt(24)
    normal.paragraph_format.widow_control = True
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_font(doc.styles["Title"], heading_font, 18, True)
    for level, size in ((1, 18), (2, 15), (3, 13), (4, 12), (5, 12), (6, 11.5)):
        style = doc.styles[f"Heading {level}"]
        set_font(style, heading_font, size, True)
        style.paragraph_format.first_line_indent = Pt(0)
        style.paragraph_format.left_indent = Pt(max(level - 3, 0) * 6)
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.keep_together = True
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        props = style._element.get_or_add_pPr()
        outline = props.find(qn("w:outlineLvl"))
        if outline is None:
            outline = element("outlineLvl")
            props.append(outline)
        outline.set(qn("w:val"), str(level - 1))
    # Remove template rules and theme residue from the title style.
    title_props = doc.styles["Title"]._element.find(qn("w:pPr"))
    if title_props is not None:
        for border in list(title_props.findall(qn("w:pBdr"))):
            title_props.remove(border)
    doc.core_properties.title = data["title"]
    doc.core_properties.subject = "项目申报书正文"
    doc.core_properties.author = ""
    doc.core_properties.last_modified_by = ""
    doc.core_properties.comments = ""
    if data.get("page_numbers", False):
        footer = page.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.paragraph_format.first_line_indent = Pt(0)
        footer.paragraph_format.space_after = Pt(0)
        field_run = footer.add_run()
        set_font(field_run, font, 10)
        field = element("fldSimple", instr="PAGE")
        field_run._r.addnext(field)
    cover = doc.add_paragraph(data["title"], "Title")
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover.paragraph_format.first_line_indent = Pt(0)
    cover.paragraph_format.space_before = Pt(100 if data.get("create_cover", False) else 0)
    cover.paragraph_format.space_after = Pt(24 if data.get("create_cover", False) else 12)
    if data.get("subtitle"):
        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle.paragraph_format.first_line_indent = Pt(0)
        set_font(subtitle.add_run(data["subtitle"]), font, 14)
    if data.get("status"):
        status = doc.add_paragraph()
        status.alignment = WD_ALIGN_PARAGRAPH.CENTER
        status.paragraph_format.first_line_indent = Pt(0)
        set_font(status.add_run(data["status"]), font, 11)
    if data.get("create_cover", False):
        doc.add_page_break()
    headings = []
    for section in data["sections"] + data.get("appendices", []):
        headings.append((section["heading"], 1))
        for block in section["blocks"]:
            if block["type"] == "heading":
                headings.append((block["text"], block.get("level", 2)))
    if data.get("create_toc", False):
        toc_title = doc.add_paragraph()
        toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        toc_title.paragraph_format.first_line_indent = Pt(0)
        set_font(toc_title.add_run("目录"), heading_font, 16, True)
        for ident, (text, level) in enumerate(headings, 1):
            if level > 2:
                continue
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.first_line_indent = Pt(0)
            paragraph.paragraph_format.left_indent = Pt(18 * (level - 1))
            paragraph.paragraph_format.space_after = Pt(4)
            paragraph.paragraph_format.line_spacing = 1.15
            add_link(paragraph, text, f"proposal_{ident}", font)
        doc.add_page_break()
    heading_index = 0

    def add_heading(text, level):
        nonlocal heading_index
        heading_index += 1
        p = doc.add_heading(text, level)
        bookmark(p, heading_index)
        return p

    def add_section(section, start_new_page=False):
        title = add_heading(section["heading"], 1)
        title.paragraph_format.page_break_before = start_new_page
        for block in section["blocks"]:
            kind = block["type"]
            if kind == "heading":
                add_heading(block["text"], block.get("level", 2))
            elif kind == "table":
                add_table(doc, block, font)
            elif kind == "bullet":
                p = doc.add_paragraph(block["text"], "List Bullet")
                p.paragraph_format.first_line_indent = Pt(0)
                for run in p.runs:
                    set_font(run, font, 12)
            else:
                doc.add_paragraph(block["text"])

    for index, section in enumerate(data["sections"]):
        add_section(section, start_new_page=index > 0 and data.get("profile", "major-program") == "major-program")
    for section in data.get("appendices", []):
        add_section(section)
    output = Path(output)
    if output.suffix.lower() != ".docx":
        raise ValueError("Output path must end in .docx")
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.proposal-word-', dir=output.parent) as folder:
        draft = Path(folder) / 'proposal.docx'
        doc.save(draft)
        reopened = Document(draft)
        if not reopened.paragraphs or not draft.stat().st_size:
            raise RuntimeError("Generated DOCX could not be reopened")
        style = scan_docx(draft)
        if not style['passed']:
            messages = [f"{item['location']}：{item['rule']}“{item['matched']}”；{item['rewrite']}" for item in style['findings']]
            raise ValueError('Word终稿语体检查需修订：\n' + '\n'.join(messages))
        draft.replace(output)
    return {"output": str(output.resolve()), "headings": len(headings), "tables": len(reopened.tables), "font_east_asia": font, "font_heading": heading_font, "toc": "static editable hyperlinks without page numbers" if data.get("create_toc", False) else "omitted", "cover": data.get("create_cover", False), "page_numbers": data.get("page_numbers", False)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Complete proposal JSON")
    parser.add_argument("--output", required=True, help="Final .docx path")
    args = parser.parse_args()
    draft = json.loads(Path(args.input).read_text(encoding="utf-8"))
    print(json.dumps(build(draft, args.output), ensure_ascii=False))


if __name__ == "__main__":
    main()
