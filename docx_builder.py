"""Compact mentor review Word export for the presentation builder.

The review document is deliberately organized around what the mentor needs to
see and edit: the actual PowerPoint slide, the editable on-slide wording,
speaker notes, and a feedback area. App-only implementation metadata is kept out
of the main review flow.
"""

from __future__ import annotations

import base64
from io import BytesIO
from typing import Any, Dict, List, Tuple
from zipfile import ZIP_DEFLATED, ZipFile

from PIL import Image
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from deck_model import APP_VERSION, identity_subtitle, identity_title, slide_output_title
from preview_utils import render_pptx_slides_to_pngs

BLUE = "1F4E79"
LIGHT_BLUE = "D9EAF7"
PALE_BLUE = "EEF5FB"
PALE_GRAY = "F5F6F7"
PINK = "FCE8E8"
WHITE = "FFFFFF"
BORDER_BLUE = "9BB8D6"
BORDER_GRAY = "B7B7B7"
DOC_FONT = "Calibri"
TEXT_DARK = RGBColor(35, 35, 35)
TEXT_MUTED = RGBColor(95, 95, 95)
TEXT_BLUE = RGBColor(31, 78, 121)
TWIPS_PER_INCH = 1440


def _safe_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _decode_payload(payload: Any) -> bytes | None:
    if not isinstance(payload, dict):
        return None
    encoded = payload.get("data_base64")
    if not encoded:
        return None
    try:
        return base64.b64decode(encoded)
    except Exception:
        return None


def _fallback_visual_bytes(slide: Dict[str, Any]) -> bytes | None:
    return _decode_payload(slide.get("uploaded_slide_preview_image")) or _decode_payload(slide.get("visual_image"))


def _shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _set_cell_borders(cell, color: str = BORDER_GRAY, size: str = "6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def _set_cell_margins(cell, top: int = 90, start: int = 110, bottom: int = 90, end: int = 110) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, value in {"top": top, "start": start, "bottom": bottom, "end": end}.items():
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _lock_table_width(table, width_inches: float) -> None:
    table.autofit = False
    table.allow_autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(width_inches * TWIPS_PER_INCH)))
    tbl_w.set(qn("w:type"), "dxa")
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    for row in table.rows:
        for cell in row.cells:
            cell.width = Inches(width_inches / len(row.cells))


def _clear_cell(cell) -> None:
    cell.text = ""
    if not cell.paragraphs:
        cell.add_paragraph()


def _set_cell_width(cell, width_inches: float) -> None:
    cell.width = Inches(width_inches)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width_inches * TWIPS_PER_INCH)))
    tc_w.set(qn("w:type"), "dxa")


def _shade_paragraph(paragraph, fill: str, border: str | None = None) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    shd = p_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        p_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    if border:
        p_bdr = p_pr.find(qn("w:pBdr"))
        if p_bdr is None:
            p_bdr = OxmlElement("w:pBdr")
            p_pr.append(p_bdr)
        for edge in ("top", "left", "bottom", "right"):
            node = p_bdr.find(qn(f"w:{edge}"))
            if node is None:
                node = OxmlElement(f"w:{edge}")
                p_bdr.append(node)
            node.set(qn("w:val"), "single")
            node.set(qn("w:sz"), "4")
            node.set(qn("w:space"), "0")
            node.set(qn("w:color"), border)


def _set_row_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = tr_pr.find(qn("w:cantSplit"))
    if cant_split is None:
        cant_split = OxmlElement("w:cantSplit")
        tr_pr.append(cant_split)


def _set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = tr_pr.find(qn("w:tblHeader"))
    if tbl_header is None:
        tbl_header = OxmlElement("w:tblHeader")
        tr_pr.append(tbl_header)
    tbl_header.set(qn("w:val"), "true")


def _add_cell_section(
    cell,
    heading: str,
    body: str,
    *,
    heading_fill: str = LIGHT_BLUE,
    body_fill: str = WHITE,
    body_font_size: float = 8.4,
    placeholder: str = "[blank]",
) -> None:
    heading_p = cell.add_paragraph()
    heading_p.paragraph_format.space_before = Pt(1)
    heading_p.paragraph_format.space_after = Pt(1)
    heading_p.paragraph_format.left_indent = Pt(3)
    heading_p.paragraph_format.right_indent = Pt(3)
    _shade_paragraph(heading_p, heading_fill, BORDER_BLUE)
    heading_run = heading_p.add_run(heading)
    heading_run.font.name = DOC_FONT
    heading_run.font.size = Pt(8.0)
    heading_run.font.bold = True
    heading_run.font.color.rgb = TEXT_BLUE if heading_fill != BLUE else RGBColor(255, 255, 255)

    body_p = cell.add_paragraph()
    body_p.paragraph_format.space_before = Pt(0)
    body_p.paragraph_format.space_after = Pt(3)
    body_p.paragraph_format.left_indent = Pt(4)
    body_p.paragraph_format.right_indent = Pt(4)
    body_p.paragraph_format.line_spacing = 1.0
    _shade_paragraph(body_p, body_fill, BORDER_BLUE)
    value = _safe_text(body) or placeholder
    run = body_p.add_run(value)
    run.font.name = DOC_FONT
    run.font.size = Pt(body_font_size)
    run.font.italic = not bool(_safe_text(body))
    run.font.color.rgb = TEXT_DARK if _safe_text(body) else TEXT_MUTED


def _write_cell_text(
    cell,
    text: Any,
    *,
    font_size: float = 9.5,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor = TEXT_DARK,
    align=WD_ALIGN_PARAGRAPH.LEFT,
) -> None:
    _clear_cell(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    lines = _safe_text(text).splitlines() or [""]
    for idx, line in enumerate(lines):
        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(line)
        run.font.name = DOC_FONT
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color


def _add_section_bar(doc: Document, text: str, fill: str = BLUE, font_size: float = 11.5) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    _lock_table_width(table, 7.35)
    cell = table.cell(0, 0)
    _shade_cell(cell, fill)
    _set_cell_borders(cell, fill)
    _set_cell_margins(cell, 75, 110, 75, 110)
    _write_cell_text(cell, text, font_size=font_size, bold=True, color=RGBColor(255, 255, 255))


def _add_labeled_box(
    doc: Document,
    heading: str,
    body: str,
    *,
    fill: str = WHITE,
    border: str = BORDER_BLUE,
    body_font_size: float = 9.2,
    placeholder: str = "[blank]",
    min_blank_lines: int = 0,
) -> None:
    table = doc.add_table(rows=2, cols=1)
    table.style = "Table Grid"
    _lock_table_width(table, 7.35)

    heading_cell = table.cell(0, 0)
    _shade_cell(heading_cell, LIGHT_BLUE)
    _set_cell_borders(heading_cell, border)
    _set_cell_margins(heading_cell, 55, 90, 55, 90)
    _write_cell_text(heading_cell, heading, font_size=9.4, bold=True, color=TEXT_BLUE)

    body_cell = table.cell(1, 0)
    _shade_cell(body_cell, fill)
    _set_cell_borders(body_cell, border)
    _set_cell_margins(body_cell, 80, 105, 80, 105)
    value = _safe_text(body) or placeholder
    if min_blank_lines > 0 and not _safe_text(body):
        value = placeholder + ("\n" * min_blank_lines)
    _write_cell_text(
        body_cell,
        value,
        font_size=body_font_size,
        italic=not bool(_safe_text(body)),
        color=TEXT_MUTED if not _safe_text(body) else TEXT_DARK,
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def _add_overview(doc: Document, deck: Dict[str, Any]) -> None:
    meta = deck.get("metadata", {}) if isinstance(deck, dict) else {}

    title_table = doc.add_table(rows=2, cols=1)
    title_table.style = "Table Grid"
    _lock_table_width(title_table, 7.40)
    title_cell = title_table.cell(0, 0)
    _shade_cell(title_cell, BLUE)
    _set_cell_borders(title_cell, BLUE)
    _set_cell_margins(title_cell, 75, 110, 75, 110)
    _write_cell_text(title_cell, "MENTOR PRESENTATION REVIEW", font_size=14, bold=True, color=RGBColor(255, 255, 255), align=WD_ALIGN_PARAGRAPH.CENTER)

    identity_cell = title_table.cell(1, 0)
    _shade_cell(identity_cell, LIGHT_BLUE)
    _set_cell_borders(identity_cell, BORDER_BLUE)
    _set_cell_margins(identity_cell, 65, 100, 65, 100)
    _write_cell_text(identity_cell, f"{identity_title(deck)}\n{identity_subtitle(deck)}", font_size=9.7, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)

    overview = doc.add_table(rows=0, cols=2)
    overview.style = "Table Grid"
    overview.autofit = False
    overview.alignment = WD_TABLE_ALIGNMENT.CENTER
    fields: List[Tuple[str, str]] = [
        ("Title / subtitle", " — ".join(filter(None, [_safe_text(meta.get("presentation_title")) or identity_title(deck), _safe_text(meta.get("presentation_subtitle"))]))),
        ("Presenter / date", " · ".join(filter(None, [_safe_text(meta.get("presenter")), _safe_text(meta.get("session_date"))]))),
        ("Audience / type", " · ".join(filter(None, [_safe_text(meta.get("audience")), _safe_text(meta.get("presentation_type"))]))),
        ("Core question", _safe_text(meta.get("core_question"))),
        ("Story arc", _safe_text(meta.get("story_arc"))),
    ]
    for label, value in fields:
        row = overview.add_row()
        _set_row_cant_split(row)
        left, right = row.cells
        _set_cell_width(left, 1.45)
        _set_cell_width(right, 5.95)
        _shade_cell(left, PALE_GRAY)
        _shade_cell(right, WHITE)
        _set_cell_borders(left, BORDER_GRAY)
        _set_cell_borders(right, BORDER_GRAY)
        _set_cell_margins(left, 48, 75, 48, 75)
        _set_cell_margins(right, 48, 75, 48, 75)
        _write_cell_text(left, label, font_size=8.3, bold=True, color=TEXT_BLUE)
        _write_cell_text(right, value or "[blank]", font_size=8.3, italic=not bool(value), color=TEXT_DARK if value else TEXT_MUTED)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)

    review = doc.add_table(rows=2, cols=1)
    review.style = "Table Grid"
    _lock_table_width(review, 7.40)
    review_heading = review.cell(0, 0)
    _shade_cell(review_heading, PINK)
    _set_cell_borders(review_heading, BORDER_BLUE)
    _set_cell_margins(review_heading, 45, 80, 45, 80)
    _write_cell_text(review_heading, "Overall mentor feedback", font_size=8.5, bold=True, color=TEXT_BLUE)
    review_body = review.cell(1, 0)
    _shade_cell(review_body, WHITE)
    _set_cell_borders(review_body, BORDER_BLUE)
    _set_cell_margins(review_body, 55, 85, 55, 85)
    _write_cell_text(review_body, "Comment on the overall story, sequencing, omissions, and alignment between objectives and content.\n\n", font_size=8.2, italic=True, color=TEXT_MUTED)

    checklist = doc.add_paragraph()
    checklist.paragraph_format.space_before = Pt(4)
    checklist.paragraph_format.space_after = Pt(1)
    checklist.paragraph_format.line_spacing = 1.0
    run = checklist.add_run("Review focus: story and sequencing · readable/necessary slides · accurate teaching points · useful transitions and speaker notes · clear interpretation of visuals/data")
    run.font.name = DOC_FONT
    run.font.size = Pt(7.8)
    run.font.color.rgb = TEXT_MUTED

    version_p = doc.add_paragraph()
    version_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    version_p.paragraph_format.space_before = Pt(0)
    version_p.paragraph_format.space_after = Pt(0)
    version_run = version_p.add_run(f"Generated with {APP_VERSION} · Track Changes is enabled.")
    version_run.font.name = DOC_FONT
    version_run.font.size = Pt(7.2)
    version_run.font.color.rgb = TEXT_MUTED


def _slide_editable_text(deck: Dict[str, Any], slide: Dict[str, Any], index: int) -> str:
    role = _safe_text(slide.get("role"))
    kind = _safe_text(slide.get("slide_kind"))
    meta = deck.get("metadata", {}) if isinstance(deck, dict) else {}
    parts: List[str] = []

    def add(label: str, value: Any) -> None:
        text = _safe_text(value)
        if text:
            parts.append(f"{label}: {text}")

    add("Title", slide_output_title(deck, slide, index))
    add("Subtitle", slide.get("subtitle"))
    add("Section label", slide.get("section_box_label"))

    if role == "Title" or kind == "title":
        add("Presenter", meta.get("presenter"))
        add("Date", meta.get("session_date"))
        add("Audience", meta.get("audience"))
        add("Presentation type", meta.get("presentation_type"))
        add("Core question", meta.get("core_question"))
        add("Story arc", meta.get("story_arc"))
    elif role == "Objectives" or kind == "objectives":
        add("Intro", slide.get("objectives_intro"))
        for number in range(1, 4):
            verb = _safe_text(slide.get(f"objective_{number}_verb"))
            sentence = _safe_text(slide.get(f"objective_{number}_text"))
            if verb or sentence:
                parts.append(f"Objective {number}: {verb}: {sentence}".replace(": :", ":"))
        add("Bottom banner", slide.get("objectives_takeaway"))
    elif role == "Take-home" or kind == "takehome":
        for number in range(1, 6):
            point = _safe_text(slide.get(f"takehome_point_{number}"))
            if point:
                parts.append(f"{number}. {point}")
    else:
        add("Slide text", slide.get("body"))

    add("Discussion prompt", slide.get("discussion_prompt"))
    return "\n".join(parts) or "[No editable on-slide wording entered in the app. Review the PowerPoint preview above.]"


def _fit_image(image_bytes: bytes, max_width: float = 2.30, max_height: float = 1.30) -> Tuple[float, float]:
    try:
        with Image.open(BytesIO(image_bytes)) as image:
            width_px, height_px = image.size
        if not width_px or not height_px:
            return max_width, max_height
        ratio = width_px / height_px
        box_ratio = max_width / max_height
        if ratio >= box_ratio:
            return max_width, max_width / ratio
        return max_height * ratio, max_height
    except Exception:
        return max_width, max_height


def _add_slide_review_table(
    doc: Document,
    deck: Dict[str, Any],
    slide: Dict[str, Any],
    index: int,
    total: int,
    preview_bytes: bytes | None,
) -> None:
    title = slide_output_title(deck, slide, index)
    role = _safe_text(slide.get("role")) or "Slide"

    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(7.40 * TWIPS_PER_INCH)))
    tbl_w.set(qn("w:type"), "dxa")
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    header = table.rows[0]
    _set_row_cant_split(header)
    _set_repeat_table_header(header)
    header_cell = header.cells[0].merge(header.cells[1])
    _shade_cell(header_cell, BLUE)
    _set_cell_borders(header_cell, BLUE)
    _set_cell_margins(header_cell, 52, 85, 52, 85)
    _write_cell_text(
        header_cell,
        f"Slide {index} of {total} · {role} · {title}",
        font_size=9.6,
        bold=True,
        color=RGBColor(255, 255, 255),
    )

    content_row = table.rows[1]
    preview_cell, text_cell = content_row.cells
    _set_cell_width(preview_cell, 2.52)
    _set_cell_width(text_cell, 4.88)
    _shade_cell(preview_cell, WHITE)
    _shade_cell(text_cell, WHITE)
    _set_cell_borders(preview_cell, BORDER_BLUE)
    _set_cell_borders(text_cell, BORDER_BLUE)
    _set_cell_margins(preview_cell, 55, 55, 55, 55)
    _set_cell_margins(text_cell, 40, 55, 40, 55)
    preview_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    text_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    _clear_cell(preview_cell)
    preview_heading = preview_cell.paragraphs[0]
    preview_heading.paragraph_format.space_after = Pt(3)
    preview_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _shade_paragraph(preview_heading, LIGHT_BLUE, BORDER_BLUE)
    preview_run = preview_heading.add_run("SLIDE PREVIEW")
    preview_run.font.name = DOC_FONT
    preview_run.font.size = Pt(7.8)
    preview_run.font.bold = True
    preview_run.font.color.rgb = TEXT_BLUE

    image_p = preview_cell.add_paragraph()
    image_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    image_p.paragraph_format.space_before = Pt(1)
    image_p.paragraph_format.space_after = Pt(2)
    image_data = preview_bytes or _fallback_visual_bytes(slide)
    if image_data:
        width, height = _fit_image(image_data)
        image_p.add_run().add_picture(BytesIO(image_data), width=Inches(width), height=Inches(height))
    else:
        image_run = image_p.add_run("[Preview unavailable]")
        image_run.font.name = DOC_FONT
        image_run.font.size = Pt(8)
        image_run.font.italic = True
        image_run.font.color.rgb = TEXT_MUTED

    _add_cell_section(
        preview_cell,
        "MENTOR FEEDBACK",
        "",
        heading_fill=PINK,
        body_fill=WHITE,
        body_font_size=8.0,
        placeholder="Add comments on accuracy, clarity, design, transitions, or suggested revisions.\n\n\n",
    )

    _clear_cell(text_cell)
    _add_cell_section(
        text_cell,
        "ON-SLIDE WORDING",
        _slide_editable_text(deck, slide, index),
        heading_fill=LIGHT_BLUE,
        body_fill=PALE_BLUE,
        body_font_size=8.1,
        placeholder="[No editable on-slide wording entered.]",
    )
    _add_cell_section(
        text_cell,
        "SPEAKER NOTES",
        _safe_text(slide.get("speaker_notes")),
        heading_fill=LIGHT_BLUE,
        body_fill=WHITE,
        body_font_size=8.2,
        placeholder="[No speaker notes entered.]",
    )

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(3)


def _enable_track_changes(docx_stream: BytesIO) -> BytesIO:
    try:
        source = BytesIO(docx_stream.getvalue())
        target = BytesIO()
        with ZipFile(source, "r") as zin, ZipFile(target, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/settings.xml":
                    xml = data.decode("utf-8")
                    if "w:trackRevisions" not in xml:
                        xml = xml.replace("</w:settings>", "<w:trackRevisions/></w:settings>")
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
        target.seek(0)
        return target
    except Exception:
        docx_stream.seek(0)
        return docx_stream


def build_mentor_review_docx(deck: Dict[str, Any], pptx_bytes: bytes | None = None) -> bytes:
    """Build the streamlined, editable mentor review document."""
    if pptx_bytes is None:
        try:
            from pptx_builder import build_pptx
            pptx_bytes = build_pptx(deck)
        except Exception:
            pptx_bytes = None

    previews = render_pptx_slides_to_pngs(pptx_bytes or b"") if pptx_bytes else []

    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.38)
    section.bottom_margin = Inches(0.38)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    styles = doc.styles
    styles["Normal"].font.name = DOC_FONT
    styles["Normal"].font.size = Pt(8.4)

    _add_overview(doc, deck)
    doc.add_page_break()

    slides = deck.get("slides", []) if isinstance(deck, dict) else []
    total = len(slides)
    for index, slide in enumerate(slides, start=1):
        preview = previews[index - 1] if index - 1 < len(previews) else None
        _add_slide_review_table(doc, deck, slide, index, total, preview)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return _enable_track_changes(output).getvalue()


def mentor_docx_contains_complete_review_fields(docx_bytes: bytes) -> bool:
    required_labels = [
        "MENTOR PRESENTATION REVIEW",
        "SLIDE PREVIEW",
        "ON-SLIDE WORDING",
        "SPEAKER NOTES",
        "MENTOR FEEDBACK",
        f"Generated with {APP_VERSION}",
    ]
    try:
        with ZipFile(BytesIO(docx_bytes), "r") as zin:
            xml = zin.read("word/document.xml").decode("utf-8", errors="ignore")
        return all(label in xml for label in required_labels)
    except Exception:
        return False


def build_plain_text_summary(deck: Dict[str, Any]) -> str:
    meta = deck.get("metadata", {}) if isinstance(deck, dict) else {}
    parts = [identity_title(deck), identity_subtitle(deck), ""]
    for key in ("presentation_subtitle", "core_question", "story_arc", "archive_notes"):
        value = _safe_text(meta.get(key))
        if value:
            parts.append(f"{key}: {value}")
    parts.append("")

    for index, slide in enumerate(deck.get("slides", []), start=1):
        parts.append(f"Slide {index}: {slide_output_title(deck, slide, index)}")
        parts.append(_slide_editable_text(deck, slide, index))
        notes = _safe_text(slide.get("speaker_notes"))
        if notes:
            parts.append(f"speaker_notes: {notes}")
        parts.append("")
    return "\n".join(parts)
