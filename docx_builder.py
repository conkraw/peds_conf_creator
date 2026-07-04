"""Mentor review Word export for the Presentation PowerPoint Builder.

The document mirrors the user's preferred Journal Club review style: portrait,
blue slide headers, gray field labels, editable text blocks, compact PowerPoint
previews, full speaker notes, and mentor comment areas.
"""

from __future__ import annotations

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

from deck_model import APP_VERSION, ensure_core_slide_order, identity_subtitle, identity_title, slide_output_title
from preview_utils import render_pptx_slides_to_pngs

BLUE = "4A90D9"
DARK_BLUE = "1F4E79"
GRAY_HEADER = "D9D9D9"
PALE_BLUE = "EEF5FB"
PALE_PINK = "FCE8E8"
WHITE = "FFFFFF"
BORDER = "666666"
DOC_FONT = "Calibri"
TEXT_DARK = RGBColor(35, 35, 35)
TEXT_MUTED = RGBColor(95, 95, 95)
TEXT_BLUE = RGBColor(31, 78, 121)
TWIPS_PER_INCH = 1440


def _safe_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _set_cell_borders(cell, color: str = BORDER, size: str = "6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def _set_cell_margins(cell, top: int = 65, start: int = 85, bottom: int = 65, end: int = 85) -> None:
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


def _set_cell_width(cell, width_inches: float) -> None:
    cell.width = Inches(width_inches)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width_inches * TWIPS_PER_INCH)))
    tc_w.set(qn("w:type"), "dxa")


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


def _set_row_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def _keep_with_next(paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    keep = p_pr.find(qn("w:keepNext"))
    if keep is None:
        keep = OxmlElement("w:keepNext")
        p_pr.append(keep)
    keep.set(qn("w:val"), "1")


def _clear_cell(cell) -> None:
    cell.text = ""
    if not cell.paragraphs:
        cell.add_paragraph()


def _write_cell_text(
    cell,
    text: Any,
    *,
    font_size: float = 8.6,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor = TEXT_DARK,
    align=WD_ALIGN_PARAGRAPH.LEFT,
) -> None:
    _clear_cell(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    lines = str(text or "").splitlines() or [""]
    for idx, line in enumerate(lines):
        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(line)
        run.font.name = DOC_FONT
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color


def _add_spacer(doc: Document, points: float = 4) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(points)


def _add_bar(doc: Document, text: str, *, fill: str = BLUE, font_size: float = 10.2) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    _lock_table_width(table, 7.30)
    cell = table.cell(0, 0)
    _shade_cell(cell, fill)
    _set_cell_borders(cell, BORDER)
    _set_cell_margins(cell, 52, 85, 52, 85)
    _write_cell_text(
        cell,
        text,
        font_size=font_size,
        bold=True,
        color=RGBColor(255, 255, 255),
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _keep_with_next(cell.paragraphs[0])


def _add_labeled_box(
    doc: Document,
    heading: str,
    body: str,
    *,
    body_fill: str = WHITE,
    body_font_size: float = 8.5,
    placeholder: str = "[blank]",
    blank_lines: int = 0,
) -> None:
    table = doc.add_table(rows=2, cols=1)
    table.style = "Table Grid"
    _lock_table_width(table, 7.30)
    _set_row_cant_split(table.rows[0])

    head = table.cell(0, 0)
    _shade_cell(head, GRAY_HEADER)
    _set_cell_borders(head, BORDER)
    _set_cell_margins(head, 45, 75, 45, 75)
    _write_cell_text(head, heading.upper(), font_size=8.2, bold=True, color=TEXT_DARK)
    _keep_with_next(head.paragraphs[0])

    body_cell = table.cell(1, 0)
    _shade_cell(body_cell, body_fill)
    _set_cell_borders(body_cell, BORDER)
    _set_cell_margins(body_cell, 48, 72, 48, 72)
    value = _safe_text(body)
    if not value:
        value = placeholder + ("\n" * blank_lines)
    _write_cell_text(
        body_cell,
        value,
        font_size=body_font_size,
        italic=not bool(_safe_text(body)),
        color=TEXT_DARK if _safe_text(body) else TEXT_MUTED,
    )
    _add_spacer(doc, 3)


def _add_title_and_guidelines(doc: Document, deck: Dict[str, Any]) -> None:
    meta = deck.get("metadata", {}) if isinstance(deck, dict) else {}

    title = doc.add_table(rows=3, cols=1)
    title.style = "Table Grid"
    _lock_table_width(title, 7.30)

    cell = title.cell(0, 0)
    _shade_cell(cell, GRAY_HEADER)
    _set_cell_borders(cell, BORDER)
    _set_cell_margins(cell, 42, 75, 42, 75)
    _write_cell_text(cell, "PRESENTATION BUILDER", font_size=9.2, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    cell = title.cell(1, 0)
    _shade_cell(cell, WHITE)
    _set_cell_borders(cell, BORDER)
    _set_cell_margins(cell, 48, 75, 48, 75)
    _write_cell_text(cell, "Mentor PowerPoint Review", font_size=14.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    cell = title.cell(2, 0)
    _shade_cell(cell, WHITE)
    _set_cell_borders(cell, BORDER)
    _set_cell_margins(cell, 46, 75, 46, 75)
    subtitle = "\n".join(filter(None, [identity_title(deck), identity_subtitle(deck)]))
    _write_cell_text(cell, subtitle, font_size=8.7, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _add_spacer(doc, 6)

    _add_bar(doc, "REVIEWER GUIDELINES")
    _add_labeled_box(
        doc,
        "Reviewer focus",
        "• Use Track Changes and/or comments so the presenter can see your suggestions.\n"
        "• Focus on clarity, accuracy, educational value, clinical reasoning, and whether the message is easy to present aloud.\n"
        "• Preserve the presenter’s voice when possible; suggest targeted edits rather than rewriting the entire presentation.\n"
        "• Flag overstatements, missing limitations, unclear applicability, jargon, or places where the takeaway is too broad.\n"
        "• Review the actual slide image and the speaker notes together so transitions and teaching points remain aligned.",
        body_font_size=8.1,
    )
    _add_labeled_box(
        doc,
        "Reviewer workflow",
        "Edit the wording below directly, or add Word comments beside sections that need discussion. The slide preview shows the actual exported PowerPoint. Speaker notes are included in full so the reviewer can assess rehearsal flow, transitions, and accuracy.",
        body_font_size=8.1,
    )

    _add_bar(doc, "PRESENTATION OVERVIEW")
    fields: List[Tuple[str, str]] = [
        ("Presentation title / subtitle", " — ".join(filter(None, [_safe_text(meta.get("presentation_title")) or identity_title(deck), _safe_text(meta.get("presentation_subtitle"))]))),
        ("Presenter / date", " · ".join(filter(None, [_safe_text(meta.get("presenter")), _safe_text(meta.get("session_date"))]))),
        ("Audience / type", " · ".join(filter(None, [_safe_text(meta.get("audience")), _safe_text(meta.get("presentation_type"))]))),
        ("Core question / tension", _safe_text(meta.get("core_question"))),
        ("Story arc", _safe_text(meta.get("story_arc"))),
    ]
    for heading, body in fields:
        _add_labeled_box(doc, heading, body, body_font_size=8.25)

    _add_labeled_box(
        doc,
        "Overall mentor notes / comments",
        "",
        body_fill=PALE_PINK,
        body_font_size=8.3,
        placeholder="[Add comments here or use Word comments in the margin.]",
        blank_lines=2,
    )

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(f"Generated with {APP_VERSION} · Track Changes is enabled.")
    run.font.name = DOC_FONT
    run.font.size = Pt(7.0)
    run.font.color.rgb = TEXT_MUTED


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
    return "\n".join(parts) or "[No editable on-slide wording entered in the app.]"


def _fit_image(image_bytes: bytes, max_width: float = 2.90, max_height: float = 1.63) -> Tuple[float, float]:
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


def _add_slide_preview_and_wording(
    doc: Document,
    deck: Dict[str, Any],
    slide: Dict[str, Any],
    index: int,
    preview_bytes: bytes | None,
) -> None:
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    _lock_table_width(table, 7.30)
    _set_cell_width(table.cell(0, 0), 3.35)
    _set_cell_width(table.cell(0, 1), 3.95)
    _set_cell_width(table.cell(1, 0), 3.35)
    _set_cell_width(table.cell(1, 1), 3.95)

    for cell, heading in zip(table.rows[0].cells, ["POWERPOINT PREVIEW", "EDITABLE ON-SLIDE WORDING"]):
        _shade_cell(cell, GRAY_HEADER)
        _set_cell_borders(cell, BORDER)
        _set_cell_margins(cell, 43, 70, 43, 70)
        _write_cell_text(cell, heading, font_size=8.0, bold=True)
        _keep_with_next(cell.paragraphs[0])
    _set_row_cant_split(table.rows[0])

    preview_cell, wording_cell = table.rows[1].cells
    for cell in (preview_cell, wording_cell):
        _shade_cell(cell, WHITE)
        _set_cell_borders(cell, BORDER)
        _set_cell_margins(cell, 55, 60, 55, 60)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    _clear_cell(preview_cell)
    p = preview_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    image_data = preview_bytes
    if image_data:
        width, height = _fit_image(image_data)
        p.add_run().add_picture(BytesIO(image_data), width=Inches(width), height=Inches(height))
    else:
        run = p.add_run("[PowerPoint preview unavailable]")
        run.font.name = DOC_FONT
        run.font.size = Pt(8)
        run.font.italic = True
        run.font.color.rgb = TEXT_MUTED

    _write_cell_text(wording_cell, _slide_editable_text(deck, slide, index), font_size=8.0)
    _set_row_cant_split(table.rows[1])
    _add_spacer(doc, 3)


def _add_slide_section(
    doc: Document,
    deck: Dict[str, Any],
    slide: Dict[str, Any],
    index: int,
    total: int,
    preview_bytes: bytes | None,
) -> None:
    role = _safe_text(slide.get("role")) or "Slide"
    title = slide_output_title(deck, slide, index)
    _add_bar(doc, f"SLIDE {index} OF {total}: {title.upper()}")

    role_p = doc.add_paragraph()
    role_p.paragraph_format.space_before = Pt(1)
    role_p.paragraph_format.space_after = Pt(2)
    run = role_p.add_run(role.upper())
    run.font.name = DOC_FONT
    run.font.size = Pt(7.7)
    run.font.bold = True
    run.font.color.rgb = TEXT_MUTED
    _keep_with_next(role_p)

    _add_slide_preview_and_wording(doc, deck, slide, index, preview_bytes)
    _add_labeled_box(
        doc,
        "Speaker notes",
        _safe_text(slide.get("speaker_notes")),
        body_font_size=7.9,
        placeholder="[No speaker notes entered.]",
    )
    _add_labeled_box(
        doc,
        "Mentor notes / comments",
        "",
        body_fill=PALE_PINK,
        body_font_size=8.2,
        placeholder="[Add comments here or use Word comments in the margin.]",
        blank_lines=0,
    )
    _add_spacer(doc, 5)


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
    """Build the compact Journal-Club-style mentor review document."""
    ensure_core_slide_order(deck)
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
    section.top_margin = Inches(0.48)
    section.bottom_margin = Inches(0.48)
    section.left_margin = Inches(0.60)
    section.right_margin = Inches(0.60)

    doc.styles["Normal"].font.name = DOC_FONT
    doc.styles["Normal"].font.size = Pt(8.4)

    _add_title_and_guidelines(doc, deck)

    slides = deck.get("slides", []) if isinstance(deck, dict) else []
    total = len(slides)
    for index, slide in enumerate(slides, start=1):
        preview = previews[index - 1] if index - 1 < len(previews) else None
        _add_slide_section(doc, deck, slide, index, total, preview)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return _enable_track_changes(output).getvalue()


def mentor_docx_contains_complete_review_fields(docx_bytes: bytes) -> bool:
    required_labels = [
        "Mentor PowerPoint Review",
        "POWERPOINT PREVIEW",
        "EDITABLE ON-SLIDE WORDING",
        "SPEAKER NOTES",
        "MENTOR NOTES / COMMENTS",
        f"Generated with {APP_VERSION}",
    ]
    try:
        with ZipFile(BytesIO(docx_bytes), "r") as zin:
            xml = zin.read("word/document.xml").decode("utf-8", errors="ignore")
        return all(label in xml for label in required_labels)
    except Exception:
        return False
