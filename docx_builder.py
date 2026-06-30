"""
Word planning-form builder for the Presentation Builder app.

This creates a DOCX companion document that behaves like the printable planning
forms in the journal-club/case-conference apps: title block, objectives,
slide-by-slide storyboard, speaker-note prompts, mentor review, and final archive
checklist.
"""

from __future__ import annotations

import io
from typing import Any, Dict, Iterable, List

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from deck_model import CHECKLIST_LABELS, split_nonempty_lines, slide_output_title

BLUE = "D9EAF7"
GREEN = "BFEDD2"
RED = "F4CCCC"
GRAY = "E7E6E6"
PALE_BLUE = "F7FBFF"
DARK_BLUE = RGBColor(31, 78, 121)


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text(cell, text: str, bold: bool = False, font_size: int = 9) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text or "")
    run.font.name = "Aptos"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def set_table_borders(table) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "B7C9D8")


def add_heading(doc: Document, text: str, color: RGBColor = DARK_BLUE) -> None:
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 2"]
    run = p.add_run(text)
    run.font.name = "Aptos Display"
    run.font.bold = True
    run.font.color.rgb = color
    run.font.size = Pt(13)


def add_small_note(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = "Aptos"
    run.font.size = Pt(8.5)
    run.font.italic = True
    run.font.color.rgb = RGBColor(90, 90, 90)


def configure_document(doc: Document) -> None:
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(9)
    styles["Heading 1"].font.name = "Aptos Display"
    styles["Heading 1"].font.size = Pt(16)
    styles["Heading 1"].font.bold = True
    styles["Heading 2"].font.name = "Aptos Display"
    styles["Heading 2"].font.size = Pt(13)
    styles["Heading 2"].font.bold = True


def add_title_block(doc: Document, deck: Dict[str, Any]) -> None:
    meta = deck.get("metadata", {})
    title = meta.get("presentation_title") or "Untitled Presentation"

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("PRESENTATION PLANNING FORM")
    run.font.name = "Aptos Display"
    run.font.bold = True
    run.font.color.rgb = DARK_BLUE
    run.font.size = Pt(15)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.font.name = "Aptos Display"
    run.font.bold = True
    run.font.size = Pt(13)

    table = doc.add_table(rows=4, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_borders(table)
    rows = [
        ("Presenter", meta.get("presenter", ""), "Date", meta.get("session_date", "")),
        ("Audience", meta.get("audience", ""), "Type", meta.get("presentation_type", "")),
        ("Core question", meta.get("core_question", ""), "Story arc", meta.get("story_arc", "")),
        ("GitHub/archive notes", meta.get("github_notes", ""), "Version", deck.get("app_version", "")),
    ]
    for row, values in zip(table.rows, rows):
        for i, value in enumerate(values):
            set_cell_text(row.cells[i], value, bold=i in (0, 2), font_size=8.5)
            if i in (0, 2):
                set_cell_shading(row.cells[i], GRAY)
            else:
                set_cell_shading(row.cells[i], PALE_BLUE)


def add_objectives(doc: Document, deck: Dict[str, Any]) -> None:
    add_heading(doc, "Objectives")
    add_small_note(doc, "Use measurable Bloom-style verbs such as describe, differentiate, apply, analyze, appraise, or design.")
    objectives_slide = next((slide for slide in deck.get("slides", []) if slide.get("role") == "Objectives"), None)
    objectives = split_nonempty_lines(objectives_slide.get("body", "") if objectives_slide else "")
    if not objectives:
        objectives = ["Objective 1:", "Objective 2:", "Objective 3:"]
    for objective in objectives:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(objective)
        run.font.name = "Aptos"
        run.font.size = Pt(9)


def slide_summary(slide: Dict[str, Any]) -> str:
    parts = []
    body = slide.get("body", "").strip()
    if body:
        parts.append(body)
    visual = slide.get("visual_plan", "").strip()
    if visual:
        parts.append(f"Visual/evidence: {visual}")
    prompt = slide.get("discussion_prompt", "").strip()
    if prompt:
        parts.append(f"Audience prompt: {prompt}")
    return "\n".join(parts)


def add_storyboard(doc: Document, deck: Dict[str, Any]) -> None:
    add_heading(doc, "Slide storyboard")
    add_small_note(doc, "Slide titles should create a narrative. Blank titles are allowed while drafting and will export as generic story-slide titles.")

    included = [slide for slide in deck.get("slides", []) if slide.get("include", True)]
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_borders(table)
    headers = ["#", "Role", "Slide title", "Main message / visual plan", "Speaker notes"]
    for i, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], header, bold=True, font_size=8.5)
        set_cell_shading(table.rows[0].cells[i], BLUE)

    for index, slide in enumerate(included, start=1):
        row = table.add_row().cells
        set_cell_text(row[0], str(index), font_size=8)
        set_cell_text(row[1], slide.get("role", ""), font_size=8)
        set_cell_text(row[2], slide_output_title(deck, slide, index), font_size=8)
        set_cell_text(row[3], slide_summary(slide), font_size=8)
        set_cell_text(row[4], slide.get("speaker_notes", ""), font_size=8)
        for cell in row:
            set_cell_shading(cell, "FFFFFF")


def add_mentor_review(doc: Document, deck: Dict[str, Any]) -> None:
    mr = deck.get("mentor_review", {})
    add_heading(doc, "Mentor review")

    table = doc.add_table(rows=4, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    rows = [
        ("Mentor", mr.get("mentor_name", ""), "Email", mr.get("mentor_email", "")),
        ("Status", mr.get("review_status", ""), "Requested", mr.get("requested_review_date", "")),
        ("Completed", mr.get("review_completed_date", ""), "Include slide", "Yes" if mr.get("include_mentor_review_slide") else "No"),
        ("Approval statement", mr.get("mentor_approval_statement", ""), "Feedback", mr.get("mentor_feedback", "")),
    ]
    for row, values in zip(table.rows, rows):
        for i, value in enumerate(values):
            set_cell_text(row.cells[i], value, bold=i in (0, 2), font_size=8.2)
            set_cell_shading(row.cells[i], GRAY if i in (0, 2) else "FFFFFF")

    checklist = mr.get("checklist", {})
    check_table = doc.add_table(rows=1, cols=2)
    check_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(check_table)
    set_cell_text(check_table.rows[0].cells[0], "Mentor checklist item", bold=True, font_size=8.5)
    set_cell_text(check_table.rows[0].cells[1], "Done", bold=True, font_size=8.5)
    set_cell_shading(check_table.rows[0].cells[0], GREEN)
    set_cell_shading(check_table.rows[0].cells[1], GREEN)
    for key, label in CHECKLIST_LABELS.items():
        cells = check_table.add_row().cells
        set_cell_text(cells[0], label, font_size=8.2)
        set_cell_text(cells[1], "☑" if checklist.get(key) else "☐", font_size=8.2)


def add_archive_check(doc: Document, deck: Dict[str, Any]) -> None:
    add_heading(doc, "Final archive checklist")
    items = [
        "Download final PowerPoint.",
        "Download this planning form if needed for local review/records.",
        "Save draft JSON so the session can be revised later.",
        "Save PowerPoint, DOCX planning form, and JSON draft to GitHub archive.",
        "Confirm speaker notes appear in PowerPoint presenter view.",
    ]
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(item)
        run.font.name = "Aptos"
        run.font.size = Pt(9)


def build_docx(deck: Dict[str, Any]) -> bytes:
    """Build the Word planning form and return DOCX bytes."""
    doc = Document()
    configure_document(doc)
    add_title_block(doc, deck)
    add_objectives(doc, deck)
    add_storyboard(doc, deck)
    add_mentor_review(doc, deck)
    add_archive_check(doc, deck)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()
