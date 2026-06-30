"""
PowerPoint builder for the Presentation Builder app.

This module is deliberately separate from app.py so the Streamlit UI stays clean.
It creates the PPTX and then post-processes the Office Open XML package to add
real PowerPoint speaker notes, because python-pptx does not currently expose a
speaker-notes authoring API.
"""

from __future__ import annotations

import html
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from deck_model import OBJECTIVE_EXAMPLES, slide_output_title, split_nonempty_lines

# Office XML namespaces.
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
ET.register_namespace("", REL_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("a", A_NS)

TITLE_BLUE = (31, 78, 121)
LIGHT_BLUE = (234, 242, 250)
PALE_BLUE = (242, 246, 250)
GRAY_TEXT = (65, 65, 65)
BODY_TEXT = (35, 35, 35)
BORDER_BLUE = (185, 205, 225)


def _rgb(color: Tuple[int, int, int]) -> RGBColor:
    return RGBColor(color[0], color[1], color[2])


def add_textbox(slide, text: str, x, y, w, h, font_size: int = 22, bold: bool = False, color: Tuple[int, int, int] = BODY_TEXT):
    box = slide.shapes.add_textbox(x, y, w, h)
    frame = box.text_frame
    frame.word_wrap = True
    frame.clear()
    p = frame.paragraphs[0]
    run = p.add_run()
    run.text = text or ""
    run.font.name = "Aptos"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    return box


def add_title_bar(slide, title: str, subtitle: str = "") -> None:
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.72))
    bar.fill.solid()
    bar.fill.fore_color.rgb = _rgb(TITLE_BLUE)
    bar.line.color.rgb = _rgb(TITLE_BLUE)

    title_box = slide.shapes.add_textbox(Inches(0.45), Inches(0.13), Inches(12.3), Inches(0.46))
    tf = title_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.name = "Aptos Display"
    run.font.size = Pt(23)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

    if subtitle:
        add_textbox(slide, subtitle, Inches(0.55), Inches(0.78), Inches(12.2), Inches(0.35), 11, False, (85, 85, 85))


def add_body_lines(slide, lines: List[str], x, y, w, h, font_size: int = 22) -> None:
    box = slide.shapes.add_textbox(x, y, w, h)
    frame = box.text_frame
    frame.word_wrap = True
    frame.clear()
    if not lines:
        lines = ["Add slide content here."]
    for idx, line in enumerate(lines):
        p = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        p.text = line
        p.level = 0
        p.font.name = "Aptos"
        p.font.size = Pt(font_size)
        p.space_after = Pt(7)


def add_footer(slide, deck: Dict[str, Any]) -> None:
    meta = deck.get("metadata", {})
    footer = " · ".join(part for part in [meta.get("presenter", ""), meta.get("session_date", ""), meta.get("audience", "")] if part)
    if footer:
        add_textbox(slide, footer, Inches(0.55), Inches(7.05), Inches(12.0), Inches(0.24), 9, False, (110, 110, 110))


def render_title_slide(prs: Presentation, deck: Dict[str, Any], slide_data: Dict[str, Any]) -> None:
    meta = deck.get("metadata", {})
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(248, 250, 252)

    title = meta.get("presentation_title") or slide_data.get("title") or "Untitled Presentation"
    presenter = meta.get("presenter", "")
    date = meta.get("session_date", "")
    audience = meta.get("audience", "")
    talk_type = meta.get("presentation_type", "")

    add_textbox(slide, title, Inches(0.75), Inches(1.22), Inches(11.8), Inches(1.3), 34, True, TITLE_BLUE)
    info_lines = [line for line in [presenter, date, audience, talk_type] if line]
    add_textbox(slide, "\n".join(info_lines), Inches(0.8), Inches(2.9), Inches(8.8), Inches(1.1), 18, False, GRAY_TEXT)

    core_question = meta.get("core_question", "")
    if core_question:
        panel = slide.shapes.add_shape(1, Inches(0.8), Inches(4.5), Inches(11.6), Inches(1.08))
        panel.fill.solid()
        panel.fill.fore_color.rgb = _rgb(LIGHT_BLUE)
        panel.line.color.rgb = _rgb(BORDER_BLUE)
        add_textbox(slide, "Core question", Inches(1.05), Inches(4.65), Inches(2.3), Inches(0.28), 13, True, TITLE_BLUE)
        add_textbox(slide, core_question, Inches(1.05), Inches(4.97), Inches(10.9), Inches(0.38), 17, False, (45, 45, 45))


def render_objectives_slide(prs: Presentation, deck: Dict[str, Any], slide_data: Dict[str, Any], index: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title_bar(slide, slide_output_title(deck, slide_data, index))
    objectives = split_nonempty_lines(slide_data.get("body", "")) or OBJECTIVE_EXAMPLES[:3]
    add_body_lines(slide, objectives, Inches(0.85), Inches(1.35), Inches(11.5), Inches(4.9), 24)
    add_footer(slide, deck)


def render_standard_slide(prs: Presentation, deck: Dict[str, Any], slide_data: Dict[str, Any], index: int) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_output_title(deck, slide_data, index)
    subtitle = slide_data.get("subtitle", "")
    add_title_bar(slide, title, subtitle)

    body_lines = split_nonempty_lines(slide_data.get("body", ""))
    visual_plan = (slide_data.get("visual_plan") or "").strip()
    discussion = (slide_data.get("discussion_prompt") or "").strip()

    if visual_plan or discussion:
        add_body_lines(slide, body_lines, Inches(0.75), Inches(1.25), Inches(7.65), Inches(4.95), 22)
        panel = slide.shapes.add_shape(1, Inches(8.75), Inches(1.2), Inches(3.85), Inches(4.95))
        panel.fill.solid()
        panel.fill.fore_color.rgb = _rgb(PALE_BLUE)
        panel.line.color.rgb = _rgb(BORDER_BLUE)
        y = 1.45
        if visual_plan:
            add_textbox(slide, "Visual / evidence plan", Inches(9.05), Inches(y), Inches(3.3), Inches(0.28), 13, True, TITLE_BLUE)
            add_textbox(slide, visual_plan, Inches(9.05), Inches(y + 0.34), Inches(3.3), Inches(1.65), 13, False, (40, 40, 40))
            y += 2.25
        if discussion:
            add_textbox(slide, "Audience prompt", Inches(9.05), Inches(y), Inches(3.3), Inches(0.28), 13, True, TITLE_BLUE)
            add_textbox(slide, discussion, Inches(9.05), Inches(y + 0.34), Inches(3.3), Inches(1.45), 13, False, (40, 40, 40))
    else:
        add_body_lines(slide, body_lines, Inches(0.85), Inches(1.3), Inches(11.6), Inches(4.95), 23)

    add_footer(slide, deck)


def render_mentor_review_slide(prs: Presentation, deck: Dict[str, Any]) -> None:
    mr = deck.get("mentor_review", {})
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title_bar(slide, "Mentor review")

    lines = [
        f"Mentor: {mr.get('mentor_name') or 'Not listed'}",
        f"Review status: {mr.get('review_status') or 'Not sent'}",
    ]
    if mr.get("review_completed_date"):
        lines.append(f"Completed: {mr['review_completed_date']}")
    if mr.get("mentor_approval_statement"):
        lines.append(f"Approval statement: {mr['mentor_approval_statement']}")

    feedback = split_nonempty_lines(mr.get("mentor_feedback", ""))
    if feedback:
        lines.append("")
        lines.append("Mentor feedback summary:")
        lines.extend(feedback[:5])

    add_body_lines(slide, lines, Inches(0.85), Inches(1.3), Inches(11.7), Inches(5.0), 20)
    add_footer(slide, deck)


def build_pptx(deck: Dict[str, Any]) -> bytes:
    """Build the presentation and return PPTX bytes."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    speaker_notes: List[str] = []
    output_index = 0
    for slide_data in deck.get("slides", []):
        if not slide_data.get("include", True):
            continue
        output_index += 1
        role = slide_data.get("role")
        if role == "Title":
            render_title_slide(prs, deck, slide_data)
        elif role == "Objectives":
            render_objectives_slide(prs, deck, slide_data, output_index)
        else:
            render_standard_slide(prs, deck, slide_data, output_index)
        speaker_notes.append(slide_data.get("speaker_notes", ""))

    if deck.get("mentor_review", {}).get("include_mentor_review_slide"):
        render_mentor_review_slide(prs, deck)
        speaker_notes.append("Internal mentor-review slide. Remove before final presentation if not intended for learners.")

    buffer = io.BytesIO()
    prs.save(buffer)
    return add_speaker_notes_to_pptx(buffer.getvalue(), speaker_notes)


# -----------------------------------------------------------------------------
# Speaker-notes Office XML injection
# -----------------------------------------------------------------------------


def _next_rid(root: ET.Element) -> str:
    nums: List[int] = []
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.get("Id", "")
        match = re.match(r"rId(\d+)$", rid)
        if match:
            nums.append(int(match.group(1)))
    return f"rId{max(nums or [0]) + 1}"


def _add_content_type_override(content_types_xml: bytes, part_name: str, content_type: str) -> bytes:
    ET.register_namespace("", CT_NS)
    root = ET.fromstring(content_types_xml)
    for override in root.findall(f"{{{CT_NS}}}Override"):
        if override.get("PartName") == part_name:
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    ET.SubElement(root, f"{{{CT_NS}}}Override", PartName=part_name, ContentType=content_type)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _notes_master_xml() -> bytes:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesMaster xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">
  <p:cSld>
    <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="914400"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Speaker notes</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
</p:notesMaster>'''.encode("utf-8")


def _notes_slide_xml(notes: str, slide_number: int) -> bytes:
    lines = notes.splitlines() or [""]
    paragraphs = []
    for line in lines:
        escaped = html.escape(line, quote=False)
        paragraphs.append(
            f'<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>{escaped}</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
        )
    paragraphs_xml = "".join(paragraphs)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs_xml}</p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{{F7021451-1387-4CA6-816F-3879F97B5CBC}}" type="slidenum"><a:rPr lang="en-US"/><a:t>{slide_number}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>'''.encode("utf-8")


def _notes_slide_rels_xml(slide_number: int) -> bytes:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{slide_number}.xml"/>
</Relationships>'''.encode("utf-8")


def add_speaker_notes_to_pptx(pptx_bytes: bytes, notes_by_slide: List[str]) -> bytes:
    """Add real PowerPoint speaker notes to a python-pptx-generated deck."""
    if not any((note or "").strip() for note in notes_by_slide):
        return pptx_bytes

    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    files["[Content_Types].xml"] = _add_content_type_override(
        files["[Content_Types].xml"],
        "/ppt/notesMasters/notesMaster1.xml",
        "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
    )

    presentation_rels_path = "ppt/_rels/presentation.xml.rels"
    pres_rels = ET.fromstring(files[presentation_rels_path])
    notes_master_rid = None
    for rel in pres_rels.findall(f"{{{REL_NS}}}Relationship"):
        if rel.get("Type", "").endswith("/notesMaster"):
            notes_master_rid = rel.get("Id")
            break
    if not notes_master_rid:
        notes_master_rid = _next_rid(pres_rels)
        ET.SubElement(
            pres_rels,
            f"{{{REL_NS}}}Relationship",
            Id=notes_master_rid,
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster",
            Target="notesMasters/notesMaster1.xml",
        )
        files[presentation_rels_path] = ET.tostring(pres_rels, encoding="utf-8", xml_declaration=True)

    presentation_path = "ppt/presentation.xml"
    pres = ET.fromstring(files[presentation_path])
    ns = {"p": P_NS, "r": R_NS}
    if pres.find("p:notesMasterIdLst", ns) is None:
        notes_master_id_list = ET.Element(f"{{{P_NS}}}notesMasterIdLst")
        notes_master_id = ET.SubElement(notes_master_id_list, f"{{{P_NS}}}notesMasterId")
        notes_master_id.set(f"{{{R_NS}}}id", notes_master_rid)
        insert_at = len(list(pres))
        for i, child in enumerate(list(pres)):
            if child.tag == f"{{{P_NS}}}sldSz":
                insert_at = i
                break
        pres.insert(insert_at, notes_master_id_list)
        files[presentation_path] = ET.tostring(pres, encoding="utf-8", xml_declaration=True)

    files["ppt/notesMasters/notesMaster1.xml"] = _notes_master_xml()
    files["ppt/notesMasters/_rels/notesMaster1.xml.rels"] = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>'''.encode("utf-8")

    for slide_number, note in enumerate(notes_by_slide, start=1):
        if not (note or "").strip():
            continue
        slide_rels_path = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
        if slide_rels_path not in files:
            continue
        slide_rels = ET.fromstring(files[slide_rels_path])
        already_linked = any(rel.get("Type", "").endswith("/notesSlide") for rel in slide_rels.findall(f"{{{REL_NS}}}Relationship"))
        if not already_linked:
            rid = _next_rid(slide_rels)
            ET.SubElement(
                slide_rels,
                f"{{{REL_NS}}}Relationship",
                Id=rid,
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
                Target=f"../notesSlides/notesSlide{slide_number}.xml",
            )
            files[slide_rels_path] = ET.tostring(slide_rels, encoding="utf-8", xml_declaration=True)

        files[f"ppt/notesSlides/notesSlide{slide_number}.xml"] = _notes_slide_xml(note, slide_number)
        files[f"ppt/notesSlides/_rels/notesSlide{slide_number}.xml.rels"] = _notes_slide_rels_xml(slide_number)
        files["[Content_Types].xml"] = _add_content_type_override(
            files["[Content_Types].xml"],
            f"/ppt/notesSlides/notesSlide{slide_number}.xml",
            "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
        )

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    return output.getvalue()
