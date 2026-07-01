"""
Pediatric Residency Presentation Builder
---------------------------------------
Streamlit front end for a standardized, story-driven PowerPoint builder.

Design choices:
- app.py is the UI only.
- deck_model.py holds schema/defaults/helpers.
- pptx_builder.py creates the PPTX and injects real speaker notes.
- docx_builder.py creates the mentor review Word document.
- github_storage.py saves/loads the archive.

There are no local JSON draft uploads/downloads. GitHub is the source of truth.
"""

from __future__ import annotations

import base64
import html
import re
from typing import Any, Dict, List

import streamlit as st

from deck_model import (
    APP_TITLE,
    ARCHIVE_DOCX_NAME,
    ARCHIVE_PPTX_NAME,
    BLOOM_HELPER,
    OBJECTIVE_EXAMPLES,
    SLIDE_ROLES,
    TALK_TYPES,
    default_deck,
    identity_subtitle,
    identity_title,
    new_slide,
    normalize_loaded_deck,
    short_label,
    split_nonempty_lines,
)
from docx_builder import build_mentor_review_docx
from github_storage import (
    GitHubStorageError,
    github_is_configured,
    github_status_message,
    list_archives_from_github,
    load_json_from_github,
    save_archive_to_github,
)
from pptx_builder import build_pptx


# -----------------------------------------------------------------------------
# Utility helpers
# -----------------------------------------------------------------------------


def count_words(text: Any) -> int:
    return len(re.findall(r"\b\w+\b", str(text or "")))


def clear_widget_state() -> None:
    for key in list(st.session_state.keys()):
        if key.startswith("widget__"):
            del st.session_state[key]


def initialize_state() -> None:
    if "deck" not in st.session_state:
        st.session_state.deck = default_deck()
    if "selected_slide_id" not in st.session_state:
        st.session_state.selected_slide_id = st.session_state.deck["slides"][0]["id"]
    if "selected_slide_radio" not in st.session_state:
        st.session_state.selected_slide_radio = st.session_state.selected_slide_id
    if "visual_uploader_nonce" not in st.session_state:
        st.session_state.visual_uploader_nonce = {}
    if "archive_path" not in st.session_state:
        st.session_state.archive_path = ""
    if "archive_results" not in st.session_state:
        st.session_state.archive_results = []


def get_selected_slide(deck: Dict[str, Any]) -> Dict[str, Any]:
    slide_id = st.session_state.selected_slide_id
    for slide in deck.get("slides", []):
        if slide.get("id") == slide_id:
            return slide
    st.session_state.selected_slide_id = deck["slides"][0]["id"]
    return deck["slides"][0]


def sync_selected_slide_from_radio() -> None:
    """Keep sidebar navigation single-click responsive.

    The radio widget stores a stable slide ID, while format_func renders the
    current human-readable label. This prevents the old double-click behavior
    that can happen when radio options are dynamic labels and the user edits a
    slide title.
    """
    selected = st.session_state.get("selected_slide_radio")
    slide_ids = [slide.get("id") for slide in st.session_state.deck.get("slides", [])]
    if selected in slide_ids:
        st.session_state.selected_slide_id = selected


def get_visual_image(slide: Dict[str, Any]) -> Dict[str, str]:
    image = slide.get("visual_image", {})
    return image if isinstance(image, dict) else {}


def visual_image_bytes(slide: Dict[str, Any]) -> bytes | None:
    image = get_visual_image(slide)
    encoded = image.get("data_base64")
    if not encoded:
        return None
    try:
        return base64.b64decode(encoded)
    except Exception:
        return None


def has_uploaded_visual(slide: Dict[str, Any]) -> bool:
    return visual_image_bytes(slide) is not None


def slide_nav_label(index: int, slide: Dict[str, Any]) -> str:
    role = slide.get("role") or "Slide"
    title = slide.get("title") or "Untitled"
    return f"{index}. {short_label(role, 16)} — {short_label(title, 30)}"


def validation_messages(deck: Dict[str, Any]) -> List[str]:
    messages: List[str] = []
    meta = deck.get("metadata", {})
    if not str(meta.get("presentation_title", "")).strip():
        messages.append("Presentation title is blank.")
    if not str(meta.get("presenter", "")).strip():
        messages.append("Presenter is blank.")
    for idx, slide in enumerate(deck.get("slides", []), start=1):
        role = slide.get("role", "Slide")
        if role == "Objectives":
            objective_count = len(split_nonempty_lines(slide.get("body", "")))
            if objective_count < 1:
                messages.append(f"Slide {idx} objectives are blank.")
        elif role != "Title" and not str(slide.get("body", "")).strip() and not has_uploaded_visual(slide):
            messages.append(f"Slide {idx} has no main slide text or uploaded visual.")
    return messages


# -----------------------------------------------------------------------------
# Styling
# -----------------------------------------------------------------------------


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .identity-strip {
            border: 1px solid #d9e6f2;
            background: #f7fbff;
            border-radius: 12px;
            padding: 0.6rem 0.85rem;
            margin: 0.4rem 0 0.9rem 0;
        }
        .identity-title { font-weight: 700; font-size: 1.02rem; }
        .identity-subtitle { color: #5c6b78; font-size: 0.88rem; margin-top: 0.12rem; }
        .helper-box {
            border-left: 4px solid #1f4e79;
            background: #f7fbff;
            padding: 0.7rem 0.9rem;
            border-radius: 8px;
            margin: 0.55rem 0 0.8rem 0;
        }
        .export-card {
            border: 1px solid #d9e6f2;
            background: #ffffff;
            border-radius: 12px;
            padding: 0.85rem;
            min-height: 10.5rem;
        }
        .small-muted { color: #5c6b78; font-size: 0.85rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_identity_strip(deck: Dict[str, Any]) -> None:
    st.markdown(
        f"""
        <div class="identity-strip">
            <div class="identity-title">{html.escape(identity_title(deck))}</div>
            <div class="identity-subtitle">{html.escape(identity_subtitle(deck))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_bloom_helper() -> None:
    with st.expander("Bloom’s taxonomy helper", expanded=False):
        cols = st.columns(3)
        for idx, (level, verbs) in enumerate(BLOOM_HELPER.items()):
            with cols[idx % 3]:
                st.caption(f"**{level}:** {verbs}")
        st.markdown("**Examples**")
        for example in OBJECTIVE_EXAMPLES:
            st.write(f"• {example}")


# -----------------------------------------------------------------------------
# Sidebar
# -----------------------------------------------------------------------------


def render_sidebar(deck: Dict[str, Any]) -> None:
    with st.sidebar:
        st.header("Slides")
        slide_ids = [slide["id"] for slide in deck["slides"]]
        id_to_label = {slide["id"]: slide_nav_label(i + 1, slide) for i, slide in enumerate(deck["slides"])}

        if st.session_state.selected_slide_id not in slide_ids:
            st.session_state.selected_slide_id = slide_ids[0]
        if st.session_state.get("selected_slide_radio") not in slide_ids:
            st.session_state.selected_slide_radio = st.session_state.selected_slide_id

        current_index = slide_ids.index(st.session_state.selected_slide_id)
        st.radio(
            "Choose slide",
            slide_ids,
            index=current_index,
            format_func=lambda sid: id_to_label.get(sid, "Slide"),
            key="selected_slide_radio",
            on_change=sync_selected_slide_from_radio,
            label_visibility="collapsed",
        )
        st.session_state.selected_slide_id = st.session_state.selected_slide_radio

        st.caption("All slides export to PowerPoint automatically.")
        st.divider()

        st.subheader("Add slides")
        new_role = st.selectbox("New slide role", SLIDE_ROLES, index=SLIDE_ROLES.index("Custom / Unknown title"))
        new_title = st.text_input("New slide title", placeholder="Leave blank if you do not know it yet")
        new_prompt = st.text_area("Optional helper prompt", height=75, placeholder="What should this slide help the presenter do?")

        selected_index = next((i for i, slide in enumerate(deck["slides"]) if slide["id"] == st.session_state.selected_slide_id), 0)
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Add after", use_container_width=True):
                slide = new_slide(new_role, new_title, new_prompt)
                deck["slides"].insert(selected_index + 1, slide)
                st.session_state.selected_slide_id = slide["id"]
                st.session_state.selected_slide_radio = slide["id"]
                st.rerun()
        with col2:
            if st.button("Add at end", use_container_width=True):
                slide = new_slide(new_role, new_title, new_prompt)
                deck["slides"].append(slide)
                st.session_state.selected_slide_id = slide["id"]
                st.session_state.selected_slide_radio = slide["id"]
                st.rerun()

        st.divider()
        st.subheader("GitHub archive")
        if github_is_configured():
            st.success(github_status_message())
        else:
            st.info(github_status_message())

        search_text = st.text_input("Search archive", placeholder="presenter, title, or date")
        if st.button("Find saved presentations", use_container_width=True):
            try:
                st.session_state.archive_results = list_archives_from_github(search_text)
                if not st.session_state.archive_results:
                    st.info("No matching saved presentations found.")
            except GitHubStorageError as exc:
                st.error(str(exc))

        results = st.session_state.get("archive_results", [])
        if results:
            label_to_path = {row["name"]: row["path"] for row in results}
            selected_archive = st.selectbox("Saved presentations", list(label_to_path.keys()))
            if st.button("Load selected", use_container_width=True):
                try:
                    payload = load_json_from_github(label_to_path[selected_archive])
                    st.session_state.deck = normalize_loaded_deck(payload)
                    st.session_state.archive_path = payload.get("archive_path", label_to_path[selected_archive])
                    st.session_state.selected_slide_id = st.session_state.deck["slides"][0]["id"]
                    st.session_state.selected_slide_radio = st.session_state.selected_slide_id
                    clear_widget_state()
                    st.success("Loaded from GitHub.")
                    st.rerun()
                except GitHubStorageError as exc:
                    st.error(str(exc))

        st.divider()
        if st.button("Start blank presentation", use_container_width=True):
            st.session_state.deck = default_deck()
            st.session_state.selected_slide_id = st.session_state.deck["slides"][0]["id"]
            st.session_state.selected_slide_radio = st.session_state.selected_slide_id
            st.session_state.archive_path = ""
            clear_widget_state()
            st.rerun()


# -----------------------------------------------------------------------------
# Editors
# -----------------------------------------------------------------------------


def widget_text(slide: Dict[str, Any], field: str, label: str, *, height: int = 120, help_text: str = "", multiline: bool = False) -> str:
    key = f"widget__{slide['id']}__{field}"
    if key not in st.session_state:
        st.session_state[key] = slide.get(field, "")
    if multiline:
        value = st.text_area(label, key=key, height=height, help=help_text)
    else:
        value = st.text_input(label, key=key, help=help_text)
    slide[field] = value
    return value


def render_title_editor(deck: Dict[str, Any], slide: Dict[str, Any]) -> None:
    meta = deck["metadata"]
    st.markdown("### Title slide")
    st.caption("These fields appear on the exported title slide. They are not repeated as a footer on the rest of the PowerPoint.")

    col1, col2 = st.columns([1.55, 1])
    with col1:
        meta["presentation_title"] = st.text_input("Presentation title", meta.get("presentation_title", ""), placeholder="Untitled presentation")
        meta["presenter"] = st.text_input("Presenter", meta.get("presenter", ""), placeholder="Presenter not entered")
        meta["audience"] = st.text_input("Audience", meta.get("audience", "Pediatric residents"))
    with col2:
        meta["session_date"] = st.text_input("Session date", meta.get("session_date", ""))
        current_type = meta.get("presentation_type", "Educational Topic")
        meta["presentation_type"] = st.selectbox(
            "Presentation type",
            TALK_TYPES,
            index=TALK_TYPES.index(current_type) if current_type in TALK_TYPES else 0,
        )
        if st.button("Replace story scaffold with this type", use_container_width=True):
            deck["slides"] = default_deck(meta["presentation_type"])["slides"]
            st.session_state.selected_slide_id = deck["slides"][0]["id"]
            st.session_state.selected_slide_radio = st.session_state.selected_slide_id
            clear_widget_state()
            st.rerun()

    meta["core_question"] = st.text_area(
        "Core question / tension",
        meta.get("core_question", ""),
        height=80,
        placeholder="What question should the audience be able to answer by the end?",
    )
    meta["story_arc"] = st.text_area(
        "Story arc",
        meta.get("story_arc", ""),
        height=80,
        placeholder="Beginning: why this matters → Middle: what we learn → End: how thinking/practice changes",
    )
    meta["archive_notes"] = st.text_area("Internal archive notes", meta.get("archive_notes", ""), height=70)

    slide["title"] = meta.get("presentation_title", "")
    slide["speaker_notes"] = widget_text(slide, "speaker_notes", "Speaker notes for title slide", height=90, multiline=True)


def render_objectives_editor(slide: Dict[str, Any]) -> None:
    st.markdown("### Objectives")
    render_bloom_helper()
    st.caption("Use measurable verbs. Example: describe, differentiate, appraise, apply, create.")
    widget_text(slide, "title", "Slide title", help_text="Usually 'Objectives'.")
    body = widget_text(slide, "body", "Objectives", height=150, multiline=True, help_text="One objective per line.")
    lines = split_nonempty_lines(body)
    st.caption(f"{len(lines)} objective(s). Aim for 2–4.")
    widget_text(slide, "speaker_notes", "Speaker notes", height=120, multiline=True)


def render_disclosures_editor(slide: Dict[str, Any]) -> None:
    st.markdown("### Disclosures")
    widget_text(slide, "title", "Slide title")
    widget_text(slide, "body", "Disclosure text", height=120, multiline=True)
    widget_text(slide, "speaker_notes", "Speaker notes", height=100, multiline=True)


def move_slide(deck: Dict[str, Any], slide: Dict[str, Any], direction: int) -> None:
    index = deck["slides"].index(slide)
    new_index = index + direction
    if new_index < 0 or new_index >= len(deck["slides"]):
        return
    deck["slides"][index], deck["slides"][new_index] = deck["slides"][new_index], deck["slides"][index]


def duplicate_slide(deck: Dict[str, Any], slide: Dict[str, Any]) -> None:
    index = deck["slides"].index(slide)
    copied = dict(slide)
    copied["id"] = new_slide(slide.get("role", "Story"))["id"]
    copied["required"] = False
    copied["title"] = f"{copied.get('title') or 'Untitled'} copy"
    deck["slides"].insert(index + 1, copied)
    st.session_state.selected_slide_id = copied["id"]
    st.session_state.selected_slide_radio = copied["id"]


def render_visual_upload(slide: Dict[str, Any]) -> None:
    """Store one optional image per slide and send it to PowerPoint/GitHub."""
    st.caption("Optional: upload a PNG/JPEG visual and it will appear on this PowerPoint slide.")
    nonce_map = st.session_state.setdefault("visual_uploader_nonce", {})
    nonce = nonce_map.get(slide["id"], 0)
    uploaded = st.file_uploader(
        "Upload visual image",
        type=["png", "jpg", "jpeg"],
        key=f"widget__{slide['id']}__visual_file__{nonce}",
        help="Best for screenshots, figures you created, diagrams, or a focused data visual. Keep it under 5 MB so the GitHub draft stays lightweight.",
    )
    if uploaded is not None:
        data = uploaded.getvalue()
        if len(data) > 5 * 1024 * 1024:
            st.error("This image is larger than 5 MB. Please compress or crop it before uploading.")
        else:
            slide["visual_image"] = {
                "filename": uploaded.name,
                "content_type": uploaded.type or "image/png",
                "data_base64": base64.b64encode(data).decode("ascii"),
            }

    image_bytes = visual_image_bytes(slide)
    image_info = get_visual_image(slide)
    if image_bytes:
        st.image(image_bytes, caption=image_info.get("filename", "Uploaded visual"), use_container_width=True)
        if st.button("Remove uploaded visual", key=f"widget__{slide['id']}__remove_visual", use_container_width=True):
            slide["visual_image"] = {}
            nonce_map[slide["id"]] = nonce + 1
            st.rerun()


def render_standard_editor(deck: Dict[str, Any], slide: Dict[str, Any]) -> None:
    slide_index = deck["slides"].index(slide) + 1
    st.markdown(f"### Slide {slide_index}")

    if slide.get("prompt"):
        st.markdown(
            f"<div class='helper-box'><strong>Helper:</strong> {html.escape(slide.get('prompt', ''))}</div>",
            unsafe_allow_html=True,
        )

    action_cols = st.columns([1, 1, 1, 1])
    with action_cols[0]:
        if st.button("Move up", use_container_width=True, disabled=slide_index == 1):
            move_slide(deck, slide, -1)
            st.rerun()
    with action_cols[1]:
        if st.button("Move down", use_container_width=True, disabled=slide_index == len(deck["slides"])):
            move_slide(deck, slide, 1)
            st.rerun()
    with action_cols[2]:
        if st.button("Duplicate", use_container_width=True):
            duplicate_slide(deck, slide)
            st.rerun()
    with action_cols[3]:
        if st.button("Delete", use_container_width=True, disabled=bool(slide.get("required", False))):
            deck["slides"].remove(slide)
            st.session_state.selected_slide_id = deck["slides"][max(0, slide_index - 2)]["id"]
            st.session_state.selected_slide_radio = st.session_state.selected_slide_id
            clear_widget_state()
            st.rerun()

    col1, col2 = st.columns([1, 1])
    with col1:
        current_role = slide.get("role", "Story")
        slide["role"] = st.selectbox(
            "Slide role",
            SLIDE_ROLES,
            index=SLIDE_ROLES.index(current_role) if current_role in SLIDE_ROLES else SLIDE_ROLES.index("Story"),
            help="This helps the app format and label the slide. It does not decide whether the slide exports.",
        )
    with col2:
        widget_text(slide, "title", "Slide title", help_text="Can be blank if you do not know the title yet.")

    widget_text(slide, "subtitle", "Optional subtitle", help_text="Use sparingly. The main title should tell the story.")
    body = widget_text(slide, "body", "Slide text", height=190, multiline=True, help_text="Use one idea per line. Short lines work best on slides.")
    st.caption(f"{count_words(body)} words. For readability, try to keep most slides under ~45 words.")

    col3, col4 = st.columns(2)
    with col3:
        widget_text(slide, "visual_plan", "Visual / evidence plan", height=125, multiline=True, help_text="Describe a figure, table, image, graph, or data point to include.")
        render_visual_upload(slide)
    with col4:
        widget_text(slide, "discussion_prompt", "Discussion prompt", height=125, multiline=True, help_text="Question to ask the audience, if useful.")

    widget_text(slide, "speaker_notes", "Speaker notes exported into PowerPoint", height=170, multiline=True)


def render_slide_editor(deck: Dict[str, Any]) -> None:
    slide = get_selected_slide(deck)
    role = slide.get("role")
    kind = slide.get("slide_kind")
    if role == "Title" or kind == "title":
        render_title_editor(deck, slide)
    elif role == "Objectives" or kind == "objectives":
        render_objectives_editor(slide)
    elif role == "Disclosures" or kind == "disclosures":
        render_disclosures_editor(slide)
    else:
        render_standard_editor(deck, slide)


# -----------------------------------------------------------------------------
# Export and archive panel
# -----------------------------------------------------------------------------


def render_export_panel(deck: Dict[str, Any]) -> None:
    """Render stacked export/archive controls in the right-side panel."""
    try:
        pptx_bytes = build_pptx(deck)
        mentor_docx_bytes = build_mentor_review_docx(deck)
    except Exception as exc:
        st.error(f"Could not build exports: {exc}")
        return

    st.markdown("### Export / archive")

    with st.container(border=True):
        st.markdown("#### Mentor Word document")
        st.caption("Give this to the mentor for comments or Track Changes. Critiques are not stored in the app.")
        st.download_button(
            "Download mentor DOCX",
            data=mentor_docx_bytes,
            file_name=ARCHIVE_DOCX_NAME,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with st.container(border=True):
        st.markdown("#### PowerPoint")
        st.caption("All slides export automatically. Speaker notes go into real PowerPoint notes.")
        st.download_button(
            "Download PPTX",
            data=pptx_bytes,
            file_name=ARCHIVE_PPTX_NAME,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    with st.container(border=True):
        st.markdown("#### GitHub archive")
        st.caption("Saves draft.json, presentation.pptx, and mentor_review.docx to GitHub.")
        if st.button("Save all to GitHub", use_container_width=True):
            try:
                results = save_archive_to_github(deck, pptx_bytes, mentor_docx_bytes, st.session_state.get("archive_path", ""))
                if results:
                    # Path looks like base/date_presenter_title/file.ext; archive folder is the parent.
                    st.session_state.archive_path = results[0].path.rsplit("/", 1)[0]
                st.success("Saved to GitHub archive.")
                for result in results:
                    if result.html_url:
                        st.caption(result.path)
            except GitHubStorageError as exc:
                st.error(str(exc))


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🩺", layout="wide")
    initialize_state()
    inject_css()

    deck = st.session_state.deck
    render_sidebar(deck)

    st.title(APP_TITLE)
    render_identity_strip(deck)

    problems = validation_messages(deck)
    if problems:
        with st.expander(f"Readiness check: {len(problems)} item(s) to review", expanded=False):
            for problem in problems:
                st.write(f"• {problem}")
    else:
        st.success("Readiness check passed. The presentation has the core fields needed for export.")

    editor_col, export_col = st.columns([2.1, 0.85], gap="large")
    with editor_col:
        render_slide_editor(deck)

    with export_col:
        render_export_panel(deck)


if __name__ == "__main__":
    main()
