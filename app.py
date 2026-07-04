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
    APP_VERSION,
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
from docx_builder import build_mentor_review_docx, mentor_docx_contains_complete_review_fields
from github_storage import (
    GitHubStorageError,
    github_is_configured,
    github_status_message,
    list_archives_from_github,
    load_json_from_github,
    save_archive_to_github,
    delete_archive_from_github,
)
from pptx_builder import build_pptx
from preview_utils import render_pptx_first_slide_to_png


def objective_verb_options() -> List[str]:
    """Build Bloom verb options locally so app.py still works if deck_model is stale."""
    verbs: List[str] = []
    for raw_verbs in BLOOM_HELPER.values():
        for raw_verb in str(raw_verbs).split(","):
            verb = raw_verb.strip().title()
            if verb and verb not in verbs:
                verbs.append(verb)
    for preferred in ["Formulate", "Appraise", "Apply"]:
        if preferred not in verbs:
            verbs.append(preferred)
    return verbs


OBJECTIVE_VERB_OPTIONS = objective_verb_options()


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


def queue_slide_selection(slide_id: str) -> None:
    """Select a slide on the next rerun without mutating the radio widget state.

    Streamlit raises an exception if code changes st.session_state for a widget
    key after that widget has already been rendered in the same run. The sidebar
    radio uses the key selected_slide_radio, so buttons such as Add after,
    Duplicate, Delete, Load, and Start blank store the requested selection here
    and the sidebar applies it before the radio is rendered on the next run.
    """
    st.session_state.selected_slide_id = slide_id
    st.session_state.pending_selected_slide_id = slide_id


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


def get_uploaded_slide_pptx(slide: Dict[str, Any]) -> Dict[str, str]:
    pptx = slide.get("uploaded_slide_pptx", {})
    return pptx if isinstance(pptx, dict) else {}


def uploaded_slide_pptx_bytes(slide: Dict[str, Any]) -> bytes | None:
    pptx = get_uploaded_slide_pptx(slide)
    encoded = pptx.get("data_base64")
    if not encoded:
        return None
    try:
        return base64.b64decode(encoded)
    except Exception:
        return None


def get_uploaded_slide_preview_image(slide: Dict[str, Any]) -> Dict[str, str]:
    image = slide.get("uploaded_slide_preview_image", {})
    return image if isinstance(image, dict) else {}


def uploaded_slide_preview_bytes(slide: Dict[str, Any]) -> bytes | None:
    image = get_uploaded_slide_preview_image(slide)
    encoded = image.get("data_base64")
    if not encoded:
        return None
    try:
        return base64.b64decode(encoded)
    except Exception:
        return None


def ensure_uploaded_slide_preview(slide: Dict[str, Any]) -> bytes | None:
    preview = uploaded_slide_preview_bytes(slide)
    if preview:
        return preview
    pptx_bytes = uploaded_slide_pptx_bytes(slide)
    if not pptx_bytes:
        return None
    preview = render_pptx_first_slide_to_png(pptx_bytes)
    if preview:
        filename = get_uploaded_slide_pptx(slide).get("filename", "uploaded_slide.pptx")
        stem = filename.rsplit('.', 1)[0]
        slide["uploaded_slide_preview_image"] = {
            "filename": f"{stem}_preview.png",
            "content_type": "image/png",
            "data_base64": base64.b64encode(preview).decode("ascii"),
        }
    return preview


def count_pptx_slides(pptx_bytes: bytes | None) -> int:
    if not pptx_bytes:
        return 0
    try:
        import zipfile
        import xml.etree.ElementTree as ET
        from io import BytesIO
        with zipfile.ZipFile(BytesIO(pptx_bytes), "r") as zf:
            root = ET.fromstring(zf.read("ppt/presentation.xml"))
        ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
        return len(root.findall("p:sldIdLst/p:sldId", ns))
    except Exception:
        return 0


def has_uploaded_visual(slide: Dict[str, Any]) -> bool:
    return visual_image_bytes(slide) is not None or uploaded_slide_pptx_bytes(slide) is not None


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
        elif role == "Take-home":
            takehome_count = len([1 for i in range(1, 6) if str(slide.get(f"takehome_point_{i}", "")).strip()])
            if takehome_count < 1 and not has_uploaded_visual(slide):
                messages.append(f"Slide {idx} take-home points are blank.")
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


def ensure_objective_fields(slide: Dict[str, Any]) -> None:
    """Backfill structured objective fields for newer objective-slide layouts."""
    lines = split_nonempty_lines(slide.get("body", ""))
    defaults = [
        ("Formulate", "Turn a bedside uncertainty into a focused clinical question."),
        ("Appraise", "Judge whether the evidence is valid, important, and clinically applicable."),
        ("Apply", "Use evidence with patient values, feasibility, and clinical judgment to make a decision."),
    ]
    if not slide.get("objectives_intro"):
        slide["objectives_intro"] = "By the end of this session, residents should be able to:"
    if not slide.get("objectives_takeaway"):
        slide["objectives_takeaway"] = "Leave with a simple script you can use tomorrow on rounds or in journal club."

    for idx in range(1, 4):
        verb_key = f"objective_{idx}_verb"
        text_key = f"objective_{idx}_text"
        if not slide.get(verb_key):
            slide[verb_key] = defaults[idx - 1][0]
        if not slide.get(text_key):
            line = lines[idx - 1] if idx - 1 < len(lines) else defaults[idx - 1][1]
            if ":" in line:
                maybe_verb, maybe_text = line.split(":", 1)
                maybe_verb = maybe_verb.strip().title()
                maybe_text = maybe_text.strip()
                if maybe_verb in OBJECTIVE_VERB_OPTIONS and maybe_text:
                    slide[verb_key] = maybe_verb
                    line = maybe_text
            slide[text_key] = line


def sync_objectives_body(slide: Dict[str, Any]) -> None:
    ensure_objective_fields(slide)
    summary_lines = []
    for idx in range(1, 4):
        verb = str(slide.get(f"objective_{idx}_verb", "")).strip()
        text = str(slide.get(f"objective_{idx}_text", "")).strip()
        if verb or text:
            summary_lines.append(f"{verb}: {text}".strip(": "))
    slide["body"] = "\n".join(summary_lines)


def ensure_takehome_fields(slide: Dict[str, Any]) -> None:
    defaults = [
        "Start with a real clinical question.",
        "Use PICO to make the question searchable.",
        "Read methods and results before trusting conclusions.",
        "Look for absolute effect, uncertainty, harms, and applicability.",
        "Make a decision, then reassess what happened.",
    ]
    lines = []
    for raw in split_nonempty_lines(slide.get("body", "")):
        cleaned = re.sub(r"^\s*(?:[-•]|\d+[.)])\s*", "", raw).strip()
        lines.append(cleaned)

    for idx in range(1, 6):
        key = f"takehome_point_{idx}"
        if not slide.get(key):
            line = lines[idx - 1] if idx - 1 < len(lines) else defaults[idx - 1]
            slide[key] = line


def sync_takehome_body(slide: Dict[str, Any]) -> None:
    ensure_takehome_fields(slide)
    slide["body"] = "\n".join(
        str(slide.get(f"takehome_point_{idx}", "")).strip()
        for idx in range(1, 6)
        if str(slide.get(f"takehome_point_{idx}", "")).strip()
    )


# -----------------------------------------------------------------------------
# Sidebar
# -----------------------------------------------------------------------------


def render_sidebar(deck: Dict[str, Any]) -> None:
    with st.sidebar:
        st.header("Slides")
        st.caption(f"Running app version: {APP_VERSION}")
        slide_ids = [slide["id"] for slide in deck["slides"]]
        id_to_label = {slide["id"]: slide_nav_label(i + 1, slide) for i, slide in enumerate(deck["slides"])}

        pending_slide_id = st.session_state.pop("pending_selected_slide_id", None)
        if pending_slide_id in slide_ids:
            st.session_state.selected_slide_id = pending_slide_id
            st.session_state.selected_slide_radio = pending_slide_id

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

        with st.expander("Add slides", expanded=False):
            new_role = st.selectbox("New slide role", SLIDE_ROLES, index=SLIDE_ROLES.index("Custom / Unknown title"))
            new_title = st.text_input("New slide title", placeholder="Leave blank if you do not know it yet")
            new_prompt = st.text_area("Optional helper prompt", height=75, placeholder="What should this slide help the presenter do?")

            selected_index = next((i for i, slide in enumerate(deck["slides"]) if slide["id"] == st.session_state.selected_slide_id), 0)
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Add after", use_container_width=True):
                    slide = new_slide(new_role, new_title, new_prompt)
                    deck["slides"].insert(selected_index + 1, slide)
                    queue_slide_selection(slide["id"])
                    st.rerun()
            with col2:
                if st.button("Add at end", use_container_width=True):
                    slide = new_slide(new_role, new_title, new_prompt)
                    deck["slides"].append(slide)
                    queue_slide_selection(slide["id"])
                    st.rerun()

        st.divider()
        with st.expander("GitHub archive", expanded=False):
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
                # Keep labels readable but unique, so duplicate presentation names do not overwrite each other.
                label_counts: Dict[str, int] = {}
                labels: List[str] = []
                for row in results:
                    base_label = row.get("name", row.get("path", "Saved presentation"))
                    label_counts[base_label] = label_counts.get(base_label, 0) + 1
                    if label_counts[base_label] == 1:
                        labels.append(base_label)
                    else:
                        labels.append(f"{base_label} [{row.get('path', '')}]")

                selected_archive = st.selectbox("Saved presentations", labels)
                selected_index = labels.index(selected_archive)
                selected_row = results[selected_index]
                selected_path = selected_row["path"]

                if st.button("Load selected", use_container_width=True):
                    try:
                        payload = load_json_from_github(selected_path)
                        st.session_state.deck = normalize_loaded_deck(payload)
                        st.session_state.archive_path = payload.get("archive_path", selected_path)
                        queue_slide_selection(st.session_state.deck["slides"][0]["id"])
                        clear_widget_state()
                        st.success("Loaded from GitHub.")
                        st.rerun()
                    except GitHubStorageError as exc:
                        st.error(str(exc))

                st.divider()
                with st.expander("Delete selected archive"):
                    st.caption("Deletes the saved draft.json, presentation.pptx, mentor_review.docx, and any other files in that archive folder.")
                    confirm_delete = st.checkbox(
                        "I understand this permanently deletes the selected GitHub archive",
                        key="widget__confirm_delete_archive",
                    )
                    delete_phrase = st.text_input(
                        "Type DELETE to confirm",
                        key="widget__delete_archive_phrase",
                        placeholder="DELETE",
                    )
                    if st.button(
                        "Delete selected from GitHub",
                        use_container_width=True,
                        disabled=not (confirm_delete and delete_phrase.strip().upper() == "DELETE"),
                    ):
                        try:
                            deleted_count = delete_archive_from_github(selected_path)
                            st.session_state.archive_results = [row for row in results if row.get("path") != selected_path]
                            if st.session_state.get("archive_path", "").strip().strip("/") == selected_path.strip().strip("/"):
                                st.session_state.archive_path = ""
                            st.success(f"Deleted {deleted_count} file(s) from GitHub.")
                            st.rerun()
                        except GitHubStorageError as exc:
                            st.error(str(exc))

        st.divider()
        if st.button("Start blank presentation", use_container_width=True):
            st.session_state.deck = default_deck()
            queue_slide_selection(st.session_state.deck["slides"][0]["id"])
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
        meta["presentation_subtitle"] = st.text_input("Presentation subtitle", meta.get("presentation_subtitle", ""), placeholder="Optional subtitle under the title")
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
            queue_slide_selection(deck["slides"][0]["id"])
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

    st.markdown("#### Optional title-slide visual")
    render_visual_upload(slide)

    slide["title"] = meta.get("presentation_title", "")
    slide["speaker_notes"] = widget_text(slide, "speaker_notes", "Speaker notes for title slide", height=90, multiline=True)


def render_objectives_editor(slide: Dict[str, Any]) -> None:
    st.markdown("### Objectives")
    st.success("Objectives card layout is active: Bloom dropdowns → numbered visual objective cards in PowerPoint.")
    ensure_objective_fields(slide)
    render_bloom_helper()
    st.caption("Choose a Bloom-style action word for each objective, then enter the explanatory sentence.")
    widget_text(slide, "title", "Slide title", help_text="Usually 'Objectives'.")
    slide["objectives_intro"] = st.text_input(
        "Intro line above objectives",
        value=slide.get("objectives_intro", "By the end of this session, residents should be able to:"),
        key=f"widget__{slide['id']}__objectives_intro",
    )

    card_cols = st.columns(3)
    for idx, col in enumerate(card_cols, start=1):
        with col:
            st.markdown(f"**Objective {idx}**")
            verb_key = f"objective_{idx}_verb"
            text_key = f"objective_{idx}_text"
            current_verb = slide.get(verb_key, "") or "Apply"
            options = OBJECTIVE_VERB_OPTIONS
            default_index = options.index(current_verb) if current_verb in options else min(len(options) - 1, max(0, options.index("Apply") if "Apply" in options else 0))
            slide[verb_key] = st.selectbox(
                f"Objective {idx} action word",
                options,
                index=default_index,
                key=f"widget__{slide['id']}__{verb_key}",
            )
            slide[text_key] = st.text_area(
                f"Objective {idx} sentence",
                value=slide.get(text_key, ""),
                height=120,
                key=f"widget__{slide['id']}__{text_key}",
                placeholder="Write the teaching objective sentence here.",
            )

    slide["objectives_takeaway"] = st.text_input(
        "Bottom takeaway banner",
        value=slide.get("objectives_takeaway", ""),
        key=f"widget__{slide['id']}__objectives_takeaway",
        placeholder="Optional summary banner at the bottom of the slide.",
    )

    sync_objectives_body(slide)
    objective_count = sum(1 for idx in range(1, 4) if str(slide.get(f"objective_{idx}_text", "")).strip())
    st.caption(f"{objective_count} objective card(s) populated.")

    st.markdown("#### Optional objective-slide visual")
    st.caption("If you upload a visual and check whole-slide mode, the visual can take over the objectives slide in the exported PowerPoint.")
    render_visual_upload(slide)
    widget_text(slide, "speaker_notes", "Speaker notes", height=120, multiline=True)


def render_takehome_editor(slide: Dict[str, Any]) -> None:
    st.markdown("### Take-home points")
    st.caption("Add up to 5 practical take-home points. In the exported PowerPoint they render as numbered points with circle markers. You can also upload an optional image for the right side of the slide.")
    ensure_takehome_fields(slide)
    widget_text(slide, "title", "Slide title", help_text="Usually something like 'Take-home points' or 'What to remember'.")
    widget_text(slide, "subtitle", "Short subtitle / setup line", help_text="Small line under the title to frame the summary.")

    for idx in range(1, 6):
        key = f"takehome_point_{idx}"
        slide[key] = st.text_input(
            f"Point {idx}",
            value=slide.get(key, ""),
            key=f"widget__{slide['id']}__{key}",
            placeholder=f"Enter take-home point {idx}.",
        )

    sync_takehome_body(slide)
    point_count = sum(1 for idx in range(1, 6) if str(slide.get(f"takehome_point_{idx}", "")).strip())
    st.caption(f"{point_count} take-home point(s) populated.")

    st.markdown("#### Optional slide visual")
    st.caption("Upload an image for the right side of the take-home slide. If you check whole-slide mode, the visual can take over the slide body instead.")
    render_visual_upload(slide)
    widget_text(slide, "discussion_prompt", "Optional discussion prompt at bottom", height=90, multiline=True, help_text="Optional closing question or prompt shown near the bottom of the take-home slide.")
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
    queue_slide_selection(copied["id"])


def render_visual_upload(slide: Dict[str, Any]) -> None:
    """Store one optional uploaded asset per slide: image or PPTX slide."""
    st.caption("Optional: upload a PNG/JPEG image or a PPTX file. Images can appear beside text or fill the slide. A PPTX upload imports the first slide as editable PowerPoint objects when possible.")
    nonce_map = st.session_state.setdefault("visual_uploader_nonce", {})
    nonce = nonce_map.get(slide["id"], 0)
    uploaded = st.file_uploader(
        "Upload image or PPTX slide",
        type=["png", "jpg", "jpeg", "pptx"],
        key=f"widget__{slide['id']}__visual_file__{nonce}",
        help="Use an image for a figure, screenshot, or diagram. Use a PPTX when you already built a polished slide and want the first slide imported as editable PowerPoint objects.",
    )
    if uploaded is not None:
        data = uploaded.getvalue()
        is_pptx = (uploaded.type == "application/vnd.openxmlformats-officedocument.presentationml.presentation") or uploaded.name.lower().endswith(".pptx")
        max_mb = 15 if is_pptx else 5
        if len(data) > max_mb * 1024 * 1024:
            st.error(f"This file is larger than {max_mb} MB. Please compress it before uploading.")
        else:
            if is_pptx:
                slide_count = count_pptx_slides(data)
                slide["uploaded_slide_pptx"] = {
                    "filename": uploaded.name,
                    "content_type": uploaded.type or "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    "slide_count": slide_count,
                    "data_base64": base64.b64encode(data).decode("ascii"),
                }
                slide["visual_image"] = {}
                slide["visual_full_slide"] = False
                slide["uploaded_slide_preview_image"] = {}
                preview_bytes = ensure_uploaded_slide_preview(slide)
                if preview_bytes:
                    st.success(f"Stored PPTX slide replacement: {uploaded.name}. First slide will be imported as editable PowerPoint objects in exports, and a preview image was generated.")
                else:
                    st.warning(f"Stored PPTX slide replacement: {uploaded.name}. The slide will still export as editable PowerPoint objects, but preview generation was not available right now.")
            else:
                slide["visual_image"] = {
                    "filename": uploaded.name,
                    "content_type": uploaded.type or "image/png",
                    "data_base64": base64.b64encode(data).decode("ascii"),
                }
                slide["uploaded_slide_pptx"] = {}
                slide["uploaded_slide_preview_image"] = {}
                st.success(f"Stored image visual: {uploaded.name}.")

    pptx_info = get_uploaded_slide_pptx(slide)
    pptx_bytes = uploaded_slide_pptx_bytes(slide)
    if pptx_bytes:
        slide_count = pptx_info.get("slide_count") or count_pptx_slides(pptx_bytes)
        slide_word = "slide" if slide_count == 1 else "slides"
        st.success(f"PPTX replacement active: {pptx_info.get('filename', 'slide.pptx')} ({slide_count or 'unknown'} {slide_word}). The first slide will be imported as editable PowerPoint objects in the exported PowerPoint. Complex animations/transitions may not import.")
        preview_bytes = ensure_uploaded_slide_preview(slide)
        if preview_bytes:
            st.image(preview_bytes, caption=f"Preview of first slide: {pptx_info.get('filename', 'slide.pptx')}", width=360)
        else:
            st.info("Preview image is not available yet for this PPTX. The editable PPTX replacement will still export.")
        st.download_button(
            "Download uploaded PPTX",
            data=pptx_bytes,
            file_name=pptx_info.get("filename", "uploaded_slide.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            key=f"widget__{slide['id']}__download_uploaded_pptx",
        )
        if st.button("Remove uploaded file", key=f"widget__{slide['id']}__remove_visual", use_container_width=True):
            slide["uploaded_slide_pptx"] = {}
            slide["uploaded_slide_preview_image"] = {}
            slide["visual_full_slide"] = False
            nonce_map[slide["id"]] = nonce + 1
            st.rerun()
        return

    image_bytes = visual_image_bytes(slide)
    image_info = get_visual_image(slide)
    if image_bytes:
        st.image(image_bytes, caption=image_info.get("filename", "Uploaded visual"), width=360)
        st.download_button(
            "Download uploaded image",
            data=image_bytes,
            file_name=image_info.get("filename", "uploaded_visual.png"),
            mime=image_info.get("content_type", "image/png"),
            use_container_width=True,
            key=f"widget__{slide['id']}__download_uploaded_image",
        )
        slide["visual_full_slide"] = st.checkbox(
            "Use this visual as a whole-slide PowerPoint visual",
            value=bool(slide.get("visual_full_slide", False)),
            key=f"widget__{slide['id']}__visual_full_slide",
            help="When checked, the uploaded image uses the full slide body instead of the half-slide visual layout.",
        )
        if st.button("Remove uploaded file", key=f"widget__{slide['id']}__remove_visual", use_container_width=True):
            slide["visual_image"] = {}
            slide["uploaded_slide_preview_image"] = {}
            slide["visual_full_slide"] = False
            nonce_map[slide["id"]] = nonce + 1
            st.rerun()
    else:
        slide["visual_full_slide"] = False

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
            queue_slide_selection(deck["slides"][max(0, slide_index - 2)]["id"])
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
    section_box_label = widget_text(slide, "section_box_label", "Optional section box label", help_text="If entered, the exported slide shows a blue section header box above the main text block (for example: Common resident questions). Leave blank for no box.")
    if section_box_label.strip():
        st.caption(f"Blue section box active: {section_box_label.strip()}")
    else:
        st.caption("No section box will be added unless this field contains text.")
    body = widget_text(slide, "body", "Slide text", height=190, multiline=True, help_text="Use one idea per line. Short lines work best on slides.")
    st.caption(f"{count_words(body)} words. For readability, try to keep most slides under ~45 words.")

    st.markdown("#### Optional slide visual")
    render_visual_upload(slide)

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
    elif role == "Take-home" or kind == "takehome":
        render_takehome_editor(slide)
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

    pptx_replacements = []
    for idx, slide in enumerate(deck.get("slides", []), start=1):
        pptx_info = get_uploaded_slide_pptx(slide)
        if pptx_info.get("data_base64"):
            pptx_replacements.append(f"Slide {idx}: {pptx_info.get('filename', 'uploaded slide.pptx')}")
    if pptx_replacements:
        st.success("Editable PPTX slide replacement active: " + "; ".join(pptx_replacements))

    with st.container(border=True):
        st.markdown("#### Mentor Word document")
        st.caption("Give this to the mentor for comments or Track Changes. Critiques are not stored in the app.")
        complete_mentor_doc = mentor_docx_contains_complete_review_fields(mentor_docx_bytes)
        if complete_mentor_doc:
            st.success(f"Complete mentor template active ({APP_VERSION}): presentation plan, core question, story arc, objectives, take-home points, visuals, and speaker notes are included.")
        else:
            st.error("The mentor DOCX did not pass the complete-template check. Redeploy all app files before downloading.")
        mentor_version = APP_VERSION.rsplit("-", 1)[-1].replace(".", "_")
        st.download_button(
            "Download mentor DOCX",
            data=mentor_docx_bytes,
            file_name=f"mentor_review_{mentor_version}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            disabled=not complete_mentor_doc,
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
                saved_file_names = [result.path.rsplit("/", 1)[-1] for result in results if result.html_url]
                if saved_file_names:
                    st.caption("Saved files: " + ", ".join(saved_file_names))
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
