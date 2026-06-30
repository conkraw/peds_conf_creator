"""
Pediatric Residency Presentation Builder
---------------------------------------
Streamlit front end for a standardized, story-driven PowerPoint builder.

Architecture mirrors the journal-club/case-conference app style:
- app.py = Streamlit UI only
- deck_model.py = shared schema/defaults/helpers
- pptx_builder.py = PowerPoint export, including real speaker notes
- docx_builder.py = printable Word planning form export
- github_storage.py = GitHub draft/archive persistence
"""

from __future__ import annotations

import html
from typing import Any, Dict, List

import streamlit as st

from deck_model import (
    APP_TITLE,
    ARCHIVE_DOCX_NAME,
    ARCHIVE_JSON_NAME,
    ARCHIVE_PPTX_NAME,
    BLOOM_HELPER,
    CHECKLIST_LABELS,
    OBJECTIVE_EXAMPLES,
    SLIDE_ROLES,
    TALK_TYPES,
    default_deck,
    load_deck_from_json,
    make_archive_slug,
    new_slide,
    short_label,
    split_nonempty_lines,
    starter_slides_for_talk_type,
    to_json_bytes,
)
from docx_builder import build_docx
from github_storage import github_config_from_secrets, github_list_json_drafts, github_load_json, save_full_archive
from pptx_builder import build_pptx


# -----------------------------------------------------------------------------
# Session state
# -----------------------------------------------------------------------------


def initialize_state() -> None:
    if "deck" not in st.session_state:
        st.session_state.deck = default_deck()
    if "selected_slide_id" not in st.session_state:
        st.session_state.selected_slide_id = st.session_state.deck["slides"][0]["id"]
    if "github_drafts" not in st.session_state:
        st.session_state.github_drafts = []


def get_selected_slide(deck: Dict[str, Any]) -> Dict[str, Any]:
    slide_id = st.session_state.selected_slide_id
    for slide in deck["slides"]:
        if slide["id"] == slide_id:
            return slide
    st.session_state.selected_slide_id = deck["slides"][0]["id"]
    return deck["slides"][0]


# -----------------------------------------------------------------------------
# Styling / repeated components
# -----------------------------------------------------------------------------


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .identity-card {
            border: 1px solid #d9e6f2;
            background: #f7fbff;
            border-radius: 12px;
            padding: 0.75rem 1rem;
            margin-bottom: 1rem;
        }
        .small-muted { color: #5c6b78; font-size: 0.9rem; }
        .helper-box {
            border-left: 4px solid #1f4e79;
            background: #f7fbff;
            padding: 0.7rem 0.9rem;
            border-radius: 8px;
            margin-bottom: 0.6rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_identity_card(deck: Dict[str, Any]) -> None:
    meta = deck["metadata"]
    title = meta.get("presentation_title") or "Untitled presentation"
    presenter = meta.get("presenter") or "Presenter not entered"
    date = meta.get("session_date") or "Date not entered"
    audience = meta.get("audience") or "Audience not entered"
    talk_type = meta.get("presentation_type") or "Presentation type not entered"
    st.markdown(
        f"""
        <div class="identity-card">
        <strong>{html.escape(title)}</strong><br>
        <span class="small-muted">{html.escape(presenter)} · {html.escape(date)} · {html.escape(audience)} · {html.escape(talk_type)}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# Sidebar and slide actions
# -----------------------------------------------------------------------------


def slide_nav_label(index: int, slide: Dict[str, Any]) -> str:
    title = slide.get("title") or "Untitled"
    role = slide.get("role") or "Slide"
    include_marker = "" if slide.get("include", True) else " [hidden]"
    return f"{index}. {short_label(role, 15)} — {short_label(title, 26)}{include_marker}"


def render_sidebar(deck: Dict[str, Any]) -> None:
    with st.sidebar:
        st.header("Slides")
        labels = [slide_nav_label(i + 1, slide) for i, slide in enumerate(deck["slides"])]
        current_index = next(
            (i for i, slide in enumerate(deck["slides"]) if slide["id"] == st.session_state.selected_slide_id),
            0,
        )
        selected_label = st.radio("Choose slide", labels, index=current_index, label_visibility="collapsed")
        new_index = labels.index(selected_label)
        st.session_state.selected_slide_id = deck["slides"][new_index]["id"]

        st.divider()
        st.subheader("Add slides")
        new_role = st.selectbox("New slide role", SLIDE_ROLES, index=SLIDE_ROLES.index("Custom / Unknown title"))
        new_title = st.text_input("New slide title", placeholder="Leave blank if you do not know it yet")
        new_prompt = st.text_area("Optional helper prompt", height=75)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Add after selected", use_container_width=True):
                slide = new_slide(new_role, new_title, new_prompt, required=False)
                deck["slides"].insert(new_index + 1, slide)
                st.session_state.selected_slide_id = slide["id"]
                st.rerun()
        with col2:
            if st.button("Add at end", use_container_width=True):
                slide = new_slide(new_role, new_title, new_prompt, required=False)
                deck["slides"].append(slide)
                st.session_state.selected_slide_id = slide["id"]
                st.rerun()

        st.divider()
        st.subheader("Draft controls")
        uploaded = st.file_uploader("Reload JSON draft", type=["json"], label_visibility="collapsed")
        if uploaded is not None:
            try:
                st.session_state.deck = load_deck_from_json(uploaded.getvalue())
                st.session_state.selected_slide_id = st.session_state.deck["slides"][0]["id"]
                st.success("Draft loaded.")
                st.rerun()
            except Exception as exc:
                st.error(f"Could not load JSON draft: {exc}")

        if st.button("Start blank presentation", use_container_width=True):
            st.session_state.deck = default_deck()
            st.session_state.selected_slide_id = st.session_state.deck["slides"][0]["id"]
            st.rerun()


# -----------------------------------------------------------------------------
# Editors
# -----------------------------------------------------------------------------


def render_metadata(deck: Dict[str, Any]) -> None:
    meta = deck["metadata"]
    with st.expander("Presentation identity", expanded=True):
        col1, col2 = st.columns([1.5, 1])
        with col1:
            meta["presentation_title"] = st.text_input("Presentation title", meta.get("presentation_title", ""))
            meta["presenter"] = st.text_input("Presenter", meta.get("presenter", ""))
            meta["audience"] = st.text_input("Audience", meta.get("audience", "Pediatric residents"))
        with col2:
            meta["session_date"] = st.text_input("Session date", meta.get("session_date", ""))
            current_type = meta.get("presentation_type", "Educational Topic")
            meta["presentation_type"] = st.selectbox(
                "Presentation type",
                TALK_TYPES,
                index=TALK_TYPES.index(current_type) if current_type in TALK_TYPES else 0,
            )
            if st.button("Replace slide scaffold with this type", use_container_width=True):
                deck["slides"] = starter_slides_for_talk_type(meta["presentation_type"])
                st.session_state.selected_slide_id = deck["slides"][0]["id"]
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
        meta["github_notes"] = st.text_area("Archive notes", meta.get("github_notes", ""), height=70)


def render_bloom_helper() -> None:
    st.markdown("**Bloom’s taxonomy helper**")
    cols = st.columns(3)
    for idx, (level, verbs) in enumerate(BLOOM_HELPER.items()):
        with cols[idx % 3]:
            st.caption(f"**{level}:** {verbs}")
    with st.expander("Objective examples", expanded=False):
        for example in OBJECTIVE_EXAMPLES:
            st.write(f"• {example}")


def move_slide(deck: Dict[str, Any], slide: Dict[str, Any], direction: int) -> None:
    index = deck["slides"].index(slide)
    new_index = index + direction
    if new_index < 0 or new_index >= len(deck["slides"]):
        return
    deck["slides"][index], deck["slides"][new_index] = deck["slides"][new_index], deck["slides"][index]


def render_slide_editor(deck: Dict[str, Any]) -> None:
    slide = get_selected_slide(deck)
    slide_index = deck["slides"].index(slide) + 1

    st.markdown(f"### Slide {slide_index}: {slide.get('role', 'Slide')}")
    if slide.get("prompt"):
        st.markdown(
            f"<div class='helper-box'><strong>Helper:</strong> {html.escape(slide.get('prompt', ''))}</div>",
            unsafe_allow_html=True,
        )

    top1, top2, top3, top4 = st.columns([1.2, 1.2, 0.8, 0.8])
    with top1:
        current_role = slide.get("role", "Story")
        slide["role"] = st.selectbox(
            "Role",
            SLIDE_ROLES,
            index=SLIDE_ROLES.index(current_role) if current_role in SLIDE_ROLES else SLIDE_ROLES.index("Story"),
            disabled=bool(slide.get("required", False)) and current_role in ["Title", "Objectives", "Disclosures"],
        )
    with top2:
        slide["include"] = st.checkbox(
            "Include in PPTX/DOCX",
            bool(slide.get("include", True)),
            disabled=bool(slide.get("required", False)),
        )
    with top3:
        if st.button("Move up", use_container_width=True, disabled=slide_index == 1):
            move_slide(deck, slide, -1)
            st.rerun()
    with top4:
        if st.button("Move down", use_container_width=True, disabled=slide_index == len(deck["slides"])):
            move_slide(deck, slide, 1)
            st.rerun()

    if slide.get("role") == "Title":
        st.info("The title slide pulls from Presentation identity. Use speaker notes below to script your opening hook.")
    else:
        title_col, subtitle_col = st.columns([1.4, 1])
        with title_col:
            slide["title"] = st.text_input(
                "Slide title",
                slide.get("title", ""),
                placeholder="Leave blank if you do not know the title yet",
            )
        with subtitle_col:
            slide["subtitle"] = st.text_input("Optional subtitle", slide.get("subtitle", ""))

    if slide.get("role") == "Objectives":
        render_bloom_helper()
        slide["body"] = st.text_area(
            "Objectives — one per line",
            slide.get("body", ""),
            height=150,
            placeholder="Describe…\nDifferentiate…\nApply…",
        )
    elif slide.get("role") == "Disclosures":
        no_disclosures = st.checkbox("No relevant financial or non-financial disclosures", value="no relevant" in slide.get("body", "").lower())
        if no_disclosures and not slide.get("body", "").strip():
            slide["body"] = "No relevant financial or non-financial disclosures."
        slide["body"] = st.text_area("Disclosure text", slide.get("body", ""), height=110)
    elif slide.get("role") != "Title":
        slide["body"] = st.text_area(
            "Slide body — keep to one main idea",
            slide.get("body", ""),
            height=180,
            placeholder="Use short lines. Each line becomes a separate paragraph in the slide body.",
        )
        col1, col2 = st.columns(2)
        with col1:
            slide["visual_plan"] = st.text_area(
                "Visual / evidence plan",
                slide.get("visual_plan", ""),
                height=110,
                help="What figure, table, case data, graph, algorithm, or image belongs here?",
            )
        with col2:
            slide["discussion_prompt"] = st.text_area(
                "Audience prompt",
                slide.get("discussion_prompt", ""),
                height=110,
                help="Optional: question to ask residents or the room.",
            )

    slide["speaker_notes"] = st.text_area(
        "Speaker notes — exported into real PowerPoint speaker notes",
        slide.get("speaker_notes", ""),
        height=180,
        help="These notes are embedded as speaker notes in the exported PPTX file.",
    )

    if not slide.get("required", False):
        if st.button("Delete this slide", type="secondary"):
            deck["slides"].remove(slide)
            st.session_state.selected_slide_id = deck["slides"][max(0, slide_index - 2)]["id"]
            st.rerun()


def render_mentor_review(deck: Dict[str, Any]) -> None:
    mr = deck["mentor_review"]
    with st.expander("Mentor review", expanded=False):
        col1, col2, col3 = st.columns([1.2, 1.2, 1])
        with col1:
            mr["mentor_name"] = st.text_input("Mentor name", mr.get("mentor_name", ""))
            mr["mentor_email"] = st.text_input("Mentor email", mr.get("mentor_email", ""))
        with col2:
            status_options = ["Not sent", "Sent for review", "Revisions requested", "Approved", "Final"]
            current_status = mr.get("review_status", "Not sent")
            mr["review_status"] = st.selectbox(
                "Review status",
                status_options,
                index=status_options.index(current_status) if current_status in status_options else 0,
            )
            mr["requested_review_date"] = st.text_input("Requested review date", mr.get("requested_review_date", ""))
        with col3:
            mr["review_completed_date"] = st.text_input("Review completed date", mr.get("review_completed_date", ""))
            mr["include_mentor_review_slide"] = st.checkbox(
                "Include mentor-review slide in PPTX",
                bool(mr.get("include_mentor_review_slide", False)),
            )

        st.markdown("**Mentor checklist**")
        for key, label in CHECKLIST_LABELS.items():
            mr["checklist"][key] = st.checkbox(label, bool(mr.get("checklist", {}).get(key, False)), key=f"mentor_{key}")

        mr["mentor_feedback"] = st.text_area("Mentor feedback / revision notes", mr.get("mentor_feedback", ""), height=130)
        mr["mentor_approval_statement"] = st.text_area(
            "Mentor approval statement",
            mr.get("mentor_approval_statement", ""),
            height=80,
            placeholder="Example: Reviewed with mentor; revisions completed; approved for presentation.",
        )


def render_story_check(deck: Dict[str, Any]) -> None:
    included = [slide for slide in deck["slides"] if slide.get("include", True)]
    missing_titles = [i + 1 for i, slide in enumerate(included) if not (slide.get("title") or "").strip() and slide.get("role") != "Title"]
    missing_notes = [i + 1 for i, slide in enumerate(included) if not (slide.get("speaker_notes") or "").strip()]
    objectives = next((slide for slide in deck["slides"] if slide.get("role") == "Objectives"), None)
    objective_lines = split_nonempty_lines(objectives.get("body", "") if objectives else "")

    with st.expander("Story and readiness check", expanded=False):
        st.write(f"Included slides: **{len(included)}**")
        st.write(f"Objectives entered: **{len(objective_lines)}**")
        if missing_titles:
            st.warning(f"Slides without a title: {', '.join(map(str, missing_titles))}. This is allowed; export will use generic story-slide titles.")
        else:
            st.success("All included non-title slides have titles.")
        if missing_notes:
            st.warning(f"Slides missing speaker notes: {', '.join(map(str, missing_notes))}.")
        else:
            st.success("All included slides have speaker notes.")


# -----------------------------------------------------------------------------
# Export / archive controls
# -----------------------------------------------------------------------------


def render_export_controls(deck: Dict[str, Any]) -> None:
    st.markdown("### Export / archive")
    slug = make_archive_slug(deck)
    pptx_bytes = build_pptx(deck)
    docx_bytes = build_docx(deck)
    json_bytes = to_json_bytes(deck)

    st.download_button(
        "Download PowerPoint",
        data=pptx_bytes,
        file_name=f"{slug}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )
    st.download_button(
        "Download Word planning form",
        data=docx_bytes,
        file_name=f"{slug}_planning_form.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
    st.download_button(
        "Download editable JSON draft",
        data=json_bytes,
        file_name=f"{slug}.json",
        mime="application/json",
        use_container_width=True,
    )

    config = github_config_from_secrets()
    with st.expander("GitHub archive", expanded=False):
        if not config:
            st.info("GitHub saving is not configured. Add secrets first.")
            st.code(
                'GITHUB_TOKEN = "ghp_your_token"\nGITHUB_REPO = "your_username/your_repo"\nGITHUB_BRANCH = "main"\nGITHUB_FOLDER = "presentation_archive"',
                language="toml",
            )
            st.caption(f"When configured, the app saves {ARCHIVE_JSON_NAME}, {ARCHIVE_PPTX_NAME}, and {ARCHIVE_DOCX_NAME} together.")
            return

        st.caption(f"Repo: {config.repo} · branch: {config.branch} · folder: {config.folder}")
        if st.button("Save JSON + PowerPoint + DOCX to GitHub", use_container_width=True):
            try:
                folder = save_full_archive(config, slug, json_bytes, pptx_bytes, docx_bytes)
                st.success(f"Saved to GitHub: {folder}")
            except Exception as exc:
                st.error(f"GitHub save failed: {exc}")

        if st.button("Refresh GitHub draft list", use_container_width=True):
            try:
                st.session_state.github_drafts = github_list_json_drafts(config)
            except Exception as exc:
                st.error(f"Could not list GitHub drafts: {exc}")

        drafts = st.session_state.get("github_drafts", [])
        if drafts:
            draft_names = [draft["name"] for draft in drafts]
            selected_name = st.selectbox("Load draft from GitHub", draft_names)
            selected = next(item for item in drafts if item["name"] == selected_name)
            if st.button("Load selected GitHub draft", use_container_width=True):
                try:
                    st.session_state.deck = github_load_json(config, selected["path"])
                    st.session_state.selected_slide_id = st.session_state.deck["slides"][0]["id"]
                    st.success("GitHub draft loaded.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Could not load selected draft: {exc}")


# -----------------------------------------------------------------------------
# Main app
# -----------------------------------------------------------------------------


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🩺", layout="wide")
    initialize_state()
    inject_css()

    deck = st.session_state.deck
    render_sidebar(deck)

    st.title(APP_TITLE)
    st.caption("Build standardized, story-driven presentations with mentor review, DOCX planning forms, GitHub archiving, and real PowerPoint speaker notes.")
    render_identity_card(deck)

    editor_col, export_col = st.columns([2.25, 0.95], gap="large")
    with editor_col:
        render_metadata(deck)
        render_slide_editor(deck)
        render_mentor_review(deck)
        render_story_check(deck)
    with export_col:
        render_export_controls(deck)


if __name__ == "__main__":
    main()
