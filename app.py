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

GitHub remains the archive source of truth, with local JSON upload available for restoring or transferring drafts.
"""

from __future__ import annotations

import base64
import hashlib
import html
import json
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
    ensure_core_slide_order,
    identity_subtitle,
    identity_title,
    new_slide,
    normalize_loaded_deck,
    short_label,
    split_nonempty_lines,
)
from github_storage import (
    GitHubStorageError,
    github_is_configured,
    github_status_message,
    list_archives_from_github,
    load_json_from_github,
    save_archive_to_github,
    save_draft_to_github,
    delete_archive_from_github,
)


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

# Persistent green success alerts are intentionally disabled to keep the app
# visually quiet. Set this to True when troubleshooting or during development.
SHOW_SUCCESS_ALERTS = False

# User-triggered archive actions still receive a brief, non-persistent toast.
SHOW_ACTION_TOASTS = True

# Lightweight autosave: only draft.json is saved. PPTX and mentor DOCX remain
# explicit actions so autosave does not slow editing by rebuilding exports.
AUTOSAVE_ENABLED = True
AUTOSAVE_ON_SLIDE_CHANGE = True
AUTOSAVE_FAILURE_TOASTS = True


def success_notice(message: str, *, action: bool = False) -> None:
    """Show optional success feedback without leaving persistent green boxes."""
    if SHOW_SUCCESS_ALERTS:
        st.success(message)
    elif action and SHOW_ACTION_TOASTS:
        st.toast(message, icon="✅")


def serialize_deck_for_export(deck: Dict[str, Any]) -> str:
    """Return a stable JSON snapshot used for export caching and staleness checks."""
    return json.dumps(deck, sort_keys=True, ensure_ascii=False, separators=(",", ":"), default=str)


def _signature_view(value: Any, parent_key: str = "") -> Any:
    """Create a lightweight export fingerprint without mutating the deck.

    Older GitHub JSON files may contain base64 assets without a stored SHA-256
    value. The previous implementation added that value while iterating through
    the same dictionary, which raised ``RuntimeError: dictionary changed size
    during iteration``. This version takes a stable snapshot of the items and
    computes any fallback digest locally.
    """
    if isinstance(value, dict):
        result: Dict[str, Any] = {}
        existing_digest = str(value.get("sha256", "") or "").strip()
        for key, item in list(value.items()):
            # Cached UI previews do not change the exported PPTX or mentor content.
            if key == "uploaded_slide_preview_image":
                continue
            if key == "data_base64":
                encoded = item if isinstance(item, str) else str(item or "")
                digest = existing_digest
                if not digest and encoded:
                    digest = hashlib.sha256(encoded.encode("ascii", errors="ignore")).hexdigest()
                result[key] = {"characters": len(encoded), "sha256": digest}
            else:
                result[key] = _signature_view(item, key)
        return result
    if isinstance(value, list):
        return [_signature_view(item, parent_key) for item in value]
    return value


def ensure_deck_asset_hashes(deck: Dict[str, Any]) -> None:
    """Backfill hashes for assets loaded from older JSON files once per load.

    Current uploads already store a SHA-256 value. Older archives often do not.
    Backfilling here keeps subsequent export-signature checks fast while leaving
    preview-only images out of the export fingerprint.
    """
    for slide in deck.get("slides", []):
        if not isinstance(slide, dict):
            continue
        for asset_key in ("visual_image", "uploaded_slide_pptx"):
            asset = slide.get(asset_key)
            if not isinstance(asset, dict):
                continue
            if str(asset.get("sha256", "") or "").strip():
                continue
            encoded = asset.get("data_base64")
            if not isinstance(encoded, str) or not encoded:
                continue
            try:
                asset["sha256"] = hashlib.sha256(base64.b64decode(encoded)).hexdigest()
            except Exception:
                # A malformed legacy asset should not prevent the presentation
                # itself from loading or the rest of the app from rendering.
                continue


def deck_export_signature(deck: Dict[str, Any]) -> str:
    ensure_deck_asset_hashes(deck)
    signature_json = json.dumps(
        _signature_view(deck),
        sort_keys=True,
        ensure_ascii=False,
        separators=(",", ":"),
        default=str,
    )
    return hashlib.sha256(signature_json.encode("utf-8")).hexdigest()


def reset_autosave_tracking(deck: Dict[str, Any], *, mark_clean: bool = True) -> None:
    """Reset autosave state after loading or starting a presentation."""
    st.session_state.autosave_file_sha = ""
    st.session_state.autosave_last_error = ""
    st.session_state.autosave_last_signature = deck_export_signature(deck) if mark_clean else ""


def autosave_current_draft(*, reason: str = "presentation change", force: bool = False) -> bool:
    """Save the current JSON draft to GitHub without rebuilding exports.

    Autosave is deliberately tied to navigation and structural changes rather
    than every keystroke. This matches the case-conference workflow while
    keeping the app responsive and avoiding excessive GitHub commits.
    """
    if not AUTOSAVE_ENABLED or not github_is_configured():
        return False
    if st.session_state.get("autosave_in_progress", False):
        return False

    deck = st.session_state.get("deck")
    if not isinstance(deck, dict):
        return False

    current_signature = deck_export_signature(deck)
    if not force and current_signature == st.session_state.get("autosave_last_signature", ""):
        return True

    st.session_state.autosave_in_progress = True
    try:
        result = save_draft_to_github(
            deck,
            st.session_state.get("archive_path", ""),
            known_sha=st.session_state.get("autosave_file_sha", ""),
            commit_message=f"Autosave presentation draft: {reason}",
        )
        st.session_state.archive_path = result.path.rsplit("/", 1)[0]
        st.session_state.autosave_file_sha = result.file_sha
        st.session_state.autosave_last_signature = current_signature
        st.session_state.autosave_last_error = ""
        return True
    except Exception as exc:
        message = str(exc)
        previous_error = st.session_state.get("autosave_last_error", "")
        # A cached GitHub file SHA may become stale if the archive changes in
        # another browser/session. Clear it so the next autosave refreshes the
        # current SHA before trying again.
        st.session_state.autosave_file_sha = ""
        st.session_state.autosave_last_error = message
        if AUTOSAVE_FAILURE_TOASTS and message != previous_error:
            st.toast("Autosave could not reach GitHub. Your current session is still intact.", icon="⚠️")
        return False
    finally:
        st.session_state.autosave_in_progress = False


@st.cache_data(show_spinner=False, max_entries=8)
def cached_build_pptx(deck_json: str) -> bytes:
    """Cache expensive PowerPoint generation across reruns and repeated requests."""
    from pptx_builder import build_pptx

    return build_pptx(json.loads(deck_json))


@st.cache_data(show_spinner=False, max_entries=8)
def cached_build_mentor_docx(deck_json: str, pptx_bytes: bytes) -> bytes:
    """Cache mentor DOCX generation, including slide-preview rendering."""
    from docx_builder import build_mentor_review_docx

    return build_mentor_review_docx(json.loads(deck_json), pptx_bytes=pptx_bytes)


@st.cache_data(show_spinner=False, max_entries=24)
def cached_render_pptx_first_slide(pptx_bytes: bytes) -> bytes | None:
    """Render an uploaded PPTX preview once, rather than on every Streamlit rerun."""
    from preview_utils import render_pptx_first_slide_to_png

    return render_pptx_first_slide_to_png(pptx_bytes)


def mentor_docx_is_complete(docx_bytes: bytes) -> bool:
    from docx_builder import mentor_docx_contains_complete_review_fields

    return mentor_docx_contains_complete_review_fields(docx_bytes)


# -----------------------------------------------------------------------------
# Utility helpers
# -----------------------------------------------------------------------------


def count_words(text: Any) -> int:
    return len(re.findall(r"\b\w+\b", str(text or "")))


def decode_uploaded_json(uploaded_file: Any) -> Dict[str, Any]:
    """Decode a locally uploaded JSON draft with clear, safe validation."""
    if uploaded_file is None:
        raise ValueError("Choose a JSON file first.")
    raw = uploaded_file.getvalue()
    if not raw:
        raise ValueError("The uploaded JSON file is empty.")
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("The uploaded file is not valid UTF-8 JSON.") from exc
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"The uploaded file is not valid JSON (line {exc.lineno}, column {exc.colno}).") from exc
    if not isinstance(payload, dict):
        raise ValueError("The JSON draft must contain a JSON object at the top level.")
    return payload


def unwrap_uploaded_json_payload(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Support plain drafts, GitHub archive payloads, and common wrappers."""
    current: Any = payload
    # Some exported files wrap the actual deck in ``content``. Support either
    # a parsed object or a JSON string without changing the standard schema.
    if isinstance(current, dict) and "content" in current and "slides" not in current and "deck" not in current:
        content = current.get("content")
        if isinstance(content, str):
            try:
                content = json.loads(content)
            except json.JSONDecodeError:
                content = None
        if isinstance(content, dict):
            current = content
    if not isinstance(current, dict):
        raise ValueError("The JSON file did not contain a presentation draft.")
    return current


def json_import_summary(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Return a small human-readable summary without fully rendering the deck."""
    unwrapped = unwrap_uploaded_json_payload(payload)
    loaded = unwrapped.get("deck", unwrapped) if isinstance(unwrapped, dict) else {}
    metadata = loaded.get("metadata", {}) if isinstance(loaded, dict) else {}
    slides = loaded.get("slides", []) if isinstance(loaded, dict) else []
    return {
        "title": str(metadata.get("presentation_title") or "Untitled presentation"),
        "presenter": str(metadata.get("presenter") or "Presenter not entered"),
        "slide_count": len(slides) if isinstance(slides, list) else 0,
        "archive_path": str(unwrapped.get("archive_path") or payload.get("archive_path") or ""),
    }


def clear_widget_state() -> None:
    for key in list(st.session_state.keys()):
        if key.startswith("widget__"):
            del st.session_state[key]


def initialize_state() -> None:
    if "deck" not in st.session_state:
        st.session_state.deck = default_deck()
    ensure_core_slide_order(st.session_state.deck)
    st.session_state.deck["app_version"] = APP_VERSION
    if "show_add_slides" not in st.session_state:
        st.session_state.show_add_slides = False
    if "show_github_archive" not in st.session_state:
        st.session_state.show_github_archive = False
    if "show_json_import" not in st.session_state:
        st.session_state.show_json_import = False
    if "json_import_nonce" not in st.session_state:
        st.session_state.json_import_nonce = 0
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
    if "prepared_pptx_bytes" not in st.session_state:
        st.session_state.prepared_pptx_bytes = None
    if "prepared_pptx_signature" not in st.session_state:
        st.session_state.prepared_pptx_signature = ""
    if "prepared_mentor_docx_bytes" not in st.session_state:
        st.session_state.prepared_mentor_docx_bytes = None
    if "prepared_mentor_signature" not in st.session_state:
        st.session_state.prepared_mentor_signature = ""
    if "autosave_last_signature" not in st.session_state:
        st.session_state.autosave_last_signature = deck_export_signature(st.session_state.deck)
    if "autosave_file_sha" not in st.session_state:
        st.session_state.autosave_file_sha = ""
    if "autosave_last_error" not in st.session_state:
        st.session_state.autosave_last_error = ""
    if "autosave_in_progress" not in st.session_state:
        st.session_state.autosave_in_progress = False


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
    current = st.session_state.get("selected_slide_id")
    if selected in slide_ids:
        if AUTOSAVE_ON_SLIDE_CHANGE and selected != current:
            autosave_current_draft(reason="slide change")
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
    preview = cached_render_pptx_first_slide(pptx_bytes)
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


def asset_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def stored_asset_sha256(asset: Dict[str, Any]) -> str:
    existing = str(asset.get("sha256", "") or "").strip()
    if existing:
        return existing
    encoded = asset.get("data_base64")
    if not encoded:
        return ""
    try:
        digest = asset_sha256(base64.b64decode(encoded))
        asset["sha256"] = digest
        return digest
    except Exception:
        return ""


def has_uploaded_visual(slide: Dict[str, Any]) -> bool:
    """Check visual presence without decoding large base64 payloads."""
    return bool(
        get_visual_image(slide).get("data_base64")
        or get_uploaded_slide_pptx(slide).get("data_base64")
    )


def slide_nav_label(index: int, slide: Dict[str, Any]) -> str:
    role = slide.get("role") or "Slide"
    title = slide.get("title") or "Untitled"
    return f"{index}. {short_label(role, 16)} — {short_label(title, 30)}"


def validation_messages(deck: Dict[str, Any]) -> List[Dict[str, str]]:
    """Return actionable readiness issues with a clear location."""
    issues: List[Dict[str, str]] = []
    meta = deck.get("metadata", {})

    if not str(meta.get("presentation_title", "")).strip():
        issues.append({
            "location": "Presentation identity → Presentation title",
            "message": "The presentation title is blank.",
            "fix": "Open the Title slide and enter a presentation title.",
            "slide_id": deck.get("slides", [{}])[0].get("id", ""),
        })

    if not str(meta.get("presenter", "")).strip():
        issues.append({
            "location": "Presentation identity → Presenter",
            "message": "The presenter name is blank.",
            "fix": "Open the Title slide and enter the presenter name.",
            "slide_id": deck.get("slides", [{}])[0].get("id", ""),
        })

    for idx, slide in enumerate(deck.get("slides", []), start=1):
        role = slide.get("role", "Slide")
        title = str(slide.get("title") or "Untitled slide")
        slide_id = str(slide.get("id") or "")
        slide_location = f"Slide {idx}: {title}"

        if role == "Objectives":
            objective_count = len(split_nonempty_lines(slide.get("body", "")))
            if objective_count < 1:
                issues.append({
                    "location": f"{slide_location} → Objectives",
                    "message": "No objectives are entered.",
                    "fix": "Add at least one Bloom-style objective.",
                    "slide_id": slide_id,
                })
        elif role == "Take-home":
            takehome_count = len([
                1 for i in range(1, 6)
                if str(slide.get(f"takehome_point_{i}", "")).strip()
            ])
            if takehome_count < 1 and not has_uploaded_visual(slide):
                issues.append({
                    "location": f"{slide_location} → Take-home points",
                    "message": "No take-home points or visual are entered.",
                    "fix": "Add at least one take-home point or upload a visual.",
                    "slide_id": slide_id,
                })
        elif role != "Title" and not str(slide.get("body", "")).strip() and not has_uploaded_visual(slide):
            issues.append({
                "location": f"{slide_location} → Slide text / visual",
                "message": "The slide has no main text or uploaded visual.",
                "fix": "Enter slide text, upload an image/PPTX slide, or delete the unused slide.",
                "slide_id": slide_id,
            })

    return issues


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

        add_slides_label = "Close add slides" if st.session_state.show_add_slides else "Add slides"
        if st.button(add_slides_label, key="toggle_add_slides_panel", use_container_width=True):
            st.session_state.show_add_slides = not st.session_state.show_add_slides
            st.rerun()

        if st.session_state.show_add_slides:
            with st.container(border=True):
                new_role = st.selectbox("New slide role", SLIDE_ROLES, index=SLIDE_ROLES.index("Custom / Unknown title"))
                new_title = st.text_input("New slide title", placeholder="Leave blank if you do not know it yet")
                new_prompt = st.text_area("Optional helper prompt", height=75, placeholder="What should this slide help the presenter do?")

                selected_index = next((i for i, slide in enumerate(deck["slides"]) if slide["id"] == st.session_state.selected_slide_id), 0)
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("Add after", use_container_width=True):
                        slide = new_slide(new_role, new_title, new_prompt)
                        deck["slides"].insert(selected_index + 1, slide)
                        autosave_current_draft(reason="slide added", force=True)
                        queue_slide_selection(slide["id"])
                        st.rerun()
                with col2:
                    if st.button("Add at end", use_container_width=True):
                        slide = new_slide(new_role, new_title, new_prompt)
                        deck["slides"].append(slide)
                        autosave_current_draft(reason="slide added", force=True)
                        queue_slide_selection(slide["id"])
                        st.rerun()

        st.divider()
        github_label = "Close GitHub archive" if st.session_state.show_github_archive else "GitHub archive"
        if st.button(github_label, key="toggle_github_archive_panel", use_container_width=True):
            st.session_state.show_github_archive = not st.session_state.show_github_archive
            st.rerun()

        if st.session_state.show_github_archive:
            with st.container(border=True):
                if not github_is_configured():
                    st.warning(github_status_message())
                elif AUTOSAVE_ENABLED:
                    st.caption("Autosave is active: draft.json saves when you change slides. PPTX and mentor DOCX save only with Save all to GitHub.")

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
                            reset_autosave_tracking(st.session_state.deck, mark_clean=True)
                            queue_slide_selection(st.session_state.deck["slides"][0]["id"])
                            clear_widget_state()
                            success_notice("Loaded from GitHub.", action=True)
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
                                success_notice(f"Deleted {deleted_count} file(s) from GitHub.", action=True)
                                st.rerun()
                            except GitHubStorageError as exc:
                                st.error(str(exc))

        st.divider()
        json_label = "Close JSON import" if st.session_state.show_json_import else "Import JSON draft"
        if st.button(json_label, key="toggle_json_import_panel", use_container_width=True):
            st.session_state.show_json_import = not st.session_state.show_json_import
            st.rerun()

        if st.session_state.show_json_import:
            with st.container(border=True):
                st.caption("Upload a presentation-builder JSON draft. Plain deck JSON and wrapped GitHub archive JSON are supported, including custom slides, speaker notes, images, and embedded PPTX replacements.")
                uploaded_json = st.file_uploader(
                    "Upload presentation JSON",
                    type=["json"],
                    key=f"widget__json_import_file__{st.session_state.json_import_nonce}",
                    help="The uploaded file opens as a new working copy by default, so autosave will not overwrite the original GitHub archive.",
                )

                payload: Dict[str, Any] | None = None
                summary: Dict[str, Any] | None = None
                if uploaded_json is not None:
                    try:
                        payload = decode_uploaded_json(uploaded_json)
                        summary = json_import_summary(payload)
                        st.caption(
                            f"{summary['title']} · {summary['presenter']} · "
                            f"{summary['slide_count']} slide(s)"
                        )
                    except ValueError as exc:
                        st.error(str(exc))

                reconnect_archive = False
                if summary and summary.get("archive_path"):
                    reconnect_archive = st.checkbox(
                        "Reconnect autosave to the archive path stored in this JSON",
                        value=False,
                        key="widget__json_import_reconnect_archive",
                        help="Leave unchecked to open a new working copy. Check only when you intentionally want future autosaves to update the original GitHub archive.",
                    )

                if st.button(
                    "Load JSON draft",
                    use_container_width=True,
                    disabled=payload is None,
                    key="load_uploaded_json_draft",
                ):
                    try:
                        unwrapped = unwrap_uploaded_json_payload(payload or {})
                        loaded_deck = normalize_loaded_deck(unwrapped)
                        if not loaded_deck.get("slides"):
                            raise ValueError("The JSON file did not contain any usable slides.")
                        ensure_deck_asset_hashes(loaded_deck)
                        st.session_state.deck = loaded_deck
                        st.session_state.archive_path = (
                            str((summary or {}).get("archive_path") or "") if reconnect_archive else ""
                        )
                        reset_autosave_tracking(st.session_state.deck, mark_clean=True)
                        st.session_state.prepared_pptx_bytes = None
                        st.session_state.prepared_pptx_signature = ""
                        st.session_state.prepared_mentor_docx_bytes = None
                        st.session_state.prepared_mentor_signature = ""
                        st.session_state.visual_uploader_nonce = {}
                        st.session_state.json_import_nonce += 1
                        st.session_state.show_json_import = False
                        queue_slide_selection(st.session_state.deck["slides"][0]["id"])
                        clear_widget_state()
                        success_notice("JSON draft loaded as a new working copy." if not reconnect_archive else "JSON draft loaded and reconnected to its GitHub archive.", action=True)
                        st.rerun()
                    except (ValueError, TypeError) as exc:
                        st.error(f"Could not load this JSON draft: {exc}")
                    except Exception as exc:
                        st.error(f"Could not load this JSON draft: {exc}")

        st.divider()
        if st.button("Start blank presentation", use_container_width=True):
            st.session_state.deck = default_deck()
            st.session_state.archive_path = ""
            reset_autosave_tracking(st.session_state.deck, mark_clean=True)
            queue_slide_selection(st.session_state.deck["slides"][0]["id"])
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
            autosave_current_draft(reason="story scaffold replaced", force=True)
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
    slide["speaker_notes"] = widget_text(slide, "speaker_notes", "Speaker notes for title slide", height=220, multiline=True)


def render_objectives_editor(slide: Dict[str, Any]) -> None:
    st.markdown("### Objectives")
    success_notice("Objectives card layout is active: Bloom dropdowns → numbered visual objective cards in PowerPoint.")
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
    widget_text(slide, "speaker_notes", "Speaker notes", height=220, multiline=True)


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
    widget_text(slide, "discussion_prompt", "Optional discussion prompt at bottom", height=220, multiline=True, help_text="Optional closing question or prompt shown near the bottom of the take-home slide.")
    widget_text(slide, "speaker_notes", "Speaker notes", height=220, multiline=True)


def render_disclosures_editor(slide: Dict[str, Any]) -> None:
    st.markdown("### Disclosures")
    widget_text(slide, "title", "Slide title")
    widget_text(slide, "body", "Disclosure text", height=220, multiline=True)
    widget_text(slide, "speaker_notes", "Speaker notes", height=220, multiline=True)


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
    """Store one optional uploaded asset without reprocessing it on every rerun."""
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
        upload_hash = asset_sha256(data)
        is_pptx = (
            uploaded.type == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            or uploaded.name.lower().endswith(".pptx")
        )
        stored_asset = get_uploaded_slide_pptx(slide) if is_pptx else get_visual_image(slide)
        already_processed = stored_asset_sha256(stored_asset) == upload_hash

        # Streamlit retains the uploaded file across reruns. Only process it when
        # the actual file changes; otherwise PPTX previews would be reconverted
        # every time the user types in any field.
        if not already_processed:
            max_mb = 15 if is_pptx else 5
            if len(data) > max_mb * 1024 * 1024:
                st.error(f"This file is larger than {max_mb} MB. Please compress it before uploading.")
            elif is_pptx:
                slide_count = count_pptx_slides(data)
                slide["uploaded_slide_pptx"] = {
                    "filename": uploaded.name,
                    "content_type": uploaded.type or "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    "slide_count": slide_count,
                    "sha256": upload_hash,
                    "data_base64": base64.b64encode(data).decode("ascii"),
                }
                slide["visual_image"] = {}
                slide["visual_full_slide"] = False
                slide["uploaded_slide_preview_image"] = {}
                with st.spinner("Generating the PPTX preview once…"):
                    preview_bytes = ensure_uploaded_slide_preview(slide)
                if not preview_bytes:
                    st.warning("The editable PPTX was stored, but its preview could not be generated. You can retry below.")
            else:
                slide["visual_image"] = {
                    "filename": uploaded.name,
                    "content_type": uploaded.type or "image/png",
                    "sha256": upload_hash,
                    "data_base64": base64.b64encode(data).decode("ascii"),
                }
                slide["uploaded_slide_pptx"] = {}
                slide["uploaded_slide_preview_image"] = {}

            autosave_current_draft(reason="visual uploaded", force=True)

    pptx_info = get_uploaded_slide_pptx(slide)
    pptx_bytes = uploaded_slide_pptx_bytes(slide)
    if pptx_bytes:
        slide_count = pptx_info.get("slide_count") or count_pptx_slides(pptx_bytes)
        preview_bytes = uploaded_slide_preview_bytes(slide)
        if preview_bytes:
            st.image(preview_bytes, caption=f"Preview of first slide: {pptx_info.get('filename', 'slide.pptx')}", width=420)
        else:
            if st.button(
                "Generate PPTX preview",
                key=f"widget__{slide['id']}__generate_pptx_preview",
                use_container_width=True,
            ):
                with st.spinner("Generating preview…"):
                    preview_bytes = ensure_uploaded_slide_preview(slide)
                if preview_bytes:
                    st.rerun()
                else:
                    st.error("The preview could not be generated. The editable PPTX can still be exported and downloaded.")

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
            autosave_current_draft(reason="visual removed", force=True)
            st.rerun()
        return

    image_bytes = visual_image_bytes(slide)
    image_info = get_visual_image(slide)
    if image_bytes:
        st.image(image_bytes, caption=image_info.get("filename", "Uploaded visual"), width=420)
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
            autosave_current_draft(reason="visual removed", force=True)
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
            autosave_current_draft(reason="slide moved", force=True)
            st.rerun()
    with action_cols[1]:
        if st.button("Move down", use_container_width=True, disabled=slide_index == len(deck["slides"])):
            move_slide(deck, slide, 1)
            autosave_current_draft(reason="slide moved", force=True)
            st.rerun()
    with action_cols[2]:
        if st.button("Duplicate", use_container_width=True):
            duplicate_slide(deck, slide)
            autosave_current_draft(reason="slide duplicated", force=True)
            st.rerun()
    with action_cols[3]:
        if st.button("Delete", use_container_width=True, disabled=bool(slide.get("required", False))):
            deck["slides"].remove(slide)
            queue_slide_selection(deck["slides"][max(0, slide_index - 2)]["id"])
            clear_widget_state()
            autosave_current_draft(reason="slide deleted", force=True)
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
    body = widget_text(slide, "body", "Slide text", height=190, multiline=True, help_text="Use one idea per line. Short lines work best on slides. When text is entered, it exports inside a lightly colored panel so the bullets feel visually connected.")
    st.caption(f"{count_words(body)} words. For readability, try to keep most slides under ~45 words.")

    st.markdown("#### Optional slide visual")
    render_visual_upload(slide)

    widget_text(slide, "discussion_prompt", "Discussion prompt", height=125, multiline=True, help_text="Question to ask the audience, if useful.")
    widget_text(slide, "speaker_notes", "Speaker notes exported into PowerPoint", height=260, multiline=True)


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


def clear_stale_prepared_exports(current_signature: str) -> None:
    """Drop stale session copies as soon as the deck changes."""
    if st.session_state.get("prepared_pptx_signature") != current_signature:
        st.session_state.prepared_pptx_bytes = None
        st.session_state.prepared_pptx_signature = ""
    if st.session_state.get("prepared_mentor_signature") != current_signature:
        st.session_state.prepared_mentor_docx_bytes = None
        st.session_state.prepared_mentor_signature = ""


def prepare_pptx_export(deck: Dict[str, Any], current_signature: str) -> bytes:
    deck_json = serialize_deck_for_export(deck)
    pptx_bytes = cached_build_pptx(deck_json)
    st.session_state.prepared_pptx_bytes = pptx_bytes
    st.session_state.prepared_pptx_signature = current_signature
    return pptx_bytes


def prepare_mentor_export(deck: Dict[str, Any], current_signature: str) -> bytes:
    pptx_bytes = st.session_state.get("prepared_pptx_bytes")
    if st.session_state.get("prepared_pptx_signature") != current_signature or not pptx_bytes:
        pptx_bytes = prepare_pptx_export(deck, current_signature)
    deck_json = serialize_deck_for_export(deck)
    mentor_bytes = cached_build_mentor_docx(deck_json, pptx_bytes)
    st.session_state.prepared_mentor_docx_bytes = mentor_bytes
    st.session_state.prepared_mentor_signature = current_signature
    return mentor_bytes


def render_export_panel(deck: Dict[str, Any]) -> None:
    """Render lazy export controls so normal editing reruns stay fast."""
    current_signature = deck_export_signature(deck)
    clear_stale_prepared_exports(current_signature)

    st.markdown("### Export / archive")
    st.caption("Exports are generated only when requested so editing and slide navigation stay responsive.")

    with st.container(border=True):
        st.markdown("#### Mentor Word document")
        st.caption("Compact PowerPoint previews, editable wording, full speaker notes, and mentor comment boxes.")
        mentor_ready = (
            st.session_state.get("prepared_mentor_signature") == current_signature
            and bool(st.session_state.get("prepared_mentor_docx_bytes"))
        )
        if st.button(
            "Refresh mentor DOCX" if mentor_ready else "Prepare mentor DOCX",
            use_container_width=True,
            key="prepare_mentor_docx",
        ):
            try:
                with st.spinner("Preparing the mentor review…"):
                    prepare_mentor_export(deck, current_signature)
                mentor_ready = True
            except Exception as exc:
                st.error(f"Could not build the mentor DOCX: {exc}")
                mentor_ready = False

        mentor_docx_bytes = st.session_state.get("prepared_mentor_docx_bytes") if mentor_ready else None
        if mentor_docx_bytes:
            complete_mentor_doc = mentor_docx_is_complete(mentor_docx_bytes)
            if not complete_mentor_doc:
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
        st.caption("Speaker notes are placed into the real PowerPoint notes section.")
        pptx_ready = (
            st.session_state.get("prepared_pptx_signature") == current_signature
            and bool(st.session_state.get("prepared_pptx_bytes"))
        )
        if st.button(
            "Refresh PPTX" if pptx_ready else "Prepare PPTX",
            use_container_width=True,
            key="prepare_pptx",
        ):
            try:
                with st.spinner("Preparing the PowerPoint…"):
                    prepare_pptx_export(deck, current_signature)
                pptx_ready = True
            except Exception as exc:
                st.error(f"Could not build the PowerPoint: {exc}")
                pptx_ready = False

        pptx_bytes = st.session_state.get("prepared_pptx_bytes") if pptx_ready else None
        if pptx_bytes:
            st.download_button(
                "Download PPTX",
                data=pptx_bytes,
                file_name=ARCHIVE_PPTX_NAME,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

    with st.container(border=True):
        st.markdown("#### GitHub archive")
        st.caption("Builds the current files only when Save is clicked, then stores the draft, PPTX, and mentor DOCX.")
        if st.button("Save all to GitHub", use_container_width=True):
            try:
                with st.spinner("Preparing current exports and saving to GitHub…"):
                    pptx_bytes = prepare_pptx_export(deck, current_signature)
                    mentor_docx_bytes = prepare_mentor_export(deck, current_signature)
                    results = save_archive_to_github(
                        deck,
                        pptx_bytes,
                        mentor_docx_bytes,
                        st.session_state.get("archive_path", ""),
                    )
                if results:
                    st.session_state.archive_path = results[0].path.rsplit("/", 1)[0]
                    st.session_state.autosave_file_sha = results[0].file_sha
                    st.session_state.autosave_last_signature = current_signature
                    st.session_state.autosave_last_error = ""
                success_notice("Saved to GitHub archive.", action=True)
                saved_file_names = [result.path.rsplit("/", 1)[-1] for result in results if result.html_url]
                if saved_file_names:
                    st.caption("Saved files: " + ", ".join(saved_file_names))
            except GitHubStorageError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Could not prepare or save the current presentation: {exc}")


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
        st.error(f"Readiness check found {len(problems)} issue(s) that may affect the export.")
        with st.expander("Show issues and locations", expanded=False):
            for number, problem in enumerate(problems, start=1):
                st.markdown(f"**{number}. {problem['location']}**")
                st.write(problem["message"])
                st.caption(f"Suggested fix: {problem['fix']}")
    else:
        # A passing readiness check is intentionally silent.
        success_notice("Readiness check passed. The presentation has the core fields needed for export.")

    editor_col, export_col = st.columns([2.1, 0.85], gap="large")
    with editor_col:
        render_slide_editor(deck)

    with export_col:
        render_export_panel(deck)


if __name__ == "__main__":
    main()
