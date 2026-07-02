"""Shared schema and helper functions for the Presentation PowerPoint Builder."""

from __future__ import annotations

import copy
import datetime as dt
import json
import re
import uuid
from typing import Any, Dict, List

APP_TITLE = "Pediatric Residency Presentation Builder"
APP_VERSION = "2026.06.30-v3.6"
ARCHIVE_JSON_NAME = "draft.json"
ARCHIVE_PPTX_NAME = "presentation.pptx"
ARCHIVE_DOCX_NAME = "mentor_review.docx"

TALK_TYPES = [
    "Educational Topic",
    "Study Review",
    "Case-Based Teaching",
    "Research / QI Update",
]

SLIDE_ROLES = [
    "Title",
    "Objectives",
    "Disclosures",
    "Introduction",
    "Story",
    "Evidence / Data",
    "Case",
    "Methods",
    "Results",
    "Discussion",
    "Application",
    "Take-home",
    "Custom / Unknown title",
]

BLOOM_HELPER = {
    "Remember": "define, list, identify, name, recall",
    "Understand": "describe, summarize, explain, classify, compare",
    "Apply": "use, demonstrate, calculate, choose, implement",
    "Analyze": "differentiate, organize, interpret, examine, contrast",
    "Evaluate": "appraise, justify, critique, prioritize, defend",
    "Create": "design, formulate, develop, construct, propose",
}

OBJECTIVE_EXAMPLES = [
    "Describe why this topic matters for pediatric practice.",
    "Differentiate the key clinical, educational, or research options.",
    "Appraise the evidence or reasoning that supports the main message.",
    "Apply the take-home points to a realistic patient, learner, or research decision.",
]

DEFAULT_DISCLOSURE = "I have no relevant financial or non-financial disclosures."


# -----------------------------------------------------------------------------
# Deck construction
# -----------------------------------------------------------------------------


def new_slide(
    role: str,
    title: str = "",
    prompt: str = "",
    required: bool = False,
    slide_kind: str = "content",
) -> Dict[str, Any]:
    """Return a new editable slide record.

    All slides are exported. There is intentionally no include/exclude flag.
    """
    return {
        "id": uuid.uuid4().hex,
        "role": role,
        "title": title,
        "subtitle": "",
        "prompt": prompt,
        "body": "",
        "visual_plan": "",
        "visual_image": {},
        "discussion_prompt": "",
        "speaker_notes": "",
        "slide_kind": slide_kind,
        "required": required,
    }


def title_slide() -> Dict[str, Any]:
    slide = new_slide("Title", "", "The title slide stores the presentation identity.", True, "title")
    slide["body"] = ""
    return slide


def objectives_slide() -> Dict[str, Any]:
    slide = new_slide("Objectives", "Objectives", "Use measurable Bloom-style verbs. Avoid vague verbs like understand or learn.", True, "objectives")
    slide["body"] = "Describe the clinical or scholarly problem and why it matters.\nAppraise the key evidence, logic, or data supporting the main message.\nApply the take-home points to a realistic clinical, educational, or research decision."
    return slide


def disclosures_slide() -> Dict[str, Any]:
    slide = new_slide("Disclosures", "Disclosures", "State relevant financial/non-financial disclosures, or state that there are none.", True, "disclosures")
    slide["body"] = DEFAULT_DISCLOSURE
    return slide


def starter_slides_for_talk_type(talk_type: str) -> List[Dict[str, Any]]:
    """Return a recommended story scaffold for the selected talk type."""
    base = [title_slide(), objectives_slide(), disclosures_slide()]

    if talk_type == "Study Review":
        base.extend(
            [
                new_slide("Introduction", "Why this study matters", "Open with the clinical or scientific gap.", True),
                new_slide("Methods", "How the investigators approached the question", "Summarize design, population, intervention/exposure, comparison, outcomes, and analysis."),
                new_slide("Results", "What they found", "Present the main result without data dumping."),
                new_slide("Discussion", "How much should we trust this?", "Appraise bias, precision, generalizability, and clinical relevance."),
                new_slide("Application", "How this changes our thinking", "Return to patient care, education, operations, policy, or the next research question."),
                new_slide("Take-home", "Take-home points", "End with 2-3 practical conclusions.", True),
            ]
        )
    elif talk_type == "Research / QI Update":
        base.extend(
            [
                new_slide("Introduction", "The problem we are trying to solve", "State the gap and why the audience should care.", True),
                new_slide("Methods", "What we did", "Describe setting, population, measures, intervention, or analysis plan."),
                new_slide("Results", "What we are seeing", "Show the main finding, run chart, table, or early signal."),
                new_slide("Discussion", "What this means", "Interpret the results and limitations."),
                new_slide("Application", "Next steps", "Clarify what feedback, decision, or action you need from the audience."),
                new_slide("Take-home", "Take-home points", "End with 2-3 practical conclusions.", True),
            ]
        )
    elif talk_type == "Case-Based Teaching":
        base.extend(
            [
                new_slide("Introduction", "Why this case matters", "Orient the audience to the clinical stakes.", True),
                new_slide("Case", "The patient story begins", "Present only the first chunk of information."),
                new_slide("Story", "What are you worried about now?", "Pause for differential, prioritization, or management decision."),
                new_slide("Evidence / Data", "Key data that changed the case", "Reveal labs, imaging, trends, or response to therapy."),
                new_slide("Application", "How we should approach this next time", "Convert the case into a generalizable approach."),
                new_slide("Take-home", "Take-home points", "End with 2-3 bedside-ready teaching points.", True),
            ]
        )
    else:
        base.extend(
            [
                new_slide("Introduction", "Why this matters", "Orient the audience: why should residents care in the first 60 seconds?", True),
                new_slide("Story", "The question we need to answer", "Create tension: clinical dilemma, knowledge gap, or decision point."),
                new_slide("Story", "Key idea 1", "Teach the first major concept."),
                new_slide("Story", "Key idea 2", "Teach the second major concept."),
                new_slide("Application", "How this changes our thinking", "Return to patient care, interpretation, or practical decision-making."),
                new_slide("Take-home", "Take-home points", "End with 2-3 actionable points.", True),
            ]
        )
    return base


def default_deck(talk_type: str = "Educational Topic") -> Dict[str, Any]:
    today = dt.date.today().isoformat()
    if talk_type not in TALK_TYPES:
        talk_type = "Educational Topic"
    return {
        "app_version": APP_VERSION,
        "metadata": {
            "presentation_title": "",
            "presenter": "",
            "session_date": today,
            "audience": "Pediatric residents",
            "presentation_type": talk_type,
            "core_question": "",
            "story_arc": "",
            "archive_notes": "",
        },
        "slides": starter_slides_for_talk_type(talk_type),
    }


# -----------------------------------------------------------------------------
# Text and persistence helpers
# -----------------------------------------------------------------------------


def split_nonempty_lines(text: Any) -> List[str]:
    return [line.strip() for line in str(text or "").splitlines() if line.strip()]


def sanitize_filename(value: str, default: str = "presentation") -> str:
    value = str(value or "").strip() or default
    value = re.sub(r"[^A-Za-z0-9._ -]+", "", value)
    value = re.sub(r"\s+", "_", value)
    return value[:90].strip("._-") or default


def short_label(text: str, max_len: int = 34) -> str:
    text = str(text or "").strip()
    return text if len(text) <= max_len else text[: max_len - 1].rstrip() + "…"


def slide_output_title(deck: Dict[str, Any], slide: Dict[str, Any], index: int) -> str:
    title = str(slide.get("title") or "").strip()
    if title:
        return title
    if slide.get("role") == "Title" or slide.get("slide_kind") == "title":
        return deck.get("metadata", {}).get("presentation_title") or "Untitled Presentation"
    if slide.get("role") == "Custom / Unknown title":
        return f"Untitled slide {index}"
    return f"{slide.get('role') or 'Slide'} {index}"


def identity_title(deck: Dict[str, Any]) -> str:
    return deck.get("metadata", {}).get("presentation_title") or "Untitled presentation"


def identity_subtitle(deck: Dict[str, Any]) -> str:
    meta = deck.get("metadata", {})
    return " · ".join(
        [
            meta.get("presenter") or "Presenter not entered",
            meta.get("session_date") or "Date not entered",
            meta.get("audience") or "Audience not entered",
            meta.get("presentation_type") or "Presentation type not entered",
        ]
    )


def make_archive_slug(deck: Dict[str, Any]) -> str:
    meta = deck.get("metadata", {})
    date = sanitize_filename(meta.get("session_date") or dt.date.today().isoformat(), "date")
    presenter = sanitize_filename(meta.get("presenter") or "presenter")
    title = sanitize_filename(meta.get("presentation_title") or "untitled_presentation")
    return f"{date}_{presenter}_{title}"


def to_json_bytes(deck: Dict[str, Any]) -> bytes:
    payload = copy.deepcopy(deck)
    payload["saved_at"] = dt.datetime.now().isoformat(timespec="seconds")
    return json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8")


def normalize_loaded_deck(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Merge a GitHub-loaded payload into the current app schema."""
    loaded = payload.get("deck", payload) if isinstance(payload, dict) else {}
    talk_type = loaded.get("metadata", {}).get("presentation_type", "Educational Topic") if isinstance(loaded, dict) else "Educational Topic"
    base = default_deck(talk_type)
    if not isinstance(loaded, dict):
        return base

    base["metadata"].update(loaded.get("metadata", {}))

    loaded_slides = loaded.get("slides")
    if isinstance(loaded_slides, list) and loaded_slides:
        normalized_slides: List[Dict[str, Any]] = []
        for raw in loaded_slides:
            if not isinstance(raw, dict):
                continue
            slide = new_slide(
                role=raw.get("role") or "Story",
                title=raw.get("title") or "",
                prompt=raw.get("prompt") or "",
                required=bool(raw.get("required", False)),
                slide_kind=raw.get("slide_kind") or "content",
            )
            slide["id"] = str(raw.get("id") or slide["id"])
            for key in ["subtitle", "body", "visual_plan", "discussion_prompt", "speaker_notes"]:
                slide[key] = raw.get(key, "")
            visual_image = raw.get("visual_image", {})
            slide["visual_image"] = visual_image if isinstance(visual_image, dict) else {}
            normalized_slides.append(slide)
        if normalized_slides:
            base["slides"] = normalized_slides
    return base
