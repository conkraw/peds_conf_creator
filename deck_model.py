"""
Shared data model for the Presentation Builder app.

This file intentionally contains no Streamlit code. It keeps the app front-end,
PowerPoint builder, Word builder, and GitHub storage modules cleanly separated.
"""

from __future__ import annotations

import copy
import datetime as dt
import json
import re
import uuid
from typing import Any, Dict, List

APP_TITLE = "Pediatric Residency Presentation Builder"
APP_VERSION = "2026.06.30-v2"
ARCHIVE_JSON_NAME = "draft.json"
ARCHIVE_PPTX_NAME = "presentation.pptx"
ARCHIVE_DOCX_NAME = "planning_form.docx"

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
    "Mentor review",
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
    "Describe the clinical or research problem and why it matters for pediatric practice.",
    "Differentiate the key diagnostic, management, or methodological options.",
    "Appraise the strengths and limitations of the evidence being presented.",
    "Apply the take-home points to a realistic clinical, educational, or research decision.",
]

CHECKLIST_LABELS = {
    "objectives_use_bloom": "Objectives use Bloom-style measurable verbs",
    "story_arc_clear": "The presentation has a clear beginning, middle, and end",
    "slide_titles_tell_story": "Slide titles tell the story rather than only labeling topics",
    "one_job_per_slide": "Each slide has one main job",
    "no_data_dump": "Slides avoid data dumping",
    "speaker_notes_complete": "Speaker notes are complete enough for rehearsal/review",
    "visuals_support_message": "Figures/tables/images support the message rather than distract",
    "take_home_points_clear": "Final take-home points are clear and actionable",
}


def new_slide(
    role: str,
    title: str,
    prompt: str = "",
    required: bool = False,
    slide_kind: str = "content",
) -> Dict[str, Any]:
    """Return a new editable slide record."""
    return {
        "id": str(uuid.uuid4()),
        "role": role,
        "title": title,
        "subtitle": "",
        "prompt": prompt,
        "body": "",
        "visual_plan": "",
        "discussion_prompt": "",
        "speaker_notes": "",
        "slide_kind": slide_kind,
        "include": True,
        "required": required,
    }


def starter_slides_for_talk_type(talk_type: str) -> List[Dict[str, Any]]:
    """Return a recommended story scaffold for the selected talk type."""
    base = [
        new_slide("Title", "", "Name the talk clearly and state the problem.", required=True, slide_kind="title"),
        new_slide("Objectives", "Objectives", "Use measurable verbs from Bloom's taxonomy.", required=True, slide_kind="objectives"),
        new_slide("Disclosures", "Disclosures", "State relevant financial/non-financial disclosures or no relevant disclosures.", required=True),
    ]

    if talk_type == "Study Review":
        base.extend(
            [
                new_slide("Introduction", "Why this study matters", "Open with the clinical or scientific gap.", required=True),
                new_slide("Methods", "How the investigators approached the question", "Summarize design, population, exposure/intervention, outcomes, and comparison.", required=False),
                new_slide("Results", "What they found", "Present the main results without data dumping.", required=False),
                new_slide("Discussion", "How much should we trust this?", "Appraise bias, confounding, precision, generalizability, and clinical relevance.", required=False),
                new_slide("Application", "How this changes our thinking", "Return to practice, policy, education, or next research question.", required=False),
                new_slide("Take-home", "Take-home points", "End with 2-3 actionable points.", required=True),
            ]
        )
    elif talk_type == "Research / QI Update":
        base.extend(
            [
                new_slide("Introduction", "The problem we are trying to solve", "State the gap and why the audience should care.", required=True),
                new_slide("Methods", "What we did", "Describe setting, population, measures, intervention, or analysis plan.", required=False),
                new_slide("Results", "What we are seeing so far", "Show the main finding, run chart, table, or early signal.", required=False),
                new_slide("Discussion", "What this means", "Interpret the results and limitations.", required=False),
                new_slide("Application", "Next steps", "Clarify what feedback, decision, or action you need from the audience.", required=False),
                new_slide("Take-home", "Take-home points", "End with 2-3 practical conclusions.", required=True),
            ]
        )
    elif talk_type == "Case-Based Teaching":
        base.extend(
            [
                new_slide("Introduction", "Why this case matters", "Orient the audience to the clinical stakes.", required=True),
                new_slide("Case", "The patient story begins", "Present only the first chunk of information.", required=False),
                new_slide("Story", "What are you worried about now?", "Pause for differential, prioritization, or management decision.", required=False),
                new_slide("Evidence / Data", "Key data that changed the case", "Reveal labs, imaging, trends, or response to therapy.", required=False),
                new_slide("Application", "How we should approach this next time", "Convert the case into a generalizable approach.", required=False),
                new_slide("Take-home", "Take-home points", "End with 2-3 bedside-ready teaching points.", required=True),
            ]
        )
    else:
        base.extend(
            [
                new_slide("Introduction", "Why this matters", "Orient the audience: why should residents care in the first 60 seconds?", required=True),
                new_slide("Story", "The question we need to answer", "Create tension: clinical dilemma, knowledge gap, or decision point.", required=False),
                new_slide("Story", "Key idea 1", "Teach the first major concept.", required=False),
                new_slide("Story", "Key idea 2", "Teach the second major concept.", required=False),
                new_slide("Application", "How this changes our thinking", "Return to patient care, interpretation, or practical decision-making.", required=False),
                new_slide("Take-home", "Take-home points", "End with 2-3 actionable points.", required=True),
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
            "github_notes": "",
        },
        "mentor_review": {
            "mentor_name": "",
            "mentor_email": "",
            "review_status": "Not sent",
            "requested_review_date": "",
            "review_completed_date": "",
            "mentor_feedback": "",
            "mentor_approval_statement": "",
            "include_mentor_review_slide": False,
            "checklist": {key: False for key in CHECKLIST_LABELS},
        },
        "slides": starter_slides_for_talk_type(talk_type),
    }


def split_nonempty_lines(text: str) -> List[str]:
    return [line.strip() for line in (text or "").splitlines() if line.strip()]


def sanitize_filename(value: str, default: str = "presentation") -> str:
    value = (value or "").strip() or default
    value = re.sub(r"[^A-Za-z0-9._ -]+", "", value)
    value = re.sub(r"\s+", "_", value)
    return value[:90].strip("._-") or default


def short_label(text: str, max_len: int = 34) -> str:
    text = (text or "").strip()
    return text if len(text) <= max_len else text[: max_len - 1].rstrip() + "…"


def slide_output_title(deck: Dict[str, Any], slide: Dict[str, Any], index: int) -> str:
    title = (slide.get("title") or "").strip()
    if title:
        return title
    if slide.get("role") == "Title":
        return deck.get("metadata", {}).get("presentation_title") or "Untitled Presentation"
    return f"Story slide {index}"


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


def load_deck_from_json(uploaded_bytes: bytes) -> Dict[str, Any]:
    loaded = json.loads(uploaded_bytes.decode("utf-8"))
    base = default_deck(loaded.get("metadata", {}).get("presentation_type", "Educational Topic"))
    base.update(loaded)
    base.setdefault("metadata", default_deck()["metadata"])
    base.setdefault("mentor_review", default_deck()["mentor_review"])
    base.setdefault("slides", default_deck()["slides"])

    # Backfill newer mentor checklist keys if a draft was created with an older version.
    checklist = base["mentor_review"].setdefault("checklist", {})
    for key in CHECKLIST_LABELS:
        checklist.setdefault(key, False)
    return base
