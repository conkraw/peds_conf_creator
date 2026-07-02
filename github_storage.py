"""GitHub persistence helpers for the Presentation PowerPoint Builder."""

from __future__ import annotations

import base64
import json
from dataclasses import dataclass
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
import streamlit as st

from deck_model import ARCHIVE_DOCX_NAME, ARCHIVE_JSON_NAME, ARCHIVE_PPTX_NAME, APP_VERSION, make_archive_slug, to_json_bytes


class GitHubStorageError(RuntimeError):
    """Raised when GitHub storage cannot save or load."""


@dataclass
class GitHubFileResult:
    path: str
    html_url: str
    commit_sha: str


def _read_github_config() -> Dict[str, str]:
    try:
        raw = st.secrets.get("github", {})
    except Exception:
        raw = {}
    return {
        "token": str(raw.get("token", "")).strip(),
        "repo": str(raw.get("repo", "")).strip(),
        "branch": str(raw.get("branch", "main")).strip() or "main",
        "base_path": str(raw.get("base_path", "presentation_archive")).strip().strip("/") or "presentation_archive",
    }


def _headers(token: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def github_is_configured() -> bool:
    cfg = _read_github_config()
    return bool(cfg["token"] and cfg["repo"] and "/" in cfg["repo"])


def github_status_message() -> str:
    cfg = _read_github_config()
    missing = []
    if not cfg["token"]:
        missing.append("github.token")
    if not cfg["repo"] or "/" not in cfg["repo"]:
        missing.append("github.repo")
    if missing:
        return "Missing Streamlit secrets: " + ", ".join(missing)
    return "GitHub archive is configured."


def _api_url(path: str) -> str:
    cfg = _read_github_config()
    api_path = quote(path.strip().lstrip("/"), safe="/")
    return f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"


def _get_existing_sha(path: str) -> Optional[str]:
    cfg = _read_github_config()
    response = requests.get(_api_url(path), headers=_headers(cfg["token"]), params={"ref": cfg["branch"]}, timeout=30)
    if response.status_code == 404:
        return None
    if response.status_code != 200:
        raise GitHubStorageError(f"Could not check GitHub file ({response.status_code}): {response.text}")
    return response.json().get("sha")


def save_file_bytes_to_github(path: str, content: bytes, commit_message: str) -> GitHubFileResult:
    cfg = _read_github_config()
    if not github_is_configured():
        raise GitHubStorageError(github_status_message())

    clean_path = path.strip().lstrip("/")
    sha = _get_existing_sha(clean_path)
    payload: Dict[str, Any] = {
        "message": commit_message,
        "content": base64.b64encode(content).decode("utf-8"),
        "branch": cfg["branch"],
    }
    if sha:
        payload["sha"] = sha

    response = requests.put(_api_url(clean_path), headers=_headers(cfg["token"]), json=payload, timeout=45)
    if response.status_code not in (200, 201):
        raise GitHubStorageError(f"GitHub save failed ({response.status_code}): {response.text}")
    data = response.json()
    content_data = data.get("content", {}) or {}
    commit_data = data.get("commit", {}) or {}
    return GitHubFileResult(path=clean_path, html_url=content_data.get("html_url", ""), commit_sha=commit_data.get("sha", ""))


def build_archive_payload(deck: Dict[str, Any], archive_path: str) -> bytes:
    payload = json.loads(to_json_bytes(deck).decode("utf-8"))
    payload["app_version"] = APP_VERSION
    payload["archive_path"] = archive_path
    return json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8")


def save_archive_to_github(deck: Dict[str, Any], pptx_bytes: bytes, mentor_docx_bytes: bytes, existing_archive_path: str = "") -> List[GitHubFileResult]:
    cfg = _read_github_config()
    if not github_is_configured():
        raise GitHubStorageError(github_status_message())

    archive_path = existing_archive_path.strip().strip("/") or f"{cfg['base_path']}/{make_archive_slug(deck)}"
    results = [
        save_file_bytes_to_github(f"{archive_path}/{ARCHIVE_JSON_NAME}", build_archive_payload(deck, archive_path), "Save presentation builder draft"),
        save_file_bytes_to_github(f"{archive_path}/{ARCHIVE_PPTX_NAME}", pptx_bytes, "Save generated presentation PPTX"),
        save_file_bytes_to_github(f"{archive_path}/{ARCHIVE_DOCX_NAME}", mentor_docx_bytes, "Save mentor review DOCX"),
    ]
    return results


def _github_get_contents(path: str) -> Any:
    cfg = _read_github_config()
    if not github_is_configured():
        raise GitHubStorageError(github_status_message())
    response = requests.get(_api_url(path), headers=_headers(cfg["token"]), params={"ref": cfg["branch"]}, timeout=30)
    if response.status_code == 404:
        return []
    if response.status_code != 200:
        raise GitHubStorageError(f"Could not list GitHub archive ({response.status_code}): {response.text}")
    return response.json()


def _search_normalize(value: Any) -> str:
    """Normalize labels/paths so searches match underscores, hyphens, and spaces."""
    import re

    text = str(value or "").lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _matches_search(search_text: str, *values: Any) -> bool:
    """Return True when every search token appears somewhere in the archive metadata."""
    tokens = _search_normalize(search_text).split()
    if not tokens:
        return True
    haystack = _search_normalize(" ".join(str(value or "") for value in values))
    return all(token in haystack for token in tokens)


def _draft_metadata_for_archive(path: str) -> Dict[str, Any]:
    """Read lightweight metadata from an archive's draft.json when available."""
    try:
        data = _github_get_contents(f"{path.strip().rstrip('/')}/{ARCHIVE_JSON_NAME}")
        if not isinstance(data, dict):
            return {}
        encoded = str(data.get("content", "")).replace("\n", "")
        if not encoded:
            return {}
        decoded = base64.b64decode(encoded).decode("utf-8")
        payload = json.loads(decoded)
        deck = payload.get("deck", payload) if isinstance(payload, dict) else {}
        meta = deck.get("metadata", {}) if isinstance(deck, dict) else {}
        if not isinstance(meta, dict):
            meta = {}
        return {
            "presentation_title": str(meta.get("presentation_title", "")),
            "presenter": str(meta.get("presenter", "")),
            "session_date": str(meta.get("session_date", "")),
            "audience": str(meta.get("audience", "")),
            "presentation_type": str(meta.get("presentation_type", "")),
            "saved_at": str(payload.get("saved_at", "")) if isinstance(payload, dict) else "",
        }
    except Exception:
        return {}


def _archive_display_name(folder_name: str, metadata: Dict[str, Any]) -> str:
    """Build a readable archive label while keeping the folder name as fallback."""
    date = str(metadata.get("session_date", "")).strip()
    presenter = str(metadata.get("presenter", "")).strip()
    title = str(metadata.get("presentation_title", "")).strip()
    parts = [part for part in [date, presenter, title] if part]
    return " · ".join(parts) if parts else folder_name


def list_archives_from_github(search_text: str = "") -> List[Dict[str, str]]:
    cfg = _read_github_config()
    rows = _github_get_contents(cfg["base_path"])
    if not isinstance(rows, list):
        return []

    archives: List[Dict[str, str]] = []
    for item in rows:
        if item.get("type") != "dir":
            continue
        folder_name = str(item.get("name", ""))
        path = str(item.get("path", ""))
        metadata = _draft_metadata_for_archive(path)

        if not _matches_search(
            search_text,
            folder_name,
            path,
            metadata.get("presentation_title", ""),
            metadata.get("presenter", ""),
            metadata.get("session_date", ""),
            metadata.get("audience", ""),
            metadata.get("presentation_type", ""),
        ):
            continue

        archives.append(
            {
                "name": _archive_display_name(folder_name, metadata),
                "path": path,
                "html_url": item.get("html_url", ""),
                "sort_key": metadata.get("saved_at") or metadata.get("session_date") or folder_name,
            }
        )
    return sorted(archives, key=lambda row: row.get("sort_key", row["name"]), reverse=True)


def load_json_from_github(path: str) -> Dict[str, Any]:
    clean = path.strip().lstrip("/")
    if not clean.endswith(".json"):
        clean = f"{clean.rstrip('/')}/{ARCHIVE_JSON_NAME}"
    data = _github_get_contents(clean)
    if not isinstance(data, dict):
        raise GitHubStorageError("Selected GitHub path did not return a JSON file.")
    encoded = str(data.get("content", "")).replace("\n", "")
    if not encoded:
        raise GitHubStorageError("GitHub JSON file did not contain content.")
    try:
        decoded = base64.b64decode(encoded).decode("utf-8")
        return json.loads(decoded)
    except Exception as exc:
        raise GitHubStorageError(f"GitHub draft is not valid JSON: {exc}") from exc
