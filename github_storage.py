"""
GitHub persistence helpers for the Presentation Builder app.

Expected Streamlit secrets:
GITHUB_TOKEN = "ghp_..."
GITHUB_REPO = "owner/repository"
GITHUB_BRANCH = "main"
GITHUB_FOLDER = "presentation_archive"
"""

from __future__ import annotations

import base64
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

import requests
import streamlit as st

from deck_model import ARCHIVE_DOCX_NAME, ARCHIVE_JSON_NAME, ARCHIVE_PPTX_NAME, load_deck_from_json


@dataclass
class GitHubConfig:
    token: str
    repo: str
    branch: str
    folder: str


def github_config_from_secrets() -> Optional[GitHubConfig]:
    try:
        token = st.secrets.get("GITHUB_TOKEN", "")
        repo = st.secrets.get("GITHUB_REPO", "")
        branch = st.secrets.get("GITHUB_BRANCH", "main")
        folder = st.secrets.get("GITHUB_FOLDER", "presentation_archive")
    except Exception:
        return None

    if not token or not repo:
        return None
    return GitHubConfig(token=token, repo=repo, branch=branch, folder=folder.strip("/"))


def github_headers(config: GitHubConfig) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {config.token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def github_api_url(config: GitHubConfig, path: str) -> str:
    safe_path = "/".join(part.strip("/") for part in path.split("/") if part.strip("/"))
    return f"https://api.github.com/repos/{config.repo}/contents/{safe_path}"


def github_get_file_sha(config: GitHubConfig, path: str) -> Optional[str]:
    response = requests.get(
        github_api_url(config, path),
        headers=github_headers(config),
        params={"ref": config.branch},
        timeout=30,
    )
    if response.status_code == 404:
        return None
    response.raise_for_status()
    data = response.json()
    return data.get("sha")


def github_upsert_bytes(config: GitHubConfig, path: str, content: bytes, message: str) -> None:
    sha = github_get_file_sha(config, path)
    payload: Dict[str, Any] = {
        "message": message,
        "content": base64.b64encode(content).decode("ascii"),
        "branch": config.branch,
    }
    if sha:
        payload["sha"] = sha
    response = requests.put(
        github_api_url(config, path),
        headers=github_headers(config),
        json=payload,
        timeout=60,
    )
    response.raise_for_status()


def github_list_json_drafts(config: GitHubConfig) -> List[Dict[str, Any]]:
    response = requests.get(
        github_api_url(config, config.folder),
        headers=github_headers(config),
        params={"ref": config.branch},
        timeout=30,
    )
    if response.status_code == 404:
        return []
    response.raise_for_status()
    items = response.json()
    if not isinstance(items, list):
        return []

    drafts: List[Dict[str, Any]] = []
    for item in items:
        if item.get("type") != "dir":
            continue
        folder_path = item.get("path", "")
        draft_path = f"{folder_path}/{ARCHIVE_JSON_NAME}"
        draft_resp = requests.get(
            github_api_url(config, draft_path),
            headers=github_headers(config),
            params={"ref": config.branch},
            timeout=30,
        )
        if draft_resp.status_code == 200:
            drafts.append({"name": item.get("name"), "path": draft_path})
    return sorted(drafts, key=lambda item: item["name"], reverse=True)


def github_load_json(config: GitHubConfig, path: str) -> Dict[str, Any]:
    response = requests.get(
        github_api_url(config, path),
        headers=github_headers(config),
        params={"ref": config.branch},
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    raw = base64.b64decode(data["content"])
    return load_deck_from_json(raw)


def save_full_archive(
    config: GitHubConfig,
    folder_slug: str,
    json_bytes: bytes,
    pptx_bytes: bytes,
    docx_bytes: bytes,
) -> str:
    """Save JSON, PPTX, and DOCX to a single GitHub archive folder."""
    folder = f"{config.folder}/{folder_slug}"
    github_upsert_bytes(config, f"{folder}/{ARCHIVE_JSON_NAME}", json_bytes, f"Save presentation draft: {folder_slug}")
    github_upsert_bytes(config, f"{folder}/{ARCHIVE_PPTX_NAME}", pptx_bytes, f"Save presentation PPTX: {folder_slug}")
    github_upsert_bytes(config, f"{folder}/{ARCHIVE_DOCX_NAME}", docx_bytes, f"Save presentation planning form: {folder_slug}")
    return folder
