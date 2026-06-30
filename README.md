# Pediatric Residency Presentation Builder

A Streamlit app for standardized, story-driven educational and study-review PowerPoint creation.

## Features

- Required slide structure: title, objectives, disclosures, introduction, story setup, core content, application, and take-home points
- Bloom's taxonomy helper for objectives
- Flexible extra story slides, including blank/unknown slide titles
- Mentor review section with checklist, feedback, approval statement, and optional mentor-review slide
- Speaker notes that export into real PowerPoint speaker notes
- Editable JSON draft download/reload
- Optional GitHub archive for both `draft.json` and `presentation.pptx`

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit secrets for GitHub saving

Add these to Streamlit Cloud secrets:

```toml
GITHUB_TOKEN = "ghp_your_token"
GITHUB_REPO = "your_username/your_repo"
GITHUB_BRANCH = "main"
GITHUB_FOLDER = "presentation_archive"
```

The app saves each presentation into:

```text
presentation_archive/YYYY-MM-DD_presenter_title/draft.json
presentation_archive/YYYY-MM-DD_presenter_title/presentation.pptx
```
