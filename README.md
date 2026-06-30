# Pediatric Residency Presentation Builder

A Streamlit app for building standardized, story-driven PowerPoint presentations with:

- separate `pptx_builder.py` PowerPoint exporter
- separate `docx_builder.py` Word planning-form exporter
- real PowerPoint speaker notes
- Bloom's taxonomy objective helper
- disclosures slide
- flexible extra slides, including unknown/blank slide titles
- mentor review checklist and optional mentor-review slide
- GitHub save/load for JSON drafts, PowerPoint files, and Word planning forms

## Files

```text
app.py              # Streamlit UI only
deck_model.py       # Shared schema, defaults, constants, helpers
pptx_builder.py     # PowerPoint builder and real speaker-note injection
docx_builder.py     # Word planning-form builder
github_storage.py   # GitHub save/load helpers
requirements.txt
README.md
```

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit secrets for GitHub

Add this to Streamlit secrets:

```toml
GITHUB_TOKEN = "ghp_your_token"
GITHUB_REPO = "your_username/your_repo"
GITHUB_BRANCH = "main"
GITHUB_FOLDER = "presentation_archive"
```

The app saves one folder per presentation:

```text
presentation_archive/YYYY-MM-DD_presenter_title/draft.json
presentation_archive/YYYY-MM-DD_presenter_title/presentation.pptx
presentation_archive/YYYY-MM-DD_presenter_title/planning_form.docx
```

## Notes

The PowerPoint builder uses `python-pptx` for slide construction and then post-processes the PPTX Office Open XML package to add real speaker notes. The notes entered in the app should appear in PowerPoint presenter view.
