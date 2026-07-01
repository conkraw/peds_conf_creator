# Pediatric Residency Presentation Builder

A Streamlit app for building standardized, story-driven PowerPoint presentations.

## What this version does

- Keeps `app.py` as the Streamlit front end.
- Uses `deck_model.py` for the shared slide schema and defaults.
- Uses `pptx_builder.py` for PowerPoint export.
- Uses `docx_builder.py` for the mentor review Word document.
- Uses `github_storage.py` for GitHub archive save/load.
- Exports real PowerPoint speaker notes.
- Uses stable slide-ID navigation so sidebar radio buttons respond on one click, even when slide titles change.
- Lets users upload one optional PNG/JPEG visual per content slide; uploaded visuals appear in the PowerPoint, mentor DOCX, and GitHub draft.
- Exports every slide automatically. There is no include/exclude checkbox.
- Keeps presentation identity on the exported title slide only, not as a footer on every slide.
- Uses the mentor review DOCX as the place for mentor critiques. There is no mentor-review form inside the app.
- Removes local JSON upload/download buttons. Draft JSON is stored only in GitHub.

## GitHub storage setup

Create `.streamlit/secrets.toml` with:

```toml
[github]
token = "ghp_your_token_here"
repo = "your_username/your_repo"
branch = "main"
base_path = "presentation_archive"
```

The app saves each presentation as:

```text
presentation_archive/YYYY-MM-DD_presenter_title/draft.json
presentation_archive/YYYY-MM-DD_presenter_title/presentation.pptx
presentation_archive/YYYY-MM-DD_presenter_title/mentor_review.docx
```

Uploaded slide visuals are stored inside `draft.json` as base64 so they reload from GitHub with the rest of the presentation. Keep images cropped/compressed when possible; the app limits uploads to 5 MB per image.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Suggested workflow

1. Build and revise the presentation in the app.
2. Download the mentor Word document and send it for comments/Track Changes.
3. Make mentor-requested revisions back in the app.
4. Export PowerPoint.
5. Save the full archive to GitHub.

## Notes

PowerPoint should only be used for minor spacing/layout tweaks after export. Major content changes should happen in the app so the GitHub draft remains the source of truth.
