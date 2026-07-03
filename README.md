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


Update v3.3: removed the Visual / evidence plan editor box; uploaded slide visuals now occupy about half of the exported PowerPoint slide.


Update v3.4: mentor review uploaded visuals are centered, constrained to fit within the table cell, and no longer display the uploaded file name.


Update v3.5: fixed Streamlit sidebar slide-selection state so Add after/Add at end/Duplicate/Delete no longer mutate the radio widget state after it renders.


Update v3.6: archive search now normalizes spaces/underscores/hyphens and searches draft metadata, so searching for names like "Conrad Krawiec" finds folders such as `Conrad_Krawiec`.


Update v3.7: added optional title-slide visual upload; when used, the PowerPoint title slide shows text on the left and the uploaded image on the right.


Update v3.8: title-slide Core question / Story arc panel now estimates wrapped text height, expands the pale-blue box, and reduces font size when needed so the box encircles the words instead of clipping or overflowing.


Update v3.9: fixed GitHub loading for larger draft.json files by falling back to GitHub download_url when the Contents API omits inline file content. This can happen after image uploads are stored in the draft.


Update v4.0: added GitHub archive deletion controls and an optional whole-slide visual layout for uploaded images.


Update v4.1: GitHub save confirmation now displays only saved file names instead of full archive paths.


Update v4.2: moved the sidebar Add slides controls into a collapsed dropdown/expander so they stay out of the way until needed.


Update v4.4: fixed startup robustness by computing Bloom objective dropdown options in app.py, kept GitHub archive collapsed in a dropdown, and preserved the visual objectives slide layout.


Update v4.5: refined the exported Objectives slide to match the visual reference more closely with a clean title on white background, intro line, divider, three numbered objective cards, and a bottom banner.


Update v4.6: added a visible app-version marker in the sidebar and a visible objectives-card-layout confirmation in the Objectives editor so deployed/stale Streamlit versions are immediately obvious.
