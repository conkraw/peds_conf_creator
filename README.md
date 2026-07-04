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


Update v4.7: the single slide upload control now accepts either an image (PNG/JPEG) or a PPTX. If a PPTX is uploaded, the app uses the first slide of that file and replaces the corresponding exported PowerPoint slide. Images still support the half-slide/whole-slide behavior.


Update v4.8: added stronger PPTX-upload status messages, export-panel confirmation when a PPTX slide replacement is active, a fallback placeholder slide if replacement fails, and retained mentor DOCX notation for PPTX slide replacements.


Update v4.9: added an optional presentation subtitle on the title slide and tightened PPTX slide replacement packaging so imported slide masters are registered and source notes/comments are not copied, reducing PowerPoint repair prompts.


Update v5.0: PPTX uploads are now rendered to a full-slide PNG and inserted as a normal image instead of splicing PowerPoint XML. This is designed to avoid PowerPoint repair/read errors. Added packages.txt with libreoffice for Streamlit Cloud conversion support.


Update v5.1: PPTX uploads now render through a more reliable LibreOffice-to-PDF route, then PyMuPDF converts the first PDF page to PNG. This avoids PowerPoint XML splicing and is more reliable on Streamlit Cloud than direct PNG conversion.


Update v5.2: PPTX uploads now import the first uploaded slide as editable PowerPoint objects instead of rasterizing to PNG. This preserves editability for text boxes, images, tables, and many charts. Complex animations/transitions and some SmartArt may not fully preserve.


Update v5.3: added a dedicated Take-home points slide editor with five structured points, a numbered-circle PowerPoint layout, optional bottom discussion prompt, and support for an optional right-side visual upload.


Update v5.4: standard content slides now support an optional section box label. If entered, a blue header box appears above the main text block (for example, "Common resident questions"). If left blank, no box is shown.


Update v5.5: fixed GitHub-loaded presentations so the optional section box label and all other newer slide fields survive normalization. Added an on-screen confirmation when the section box is active and included the label in the mentor DOCX.


Update v5.6: expanded the mentor review Word document to include every presentation-level planning field (including core question, story arc, archive notes, and audience/type) and every slide-specific field, including helper prompts, structured objectives, five take-home points, section box labels, visuals, discussion prompts, and speaker notes.

Update v5.7: the mentor DOCX now has an unmistakable v5.7 generation marker, a complete presentation-plan/story-arc section, explicit structured objective and take-home fields, future-proof inclusion of additional metadata/slide fields, and an internal completeness check. The app disables the mentor download if the old/incomplete DOCX builder is still deployed, and downloads the file as mentor_review_v5_7.docx to avoid confusion with older downloads.


Update v5.8: PPTX uploads now generate a preview image for the app and the mentor DOCX, while still exporting as editable PowerPoint objects. Uploaded images and uploaded PPTX files can also be downloaded again from the slide editor.


Update v5.9: reduced uploaded image and PPTX preview sizes in the app and mentor DOCX. Source downloads and editable PPTX replacement behavior are unchanged.


Update v5.10: increased uploaded image and PPTX preview sizes slightly in both the app and mentor DOCX while keeping them compact.


Update v5.11: new presentation scaffolds now place Disclosures immediately after the Title slide, followed by Objectives.


Update v5.12: increased all speaker-note text areas in the app. Standard slides use a 260 px note box; title, objectives, disclosures, and take-home slides use a 220 px note box.


Update v5.13: slide text on standard content and disclosure slides now appears inside a compact light-blue rounded panel with a subtle border. The panel automatically grows to fit the bullet text instead of leaving bullets floating on the slide.


Update v5.15: slide text panels no longer automatically add bullet dots. Existing leading bullet/dash characters are stripped so statements appear as clean lines inside the colored box.


Update v5.16: redesigned the mentor DOCX into a streamlined review packet. It now begins with a compact presentation overview, then shows the actual rendered PowerPoint slide, editable on-slide wording, speaker notes, and a mentor feedback box for every slide. App-only implementation metadata and the Bloom reference table were removed from the main review flow.


Update v5.17: disclosure slide text is now centered inside its content panel, without left alignment or bullets. This applies only to the Disclosures slide.


Update v5.18: redesigned the mentor review as a compact portrait table document. Slide previews are smaller, editable slide wording and speaker notes appear beside the preview, mentor feedback remains editable, and slides flow continuously instead of forcing one or two pages per slide.
