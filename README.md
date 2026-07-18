

## v5.25
- Added a button-activated **Import JSON draft** panel in the sidebar.
- Supports plain presentation-builder JSON, wrapped GitHub archive JSON, and common `content` wrappers.
- Restores custom slides, speaker notes, images, and embedded editable PPTX slide replacements.
- Opens uploads as a new working copy by default so autosave does not overwrite the original archive.
- Offers an explicit opt-in to reconnect autosave to the archive path stored in the JSON.
- Includes validation and clear error messages for empty, malformed, or incompatible JSON files.

Update v5.26: added one optional presentation-level reference file upload on the Title slide. Any file type can be uploaded (PDF, DOCX, spreadsheet, image, text, etc.). The source file is saved separately inside the presentation's GitHub archive, can be downloaded again from the app, and is not inserted into the PowerPoint or mentor DOCX. The draft keeps only lightweight metadata after the file is archived so autosave remains responsive.


Update v5.28: clarified the difference between reference files and slide visuals. Reference files are now labeled as not inserted into PowerPoint, title-slide visual upload is clearly labeled as appearing on Slide 1, reference images can be copied into the title-slide visual with one button, and prepared exports are invalidated whenever visual/reference assets change.


## v5.28 title-image reliability fix

- Slide 1 now has a dedicated presentation-level image field instead of sharing the generic per-slide visual state.
- Older JSON title images migrate automatically.
- A newly uploaded title image clears stale title-slide images and PPTX replacements.
- The app shows the exact active filename and preview before export.
- The title-image uploader resets immediately after a successful replacement.


Update v5.29: hardened Slide 1 image replacement. Title image uploads now update both the dedicated title-image metadata and the legacy Slide 1 visual field, clear export caches, invalidate prepared files, and use download buttons keyed to the current deck signature to avoid stale PPTX downloads.


## v5.30

- Moved the presentation reference-file uploader into a dedicated app-only navigation section after the final slide.
- With a 19-slide deck, it appears as **20. Reference file — app only**.
- The reference section is not added to `deck["slides"]` and therefore cannot appear in the PowerPoint or mentor DOCX.
- Removed the reference uploader from the Title-slide editor so it is easier to find and cannot be confused with the Slide 1 visual uploader.
