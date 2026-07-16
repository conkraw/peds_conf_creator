

## v5.25
- Added a button-activated **Import JSON draft** panel in the sidebar.
- Supports plain presentation-builder JSON, wrapped GitHub archive JSON, and common `content` wrappers.
- Restores custom slides, speaker notes, images, and embedded editable PPTX slide replacements.
- Opens uploads as a new working copy by default so autosave does not overwrite the original archive.
- Offers an explicit opt-in to reconnect autosave to the archive path stored in the JSON.
- Includes validation and clear error messages for empty, malformed, or incompatible JSON files.

Update v5.26: added one optional presentation-level reference file upload on the Title slide. Any file type can be uploaded (PDF, DOCX, spreadsheet, image, text, etc.). The source file is saved separately inside the presentation's GitHub archive, can be downloaded again from the app, and is not inserted into the PowerPoint or mentor DOCX. The draft keeps only lightweight metadata after the file is archived so autosave remains responsive.
