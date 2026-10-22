---
id: SPEC-024
title: Images — insert, resize and wrap
status: draft
area: structure
depends_on: [SPEC-016, SPEC-022]
---
## Why
Documents contain pictures — logos, screenshots, diagrams. Images are also the
first binary content in the package, so this spec establishes how media parts
are read, stored and written for both formats.

## Acceptance criteria
- [ ] "Insert image" accepts a file picker, drag-and-drop onto the canvas, and paste from the clipboard, for PNG, JPEG, GIF and SVG
- [ ] The inserted image appears at the caret, scaled down to fit the text area width if larger, at its natural aspect ratio
- [ ] Clicking an image selects it, showing eight resize handles and a floating image toolbar
- [ ] Dragging a corner handle resizes proportionally; dragging an edge handle resizes one dimension; Shift constrains to aspect ratio
- [ ] The image toolbar offers wrap mode: in line with text, wrap text (square), and behind/in front of text, with the rendering matching the chosen mode
- [ ] An "Image options" panel sets alt text (description), size in cm with a "lock aspect ratio" toggle, and a "Reset size" button
- [ ] Alt text is written to `.docx` (`wp:docPr/@descr`) and `.odt` (`svg:desc`) and read back on open
- [ ] Images are stored in the package as media parts with correct content types and relationships/manifest entries, not as base64 data URIs in the XML
- [ ] Images in opened `.docx` and `.odt` fixtures render at the size the source application shows, including one anchored (floating) image and one inline image
- [ ] Deleting a selected image removes it and its media part on save; an image used twice is not deleted from the package while one use remains
- [ ] Images larger than 10 MB are rejected with "That image is too large (limit 10 MB)"

## Out of scope
Cropping, rotation, borders and image effects. Charts and drawing shapes. Text
boxes. Image compression on save. Linked (external URL) images.
