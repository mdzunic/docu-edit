---
id: SPEC-048
title: Font fallback and embedding
status: draft
area: platform
depends_on: [SPEC-013, SPEC-034]
---
## Why
A document written in Calibri on Windows and opened on a Mac reflows: different
metrics, different line breaks, different page count. Font handling is the last
big source of "it looks different than in Word", and the fix is metric-
compatible substitution plus honest disclosure of what we swapped.

## Acceptance criteria
- [ ] The app bundles metric-compatible substitutes for the common Microsoft fonts (Carlito for Calibri, Caladea for Cambria, Liberation Sans/Serif/Mono for Arial/Times New Roman/Courier New) and uses them automatically when the original is unavailable
- [ ] Substituted text keeps the original font name in the document model and writes the original name on save — substitution is a rendering decision, never a content change
- [ ] Line breaks and page count for a `.docx` fixture using Calibri match Word's within one line per page
- [ ] A "Fonts" panel lists every font used in the document, marking each as available, substituted (naming the substitute), or missing
- [ ] Text in a genuinely missing font falls back to a documented default and appears in the Fonts panel as missing
- [ ] Fonts embedded in an opened `.docx` (`w:embedRegular` and friends) or `.odt` are extracted and used for rendering, subject to their embedding permission bits, and are preserved on save
- [ ] A "Embed fonts in this file" option on save writes subset embedded fonts for the fonts used, and states the size cost before saving
- [ ] Fonts whose licence bits forbid embedding are skipped, with a named list shown to the user
- [ ] Font loading does not cause a flash of unstyled text — the page shows measured, correct layout on first paint or a skeleton until fonts resolve
- [ ] PDF export (SPEC-034) uses the same resolved fonts, so a PDF matches the on-screen layout

## Out of scope
Uploading arbitrary user fonts. Web font services. Complex-script shaping
beyond what the browser provides. CJK font subsetting quality tuning.
