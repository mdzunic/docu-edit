---
id: SPEC-019
title: Paragraph styles and the style picker
status: draft
area: formatting
depends_on: [SPEC-013, SPEC-014]
---
## Why
Styles are how long documents stay consistent, how a table of contents knows
what a heading is, and how a screen reader navigates. Direct formatting alone
produces documents that cannot be restructured.

## Acceptance criteria
- [ ] The toolbar has a style picker showing the current paragraph's style name
- [ ] The picker lists Normal, Title, Subtitle, Heading 1–6, Quote and Code, each preview rendered in its own style
- [ ] Choosing a style applies it to every paragraph in the selection and updates rendering immediately
- [ ] Ctrl/Cmd+Alt+1..6 apply Heading 1–6; Ctrl/Cmd+Alt+0 applies Normal
- [ ] Styles opened from a document are read from `styles.xml`, including inheritance (`w:basedOn` / `style:parent-style-name`), and the picker lists the document's own styles rather than a hard-coded set
- [ ] Direct formatting layered on a styled paragraph is preserved and takes precedence in rendering, and both survive a save
- [ ] A style entry's context menu offers "Update <style> to match" which redefines the style from the current paragraph's formatting and re-renders every paragraph using it
- [ ] "Reset style" removes direct formatting from the selection, leaving the style's own formatting
- [ ] Applying a heading style makes the paragraph a real heading in the saved file: `w:pStyle` referencing the built-in heading style for `.docx`, `text:h` with the right outline level for `.odt`
- [ ] Rendered headings expose correct heading semantics (h1–h6) to assistive technology

## Decisions
If a document has styles defined but lacks a 'Normal' style, does the picker show a fallback 'Normal' entry alongside the document's styles, or is 'Normal' only present when the document has zero styles? AC-4 implies the latter, but AC-7 implies a fallback Normal is always available. - Normal is always available.

## Out of scope
Creating and managing custom styles (SPEC-042). Character styles and list
styles. Style sets / theme switching.

## Note (SPEC-074)
Borders and shading join the properties a paragraph style carries: a
paragraph inherits its style's `borders`/`shadingFill` until it sets its own,
which falls out of this spec's existing baked-formatting model — the
direct-formatting diff compares `borders` structurally (an identity check
would report every freshly-baked object as direct formatting).
