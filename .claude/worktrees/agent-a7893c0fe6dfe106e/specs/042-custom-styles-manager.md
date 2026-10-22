---
id: SPEC-042
title: Custom styles manager
status: draft
area: formatting
depends_on: [SPEC-019, SPEC-032]
---
## Why
SPEC-019 gave users the built-in styles. Teams with a house template need their
own — and need them to survive a round trip so the document still matches the
brand when it comes back from Word.

## Acceptance criteria
- [ ] A "Styles" side pane lists all styles in the document grouped by type (paragraph, character, list, table), showing which are in use
- [ ] "New style" creates a style from the current selection's formatting, with a name, type, "based on" parent and "style for following paragraph"
- [ ] "Modify style" edits any of those properties plus the full formatting set (font, paragraph, numbering, borders, shading) and re-renders every paragraph using it
- [ ] Renaming a style updates all references; deleting a style asks for confirmation and reverts its paragraphs to the parent style
- [ ] Character styles can be applied to a selection independently of the paragraph style, and both are shown for the caret
- [ ] Style inheritance is respected: changing a parent changes children that do not override the property
- [ ] A name that collides with an existing style is rejected with an inline error
- [ ] Custom styles round-trip through `.docx` (`styles.xml` with `w:styleId`, `w:name`, `w:basedOn`, `w:next`) and `.odt` (`style:style` with `style:parent-style-name`, `style:next-style-name`), keeping user-visible names identical
- [ ] Styles from an opened document that the editor cannot fully model are still listed, applied by reference, and written back unchanged
- [ ] "Import styles from another document" merges styles from a chosen `.docx`/`.odt`, listing conflicts and letting the user pick per style

## Out of scope
Theme colours and font schemes. Table styles with banding. A shared
organisation-wide style library.

## Note (SPEC-074)
The style editor gains a "Borders and shading…" control on paragraph styles,
which opens the very same dialog `Format → Borders and shading…` opens for a
paragraph — so a style and a paragraph can never disagree about what a
border is. The result lands in the style's `paragraphFormatting`.
