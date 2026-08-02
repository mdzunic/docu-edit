---
id: SPEC-035
title: Spell check and document language
status: draft
area: editing
depends_on: [SPEC-011]
---
## Why
Red squiggles are table stakes, and the browser's built-in checker cannot be
trusted to see our content model or to know the document's language. Language
is also a required attribute in both file formats and affects hyphenation and
screen readers.

## Acceptance criteria
- [ ] Misspelled words are underlined with a red wavy line as the user types, with a check that runs off the typing path and does not delay input
- [ ] Right-clicking a misspelled word offers up to five suggestions, "Ignore", "Ignore all" and "Add to dictionary"
- [ ] Choosing a suggestion replaces the word, preserving its formatting and its capitalisation pattern
- [ ] The custom dictionary persists across sessions and is editable from a settings dialog
- [ ] A language selector sets the language for the document, and for the selection when text is selected
- [ ] The dictionary loaded matches the text's language; text marked as a different language is checked against that language
- [ ] Text marked "Do not check spelling" is skipped; URLs, email addresses and code-styled text are skipped by default
- [ ] Language is written to `.docx` (`w:lang`) and `.odt` (`fo:language` / `fo:country`), read on open, and reflected in the `lang` attribute of the rendered text for assistive technology
- [ ] A status bar control shows the current language and opens the selector
- [ ] Spell checking a 200-page document does not block scrolling; only the visible region plus a margin is checked eagerly
- [ ] Dictionaries load on demand and at least English (US/UK), German, French, Spanish and Czech are available

## Out of scope
Grammar and style checking. Autocorrect of misspellings while typing.
Thesaurus. Hyphenation.
