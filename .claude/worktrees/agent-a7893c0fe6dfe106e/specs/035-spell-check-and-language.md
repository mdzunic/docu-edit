---
id: SPEC-035
title: Spell check with user-supplied dictionaries, and document language
status: draft
area: editing
depends_on: [SPEC-011, SPEC-015]
---
## Why
Red squiggles are table stakes, and the browser's built-in checker cannot be
trusted to see our content model or to know the document's language. Language is
also a required attribute in both file formats and affects hyphenation and
screen readers.

The product ships **no dictionaries**. Word lists are large, licensed
individually per language, and are not ours to redistribute; the earlier version
of this spec required five of them and could not be built, because there is no
dictionary dependency available here and hand-writing word lists is not
engineering. So the feature is the *machinery* — the checker, the suggestions,
the language model, the format round trip — and the user brings the word list.
That is a smaller product promise and an honest one: a user who wants German
spell checking downloads a German dictionary once and the feature works.

## Acceptance criteria
- [ ] With no dictionary installed, nothing is ever underlined and no spelling UI appears; the status bar language control shows the document language and that no dictionary is installed, and opens the dictionary manager
- [ ] The dictionary manager lists installed dictionaries with their language tag and word count, and adds one from a file the user chooses, tagged with a language they pick or confirm
- [ ] Two file formats are accepted: a plain UTF-8 word list (one word per line) and a Hunspell `.dic` file, of which only the word forms are read — affixes are not expanded and the `.aff` file is neither required nor used
- [ ] A file that is not valid UTF-8, is empty, or yields no words is rejected with a message naming the reason, and nothing is stored
- [ ] Installed dictionaries survive a reload and can be removed from the same dialog; removing one immediately clears the squiggles it produced
- [ ] With a dictionary installed for the document's language, misspelled words are underlined with a red wavy line as the user types, with a check that runs off the typing path and does not delay input
- [ ] Right-clicking a misspelled word offers up to five suggestions, "Ignore", "Ignore all" and "Add to dictionary"
- [ ] Choosing a suggestion replaces the word, preserving its formatting and its capitalisation pattern, as one undo step
- [ ] "Add to dictionary" persists across sessions in a custom word list editable from the dictionary manager; "Ignore" and "Ignore all" apply to the current session only
- [ ] A language selector sets the language for the document, and for the selection when text is selected
- [ ] Text carrying a language other than the document's is checked against that language's dictionary if one is installed, and skipped silently if not
- [ ] Text marked "Do not check spelling" is skipped; URLs, email addresses and code-styled text are skipped by default
- [ ] Language is written to `.docx` (`w:lang`) and `.odt` (`fo:language` / `fo:country`), read on open, and reflected in the `lang` attribute of the rendered text for assistive technology
- [ ] Spell checking a 200-page document does not block scrolling: only the visible region plus a margin is checked eagerly, and importing a 100 000-word dictionary does not freeze the editor
- [ ] **No word list ships in this repository.** The diff adds no dictionary data file and no word-list literal in source; test fixtures contain at most 50 words in total, and the unit tests build their dictionaries from those fixtures

## Out of scope
Bundled dictionaries in any language — the product ships with none and the
repository contains none. Hunspell affix expansion and `.aff` parsing.
Downloading dictionaries from a URL or a bundled catalogue. Grammar and style
checking. Autocorrect of misspellings while typing. Thesaurus. Hyphenation.

## Reference

### The three decisions this spec makes, so nothing has to be guessed

1. **"Code-styled text" means the style system, not the font.** A run is
   code-styled when its paragraph's `styleId` is `Code`/`CodeBlock` or the run
   carries a character style named `Code`. Choosing a monospace font by hand does
   not make text code.
2. **"Ignore" and "Ignore all" are session-only.** They live in memory and are
   gone on reload. Only "Add to dictionary" persists.
3. **"Visible region plus a margin" is the viewport plus 500 px above and below.**
   Text outside it is checked lazily as it scrolls into range.

### Model addition

`RunFormatting` in `src/lib/documentModel.ts` gains one field, and nothing else
in the model changes:

```ts
export interface RunFormatting {
  // …existing fields (bold, italic, fontFamily, highlight, linkHref, …)
  /** BCP 47 language tag for this run, e.g. "en-GB", "de-DE" (SPEC-035).
   * A plain scalar, like linkHref, so run coalescing and formatting equality
   * keep comparing with === per key — see textEditing's sameFormatting. */
  lang?: string;
  /** True when this run is marked "do not check spelling" (docx `w:noProof`,
   * odt `fo:language="zxx"`). Skipped by the checker (SPEC-035). */
  noProof?: boolean;
}
```

The document's default language is a document-scoped field, not a run field;
runs without `lang` are checked against it. A selection-scoped language change
sets `lang` on the runs it spans through the existing run-formatting command
(SPEC-013), so it splits and coalesces runs exactly as bold does.

### Dictionary storage

A **separate** IndexedDB database, so the document library's schema and version
are untouched:

- database `docu-edit-dictionaries`, version 1, object store `dictionaries`,
  `keyPath: "id"`
- record: `{ id: string, lang: string, name: string, words: string[], addedAt: number }`
  where `id` is the language tag for an imported dictionary and the literal
  `"custom"` for the user's own added words
- the wrappers follow `src/lib/idb.ts` exactly, including its rule that every
  function rejects when IndexedDB is missing, blocked or over quota, and that
  callers treat any rejection as "storage unavailable" and keep working in
  memory for the session

### Dictionary file parsing

Both formats reduce to a `string[]` of lowercase word forms:

- **Plain word list** — split on newlines, trim, drop blank lines and lines
  starting with `#`.
- **Hunspell `.dic`** — the first line is a word count and is discarded if it
  parses as an integer. On every other line, take the text before the first `/`
  (the affix flags after it are discarded, since affixes are not expanded), then
  trim. A leading UTF-8 BOM is stripped. Lines containing only flags are dropped.

Detection is by content, not by file extension: if the first non-empty line
parses as an integer on its own, treat the file as `.dic`. Both paths then
lowercase, de-duplicate and sort.

Because affixes are not expanded, an inflected form absent from the `.dic` is
reported as misspelled. That is a stated limitation of this spec, not a defect —
the dictionary manager says so next to the word count.

### Checking, suggestions and capitalisation

- **Tokenising** is over run text with the run's resolved language. A token is a
  maximal sequence of letters plus interior apostrophes and hyphens. Tokens that
  match a URL or an email address, or that sit in a `noProof` or code-styled run,
  are skipped without lookup.
- **Lookup** is case-insensitive: a token is misspelled when neither it nor its
  lowercase form is in the language's dictionary or the custom list.
- **Suggestions** are the dictionary words within Damerau–Levenshtein distance 2,
  ranked by distance then by length, capped at five.
- **Capitalisation pattern** is preserved when a suggestion is applied: `all
  lower` → as-is, `Title` → capitalise the first letter, `ALL UPPER` → uppercase.
  Any other pattern is left as the suggestion supplies it.

### Dependency contracts (restated per SPEC-000 rule 2)

- **SPEC-011** (`src/lib/textEditing.ts`): mutations take `(model, caret, …)` and
  return the new caret; the DOM is never the source of truth. Replacing a
  misspelled word uses `deleteRange(a, b)` then `insertText`, with the caret
  coordinate type `{ section, block, offset }`.
- **SPEC-013** (`src/lib/documentModel.ts`): setting a run format over a
  selection goes through the existing run-formatting command and `PendingFormat`
  patch semantics; `lang` follows the same per-key `===` comparison as
  `linkHref`.
- **SPEC-015** (`src/lib/undoHistory.ts`): applying a suggestion calls
  `recordEdit` once, so it is a single undo step. `MAX_HISTORY_STEPS` is 200.
- **Equality**: the docx/odt language round-trip tests use `equalStructural` from
  `src/lib/equality.ts`; byte-level preservation of untouched parts is
  `equalBytes`, per SPEC-008.
- **The checker is framework-free.** It lives in `src/lib/` as plain TypeScript
  with no Vue imports and is unit-testable without mounting a component, per
  CLAUDE.md.
