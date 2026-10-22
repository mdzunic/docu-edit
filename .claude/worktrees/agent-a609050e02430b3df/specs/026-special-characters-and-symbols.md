---
id: SPEC-026
title: Special characters and symbol picker
status: draft
area: editing
depends_on: [SPEC-011, SPEC-015, SPEC-021]
---
## Why
Every real document eventually needs a — an ™ or a °, and typing them is
awkward on most keyboards. Small feature, constant use, and it forces correct
handling of characters outside the Basic Multilingual Plane.

## Acceptance criteria
- [ ] An "Insert → Special characters" dialog shows a searchable grid of characters grouped by the eight categories in "Character catalogue" below, containing exactly the code points listed there and no others
- [ ] Searching filters the grid by character name, by keyword alias, or by hex code point in any of the forms `2014`, `u+2014`, `0x2014`, `#2014` (case-insensitive, leading zeros optional) — matching is case-insensitive substring matching on names and keywords, exact matching on code points
- [ ] Clicking a character inserts it at the caret and keeps the dialog open for another insertion
- [ ] Recently used characters appear in a "Recent" row, most recent first, de-duplicated, capped at `RECENT_CHARACTERS_LIMIT`, persisted in `localStorage` under `docu-edit:recent-special-characters` so they survive a reload
- [ ] Non-breaking space (Ctrl/Cmd+Shift+Space), em dash (U+2014), en dash (U+2013) and non-breaking hyphen (U+2011) have direct insert entries, and the non-breaking space shortcut is registered in `SHORTCUTS` (SPEC-021) so it appears in the shortcut reference
- [ ] Typing `-` immediately after a `-` replaces both with an en dash (U+2013); typing a third `-` immediately after that replaces the en dash with an em dash (U+2014) — exactly the trigger rules in "Dash substitution" below
- [ ] Typing `'` or `"` inserts the typographic quote chosen by the rules in "Smart quote substitution" below, which depend only on the character immediately before the insertion point
- [ ] Each substitution is recorded as its own undo entry after the plain insertion, so one Ctrl/Cmd+Z restores the literal typed characters (`--`, `'`, `"`) with the caret after them, and the substitution does not re-fire at that caret position until the caret moves
- [ ] An autocorrect setting in the app menu turns the smart-quote and dash substitutions off independently, persisted in `localStorage` under `docu-edit:autocorrect` as `{ smartQuotes, dashes }`, both defaulting to `true`
- [ ] Substitutions run on typed input only — never on paste (SPEC-016), find-and-replace (SPEC-017), the character picker, or a document load
- [ ] Characters outside the BMP (e.g. 𝑥, 😀) insert, render, select and delete as single units: caret motion, Shift-extension, Backspace and Delete step by extended grapheme cluster as defined in "Grapheme-aware caret motion" below, so a surrogate pair is never split and an emoji ZWJ/variation-selector/skin-tone sequence is never partly deleted
- [ ] A non-breaking space prevents a line break at that point in rendering, and a document containing U+00A0, U+202F, U+2011, U+00AD and a non-BMP character survives open → save → open with `equalStructural` in both `.docx` and `.odt`, written as the characters named in "File format encoding" below

## Out of scope
Equation/formula editing. Handwriting or drawing input. Custom autocorrect
dictionaries. Extending the catalogue at runtime, or loading the full Unicode
character database — the catalogue is the fixed, hand-authored table below.
Automatic prime substitution (`'` → U+2032). Nesting-aware or language-aware
quote styles (German `„…"`, French `«…»`): the smart-quote rule is local and
produces English quotes only. Font fallback for a character the body font
lacks (SPEC-048).

## Reference

Dependency specs: `specs/002-create-blank-document.md` (canonical document
model and the equality relations), `specs/011-rich-text-editing-core.md`
(caret, selection, text mutation), `specs/015-undo-and-redo.md` (undo
entries and grouping), `specs/021-keyboard-shortcuts-and-reference.md`
(shortcut registry).

### What this spec does *not* change about the model
The canonical model is defined by SPEC-002 and `docs/document-model.md`. This
spec adds no node kind, no field and no unit. The three facts it relies on,
restated inline:

- A paragraph's text lives in `Run.text`, a JavaScript string — a sequence of
  **UTF-16 code units**. A non-BMP character is stored as its surrogate pair:
  one code point, two code units. There is no escaping, entity encoding or
  numeric character reference at the model layer; U+00A0 is stored as the
  literal character U+00A0 and never as `&nbsp;` or U+0020.
- `CaretPosition.offset` (`src/lib/textEditing.ts`) stays a **UTF-16 code-unit
  index** into the paragraph's plain text (`blockText`). Grapheme awareness is
  a property of caret *motion*, not of the offset's meaning — see below.
- Round-trip assertions name a relation from SPEC-002: `equalStructural` for
  same-format round trips (open → save → open), `equalSemantic` plus an
  exclusion list for cross-format conversion. `equalBytes` is not used here.

### Character catalogue
`src/lib/specialCharacters.ts` exports the catalogue as plain data:

```ts
interface SpecialCharacter {
  cp: number            // Unicode code point
  name: string          // the Unicode 15.1 character name, e.g. "EM DASH"
  keywords: string[]    // lowercase search aliases; [] when the name suffices
}
interface CharacterCategory { id: CategoryId; label: string; chars: SpecialCharacter[] }
type CategoryId = 'punctuation' | 'currency' | 'maths' | 'arrows'
                | 'latin' | 'greek' | 'symbols' | 'emoji'
export const CHARACTER_CATEGORIES: CharacterCategory[]
export const RECENT_CHARACTERS_LIMIT = 24
```

Entries within a category are ordered by code point, except Emoji, which keeps
the thematic order listed below (faces, hands, celebration, status, documents,
tools, objects) because code-point order scatters related emoji. `name` is the Unicode
15.1 character name verbatim (displayed title-cased); `keywords` carry the
words a user is likely to type that the formal name lacks — at minimum
`em dash`→`dash`, U+00A0→`nbsp`, `non-breaking space`, U+2026→`dots`,
U+00B0→`degree`, U+2122→`trademark`, U+00A9→`copyright`, U+2713→`tick`,
`check`. **552 characters total**, exactly as follows.

**Punctuation — 48.**
```
00A0 00A1 00A7 00AB 00AD 00B6 00B7 00BB 00BF 2002
2003 2009 200B 2010 2011 2012 2013 2014 2015 2016
2018 2019 201A 201B 201C 201D 201E 201F 2020 2021
2022 2023 2025 2026 2027 202F 2030 2031 2032 2033
2034 2039 203A 203B 203D 2043 2044 2116
```

**Currency — 38.** `0024 00A2 00A3 00A4 00A5 0192`, then the whole Currency
Symbols block `20A0`–`20BF` (32).

**Maths — 64.**
```
00AC 00B1 00B5 00BC 00BD 00BE 00D7 00F7 2153 2154
215B 215C 215D 215E 2200 2202 2203 2204 2205 2206
2207 2208 2209 220B 220F 2211 2212 2213 2215 2217
2218 2219 221A 221B 221D 221E 2220 2221 2225 2226
2227 2228 2229 222A 222B 222C 2234 2235 223C 2245
2248 2260 2261 2264 2265 226A 226B 2282 2283 2286
2287 2295 2297 22A5
```

**Arrows — 32.** The contiguous range `2190`–`21AF`.

**Latin — 190.** `00C0`–`00FF` **excluding** `00D7` and `00F7` (which are in
Maths) — 62 — plus the whole Latin Extended-A block `0100`–`017F` (128).

**Greek — 52.** `0391`–`03A1` and `03A3`–`03A9` (24 capitals, skipping the
unassigned `03A2`), `03B1`–`03C9` (25 smalls, including final sigma `03C2`),
plus the variant letters `03D1 03D5 03D6`.

**Symbols — 64.**
```
00A9 00AE 00B0 2103 2105 2109 2113 2117 2118 2120
2122 2126 212B 2139 2300 2302 2318 2325 2326 232B
23CE 25A0 25A1 25AA 25B2 25B6 25BC 25C0 25CB 25CF
25E6 2600 2601 2602 260E 2610 2611 2612 2618 261E
2620 262E 263A 2660 2663 2665 2666 266A 266B 2691
26A0 2713 2714 2717 2718 2720 2726 272A 2731 2736
273F 2744 2756 2764
```

**Emoji — 64.** Deliberately mostly non-BMP — this category is what exercises
AC-11 and AC-12.
```
1F600 1F601 1F602 1F609 1F60A 1F60D 1F610 1F621 1F622 1F62D
1F634 1F642 1F914 1F44B 1F44C 1F44D 1F44E 1F44F 1F4AA 1F64F
1F91D 1F381 1F388 1F389 1F3C6 1F947 2B50 2728 1F525 1F4A1
1F4AF 2757 2753 2705 274C 26A1 1F4CC 1F4CE 1F4C1 1F4C4
1F4CA 1F4C8 1F4C9 1F4C5 23F0 1F50D 1F512 1F511 1F6E0 2699
1F680 1F41B 1F4BB 1F4F1 2709 1F4DE 1F3E0 1F30D 1F31E 1F327
2615 1F355 1F3B5 1F49A
```

Characters with no visible glyph — `00A0 00AD 200B 2002 2003 2009 202F` —
render in the grid as a dotted placeholder box carrying the character's name
as its accessible label, not as blank cells.

A unit test asserts the per-category counts above, that every `cp` is unique
across the whole catalogue, and that every `name` is non-empty.

### Smart quote substitution
Substitution happens at the moment the straight quote is typed. Let `prev` be
the code point immediately before the insertion point **within the same
paragraph** — for a non-empty selection, the code point before the selection
start — or *none* at a paragraph start. Define two sets:

- `OPENING_BEFORE` = whitespace (`0020 0009 00A0 2000`–`200A` `202F 205F
  3000`), a line break (a `LineBreak` inline, i.e. Shift+Enter), and the
  characters `( [ { < ¿ ¡ « ‹ / \ – — ‑ …` plus the opening quotes `‘ “`.
- Everything else — letters, digits, closing brackets, and all remaining
  punctuation — is `CLOSING_BEFORE`.

Then, for a typed `"`:
- `prev` is *none* or in `OPENING_BEFORE` → **U+201C** (left double quote)
- otherwise → **U+201D** (right double quote)

For a typed `'`:
- `prev` is *none* or in `OPENING_BEFORE` → **U+2018** (left single quote)
- otherwise → **U+2019** (right single quote) — this is the rule that makes
  `don't` and `parents'` correct, since `prev` is a letter in both

One retroactive correction, for the decade idiom: if a **digit** is typed
immediately after a U+2018 this feature just inserted, that U+2018 is replaced
with U+2019, so `'90s` gives `’90s`. The replacement is part of the digit's
edit and is undone with it.

The rule is deliberately **purely local**: no paired-quote state is tracked, so
a quote typed into the middle of existing text is decided by its left
neighbour alone. Nothing in this spec converts an already-typed quote when
surrounding text later changes.

### Dash substitution
Triggered only by typing `-` (U+002D):

- The typed `-` makes the text immediately before the caret end in `--`, and
  the character before that pair is **not** `-` → replace both with **U+2013**.
- The typed `-` immediately follows a U+2013 **that this feature produced at
  this caret position** → replace the U+2013 with **U+2014**.

A `-` in any other position is inserted literally. Neither trigger fires when
the caret has moved since the previous insertion.

### Undo behaviour
The plain insertion and the substitution are two entries: the typed character
is inserted first, then `breakGroup(docId)` (`src/lib/undoHistory.ts`) is
called, then the substitution is recorded as its own `HistoryEntry`. So the
first Ctrl/Cmd+Z after `--` restores exactly `--`, and the second undoes
whatever typing group preceded it. After such an undo, the same substitution is
suppressed for that caret position until the caret moves, so redoing the keypress
is not required to keep the literal text.

### Grapheme-aware caret motion
Left/Right arrow, Shift+arrow, Backspace and Delete move or delete by
**extended grapheme cluster**, computed with `Intl.Segmenter` (`{ granularity:
'grapheme' }`) over the paragraph's `blockText`. Where `Intl.Segmenter` is
unavailable the fallback steps by **code point** (never by lone surrogate).
Offsets stay code-unit indices; only the step size changes. Word-wise motion
(Ctrl/Cmd+arrow), double-click and triple-click selection keep the SPEC-011
behaviour.

Concretely, `deleteBackward` in `src/lib/textEditing.ts` currently steps
`offset - 1`, which splits a surrogate pair; it and `deleteForward` are the
two functions this changes.

### File format encoding
Both writers already emit UTF-8 and iterate text without splitting surrogate
pairs. The normative rules, and the regressions the tests guard:

**`.docx`** (`src/lib/docx/writer.ts`, `reader.ts`)
- Non-BMP characters are written as literal UTF-8 in `w:t` (four bytes), never
  as surrogate-pair numeric character references.
- U+00A0 and U+202F are literal characters in `w:t`. The existing
  `xml:space="preserve"` rule stands (`buffer.trim() !== buffer`, which treats
  U+00A0 as trimmable and so preserves it — correct, if incidentally so).
- U+2011 is written as `<w:noBreakHyphen/>` and U+00AD as `<w:softHyphen/>`,
  and the reader maps both back to their characters — the same inverse-pair
  discipline `runContentXml` already uses for `\t`→`<w:tab/>` and
  `\n`→`<w:br/>`.

**`.odt`** (`src/lib/odf/writer.ts`, `reader.ts`)
- ODF whitespace collapsing applies to U+0020, tab, CR and LF only. U+00A0 and
  U+202F must therefore stay **literal characters** and must never become
  `<text:s/>`, and must not update `WhitespaceState.lastChar` as if they were
  spaces. A unit test asserts that a run whose text is
  `"a\u00A0\u00A0b"` writes two literal U+00A0 characters and no `text:s`.
- U+2011 and U+00AD are literal characters; ODF has no element form for them.

**Round-trip tests.** A fixture paragraph whose text is, written as escapes so
the invisible characters are unambiguous:

```
"a b c‑d­e\u{1D465}f\u{1F600}g"
```

It round-trips open → save → open under `equalStructural` in each format, and
converts `.docx` → `.odt` (and back) under `equalSemantic` with an empty
exclusion list — every character in it is representable in both formats.

### Files
- `src/lib/specialCharacters.ts` + `.spec.ts` — the catalogue and its search
- `src/lib/autocorrect.ts` + `.spec.ts` — the quote/dash rules as pure
  functions over (`prev`, typed character, settings), no Vue imports
- `src/components/SpecialCharacterDialog.vue` (+ `.spec.ts`)
- `e2e/SPEC-026.spec.ts` — picker, search, recent row, autocorrect toggle,
  undo of a substitution, and non-BMP insert/select/delete
