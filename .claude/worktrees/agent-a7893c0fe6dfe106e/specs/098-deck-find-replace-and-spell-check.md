---
id: SPEC-098
title: Deck find and replace, and spell check on slides
status: draft
area: editing
complexity: normal
story_points: 8
depends_on: [SPEC-017, SPEC-035, SPEC-079, SPEC-081, SPEC-082, SPEC-083, SPEC-091]
---
## Why
SPEC-081's out-of-scope paragraph ends with "find and replace, spell check on
slides", and that is the hole every real deck hits first. A sixty-slide deck
renders one slide at a time; the client's name is on eleven of them and there
is no way to find the word, let alone change it. The word processor has had
both tools since SPEC-017 and SPEC-035, and every part of them that is not
about `DocumentModel` is already product-neutral — the find bar is a controlled
component, the checker takes a string and a dictionary, the suggestion menu
takes coordinates and words. This spec generalises those surfaces rather than
forking them, and adds the deck-shaped parts: a hit list grouped by slide,
navigation that selects a slide *and* a frame *and* a range, and a replace that
dirties exactly the slides it touched.

The scope statement is half the feature. A slide carries text this editor can
read but must not rewrite: a `GroupFrame`'s children render from `children` but
save from `GroupFrame.xml`, and a `ForeignFrame` is opaque markup by
definition. Offering hits whose replacement is silently discarded on the next
save is worse than not offering them, so the bar states its scope on screen.
**This file is the normative source for deck search scope, the deck hit model,
and slide-run language.** It also fixes a real loss: `runXml`
(`src/lib/presentation/pptx/serialize.ts:235`) hard-codes `<a:rPr
lang="en-US"…>` on every rebuilt run, so a German deck edited here comes back
declaring English.

## Acceptance criteria
- [ ] `SlideRun` gains `lang?: string` (source-verbatim tag, never re-normalised on read — `normalizeLanguageTag` is used only for dictionary lookup) and `noProof?: true`; `PresentationModel` gains `language?: string` (deck checking default, absent = `DEFAULT_LANGUAGE` from `src/lib/core/language.ts`); all three are plain JSON so `roundTrip.spec.ts`'s `JSON.parse(JSON.stringify(model))` `toEqual(model)` assertion still holds
- [ ] `src/lib/core/findQuery.ts` (new) holds `FindOptions`, `CompiledQuery`, `QueryError`, `INVALID_REGEX_MESSAGE`, `isQueryError`, `escapeRegExp` and `compileQuery`, moved **unchanged, names included**, out of `src/lib/findReplace.ts`, which re-exports them; `src/lib/findReplace.spec.ts` is untouched and still green, proving the move behaviour-preserving
- [ ] `src/lib/presentation/deckSearch.ts` (framework-free; imports only presentation siblings and `lib/core`) exports `DeckHit`, `findInDeck(model, compiled): DeckHit[]`, `hitsBySlide(hits)`, `hitIndexAtOrAfter(hits, slideIndex, frameId)`, `replaceHit(model, hit, replacement): boolean` and `replaceAllInDeck(model, hits, replacement): number` per the Reference, hits ordered by slide index, then z-order, then paragraph, then offset
- [ ] Search covers, per slide: every `TextFrame.paragraphs`, every `ShapeFrame.paragraphs`, every non-covered `TableFrame` cell's `paragraphs` (SPEC-091; where that spec has not landed the `table` branch is simply absent and those tables count as unsearchable foreign frames) and `Slide.notes`; it never reads `GroupFrame.children`, `ForeignFrame.xml`, `Slide.sourceXml`, `Slide.preserved`, `PresentationModel.opaque` or `.media`, pinned by a spec whose model defines throwing getters for `media`/`opaque`
- [ ] The toolbar gains a Find toggle (`slides-find-toggle`, in the right-aligned `tb-slide-actions` group beside Transition/Animate); Cmd/Ctrl+F opens the bar and Cmd/Ctrl+H opens it with the replace row revealed, both via `findShortcut` (`find.open`, `find.openReplace`) from `src/lib/shortcuts.ts`; Escape closes it and returns focus to the stage
- [ ] The bar is `src/components/FindBar.vue` **generalised, not forked**: two new props, `scopeLabel` (default `"Body text only"`) and `scopeTitle` (default `"Search covers body text only. Headers, footers and footnotes are excluded."` — the string the span hard-codes today at `FindBar.vue:258`, so both of `FindBar.spec.ts`'s AC-28 assertions and DocumentView are unchanged), render in the existing `find-scope` span; the deck passes `"Slides, shapes, tables and notes"` and a title naming the exclusions and counting them from the loaded model
- [ ] A results list (`deck-search-results`) below the bar groups hits by slide: one `deck-search-group` per slide with a `deck-search-slide-chip` carrying the 1-based number, then one `deck-search-hit` button per hit showing a ±40-character snippet with the match emphasised, the frame's name (`frameAccessibleName`) and a scope tag ("Notes", "Table r2c3"); `deck-search-empty` when nothing matches, `deck-search-scope` restating what was not searched
- [ ] Activating a hit (click, Enter, find-next/prev) sets `currentIndex` to its slide, calls `selection.select(frameId)`, starts a text-editing session on that frame and selects the range through `setSelectionRange(host, {para, offset}, {para, offset})` from `src/components/presentation/slideTextDom.ts`; a notes hit opens the notes panel and selects the range in the `slides-notes-editor` textarea instead; wrapping past the last hit shows the "Wrapped to top" hint; hits outside an editing session paint as rects in `SlideProofingLayer.vue` under the SPEC-017 testids `find-highlights` / `find-highlight` (`.current` on the current one), re-measured on `stageScale` and slide change
- [ ] "Match case", "Whole words only" and "Use regular expressions" behave exactly as SPEC-017 defines (same `compileQuery`), persist for the session in the composable, and an invalid expression shows `INVALID_REGEX_MESSAGE` in `find-status` and runs nothing
- [ ] "Replace" rewrites the current hit and steps to the next; "Replace all" reports "Replaced N occurrences" and is **one** undo entry — one `record("Replace all")`, every mutation, one `commit()` — with per-paragraph hits applied back to front so earlier offsets stay valid, the replacement inheriting the marks of the run containing the hit's `start` (SPEC-017's rule)
- [ ] Dirty marking is per slide and per scope: a frame or cell replacement calls `markSlideEdited(slide, model.format)`; a notes-only replacement sets `slide.notesDirty = true` on pptx and calls `markSlideEdited(slide, "odp")` on odp (because `writeOdp` never consults `notesDirty`); slides with no hit keep both flags absent and their parts stay byte-verbatim on the next save
- [ ] Spell check on the stage: with a dictionary available for a run's effective language, misspelled words carry a red wavy underline drawn by `SlideProofingLayer.vue` with testids `spell-underlines` / `spell-underline`; right-click opens `src/components/SpellContextMenu.vue` unchanged (`spell-suggestion` ×≤5, `spell-ignore`, `spell-ignore-all`, `spell-add-to-dictionary`); applying a suggestion preserves the word's capitalisation pattern and is one undo entry labelled "Correct spelling"
- [ ] `src/components/LanguageStatusControl.vue` mounts unchanged in the deck header (or in `DeckHeader.vue` where SPEC-090's extraction has landed): it sets the deck language when nothing is selected and the selection's `lang` when text is selected, "Do not check spelling" sets `noProof` on the selection, and "Custom dictionary…" opens `CustomDictionaryDialog.vue` unchanged; only the current slide's frames and its notes are ever checked, off the main thread through `checkTextsInWorker`
- [ ] Wire round trip: pptx `a:rPr@lang` / `@noProof="1"` and odp `style:text-properties/@fo:language` + `@fo:country` (`zxx` for `noProof`) read into and write out of the model per the Reference; a deck given per-run languages saves → reopens with identical tags (field-level `toEqual`), untouched parts stay `equalBytes` (`src/lib/core/equality.ts`) part by part, and a rebuilt slide whose runs carry no `lang` produces today's exact bytes
- [ ] Performance, measured not assumed: the clock starts when the 150 ms input debounce fires and the walk begins — never at the keystroke — and from that instant a 60-slide deck's first hits render within 150 ms and the full walk settles within 500 ms; the walk triggers no slide re-render (no `commit()`, no re-identification — `findInDeck` is pure) so filmstrip thumbnails never repaint, no `MediaPart` is base64-decoded, and decks over 40 slides walk in batches of 10 through `scheduleIdleWork` (`src/lib/core/idleScheduler.ts`) with `deck-search-progress` shown until settled
- [ ] e2e (`e2e/SPEC-098.spec.ts`): open `e2e/fixtures/mini-deck.pptx` → Cmd/Ctrl+F → type "Open" → two groups, slide 1 (a Notes hit, "Open with the Q3 highlight") and slide 2 ("Open Berlin office") → activate the slide-2 hit → slide 2 shows with that frame selected and the word selected → search "Grouped" → `deck-search-empty` plus the scope note (slide 3's grouped text is deliberately not searched) → expand, replace all "Open" → "Launch" → "Replaced 2 occurrences" → one undo restores both → redo, download, reopen → both strings changed

## Out of scope
Search across decks or the library; find by formatting, effect, layout or
theme; regex capture groups in the replacement (SPEC-017's stated follow-up);
master, layout and notes-master text (SPEC-089 models layout *slots*, never
layout text); `GroupFrame` children, `ForeignFrame` markup and any
search-inside-XML mode; a replace-within-selection scope; grammar and style
checking, thesaurus, hyphenation and autocorrect on slides (SPEC-035 excludes
them for text too); bundled dictionaries in any language; east-asian and
complex-script language variants on a run; a search surface inside playback
(SPEC-084 owns the play view); the deck language reaching the file as a
deck-level default — see the wire section.

## Reference

### Dependencies restated
- **`specs/017-find-and-replace.md`** — `src/lib/findReplace.ts` exports
  `FindOptions { matchCase; wholeWords; regex }`, `compileQuery(query,
  options): CompiledQuery | QueryError | null` (empty query → `null`; only
  regex mode can fail; `wholeWords` wraps the source in `\b(?:…)\b`; flags
  `g`/`gi`), `escapeRegExp`, `isQueryError`, `INVALID_REGEX_MESSAGE =
  "Invalid regular expression"`. `src/components/FindBar.vue` is controlled:
  props `visible, expanded, query, replaceText, matchCase, wholeWords,
  useRegex, matchCount, currentIndex, error, hint`; emits `update:query,
  update:replaceText, update:matchCase, update:wholeWords, update:useRegex,
  toggle-expand, next, prev, replace, replace-all, close`; testids `find-bar,
  find-query, find-count, find-status, find-prev, find-next, find-close,
  find-expand, find-match-case, find-whole-words, find-regex, find-scope,
  find-replace-row, find-replace-input, find-replace, find-replace-all`.
- **`specs/035-spell-check-and-language.md`** — no dictionary ships beyond the
  small curated lists in `src/lib/spellcheck/dictionaries/`; Ignore and Ignore
  all are session-only, Add to dictionary persists. `src/lib/spellcheck/`
  (`checkText(text, dictionary): SpellIssue[]`, `tokenizeForSpellCheck` —
  already skips URLs and emails, `createDictionary`, `getSuggestions`,
  `MAX_SUGGESTIONS = 5`, `detectCasePattern`/`applyCasePattern`,
  `loadDictionary(tag)`), `src/lib/customDictionary.ts` and
  `src/lib/workers/spellCheckClient.ts` (`checkTextsInWorker(items)`) are
  already framework-free and product-neutral. **No hoist is proposed for any
  of them**: their consumer `src/views/presentation/useSlideSpellCheck.ts` is
  a composition point that may import anything, and moving them into
  `src/lib/core` would force `core/boundary.spec.ts` to recurse (it reads
  `readdirSync(coreDir)` flat today) for no gain. `src/lib/core/language.ts`
  is already core: `DEFAULT_LANGUAGE = "en-US"`, `SUPPORTED_LANGUAGES`,
  `normalizeLanguageTag`, `isDictionaryAvailable`, `splitLanguageTag`,
  `composeLanguageTag`, `effectiveLanguage(runLang, docLang)`,
  `languageLabel`, `asLanguageOptions`, `ODF_NO_LINGUISTIC_CONTENT = "zxx"`.
- **`specs/081-slide-text-editing.md`** — `presentation/text.ts`:
  `TextPoint { para; offset }`, `TextRange { start; end }`, `normalizeRange`,
  `RunFormatPatch` (`false` REMOVES a mark), `SelectionFormatState` of
  `FieldState<T> = T | "mixed" | undefined`, and `markSlideEdited(slide,
  format)` = `slide.dirty = true` plus `delete slide.sourceXml` for odp — the
  single dirtying primitive every editing command ends with.
  `components/presentation/slideTextDom.ts`: `CaretAddress { para; offset }`,
  `setSelectionRange(host, start, end)`, `selectionRangeOf(host)`. Slide-part
  serialisation is SPEC-081's; this spec touches only `runXml` and
  `StyleRegistry.textStyle`.
- **`specs/082-…`** — frame identity (`SlideFrame.id`, minted at read, never
  serialised), `useFrameSelection`'s `{ selectedFrameId, selectedFrame,
  select, clear, cycle }`, and `frameNoun`/`frameAccessibleName` from
  `useFrameManipulation.ts` for the results list.
- **`specs/083-speaker-notes.md`** — `Slide.notes?: string[]`, plain
  paragraphs, trailing empties stripped, `[]` never stored; `useNotes`'s
  `onInput` sets `notesDirty` on pptx but calls `markSlideEdited(slide,
  "odp")` on odp.
- **`specs/091-tables-on-slides.md`** — `TableFrame { kind: "table"; rows:
  SlideTableRow[]; columnWidthsCm: number[]; … }` with
  `rows[r].cells[c].paragraphs: SlideParagraph[]`, `[]` on a covered slot.
- **Model** (`presentation/model.ts`, SPEC-079) — `SlideRun { text; bold?;
  italic?; underline?; fontFamily?; sizePt?; color? }`, `SlideParagraph
  { runs; align?; bullet?; level? }`. `paragraphText(paragraph)` in
  `slideModel.ts` joins a paragraph's run text and is the offset basis for
  every `start`/`end` below.

### The deck hit model (`src/lib/presentation/deckSearch.ts`)
```ts
export type DeckHitScope = "frame" | "cell" | "notes";
export interface DeckHit {
  slideIndex: number;
  frameId?: string;                // absent only for a notes hit
  row?: number; column?: number;   // "cell" scope only (SPEC-091)
  para: number;                    // index into the frame/cell/notes list
  start: number; end: number;      // half-open, into paragraphText(para)
  scope: DeckHitScope;
  snippet: string;                 // paragraph text trimmed to ±40 chars
  snippetOffset: number;           // where `start` falls inside `snippet`
}
```
A hit carries no frame *index*: frames are addressed by `frameId`, which
survives an undo snapshot (`structuredClone(model.slides)`), so a hit list
stays valid across a replace. `replaceHit` returns `false` and changes nothing
when the hit no longer matches. `replaceAllInDeck` groups by
slide/frame/paragraph, applies within a paragraph back to front, calls the
right dirtying primitive once per slide, and returns the count. Run-level
rewriting goes through a new pure `replaceParagraphText(paragraph, start, end,
replacement)` in `presentation/text.ts`, built on that module's existing
private `expand`/`coalesce` pair, so the replacement takes the marks of the
run containing `start` and neighbouring identical runs re-coalesce.
`find-scope`'s deck title names the exclusions and counts them from the loaded
model — group children, foreign markup (tables included where SPEC-091 has not
landed), master and layout text; `deck-search-scope` repeats that sentence
under an empty result set, which is exactly when it is needed.

### Language on a run, and the two wire mappings
`RunFormatPatch` gains `lang?: string`, `clearLang?: true` and
`noProof?: boolean` (`false` removes, per that type's convention); the private
`RunMarks` gains both fields so runs differing only in language stop
coalescing; `SelectionFormatState` gains `lang: FieldState<string>` and
`noProof: FieldState<boolean>` so the control can render "mixed".

- **pptx** — read `a:rPr@lang` → `run.lang` and `a:rPr@noProof="1"` →
  `run.noProof`, stored verbatim. `runXml` becomes
  `<a:r><a:rPr lang="${escapeXmlAttribute(run.lang ?? DEFAULT_LANGUAGE)}"${noProof}${sz}${b}${i}${u}>…`
  — byte-identical to today whenever `run.lang` is absent, because the
  hard-coded literal it replaces *is* `DEFAULT_LANGUAGE`. `err`, `dirty` and
  `smtClean` stay unmodelled and unemitted; on an untouched slide they survive
  verbatim like everything else.
- **odp** — `@fo:language` + `@fo:country` compose into one tag through
  `composeLanguageTag`; `fo:language="zxx"` (core's
  `ODF_NO_LINGUISTIC_CONTENT`) is `noProof`. `StyleRegistry.textStyle(run)`
  emits the pair — split by `splitLanguageTag` — **only** when `run.lang` is
  set or `run.noProof` is true, so every existing `odp/serialize.spec.ts`
  expectation stays byte-identical and the `de-T{n}` key simply widens.
- **The deck default is deliberately not serialised.** `model.language` is set
  at read to the most frequent non-empty `run.lang` in the deck (absent when
  none) and drives checking and the control's label only. The candidate slots
  — pptx `p:defaultTextStyle/a:lvl1pPr/a:defRPr@lang` and odp `styles.xml`'s
  `style:default-style` text properties — are **not** specified here: their
  exact shape is not something to write from memory. A later spec wanting them
  goes through SPEC-086's golden-fixture-first process (hand-authored
  canonical bytes as a `*Golden.ts` constant, the serialiser asserted equal to
  them, cross-app verification as manual QA in the PR, not CI).

### Performance rules (normative)
1. `findInDeck` is a pure walk over plain model arrays: no view call, no
   `commit()`, no `bumpCurrentSlide()`, no mutation. No object identity
   changes, so neither the stage nor the filmstrip re-renders — SPEC-086's
   reactivity rule read backwards.
2. It touches text only (`paragraphs`, `runs[].text`, `notes`). `model.media`,
   `model.opaque`, `mediaDataUri`, `decodedPartBytes` and every `xml` string
   are outside its import graph, so no base64 is decoded and no XML parsed.
3. Query input debounces 150 ms; above 40 slides the walk is chunked at 10
   slides per `scheduleIdleWork` slice, appending as it goes so the first
   group paints immediately.
4. Spell checking covers the current slide's frames plus its notes only, in
   `checkTextsInWorker`, scheduled on idle; dictionaries load lazily per
   language through `loadDictionary(tag)`, which caches them.
5. Navigating a hit changes `currentIndex` and the selection: exactly one
   slide render, no mutation, no undo entry.

### View wiring and the size ratchet
`src/views/presentation/useDeckSearch.ts` — deps `{ model, currentIndex,
record, commit, selectFrame, startEditing, openNotes }` — returns `{ open,
expanded, query, replaceText, matchCase, wholeWords, useRegex, error, hint,
hits, groups, currentHitIndex, hitCount, scopeLabel, scopeTitle, indexing,
openFind, closeFind, toggleExpand, next, prev, goToHit, replaceCurrent,
replaceAll, onGlobalKeydown }`. `useSlideSpellCheck.ts` — deps `{ model,
currentSlide, currentIndex, record, commit, selection, notesText }` — returns
`{ underlineRects, contextMenu, deckLanguage, languageOptions,
selectionLanguage, selectionNoProof, openContextMenuAt, applySuggestion,
ignore, ignoreAll, addToDictionary, setDeckLanguage, setSelectionLanguage,
setNoProof, dictionaryOpen, openDictionary, closeDictionary, remeasure,
reset }`. Both obey `record(label)` → mutate → `commit()`, one undo entry per
gesture, derived state as fresh-array computeds; labels in the existing voice:
"Replace", "Replace all", "Correct spelling", "Language", "Do not check
spelling". Components: `DeckSearchResults.vue` (props exactly `{ controller:
{ type: Object, required: true } }`, emits nothing, reads `c.groups.value`
like `TransitionsPanel.vue`) and `SlideProofingLayer.vue` (props
`{ highlightRects, underlineRects, scale }`, emits `spell-context`), typed
pass-throughs in a sibling `proofingRects.ts` — the text side's
`SpellUnderlineRect` carries `section`/`block`, coordinates a slide has no
equivalent of, so it is not shared. The find bar is transient state, not a
panel: **no `preferences.ts` key is added**, unlike Notes/Transitions/
Animations.

`PresentationView.vue` gains two `use*({ … })` calls, the `<FindBar>` dock,
`<DeckSearchResults>`, `<SlideProofingLayer>`, `<SpellContextMenu>`,
`<LanguageStatusControl>` and a fourth window `keydown` listener
(`search.onGlobalKeydown`). To pay for them, implementation first extracts
`initialModel`, `loadPresentation`, the
`initLibrary().then(loadPresentation)` re-sync and the `presentationId` watch
into `src/views/presentation/usePresentationLoad.ts` (deps `{ presentationId,
model, currentIndex, router, resetUndo, returnIndex, recomputeScale }`), then
ratchets `src/views/PresentationView.size.spec.ts`'s `LINE_BUDGET` **down** by
at least 15 from whatever it stands at when this lands (550 → 535 today), with
a ledger row naming this spec and its extraction appended. The budget never
rises.

### Tests
- `src/lib/core/findQuery.spec.ts` (new, over the moved functions);
  `src/lib/findReplace.spec.ts` unchanged and still green — coverage is
  preserved, never weakened.
- `src/lib/presentation/deckSearch.spec.ts` — ordering, every scope,
  group/foreign/covered-cell exclusion, the throwing-getter media assertion,
  back-to-front replace-all, per-slide dirty and notes-dirty marking per
  format, `replaceHit` returning `false` on a stale hit.
- `src/lib/presentation/text.spec.ts` — `replaceParagraphText` mark
  inheritance and re-coalescing; `lang`/`noProof` through `RunFormatPatch`,
  `RunMarks` and `selectionFormatState` ("mixed").
- `src/lib/presentation/pptx/serialize.spec.ts` and `odp/serialize.spec.ts` —
  the `a:rPr` and `style:text-properties` deltas, including the no-language
  case producing today's exact bytes.
- `src/lib/presentation/languageRoundTrip.spec.ts` (new, shaped like
  `transitionsRoundTrip.spec.ts`) — per-run `lang`/`noProof` survive save →
  reopen in both formats by field-level `toEqual`; untouched parts by
  `equalBytes` from `src/lib/core/equality.ts`; a replaced slide's part is
  byte-stable written twice. No semantic relation is claimed:
  `equalSemantic`/`resolve` in `src/lib/equality.ts` are typed
  `(model: DocumentModel)` and are not usable on a `PresentationModel`.
  `src/lib/presentation/roundTrip.spec.ts` is untouched.
- `src/views/presentation/useDeckSearch.spec.ts`, `useSlideSpellCheck.spec.ts`,
  `src/components/presentation/SlideProofingLayer.spec.ts`,
  `src/components/FindBar.spec.ts` (the two new props, existing cases
  unchanged), `src/views/PresentationView.size.spec.ts` (lowered budget +
  ledger row).
- `e2e/SPEC-098.spec.ts` per the last AC, over the checked-in
  `e2e/fixtures/mini-deck.pptx`. No new fixture is built, so
  `scripts/build-presentation-fixtures.ts` is unchanged.

### Specs this changes
- SPEC-081's out-of-scope line "find and replace, spell check on slides" is
  satisfied; its `runXml` gains the language attribute and its odp
  `StyleRegistry` a wider text-style key.
- SPEC-017 keeps ownership of the bar's behaviour; this spec adds the two
  scope props and makes the component explicitly shared by both products.
- SPEC-035 keeps ownership of the checker, dictionaries, suggestion ranking
  and the language model; a slide run resolves its language through the same
  `effectiveLanguage` chain, with `PresentationModel.language` standing in for
  `DefaultStyles.language`.
- `specs/000-roadmap.md` gains this spec's phase row; nothing else changes.
