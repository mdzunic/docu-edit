---
id: SPEC-115
title: Design a presentation from a description (experimental)
status: draft
area: structure
complexity: normal
story_points: 8
depends_on: [SPEC-037, SPEC-080, SPEC-083, SPEC-089, SPEC-094, SPEC-106, SPEC-113]
---
## Why
The hardest part of a deck is the first ten minutes: how many slides, in what
order, saying what, looking how. Everything after that is typing. SPEC-106
gave this product a way to build a deck from a plain JSON description without
touching a byte of PPTX; SPEC-113 gave it a single, consented, auditable way
to ask a language model a question. Put together, a user can describe the
deck they want — "a 10-slide update for the board on the Rotterdam migration,
one number per slide, restrained, our blue" — and get a real, editable deck
in the library.

The design that makes this safe is the same one that makes it good: **the
model's only output is a `DeckDescription`, validated by the code that
already exists, and rejected if it is not one.** Nothing generated is
executed, nothing is written as XML by a model, and nothing reaches the file
format that the builder would not have accepted from a hand-written JSON
file. A wrong answer produces a bad deck, which the user edits or deletes —
never a corrupt package, never a surprising network call, never a rewritten
document.

The second design decision is that the feature works with no model at all.
A deterministic outline-to-deck designer ships alongside, drives the tests,
and gives an answer to the user who says no to the whole idea.

## Acceptance criteria
- [ ] The feature is gated by SPEC-113's `deck-designer` experiment: flag off means no entry point renders, no module loads and no request is possible; flag on shows the entry points with SPEC-113's `experiment-badge`
- [ ] Two entry points, both creating something new and neither overwriting anything: the library's New menu gains "Presentation from a description…" (`new-deck-from-description`), and a deck's own menu gains "Redesign…" (`deck-redesign`) which changes **only** theme, layout assignment and slide order on the open deck as one undoable entry (SPEC-080's `record` → mutate → `commit()`), never rewriting slide text
- [ ] `DeckDesignerDialog.vue` (`deck-designer`) is a three-step flow with visible steps: **Brief** — a prompt field (`designer-brief`), audience/length/tone selects, an optional source (`designer-source`: nothing, pasted text, or a library document or deck chosen by id) and a live "what will be sent" byte count; **Plan** — an editable outline; **Build** — progress and the result. Back is available from every step and nothing is stored before the last one
- [ ] The plan step is a real review gate, not a spinner: the model returns a `DeckPlan` (per the Reference — deck title, theme name, slide size, and per slide a title, a layout id, 0–8 bullet intents and an optional note), rendered as an editable list where the user can edit any text, change any layout from a select, reorder with `Alt+↑`/`Alt+↓`, delete a slide and add one, and only then press "Build deck" — the plan the builder receives is what the user last saw, not what the model first said
- [ ] The design vocabulary is **closed** and comes from the product, not the model: layout kinds from SPEC-089's `SlideLayoutKind` union (`title`, `titleAndBody`, `sectionHeader`, `twoContent`, `titleOnly`, `blank` — never `other`), theme names from SPEC-094's `BUILT_IN_THEMES`, slide size from SPEC-090's presets — each enumerated in the request and each validated on return, with an unknown value replaced by the nearest default and recorded as a drop; a planned kind is resolved to the deck's own `layoutKey` through `layoutFor`/`STANDARD_LAYOUTS`, so a model that names a layout this product does not have cannot produce a deck that references one
- [ ] The build goes through SPEC-106 and adds no serialisation: the reviewed plan becomes a `DeckDescription`, `deckFromDescription` validates it, `PresentationBuilder` produces the `PresentationModel`, and the deck is stored through the existing presentation path — every validation failure is a `DocumentGenerationError` naming the offending path, shown to the user with the raw model output available in a `designer-raw` disclosure, and **nothing is stored when validation fails**
- [ ] Malformed output gets exactly one repair attempt, then stops: a response that is not parseable JSON, or that fails `deckFromDescription`, is sent back once with the validator's own error message appended and a demand for corrected JSON; a second failure ends the flow with the error, the raw text and a "Try again" that starts a new request — no loops, no silent retries, no partial deck
- [ ] Text is treated as text at every hop: the model's strings become `SlideRun` text through SPEC-106's `slideRuns.ts` and are never rendered as HTML, never interpreted as Markdown beyond SPEC-106's own grammar, never used as a URL, a file name or a route, and never followed as an instruction — an instruction inside a returned string is content, and the spec says so because the input to this feature can be a document the user did not write
- [ ] **The model-free path is a first-class feature, not a stub:** `src/lib/design/outlineDesigner.ts` exports `planFromOutline(source, options): DeckPlan` — a deterministic, pure designer that turns a text document's SPEC-037 heading outline (or pasted text's blank-line-separated blocks) into a plan: one section slide per Heading 1, one content slide per Heading 2 with its following paragraphs reduced to at most six bullets of at most 12 words, a title slide from the document's title, layouts chosen by a stated rule table, and a theme chosen from the deck's existing one or the default; it is offered as "Outline this document" whenever the provider is off, unreachable or declined
- [ ] The result opens where the user can judge it: the built deck is stored as a **new** library entry titled from the plan under SPEC-003's duplicate numbering and the app navigates to `/slides/:id` with the first slide selected; a toast offers "Undo" which trashes the new deck (SPEC-067's `trashDocument`), and the source document — if any — is untouched, asserted against a frozen model
- [ ] Provenance is recorded and visible: the deck's custom properties (SPEC-041's `customProperties`, carried by `core/opcProperties.ts` into both formats) gain `GeneratedBy` (this application), `GeneratedAt` and `GeneratedModel` (the model id, or `outline-designer` for the model-free path) — and **not** the prompt, unless `designer-keep-prompt` is ticked, which stores it in a `GeneratedPrompt` property and says that it will travel with the file
- [ ] Everything that goes over the wire is the user's choice and is shown before it goes: SPEC-113's consent dialog carries the exact request body, the source document's text is included only when the user picked one, it is truncated at `MAX_SOURCE_CHARS = 40_000` with the truncation stated in the dialog, and the request carries no library listing, no other document, no file names beyond the chosen source's title and no identifying information about the user
- [ ] Streaming is used where it helps and never where it misleads: the plan step streams tokens into the outline as plain text so a slow local model shows progress, the build step does not stream (a partial `DeckDescription` is not a deck), `designer-cancel` aborts through the `AbortSignal` at any point leaving nothing stored, and a timeout reports SPEC-113's `ModelTimeoutError` with the configured limit named
- [ ] Tests: `src/lib/design/outlineDesigner.spec.ts` — the rule table over a document with nested headings, a document with none, an empty document and a 300-heading document (bounded), all pure and exact; `src/lib/design/planToDescription.spec.ts` — every plan field mapped, the closed-vocabulary substitutions and their drop records, and a plan that validates producing a deck whose slide count and titles match; `src/views/design/useDeckDesigner.spec.ts` against a stubbed gateway — the repair-once-then-stop path, cancellation storing nothing, the flag-off path, the frozen-source assertion, and provenance properties present with the prompt absent by default; e2e (`e2e/SPEC-115.spec.ts`): with the `deck-designer` flag on and `page.route` stubbing the endpoint with a canned valid response → New › "Presentation from a description…" → type a brief → consent dialog shows the payload → Send → the plan lists 5 slides → edit slide 3's title and delete slide 4 → Build → `/slides/:id` opens with 4 slides, slide 3 carrying the edited title → the library shows one new deck and the source document unchanged → repeat with the stub returning invalid JSON → one repair request is observed, then the error and the raw text, and no new library entry → turn the provider off → "Outline this document" over a `.docx` fixture produces a deck with no network request at all

## Out of scope
Generating slide **images**, illustrations, icons or photographs — no image
model is involved and no image is fetched. Writing the body copy of an
existing deck, rewriting slides in place, summarising a document into prose,
translating, or any general "AI writing" surface in the editor: this feature
produces a deck structure and its text once, at creation. Charts and tables
as generated content (SPEC-091/092 own them; a plan that asks for one gets a
bullet slide and a drop record). Speaker-note generation beyond one optional
note per slide. Applying a brand: uploading a logo, extracting a palette from
one, or matching a corporate template — SPEC-094's named themes are the whole
design vocabulary. Multi-turn conversation, follow-up refinement of a built
deck, or a chat surface anywhere in the product. Evaluating or scoring the
result. Caching, replaying or sharing model responses. Any use of a model for
the word processor — text documents are untouched by this spec.

## Reference

### Dependencies restated
- `specs/106-…` — `PresentationBuilder`, `DeckDescription` +
  `deckFromDescription(description): PresentationModel` (validated through
  `generate/validate.ts`, every failure a `DocumentGenerationError` naming
  the path), `slideRuns.ts` mapping a `RunSpec` onto a `SlideRun`, the
  `DeckDropRecord { kind, location, reason }` shape over its own closed
  union, and the rule that generation **adds no serialisation code** — a
  construct the writers cannot emit is fixed in the writer. Everything this
  spec builds goes through that surface; it introduces no second path to a
  `.pptx`.
- `specs/113-…` — `isExperimentEnabled("deck-designer")`, the gateway's
  `complete`/`stream` with their `AbortSignal`, timeouts and typed errors,
  the per-origin consent dialog showing the exact payload, and the rule that
  a model's output is untrusted data.
- `specs/089-…` — `SlideLayoutKind`, `SlideLayout.key`, `STANDARD_LAYOUTS`,
  `layoutFor(model, slide)` and `applySlideLayout`, which is what "Redesign"
  calls per slide; `specs/094-…` — `BUILT_IN_THEMES` and `DEFAULT_THEME`; `specs/090-…` slide-size presets; `specs/083-…` `Slide.notes?:
  string[]`; `specs/080-…` the undo order and the presentation store path;
  `specs/037-…` the document outline (`outline.ts`) the model-free designer
  reads; `specs/041-…` custom document properties, carried by
  `core/opcProperties.ts` into both formats.

### The plan
```ts
// src/lib/design/plan.ts  (product-neutral in shape; lives beside the designer)
export interface DeckPlan {
  title: string;
  theme: string;                 // must be a BUILT_IN_THEMES name (SPEC-094)
  size: "16:9" | "4:3";
  slides: PlannedSlide[];        // 1..MAX_PLANNED_SLIDES (60)
}

export interface PlannedSlide {
  layout: SlideLayoutKind;       // SPEC-089's closed union, minus "other"
  title: string;
  bullets: string[];             // 0..8, each <= 200 characters
  note?: string;
}

export function validatePlan(value: unknown): { plan: DeckPlan; dropped: DeckDropRecord[] };
export function planToDescription(plan: DeckPlan): DeckDescription;
```
`validatePlan` is where the closed vocabulary is enforced: an unknown layout
or theme is replaced by the default and recorded; a slide count over the cap
is truncated and recorded; a bullet over the length cap is truncated and
recorded. The user sees the drops as a "What we changed" list on the plan
step, before building.

### The request
One system prompt describing the output contract and enumerating the
available layouts and themes, and one user message carrying the brief, the
chosen length and tone, and the truncated source text when there is one.
`json: true`. The system prompt is a checked-in constant with a unit test
asserting the enumerations are generated from `SPEC-089`/`094`'s own
registries rather than typed twice — a vocabulary that drifts from the code
is how a validator starts rejecting everything.

The response is parsed with `JSON.parse` inside a `try`, and nothing else.
No `eval`, no `Function`, no dynamic import, no fetch of anything the
response names.

### The model-free designer
```ts
// src/lib/design/outlineDesigner.ts
export function planFromOutline(source: OutlineSource, options?: { maxSlides?: number }): DeckPlan;
```
Rules, in order, each tested:
1. A title slide from the document title (and its first Heading 1's following
   paragraph as a subtitle, when there is one).
2. One section slide per Heading 1 that has Heading 2 children; a Heading 1
   with none becomes a content slide.
3. One content slide per Heading 2, its bullets taken from the following
   paragraphs and list items, at most 6, each cut at 12 words on a word
   boundary with an ellipsis.
4. Headings deeper than 2 fold into their parent's bullets, indented one
   level (SPEC-106's `{ text, level }` rows).
5. Layout per slide: `title` for 1, `sectionHeader` for 2, `titleAndBody`
   for a slide with bullets, `titleOnly` for one without.
6. Theme: the source deck's, else the product default. Size: 16:9.

It is pure, deterministic and needs neither a model nor a network, which is
why it is also the designer the e2e path exercises for the no-provider case.

### Why the model never writes the file
Three properties fall out of routing everything through
`deckFromDescription`: a malformed response cannot produce a malformed
package (the validator refuses it first); the deck a user gets is exactly the
deck SPEC-106's API would have produced from the same JSON, so there is one
behaviour to test and one to support; and the feature's blast radius is a bad
first draft in a new library entry, which is the same blast radius as any
other "New from template". The alternative — asking a model for PPTX, or for
code — has none of those properties.

### Tests
- `src/lib/design/plan.spec.ts`, `outlineDesigner.spec.ts`,
  `planToDescription.spec.ts` as in the acceptance criteria — all pure, all
  model-free, all fast.
- `src/lib/design/prompt.spec.ts` — the enumerations in the system prompt are
  generated from `SlideLayoutKind` and `BUILT_IN_THEMES` and match them
  exactly.
- `src/views/design/useDeckDesigner.spec.ts` against a stubbed gateway, as in
  the acceptance criteria.
- `e2e/SPEC-115.spec.ts` as in the acceptance criteria. Per SPEC-113, the
  endpoint is stubbed with `page.route`; **no test in this repo may require a
  running model.**

### Specs this changes
- `specs/113-…` — `EXPERIMENTS` gains the `deck-designer` entry.
- `specs/106-…` — no contract change; this spec is a caller. Any construct it
  needs and the builder lacks is added there, not here.
- `specs/067-…` — the New menu gains one entry; `specs/080-…` — the deck menu
  gains "Redesign…", which records one undo entry.
- `specs/000-roadmap.md` gains this spec's phase row.
