---
id: SPEC-083
title: Speaker notes
status: draft
area: editing
depends_on: [SPEC-081]
---
## Why
Real decks carry a second document under the slides: the presenter's own
notes. Both formats model them (`.pptx` notes-slide parts, `.odp`
`presentation:notes` inside each page); today the prep milestone preserves
them as opaque bytes and shows nothing. This spec reads them, shows them
in a notes panel under the stage, lets them be edited as plain text, and
writes them back — the last per-slide content channel before playback
(SPEC-084 does not use notes; a presenter view that does is future work,
but it can only exist once notes are modelled) and a required input for
Keynote import (SPEC-087 imports notes into this field).

## Acceptance criteria
- [ ] `Slide` gains `notes?: string[]` (plain paragraphs, empty array never stored — absent means none); both readers populate it: pptx from the slide's notes-slide part's body placeholder text, odp from the page's `presentation:notes` text frame — formatting flattened to text, paragraph breaks kept
- [ ] `/slides/:id` gains a notes panel below the stage: a toggle button in the toolbar ("Notes", pressed-state, testid `slides-notes-toggle`) and Cmd/Ctrl+Shift+H toggle it; the panel shows "Click to add speaker notes" when empty and the collapsed/expanded choice persists per device (a `presentationNotesOpen` preference via the SPEC-066 preferences store; default collapsed)
- [ ] The panel is a plain multi-line text editor for the current slide's notes (textarea semantics, label "Speaker notes for slide N"); edits update the model per input, coalesce into undo bursts exactly like SPEC-081 typing, and switching slides commits the burst
- [ ] Notes edits mark the slide dirty; the pptx writer rebuilds only the affected notes-slide parts (the slide part itself stays byte-verbatim when only notes changed), and the odp writer rebuilds the page with its `presentation:notes` regenerated while the page's other preserved XML stays put
- [ ] A slide that never had notes gains, on first notes text in a `.pptx` save: a notes-slide part named `ppt/notesSlides/notesSlideN.xml` (N unique), its `.rels` (→ its slide, → the notes master), the slide's `.rels` gaining the notesSlide relationship, Content-Types overrides, and — when the package has no notes master at all — a minimal `ppt/notesMasters/notesMaster1.xml` (+ rels + Content-Types + `p:notesMasterIdLst` splice into `presentation.xml`) so Office opens the result without repair
- [ ] Deleting all notes text removes the model field; the pptx writer then omits the notes-slide part and its relationship (a package that never needed notes machinery is not left with orphan parts referencing empty notes); the odp writer omits `presentation:notes`
- [ ] Notes survive the round trip in both formats: edit → save → reopen shows the text (unit round trip); a deck whose notes were never touched re-emits its notes parts byte-verbatim
- [ ] Notes text appears in no slide rendering, no outline (`slideOutline` is body-only, unchanged), and no export surface until a spec says otherwise
- [ ] The fixture decks gain a slide with notes in both formats so the readers' extraction is pinned to exact strings
- [ ] e2e (`e2e/SPEC-083.spec.ts`): open the pptx fixture → toggle the notes panel → fixture notes text visible → edit it → switch slides and back (text kept per slide) → download, reopen → edited notes visible

## Out of scope
Formatted notes (bold etc. flatten to plain text on first edit of that
slide's notes; untouched notes keep their original formatting bytes via
the verbatim path). A presenter/rehearsal view. Printing notes. Notes in
the HTML export (SPEC-088 owns that decision). Notes-page layout editing
(the generated notes-slide uses a fixed body-placeholder layout).
Importing notes from formats other than the two native ones (SPEC-087
covers Keynote).

## Reference

### Dependencies restated
- SPEC-079 model + `preserved` capture (odp `presentation:notes` moves
  from `preserved` to the modelled field in this spec — the reader stops
  capturing it as preserved XML and parses it instead; pptx never put
  notes in `preserved` since they live in sibling parts).
- SPEC-080/081 writers: substitution save path, dirty rules, splice
  utilities (Content-Types/rels/manifest), serialiser conventions,
  `usePresentationUndo`, write-through persistence.
- SPEC-066 preferences store for the panel's persistence key.

### Reading
pptx: from each slide part's `.rels`, follow
`…/relationships/notesSlide`; in the notes part, the text of the
`p:sp` whose `p:ph type="body"` (the notes body placeholder) — other
shapes (slide image placeholder, headers) are ignored; paragraphs map
1:1 to `notes` entries. odp: the `draw:frame`
`presentation:class="notes"` inside `presentation:notes`; its
`draw:text-box` paragraphs map to entries (`draw:page-thumbnail` is
ignored). Both readers strip trailing empty paragraphs.

### Writing
pptx notes-slide serialisation (extends `presentation/pptx/serialize.ts`):
```
<p:notes …ns…><p:cSld><p:spTree>
  (scaffold)
  <p:sp>  — slide-image placeholder: p:ph type="sldImg", fixed geometry
  <p:sp>  — body: p:ph type="body" idx="1", fixed geometry (
            1.32cm margins on a 6.86×9.14cm-ish notes page — use the
            notes master's size: default p:notesSz 6858000×9144000 EMU)
    one a:p per notes entry, single run, no explicit rPr
</p:spTree></p:cSld></p:notes>
```
A slide whose notes part existed and only whose text changed: rebuild
that part with the same name, keeping its `.rels` verbatim. The minimal
notes master mirrors the from-scratch slide master's construction
(SPEC-080) with `p:notesMasterIdLst` spliced into `presentation.xml` —
this is the one case where a notes edit touches `presentation.xml`; the
splice preserves the rest of the part byte-verbatim like SPEC-080's
`sldIdLst` splice.

odp: `serializePage` gains, before `</draw:page>` (and before preserved
`anim:*`):
```
<presentation:notes draw:style-name={original or generated}>
  <draw:frame presentation:class="notes" …fixed geometry…>
    <draw:text-box><text:p>…</text:p>…</draw:text-box>
  </draw:frame>
</presentation:notes>
```
For pages whose notes were untouched but which dirty for other reasons,
the original `preserved["presentation:notes"]` — captured before this
spec — no longer exists; the reader now round-trips notes through the
model, and the golden serialiser tests pin that a parse→serialise cycle
of an Impress-authored notes block loses only formatting, never text.

### View
`components/presentation/NotesPanel.vue` + a small
`views/presentation/useNotes.ts` wiring composable. The panel sits in a
vertical flex with the stage; the stage's fit-scaling (SPEC-079
ResizeObserver) shrinks accordingly — filmstrip untouched. Panel height
fixed (7 lines) with its own scroll; no resizer in this spec.

### Tests
- Reader specs: extraction from both fixture decks (exact strings),
  trailing-empty-paragraph stripping, notes-less slides.
- Serialiser specs: golden notes-slide XML; first-notes package surgery
  (rels, Content-Types, notes master creation, `p:notesMasterIdLst`
  splice); removal path leaves no orphans (walk the rebuilt package's
  rels: no target missing, no notesSlide override without a part).
- Round trip per AC, plus: notes-only edit keeps the slide part
  byte-verbatim (pptx).
- View spec: toggle + preference persistence, per-slide binding, burst
  undo.
- e2e per AC.

### Specs this changes
- SPEC-079's `preserved` key list drops `presentation:notes` (odp) —
  superseded by the modelled field; `p:timing`/`p:transition` handling
  is unchanged.
- SPEC-087 (Keynote import) and any future presenter view consume
  `Slide.notes` as defined here.
