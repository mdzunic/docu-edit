---
id: SPEC-101
title: Comments on slides — review threads, markers and both wire formats
status: draft
area: collab
complexity: normal
story_points: 8
depends_on: [SPEC-038, SPEC-079, SPEC-080, SPEC-081, SPEC-082, SPEC-084, SPEC-085, SPEC-086]
---
## Why
A deck is reviewed far more often than it is authored, and the presentation product
has no review surface at all. SPEC-038 gave the word processor threads, replies and
both wire mappings; the slides side has nothing. The bytes do survive —
`ppt/comments/comment1.xml` and `ppt/commentAuthors.xml` land in
`PresentationModel.opaque`, an `office:annotation` inside `draw:page` survives inside
`Slide.sourceXml` — so today's behaviour is silent invisibility: three reviewers'
work is carried and never shown, and the first `.odp` edit rebuilds the page from a
model that never knew the annotations existed.

The shape mirrors the text side rather than forking it: the thread/message/resolved
vocabulary, the `"{author} ({createdAt}): {text}"` reply encoding both `.docx` and
`.odt` already use, the initials rule and the settable author identity are all
product-neutral, `src/lib/core/commentIdentity.ts` is **already** in core and the
rest moves down beside it. What is new is the anchor: a slide comment pins to the
slide, to a point on it, or to a frame — not to a character range.

The wire side is fiddly. pptx has **two comment generations** (ECMA-376's `p:cmLst`
and PowerPoint 365's `p188:cmLst`) that can both be present in one package, and ODF
has neither replies nor a standard resolved flag; every uncertain byte below is
routed through SPEC-086's golden-fixture-first process instead of being invented.
**This file is the normative source for the slide comment model, its anchor
semantics, the two pptx generations' read/author rules and the odp
`office:annotation` mapping.**

## Acceptance criteria
- [ ] `Slide` gains `comments?: SlideComment[]`, `commentSource?: SlideCommentSource` (pptx retention) and `commentsDirty?: true` per the Reference — author/initials/text/`createdAt`/`resolved`, ordered `messages` with `messages[0]` the opening comment, and `anchor` one of `{kind:"slide"}` / `{kind:"point";xCm;yCm}` / `{kind:"frame";frameId;dxCm;dyCm}`; absent `comments` and `[]` both mean "no comments", ids are `commentId`/`messageId` (never `id` — `equalStructural`'s `IGNORED_KEYS` drops `"id"`) and are **derived** from the wire on read so re-parsing a written file yields identical ids
- [ ] The product-neutral half moves into a new `src/lib/core/comments.ts` by the Reference's two-phase move — `CommentMessage` (from `src/lib/documentModel.ts`), `replyParagraphText`/`parseReplyParagraphText`/`canModifyMessage` (from `src/lib/comments.ts`) and `commentInitials` (the private `initialsOf` at `src/lib/docx/writer.ts:1772`) — with `src/lib/core/commentIdentity.ts` reused as-is; both boundary specs stay green and every existing text-side comment test passes unchanged, expectations preserved and never weakened
- [ ] pptx read: for each slide, the `.rels` `comments` relationship yields a legacy `p:cmLst` part and the modern relationship a `p188:cmLst` part; authors resolve through `ppt/commentAuthors.xml` (legacy `p:cmAuthor@id/@name/@initials`) and `ppt/authors.xml` (modern) — when both parts exist the **modern** set is modelled and the legacy part stays verbatim, `commentSource` records `{partName, generation, otherPartName?}`, and anything out of subset leaves the slide with **no** `comments` and every part byte-verbatim (comments never half-parse)
- [ ] odp read: `office:annotation` children of `draw:page` become comments — `dc:creator` → author, `dc:date` → `createdAt`, `svg:x`/`svg:y` (via `odfLengthToCm` in `src/lib/core/odfXml.ts`) → the anchor point, absent coordinates → `{kind:"slide"}`, first `text:p` → the opening message and each later `text:p` through `parseReplyParagraphText` → a reply (a non-matching paragraph folds into the previous message's text, exactly as the `.odt` reader does); reading sets nothing dirty and does not touch `sourceXml`
- [ ] Anchors resolve through `src/lib/presentation/comments.ts`: a decoded point inside the topmost non-`foreign` frame's rect becomes `{kind:"frame"}` with `dxCm`/`dyCm` from that frame's top-left, so moving or resizing the frame carries the marker; deleting the frame (SPEC-082 `deleteFrame`) degrades the anchor to `{kind:"point"}` at its last absolute position and **never** deletes the comment; a point at or outside the slide origin (`x<=0 && y<=0`) reads as `{kind:"slide"}`
- [ ] A `Comments` toolbar toggle (`slides-comments-toggle`, in `PresentationToolbar.vue`'s right-aligned `div.tb-group.tb-slide-actions` beside `slides-transition-toggle`, prop `commentsOpen`, emit `comments-toggle`, keyboard-operable, fitting the 720px e2e viewport) opens `SlideCommentsPanel.vue` (`slides-comments-panel`, `<section class="comments-panel">`, `aria-label="Comments"`, `max-height: 45%`, `overflow-y: auto`), a dumb `:controller`-taking component like `TransitionsPanel.vue`; its open state persists as `presentationCommentsOpen` in `src/lib/preferences.ts` (interface ~line 88, default `false` ~line 148, `isBoolean` ~line 211)
- [ ] The panel groups threads **by slide** in deck order (`slide-comment-group`, heading "Slide N", the current slide's group expanded and scrolled into view), each thread `slide-comment-thread-${commentId}` listing `slide-comment-message-${messageId}` rows (author, initials chip, `formatTimestamp` from `src/lib/format.ts`, text), with `slide-comment-reply-input` + `slide-comment-reply-submit`, `slide-comment-resolve`/`slide-comment-reopen`, `slide-comment-edit`, `slide-comment-delete` (own messages only, via `canModifyMessage`), a filter `slide-comment-filter` (`all` | `unresolved`, default `all`), the identity field `slide-comment-author-name` (`loadCommentAuthor`/`saveCommentAuthor`), `slide-comment-add` + `slide-comment-composer`/`-input`/`-submit`, and `slide-comment-empty` when nothing matches; comment text renders as plain text everywhere — never as markup, never `v-html` — with a security test feeding `<img src=x onerror=…>` through the panel, the marker tooltip and the badge title
- [ ] Markers render on the stage in a new `SlideCommentLayer.vue` overlay inside `stage-workspace` (a sibling of `AnimationStageLayer`, positioned through `cmToStagePx` from `src/views/presentation/stageMath.ts`): a numbered `slide-comment-marker` / `slide-comment-marker-${index}` button per thread at its anchor point (slide-anchored threads stack at the slide's top-left inset), showing the author's initials, dimmed when resolved, hidden while the panel is closed; clicking a marker focuses its thread and clicking a thread highlights its marker and its anchoring frame; `SlideFilmstrip.vue` shows `filmstrip-comment-badge` on any slide with an unresolved thread, its count from `unresolvedCommentCount(slide)`
- [ ] Add / reply / edit / resolve / reopen / delete run through `src/lib/presentation/comments.ts` (framework-free, presentation siblings + `lib/core` only); each records exactly one undo entry before mutating and one `commit()` after — labels "Add comment", "Reply to comment", "Edit comment", "Resolve comment", "Reopen comment", "Delete comment" — and ends with `markCommentsEdited(slide, format)`, which sets `slide.commentsDirty = true` on pptx and calls `markSlideEdited(slide, "odp")` on odp (the exact guard `useNotes.ts:79` uses, because `writeOdp` consults no `*Dirty` flag beyond `dirty`/`structureDirty`); a first edit on a slide that carried **both** pptx generations instead parks in `pendingReplace` (SPEC-085/086's discipline) behind a `role="alertdialog"` `aria-label="Replace the older copy of these comments"` with `slide-comment-replace-confirm`/`-confirm-button`/`-cancel`, confirming re-authoring **both** parts from the one model so they can never diverge and cancelling changing nothing while bumping `formNonce`
- [ ] pptx write: `commentsDirty` joins `writePptx`'s verbatim guard (`writer.ts:80` becomes `!structureDirty && !anyDirty && !anyNotes && !anyComments`) and adds a `planComments` slot to `buildTextEditSubstitution` and `buildSubstitution` shaped exactly like `planNotes` in `pptx/notesWriter.ts` — the comment part substituted or minted as `ppt/comments/comment{n}.xml` past the highest existing, the slide `.rels` gaining the relationship, `[Content_Types].xml` gaining one `Override`, and `ppt/commentAuthors.xml` **spliced** (an existing `p:cmAuthor` with a matching `@name` is reused; otherwise one is appended at `max(@id)+1`) so authors with no live comments survive; the slide's own part stays byte-verbatim when only comments changed, and a deleted slide takes its comment part and relationship into `removed`
- [ ] odp write: `serializePage` (`src/lib/presentation/odp/serialize.ts:188`) emits the slide's `office:annotation` elements between `{frames}` and `{notes}`, declaring `xmlns:dc` locally the way `odf/writer.ts#annotationStartXml` does, and a page with no modelled comments emits none — so a deck that never touched comments still splices its pages byte-identically per the `odpContentEnvelope` invariant
- [ ] Comments never reach playback or export: `PresentationPlayView.vue`, `useBuildPreview` and SPEC-088's HTML export render no markers, no badges and no comment text, and a slide's comments do not participate in `slideBuildGroups`, `slideOutline` or `unconvertibleFrameCount`; `deleteSlide` takes its comments with it and undo restores both (they live in `usePresentationUndo`'s `structuredClone(model.slides)` snapshot already); `duplicateSlide` and a slide imported by SPEC-099 arrive with `comments` and `commentSource` **dropped** (review state does not travel with copied content, and the source deck's author ids do not transfer); `stripRetention` drops `commentSource`/`commentsDirty` and keeps `comments`
- [ ] Round trip: a no-edit download stays byte-identical part-for-part under `equalBytes` from `src/lib/core/equality.ts` after `unzip` (there is no presentation-shaped `equalSemantic` — `resolve`/`equalSemantic` in `src/lib/equality.ts` are typed against `DocumentModel` — so the modelled relation here is `equalStructural` from `src/lib/equality.ts` plus field-level `toEqual`); add/reply/resolve → save → reopen reproduces every modelled field in both formats, writing the reopened model twice is byte-stable for the comment parts, and the **one sanctioned delta** is that the legacy pptx `p:cm` form cannot carry `resolved` — a resolved thread on a legacy-authored slide reopens unresolved, and the test pins exactly that
- [ ] Fixtures: `src/lib/presentation/fixtures/build.ts` gains `buildPptxCommentsFixtureBytes`/`buildOdpCommentsFixtureBytes` + `EXPECTED_FIXTURE_COMMENTS` under a new banner (the four mini-deck slides are never mutated), `scripts/build-presentation-fixtures.ts` writes `e2e/fixtures/spec101-comments.pptx` and `.odp`, and each deck carries a legacy thread with a reply, a slide-anchored thread, a resolved thread, one slide carrying **both** pptx generations, and one deliberately out-of-subset comment part pinning the preserve-don't-parse path
- [ ] `PresentationView.vue` gains exactly one `useSlideComments({ model, currentSlide, currentIndex, record, commit, selection: stageEdit.selection })` wiring call plus its template hookup, paid for by extracting the stage overlay stack (`AnimationStageLayer` + `SlideTransitionLayer` + `SlideCommentLayer`) into `src/components/presentation/StageOverlays.vue`; `PresentationView.size.spec.ts`'s `LINE_BUDGET` ratchets **down** (to at most 545, lower if a sibling spec already lowered it — the budget only ever decreases) with a new ledger row naming this spec
- [ ] e2e (`e2e/SPEC-101.spec.ts`): open `e2e/fixtures/spec101-comments.pptx` → panel shows the reviewer's thread with its reply, grouped under "Slide 2", and a marker sits over the body frame → reply as "You" → resolve the second thread and see it leave the unresolved filter and its filmstrip badge drop → add a new comment on slide 3 by clicking the stage and typing → Present and confirm no marker or badge is visible in playback → exit, download, reopen → all three threads, their authors, timestamps and anchors come back, and the deck's out-of-subset comment part is still byte-identical; the read half repeats against `spec101-comments.odp`

## Out of scope
@-mentions, notifications and comment sync between users (SPEC-044's territory).
Rich text, images or links inside a comment (plain text only, as SPEC-038 fixed).
Track changes on slides — neither ODF nor PresentationML has a `w:ins`/`w:del`
equivalent for slide content and none is modelled here. Comment-anchored **text
ranges** and per-paragraph anchors inside a `TextFrame` (the anchor is the frame).
Authoring the modern pptx form for a slide that arrived legacy-only, or converting
between generations beyond the both-present confirm. Notes-slide, layout and master
comments (their parts stay verbatim `OpaquePart`s). Printing or exporting comments
(SPEC-088 renders none). Reviewer colour assignment, comment counts on the library
card (SPEC-067) and any per-deck "review complete" state.

## Reference

### Dependencies restated
- `specs/038-comments.md` — `CommentThread { threadId; messages: CommentMessage[];
  resolved: boolean }` and `CommentMessage { messageId; author; text; createdAt }`
  (ISO 8601) in `src/lib/documentModel.ts`; `src/lib/comments.ts`'s
  `replyParagraphText(message)` → `` `${author} (${createdAt}): ${text}` ``,
  `parseReplyParagraphText(text)` → `{author, createdAt, text} | null`,
  `canModifyMessage(message, author)`; identity in `src/lib/core/commentIdentity.ts`
  (**already core**) — `loadCommentAuthor()`, `saveCommentAuthor(name)`,
  `DEFAULT_COMMENT_AUTHOR = "You"`, `COMMENT_AUTHOR_STORAGE_KEY`. UI precedent
  `src/components/CommentsPane.vue`.
- `specs/079-slide-canvas-rendering.md` — `Slide { id; frames; background?; notes?;
  dirty?; notesDirty?; transition?; animations?; partName?; sourceXml?; preserved? }`,
  `SlideFrame` = text|image|shape|group|foreign with cm geometry (`xCm`, `yCm`,
  `widthCm`, `heightCm`); lengths cm, sizes pt, colours `#rrggbb`; frame ids minted
  at read and never serialised; `preserved` omitted when the record would be empty.
- `specs/080-create-presentation-and-slide-management.md` — `stripRetention`,
  `duplicateSlide`/`deleteSlide`, and `usePresentationUndo.ts`: `record(label)`
  **before** the mutation, `commit()` after, `Snapshot = structuredClone(model.slides)`
  + `structureDirty`, `MAX_UNDO_STEPS = 100`, one entry per commit.
  `specs/084-slideshow-playback.md` — `usePlayback` and `PlaybackStep { slideIndex;
  buildIndex }`, unchanged by this spec.
- `specs/081-slide-text-editing.md` — `serializeSlideXml(slide, rels, slidePartPath)`
  (`pptx/serialize.ts`), `serializePage(slide, registry, usedNames, master,
  styleSource)` (`odp/serialize.ts`) and `markSlideEdited(slide, format)`
  (`text.ts:340` — sets `dirty`, and on odp deletes `sourceXml`).
  `specs/082-slide-objects-images-shapes-text-boxes.md` — frame identity,
  `frames.ts`'s `deleteFrame`/`moveFrame`/`resizeFrame`/`duplicateFrame` and
  `frameNoun`/`frameAccessibleName` from `useFrameManipulation.ts`.
- `specs/085-slide-transitions.md` / `specs/086-slide-animations.md` — the
  golden-fixture-first process for unconfirmed bytes (a hand-authored golden module,
  the serialiser asserted equal to it, a tolerant reader, the four-test shape), the
  dumb `:controller` panel contract, and the `pendingReplace` → confirm dialog →
  `formNonce` re-mount discipline this spec's both-generations confirm reuses.
- Core consumed unchanged: `core/xml.ts` (`parseXml`, `serializeElement`,
  `parseRelationships`, `resolvePartPath`, `relsPathFor`, `escapeXmlAttribute`),
  `core/odfXml.ts` (`odfLengthToCm`), `core/packageParts.ts` (`OpaquePart`,
  `decodedPartBytes`), `core/equality.ts` (`equalBytes`).

### The model (`src/lib/presentation/model.ts`)
```ts
export type SlideCommentAnchor =
  | { kind: "slide" }
  | { kind: "point"; xCm: number; yCm: number }
  /** dx/dy offset the marker from the frame's top-left, so it travels with
   *  the frame. Flattened to an absolute point on write. */
  | { kind: "frame"; frameId: string; dxCm: number; dyCm: number };

/** `initials`: up to 3 from `commentInitials(author)` unless the source
 *  carried its own. `CommentMessage` comes from `core/comments.ts`. */
export interface SlideCommentMessage extends CommentMessage { initials?: string }

export interface SlideComment {
  commentId: string;                 // derived from the wire; never `id`
  messages: SlideCommentMessage[];   // [0] opens the thread; [1..] are replies
  resolved: boolean;
  anchor: SlideCommentAnchor;
}
export interface SlideCommentSource {   // pptx retention; absent on odp
  partName: string;                     // "ppt/comments/comment2.xml"
  generation: "legacy" | "modern";
  otherPartName?: string;               // the generation kept verbatim
}
```
`Slide` gains `comments?: SlideComment[]`, `commentSource?: SlideCommentSource`,
`commentsDirty?: true`; everything stays plain JSON, so `roundTrip.spec.ts`'s
`JSON.parse(JSON.stringify(model))` `toEqual(model)` still holds. Format-neutral
commands live in `src/lib/presentation/comments.ts`: `createCommentId`,
`createCommentMessageId`, `unresolvedCommentCount(slide)`,
`commentAnchorPointCm(slide, comment)`, `resolveCommentAnchor(slide, xCm, yCm)`
(topmost non-`foreign` frame hit wins), `addSlideComment`, `addCommentReply`,
`editCommentMessage`, `deleteCommentMessage`, `deleteSlideComment`,
`setCommentResolved`, `detachCommentsForFrame`, `dropSlideComments`,
`markCommentsEdited(slide, format)`.

### Core hoist (two-phase move)
Phase 1 — create `src/lib/core/comments.ts` holding `CommentMessage`,
`replyParagraphText`, `parseReplyParagraphText`, `canModifyMessage` and
`commentInitials`, with `src/lib/documentModel.ts` and `src/lib/comments.ts`
re-exporting them so **no** existing importer, spec or fixture changes; it imports
nothing, so `src/lib/core/boundary.spec.ts` stays green. Phase 2 —
`src/lib/docx/writer.ts` imports `commentInitials` and drops its private
`initialsOf`, and `src/lib/presentation/comments.ts` imports from `../core/comments`,
satisfying `src/lib/presentation/boundary.spec.ts`. Behaviour is unchanged throughout.

### pptx wire mapping — two generations
Legacy (ECMA-376, the interoperable default we author): part
`ppt/comments/comment{n}.xml`, content type
`application/vnd.openxmlformats-officedocument.presentationml.comments+xml`,
relationship type `…/officeDocument/2006/relationships/comments` from the slide part;
root `<p:cmLst>` of `<p:cm authorId dt idx>` with `<p:pos x= y=/>` and
`<p:text>…</p:text>`. Authors: `ppt/commentAuthors.xml`, `<p:cmAuthorLst>` of
`<p:cmAuthor id name initials lastIdx clrIdx/>`, related from `ppt/presentation.xml`.
`commentId` derives as `` `cm-${authorId}-${idx}` ``; replies encode through
`replyParagraphText` (legacy `p:cm` has no thread structure) and `resolved` has no
legacy representation — see the round-trip AC's sanctioned delta.

Modern (PowerPoint 365; read always, authored only for a slide that arrived modern):
part `ppt/comments/modernComment_*.xml`, root `p188:cmLst` in namespace
`http://schemas.microsoft.com/office/powerpoint/2018/8/main`, `p188:cm` carrying a
GUID id, an author GUID, a created timestamp, `p188:pos`, an `a:`-namespaced text
body and a `p188:replyLst` of real replies; authors in `ppt/authors.xml`.

**Uncertain wire details, routed through SPEC-086's golden-fixture-first process and
pinned by `src/lib/presentation/pptx/commentsGolden.ts` — do not invent bytes:** the
unit of `p:pos/@x`/`@y` (the schema types it `a:CT_Point2D`, i.e. EMU, yet real
PowerPoint files carry values two orders of magnitude too small for that), the exact
modern element/attribute names, its content type and relationship type, and whether
the modern form carries a `resolved`/`status` attribute. Until a golden pins one the
reader treats it as out of subset and preserves; a decoded point outside the slide
bounds degrades to `{kind:"slide"}` rather than placing a marker off-stage. Module
contract mirrors `pptx/timing.ts` exactly, private `const ABORT =
Symbol("out-of-subset")` included: `writePptxComments(comments, ctx)` → `""` for
none, `readPptxComments(cmLstEl, authors, resolveAnchor)` → `undefined` out of subset.

### odp wire mapping
`<office:annotation svg:x svg:y svg:width svg:height office:name>` as a child of
`draw:page`, with `<dc:creator>`, `<dc:date>` and one `<text:p>` per message — the
opening message verbatim, then each reply through `replyParagraphText`, exactly the
encoding `src/lib/odf/writer.ts#annotationStartXml` already produces for `.odt`.
There is no `office:annotation-end` on a slide (the anchor is a point, not a range);
a slide-anchored comment omits `svg:x`/`svg:y`; lengths go through `odfLengthToCm`;
`commentId` derives from `office:name`, else from the page-order index. Resolved
state has **no standard ODF representation**: the intended encoding is
`loext:resolved="true"` (`urn:oasis:names:tc:opendocument:xmlns:loext:1.0`, declared
locally like `xmlns:dc`) — **unconfirmed**, so pinned by
`src/lib/presentation/odp/commentsGolden.ts` plus a manual Impress round trip
recorded in the PR, and if that fails odp inherits the legacy pptx form's sanctioned
delta. Module contract mirrors `odp/anim.ts`: `writeOdpComments(comments): string`,
`readOdpComments(elements): SlideComment[] | undefined`.

### Writer rules
pptx (`src/lib/presentation/pptx/writer.ts`): `const anyComments = model.slides.some(s =>
s.commentsDirty === true)` lands beside `anyDirty`/`anyNotes` (lines 78-79) and joins
the verbatim guard at line 80; `planComments(input):
CommentsPlan` mirrors `planNotes` in `pptx/notesWriter.ts` — `{ parts, relationships,
contentTypeAdditions, removed }` — and its parts slot into the assembly precedence
`[Content_Types].xml` → `notes.parts` → `comments.parts` → `rebuilt` →
`media.override` → original bytes, in both `buildTextEditSubstitution` and
`buildSubstitution`. `ppt/commentAuthors.xml` is spliced, never regenerated. A slide
whose only change is comments keeps its own part, `presentation.xml`, its rels and
`[Content_Types].xml` byte-verbatim unless a part was minted. odp
(`src/lib/presentation/odp/writer.ts`): a comment edit routes through
`markSlideEdited(slide, "odp")`, so `structural` is already true,
`rebuildContentXml` re-emits the page and `serializePage` emits the annotations
between `{frames}` and `{notes}` — no new substituted part, `styles.xml` and
`META-INF/manifest.xml` untouched.

### UI
`src/views/presentation/useSlideComments.ts` — deps `{ model, currentSlide,
currentIndex, record, commit, selection }`; returns `open, toggle, groups, threads,
filter, setFilter, activeCommentId, focusThread, authorName, setAuthorName,
composerOpen, openComposer, cancelComposer, submit, reply, edit, remove, resolve,
reopen, markers, badgeCountFor, pendingReplace, confirmReplace, cancelReplace,
formNonce`. Derived state is exposed as fresh-array computeds
(`(currentSlide.comments ?? []).slice()`) per `useSlideRender.ts`'s reactivity rule.
`SlideCommentsPanel.vue` and `SlideCommentLayer.vue` declare runtime
`defineProps({ controller: { type: Object, required: true }, … })` (espree rejects
TS-only syntax in `.vue` scripts); typed transforms (`commentMarkerPoint`,
`commentBadgeLabel`) go in `src/components/presentation/slideRender.ts`.

### Tests
- `src/lib/core/comments.spec.ts` — the hoisted encoding and initials rule; existing
  `src/lib/comments.spec.ts`, `docx/comments.spec.ts` and `odf/comments.spec.ts` pass
  unchanged, proving the move behaviour-preserving.
- `src/lib/presentation/comments.spec.ts` — anchor resolution and the frame-delete
  degrade, add/reply/edit/resolve/delete, `markCommentsEdited`'s per-format guard,
  `unresolvedCommentCount`, drop-on-duplicate.
- `src/lib/presentation/pptx/comments.spec.ts` and `odp/comments.spec.ts` in
  SPEC-086's four-test golden shape: golden equality, `""` for nothing, golden →
  model, a PowerPoint/Impress-flavoured tolerant variant, out-of-subset → `undefined`.
- `pptx/writer.spec.ts` / `odp/serialize.spec.ts` — part minting, rels and
  Content-Types splices, the authors-part splice, the comments-only-change
  byte-verbatim slide part, the annotation slot's position.
- New `src/lib/presentation/commentsRoundTrip.spec.ts` per the round-trip AC
  (`equalBytes` per part after `unzip`, `equalStructural` over the modelled subset,
  the pinned resolved-state delta); nothing in `roundTrip.spec.ts` is weakened.
  `useSlideComments.spec.ts`, a `SlideCommentsPanel.vue` security test for the
  plain-text rule, `PresentationView.size.spec.ts` at its new budget, and
  `e2e/SPEC-101.spec.ts` per the last AC.

### Specs this changes
- `specs/038-comments.md` — its message type and reply encoding move to
  `src/lib/core/comments.ts`; contract, ACs and coverage otherwise unchanged.
- `specs/079-slide-canvas-rendering.md` — an `office:annotation` inside `draw:page` is
  no longer merely carried in `sourceXml`, closing its "no comments" gap;
  `specs/080-create-presentation-and-slide-management.md` — `duplicateSlide` drops
  comments and `stripRetention` drops `commentSource`/`commentsDirty`;
  `specs/081-slide-text-editing.md` — `serializePage` gains the annotation slot;
  `specs/088-export-presentation-to-html.md` — confirmed to export no comments.
- SPEC-099 (reuse slides from another deck) inherits this file's rule: a reused slide
  arrives with its comments dropped. `specs/000-roadmap.md` gains this spec's phase
  row; nothing else changes.
