// Find and replace over a deck (SPEC-098 — the normative source for deck
// search scope and the deck hit model). Framework-free like the rest of this
// directory: a pure walk over the model's plain text arrays producing hits,
// plus the two replace operations that rewrite exactly the paragraphs a hit
// names and dirty exactly the slides they touch. The view
// (src/views/presentation/useDeckSearch.ts) owns the bar, the debounce, the
// idle-sliced walk, navigation and the highlight overlay.
//
// Scope: per slide, every text frame's and shape frame's paragraphs, every
// non-covered table cell's paragraphs and the speaker notes. A slide also
// carries text this editor can read but must not rewrite — a group's
// children render from `children` but save from `GroupFrame.xml`, and a
// foreign frame is opaque markup — so offering hits there would offer
// replacements the next save silently discards. The walk therefore never
// reads `GroupFrame.children`, `ForeignFrame.xml`, `Slide.sourceXml`,
// `Slide.preserved`, `PresentationModel.opaque` or `.media` (and so never
// decodes base64 or parses XML), and the bar states what it left out.
//
// Layering rule (presentation/boundary.spec.ts): imports only presentation
// siblings and src/lib/core.
import type { CompiledQuery } from "../core/findQuery";
import type { PresentationModel, Slide, SlideParagraph } from "./model";
import { paragraphText } from "./slideModel";
import { markSlideEdited, replaceParagraphText } from "./text";

export type DeckHitScope = "frame" | "cell" | "notes";

/** One match. Frames are addressed by `frameId`, which survives an undo
 * snapshot (`structuredClone(model.slides)`), never by index — so a hit list
 * stays valid across a replace. */
export interface DeckHit {
  slideIndex: number;
  /** Absent only for a notes hit. */
  frameId?: string;
  /** "cell" scope only (SPEC-091). */
  row?: number;
  column?: number;
  /** Index into the frame's / cell's paragraphs or the slide's notes lines. */
  para: number;
  /** Half-open, into `paragraphText(para)` (a notes line's own text). */
  start: number;
  end: number;
  scope: DeckHitScope;
  /** The paragraph text trimmed to ±SNIPPET_CONTEXT characters around the
   * match ("…" marking a cut). */
  snippet: string;
  /** Where `start` falls inside `snippet`. */
  snippetOffset: number;
}

/** One slide's hits, in hit order, each with its index into the full list. */
export interface DeckHitGroup {
  slideIndex: number;
  hits: { hit: DeckHit; index: number }[];
}

/** Characters of context kept either side of a match in its snippet. */
export const SNIPPET_CONTEXT = 40;

/** What the deck's find bar says it searches (its `find-scope` span). */
export const DECK_SCOPE_LABEL = "Slides, shapes, tables and notes";

function snippetOf(text: string, start: number, end: number): { snippet: string; snippetOffset: number } {
  const from = Math.max(0, start - SNIPPET_CONTEXT);
  const to = Math.min(text.length, end + SNIPPET_CONTEXT);
  const lead = from > 0 ? "…" : "";
  const tail = to < text.length ? "…" : "";
  return { snippet: `${lead}${text.slice(from, to)}${tail}`, snippetOffset: start - from + lead.length };
}

/** Every non-empty match of `regExp` in `text`, in order. A zero-width match
 * is stepped past (SPEC-017's rule), so a pathological query yields nothing
 * rather than an endless loop. */
function eachMatch(text: string, regExp: RegExp, found: (start: number, end: number) => void): void {
  regExp.lastIndex = 0;
  let match = regExp.exec(text);
  while (match !== null) {
    if (match[0].length === 0) {
      regExp.lastIndex += 1;
    } else {
      found(match.index, match.index + match[0].length);
    }
    match = regExp.exec(text);
  }
}

/** Every hit on one slide: its frames in z-order (a table's cells in row
 * order), then paragraph, then offset — and its notes last. */
export function findInSlide(model: PresentationModel, slideIndex: number, compiled: CompiledQuery): DeckHit[] {
  const slide = model.slides[slideIndex];
  const hits: DeckHit[] = [];
  if (!slide) return hits;
  const scan = (texts: string[], at: Omit<DeckHit, "para" | "start" | "end" | "snippet" | "snippetOffset">): void => {
    texts.forEach((text, para) => {
      eachMatch(text, compiled.regExp, (start, end) => hits.push({ ...at, para, start, end, ...snippetOf(text, start, end) }));
    });
  };
  const texts = (paragraphs: SlideParagraph[]): string[] => paragraphs.map(paragraphText);
  for (const frame of slide.frames) {
    // Only the kinds whose text this editor writes back; groups, foreign
    // frames, media, images and charts are never opened.
    if (frame.kind === "text" || frame.kind === "shape") {
      scan(texts(frame.paragraphs ?? []), { slideIndex, frameId: frame.id, scope: "frame" });
    } else if (frame.kind === "table") {
      frame.rows.forEach((row, r) =>
        row.cells.forEach((cell, c) => {
          if (!cell.covered) scan(texts(cell.paragraphs), { slideIndex, frameId: frame.id, row: r, column: c, scope: "cell" });
        }),
      );
    }
  }
  scan(slide.notes ?? [], { slideIndex, scope: "notes" });
  return hits;
}

/** Every hit in the deck, ordered by slide index, then z-order, then
 * paragraph, then offset. Pure: no mutation, no re-identification. */
export function findInDeck(model: PresentationModel, compiled: CompiledQuery): DeckHit[] {
  const hits: DeckHit[] = [];
  for (let index = 0; index < model.slides.length; index++) {
    for (const hit of findInSlide(model, index, compiled)) hits.push(hit);
  }
  return hits;
}

/** The hits grouped by slide (one group per slide with a hit, in order). */
export function hitsBySlide(hits: readonly DeckHit[]): DeckHitGroup[] {
  const groups: DeckHitGroup[] = [];
  hits.forEach((hit, index) => {
    const last = groups[groups.length - 1];
    if (last && last.slideIndex === hit.slideIndex) last.hits.push({ hit, index });
    else groups.push({ slideIndex: hit.slideIndex, hits: [{ hit, index }] });
  });
  return groups;
}

/** The first hit at or after a position — inside `frameId` on that slide when
 * it has one, else the slide's first hit or the next slide's — wrapping to
 * the first hit when none follow; -1 with no hits. */
export function hitIndexAtOrAfter(hits: readonly DeckHit[], slideIndex: number, frameId?: string): number {
  if (hits.length === 0) return -1;
  if (frameId !== undefined) {
    const inFrame = hits.findIndex((hit) => hit.slideIndex === slideIndex && hit.frameId === frameId);
    if (inFrame >= 0) return inFrame;
  }
  const after = hits.findIndex((hit) => hit.slideIndex >= slideIndex);
  return after >= 0 ? after : 0;
}

/** The paragraph list a frame/cell hit points into, or null when the frame,
 * cell or kind is gone. */
function hitParagraphs(slide: Slide, hit: DeckHit): SlideParagraph[] | null {
  const frame = slide.frames.find((each) => each.id === hit.frameId);
  if (hit.scope === "frame" && (frame?.kind === "text" || frame?.kind === "shape")) {
    return frame.paragraphs ?? null;
  }
  if (hit.scope === "cell" && frame?.kind === "table") {
    const cell = frame.rows[hit.row ?? -1]?.cells[hit.column ?? -1];
    return cell && !cell.covered ? cell.paragraphs : null;
  }
  return null;
}

/** The text a hit's paragraph holds now, or null when it no longer exists. */
function currentText(slide: Slide, hit: DeckHit): string | null {
  if (hit.scope === "notes") return slide.notes?.[hit.para] ?? null;
  const paragraph = hitParagraphs(slide, hit)?.[hit.para];
  return paragraph ? paragraphText(paragraph) : null;
}

/** Whether the hit's range still holds the text it matched (its snippet's
 * emphasised slice) — a stale hit is never rewritten. */
export function isHitCurrent(model: PresentationModel, hit: DeckHit): boolean {
  const slide = model.slides[hit.slideIndex];
  const text = slide ? currentText(slide, hit) : null;
  const matched = hit.snippet.slice(hit.snippetOffset, hit.snippetOffset + hit.end - hit.start);
  return text !== null && hit.end <= text.length && text.slice(hit.start, hit.end) === matched;
}

/** Rewrites one current hit in place (no dirtying). */
function rewrite(slide: Slide, hit: DeckHit, replacement: string): void {
  if (hit.scope === "notes" && slide.notes) {
    const line = slide.notes[hit.para];
    slide.notes[hit.para] = line.slice(0, hit.start) + replacement + line.slice(hit.end);
    return;
  }
  const paragraph = hitParagraphs(slide, hit)?.[hit.para];
  if (paragraph) replaceParagraphText(paragraph, hit.start, hit.end, replacement);
}

/** Dirties a slide per scope: a frame/cell change rebuilds the slide; a notes
 * change rebuilds only the notes-slide part on pptx (`notesDirty`), and the
 * whole page on odp, whose writer never consults `notesDirty`. Notes emptied
 * entirely are dropped (`[]` is never stored, SPEC-083). */
function markTouched(model: PresentationModel, slide: Slide, frames: boolean, notes: boolean): void {
  if (frames) markSlideEdited(slide, model.format);
  if (!notes) return;
  if (slide.notes && slide.notes.every((line) => line.length === 0)) delete slide.notes;
  if (model.format === "odp") markSlideEdited(slide, "odp");
  else slide.notesDirty = true;
}

/** Replaces one hit (the bar's "Replace"). False, changing nothing, when the
 * hit no longer matches. The replacement takes the marks of the run
 * containing the hit's start. */
export function replaceHit(model: PresentationModel, hit: DeckHit, replacement: string): boolean {
  const slide = model.slides[hit.slideIndex];
  if (!slide || !isHitCurrent(model, hit)) return false;
  rewrite(slide, hit, replacement);
  markTouched(model, slide, hit.scope !== "notes", hit.scope === "notes");
  return true;
}

/** Replaces every current hit ("Replace all"), grouped by slide, target and
 * paragraph and applied back to front within a paragraph so earlier offsets
 * stay valid; each touched slide is dirtied once, per scope. Returns how many
 * were replaced. The caller wraps it in one undo entry. */
export function replaceAllInDeck(model: PresentationModel, hits: readonly DeckHit[], replacement: string): number {
  const bySlide = new Map<number, Map<string, DeckHit[]>>();
  for (const hit of hits) {
    const targets = bySlide.get(hit.slideIndex) ?? new Map<string, DeckHit[]>();
    bySlide.set(hit.slideIndex, targets);
    const key = `${hit.scope}|${hit.frameId ?? ""}|${hit.row ?? ""}|${hit.column ?? ""}|${hit.para}`;
    targets.set(key, [...(targets.get(key) ?? []), hit]);
  }
  let count = 0;
  for (const [slideIndex, targets] of bySlide) {
    const slide = model.slides[slideIndex];
    if (!slide) continue;
    let frames = false;
    let notes = false;
    for (const paragraphHits of targets.values()) {
      for (const hit of paragraphHits.toSorted((a, b) => b.start - a.start)) {
        if (!isHitCurrent(model, hit)) continue;
        rewrite(slide, hit, replacement);
        count += 1;
        if (hit.scope === "notes") notes = true;
        else frames = true;
      }
    }
    if (frames || notes) markTouched(model, slide, frames, notes);
  }
  return count;
}

function plural(count: number, one: string, many: string): string {
  return `${count} ${count === 1 ? one : many}`;
}

/** The find bar's scope title (and the empty state's `deck-search-scope`
 * note): what was searched, and the exclusions counted from the loaded model
 * — group children, foreign markup, master and layout text. Reads only the
 * top-level frames' `kind`. */
export function deckScopeTitle(model: PresentationModel | null): string {
  let groups = 0;
  let foreign = 0;
  for (const slide of model?.slides ?? []) {
    for (const frame of slide.frames) {
      if (frame.kind === "group") groups += 1;
      else if (frame.kind === "foreign") foreign += 1;
    }
  }
  const skipped: string[] = [];
  if (groups > 0) skipped.push(`text inside ${plural(groups, "group", "groups")}`);
  if (foreign > 0) skipped.push(plural(foreign, "embedded object", "embedded objects"));
  skipped.push("master and layout text");
  const list = skipped.length > 1 ? `${skipped.slice(0, -1).join(", ")} and ${skipped[skipped.length - 1]}` : skipped[0];
  return `Search covers slide text, shapes, tables and speaker notes. Not searched: ${list}.`;
}
