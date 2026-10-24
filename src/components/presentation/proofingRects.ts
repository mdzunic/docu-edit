// The slide stage's proofing rects (SPEC-098): typed pass-throughs for
// SlideProofingLayer.vue's untyped Array props (espree bars the `PropType`
// annotation in .vue scripts) and the DOM measurement both proofing
// composables share — a model text range (frame or table cell, paragraph,
// character offsets) to rects in the stage surface's own unscaled px. The
// layer paints INSIDE the scaled surface, so the surface's CSS transform
// re-applies the zoom and every rect stays registered with its text. Kept
// beside the renderer (a rendering concern); the text side's
// SpellUnderlineRect carries `section`/`block`, coordinates a slide has no
// equivalent of, so it is not shared.

/** A box in stage-surface px (unscaled). */
export interface ProofRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** A find hit's highlight; `current` marks the current hit (SPEC-017's look). */
export interface HighlightRect extends ProofRect {
  current: boolean;
}

/** A misspelled word's glyph box — the squiggle paints at its bottom edge and
 * a right-click anywhere in it opens the suggestions; `issue` indexes the
 * spell checker's issue list. */
export interface SpellUnderlineRect extends ProofRect {
  issue: number;
  word: string;
}

/** Where a text range sits on the stage: a top-level frame (and, for a
 * table, its cell), a paragraph, and a half-open character range. */
export interface StageTextRange {
  frameId?: string;
  row?: number;
  column?: number;
  para: number;
  start: number;
  end: number;
}

export function asHighlightRects(value: unknown): HighlightRect[] {
  return value as HighlightRect[];
}

export function asSpellUnderlineRects(value: unknown): SpellUnderlineRect[] {
  return value as SpellUnderlineRect[];
}

/** The rendered paragraph element a range addresses: the live editing host's
 * block when a session is open there, else the stage's `.paragraph`. */
function paragraphElement(surface: HTMLElement, range: StageTextRange): Element | null {
  const frame = Array.from(surface.querySelectorAll<HTMLElement>("[data-frame-id]")).find((el) => el.dataset.frameId === range.frameId);
  if (!frame) return null;
  const container =
    range.row !== undefined && range.column !== undefined
      ? frame.querySelector(`[data-testid="slide-table-cell-${range.row}-${range.column}"]`)
      : frame;
  if (!container) return null;
  const host = container.querySelector(".ce-host");
  if (host) return host.children[range.para] ?? null;
  return container.querySelectorAll(".paragraph")[range.para] ?? null;
}

/** The (text node, offset) pairs for a paragraph's character offsets — list
 * markers (`.bullet-marker`) are not paragraph text and are skipped. */
function domPoint(paragraph: Element, offset: number): { node: Node; offset: number } | null {
  const doc = paragraph.ownerDocument;
  const walker = doc.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT, {
    acceptNode: (node) => (node.parentElement?.closest(".bullet-marker") ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  let remaining = offset;
  let last: Node | null = null;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const length = node.textContent?.length ?? 0;
    if (remaining <= length) return { node, offset: remaining };
    remaining -= length;
    last = node;
  }
  return last ? { node: last, offset: last.textContent?.length ?? 0 } : null;
}

/** A model text range's on-screen rects, in surface px. Empty when the frame
 * or paragraph is not rendered (another slide, a covered cell, jsdom). */
export function measureStageRange(surface: HTMLElement, range: StageTextRange): ProofRect[] {
  const paragraph = paragraphElement(surface, range);
  const from = paragraph ? domPoint(paragraph, range.start) : null;
  const to = paragraph ? domPoint(paragraph, range.end) : null;
  if (!from || !to) return [];
  const box = surface.getBoundingClientRect();
  // The surface is CSS-scaled; its layout width is the unscaled size.
  const scale = surface.offsetWidth > 0 ? box.width / surface.offsetWidth : 1;
  try {
    const dom = surface.ownerDocument.createRange();
    dom.setStart(from.node, from.offset);
    dom.setEnd(to.node, to.offset);
    return Array.from(dom.getClientRects())
      .filter((rect) => rect.width > 0 && rect.height > 0)
      .map((rect) => ({
        left: (rect.left - box.left) / scale,
        top: (rect.top - box.top) / scale,
        width: rect.width / scale,
        height: rect.height / scale,
      }));
  } catch {
    // A Range without getClientRects (some embedding environments): nothing
    // paints this pass rather than a crash.
    return [];
  }
}

/** A client point in surface px (the inverse of the measurement above). */
export function surfacePoint(surface: HTMLElement, clientX: number, clientY: number): { x: number; y: number } {
  const box = surface.getBoundingClientRect();
  const scale = surface.offsetWidth > 0 ? box.width / surface.offsetWidth : 1;
  return { x: (clientX - box.left) / scale, y: (clientY - box.top) / scale };
}
