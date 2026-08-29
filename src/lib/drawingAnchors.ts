// Explicit anchors for floating drawings (SPEC-078), framework-free like
// every lib module. A floating Shape/Chart and its anchor paragraph pair by
// sharing an `anchorOrdinal` (unique per paragraph document-wide); a drawing
// without one — every pre-SPEC-078 document — resolves by the adjacency
// fallback below. This module is the ONLY place an anchor is derived: the
// anchor marker, the re-anchor drag, the renderer's hoisting and both
// writers all call resolveAnchorBlockIndex rather than re-implementing
// adjacency, so they cannot disagree about which paragraph a drawing
// belongs to (the `[para, d1, d2]` ambiguity this spec exists to kill).
import {
  isChart,
  isConnector,
  isImage,
  isShape,
  type Block,
  type Chart,
  type Connector,
  type DocumentModel,
  type Image,
  type Paragraph,
  type Shape,
} from "./documentModel";

/** A drawing that sits outside the text flow — the blocks that carry an
 * anchor pairing, and the ones the renderer hoists to their anchor
 * paragraph's own origin.
 *
 * Image joined this list when it gained real placement: before that it had no
 * anchor-relative offsets to measure (SPEC-078 called it out of scope), so
 * pairing and hoisting would have moved nothing. Now a floating picture's
 * `offsetXCm`/`offsetYCm` measure from exactly the same origin a shape's do,
 * which only holds if it hoists the same way. */
export function isFloatingDrawing(block: Block | undefined): block is Shape | Chart | Connector | Image {
  return (isShape(block) || isChart(block) || isConnector(block) || isImage(block)) && block.wrap !== "inline";
}

/** Any drawing block — what the adjacency fallback walks over: in
 * `[para, image, shape]` the shape's nearest paragraph is `para`, not
 * nothing. */
function isDrawingBlock(block: Block | undefined): boolean {
  return isShape(block) || isChart(block) || isImage(block);
}

/** The block index of the paragraph a drawing at `drawingIndex` is anchored
 * to, or null when there is none to resolve.
 *
 * By ordinal first: the FIRST paragraph in the section carrying the
 * drawing's own `anchorOrdinal`. First-in-section is deliberate — a
 * paragraph split briefly leaves the value reachable from both halves in a
 * pure splitter's caller window, and first-match picks the top half, the
 * piece that keeps the paragraph's top edge offsets are measured from.
 *
 * Else by adjacency (a drawing with no ordinal, or a dangling reference):
 * walk backward over consecutive drawing blocks to the nearest paragraph,
 * failing that walk forward the same way, failing both null. This subsumes
 * SPEC-051's one-step inference — the walk skips a neighbouring drawing
 * instead of giving up on it, so `[para, d1, d2]` resolves d2 to `para`
 * even for legacy documents. */
export function resolveAnchorBlockIndex(blocks: readonly Block[], drawingIndex: number): number | null {
  const drawing = blocks[drawingIndex];
  if (!isShape(drawing) && !isChart(drawing)) {
    return null;
  }
  if (drawing.anchorOrdinal !== undefined) {
    const byOrdinal = paragraphIndexForOrdinal(blocks, drawing.anchorOrdinal);
    if (byOrdinal !== null) {
      return byOrdinal;
    }
  }
  for (let i = drawingIndex - 1; i >= 0; i--) {
    if (blocks[i].type === "paragraph") {
      return i;
    }
    if (!isDrawingBlock(blocks[i])) {
      break;
    }
  }
  for (let i = drawingIndex + 1; i < blocks.length; i++) {
    if (blocks[i].type === "paragraph") {
      return i;
    }
    if (!isDrawingBlock(blocks[i])) {
      break;
    }
  }
  return null;
}

/** The block index of the FIRST paragraph in `blocks` carrying `ordinal`,
 * or null — the explicit half of resolveAnchorBlockIndex, exported on its
 * own for the writers, which absorb a drawing into its anchor paragraph
 * only when the pairing is explicit (a dangling or absent ordinal keeps the
 * legacy own-paragraph wire form byte-for-byte — SPEC-078's compatibility
 * criterion). */
export function paragraphIndexForOrdinal(blocks: readonly Block[], ordinal: number): number | null {
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type === "paragraph" && block.anchorOrdinal === ordinal) {
      return i;
    }
  }
  return null;
}

/** A fresh anchorOrdinal not used by any paragraph OR drawing in the
 * document — drawings included so a dangling reference can never collide
 * with a new pairing. Mirrors shapes.ts's freshGroupOrdinal. */
export function mintAnchorOrdinal(model: DocumentModel): number {
  let max = 0;
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "paragraph" && block.anchorOrdinal !== undefined) {
        max = Math.max(max, block.anchorOrdinal);
      } else if ((isShape(block) || isChart(block)) && block.anchorOrdinal !== undefined) {
        max = Math.max(max, block.anchorOrdinal);
      }
    }
  }
  return max + 1;
}

/** Whether any Shape/Chart in the document references `ordinal`, other than
 * the one at the excluded address (the drawing being deleted, re-anchored
 * or made inline). */
export function anchorOrdinalReferenced(
  model: DocumentModel,
  ordinal: number,
  except?: { section: number; block: number },
): boolean {
  for (let s = 0; s < model.sections.length; s++) {
    const blocks = model.sections[s].blocks;
    for (let b = 0; b < blocks.length; b++) {
      if (except && except.section === s && except.block === b) {
        continue;
      }
      const block = blocks[b];
      if ((isShape(block) || isChart(block)) && block.anchorOrdinal === ordinal) {
        return true;
      }
    }
  }
  return false;
}

/** One section's containment plan for a writer (SPEC-078): which floating
 * drawings are absorbed into their anchor paragraph instead of being
 * emitted at their own block slot, and on which side of the paragraph's
 * runs each lands — before them when the drawing precedes its anchor in the
 * block list, after them otherwise, in block order. Absorption requires an
 * EXPLICIT pairing that resolves (`paragraphIndexForOrdinal`); a drawing
 * with no ordinal, or a dangling one, keeps the legacy own-paragraph wire
 * form byte-for-byte. Shared by both writers so `.docx` and `.odt` cannot
 * disagree about what absorbs. */
export interface AnchoredDrawingPlan {
  /** Drawing block indices emitted inside a paragraph, not at their slot. */
  absorbed: Set<number>;
  /** Anchor paragraph block index -> its absorbed drawings' indices. */
  byParagraph: Map<number, { leading: number[]; trailing: number[] }>;
}

export function planAnchoredDrawings(blocks: readonly Block[]): AnchoredDrawingPlan {
  const absorbed = new Set<number>();
  const byParagraph = new Map<number, { leading: number[]; trailing: number[] }>();
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (!isFloatingDrawing(block) || block.anchorOrdinal === undefined) {
      continue;
    }
    const anchorIndex = paragraphIndexForOrdinal(blocks, block.anchorOrdinal);
    if (anchorIndex === null) {
      continue;
    }
    absorbed.add(i);
    let entry = byParagraph.get(anchorIndex);
    if (!entry) {
      entry = { leading: [], trailing: [] };
      byParagraph.set(anchorIndex, entry);
    }
    (i < anchorIndex ? entry.leading : entry.trailing).push(i);
  }
  return { absorbed, byParagraph };
}

/** A copy of `paragraph` carrying `ordinal` — the pure builder commands
 * splice in when pairing (block objects are shared across cloneContainers'd
 * models, so the original is never mutated). */
export function anchoredParagraph(paragraph: Paragraph, ordinal: number): Paragraph {
  return { ...paragraph, anchorOrdinal: ordinal };
}

/** The ordinal to pair a drawing on `paragraph` with: the paragraph's own
 * when it already anchors something, else a fresh mint — so two drawings
 * inserted onto one paragraph share a single pairing. */
export function ordinalForParagraph(model: DocumentModel, paragraph: Paragraph): number {
  return paragraph.anchorOrdinal ?? mintAnchorOrdinal(model);
}

/** Strips `ordinal` from whichever paragraph carries it when no drawing
 * references it any more (a paragraph carries `anchorOrdinal` only while it
 * anchors something — SPEC-078's release rule). `model` must already be
 * cloneContainers'd: the paragraph is replaced by splice, never mutated.
 * `except` names a drawing to disregard while counting referrers (the one
 * being deleted or re-pointed, which still carries the old value). */
export function releaseAnchorOrdinal(
  model: DocumentModel,
  ordinal: number | undefined,
  except?: { section: number; block: number },
): void {
  if (ordinal === undefined || anchorOrdinalReferenced(model, ordinal, except)) {
    return;
  }
  for (const section of model.sections) {
    const blocks = section.blocks;
    for (let b = 0; b < blocks.length; b++) {
      const block = blocks[b];
      if (block.type === "paragraph" && block.anchorOrdinal === ordinal) {
        const released: Paragraph = { ...block };
        delete released.anchorOrdinal;
        blocks.splice(b, 1, released);
        return;
      }
    }
  }
}
