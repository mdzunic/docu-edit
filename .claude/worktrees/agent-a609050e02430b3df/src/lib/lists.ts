// Bulleted/numbered list logic over the document model (SPEC-018),
// framework-free like every lib module. Markers are never stored as text:
// the displayed marker of every item is derived here from the block
// sequence, which is what makes numbering continuous across consecutive
// items (AC-10) and restart after an interruption (AC-11) without any
// bookkeeping in the editing commands.
//
// The marker styles, ordinal formatting and level planning moved to
// core/lists.ts (product-neutral); imported for the scans below and
// re-exported so existing import paths keep working. This module keeps the
// half that walks the document model's block sequence.
import {
  isBreak,
  isChart,
  isConnector,
  isImage,
  isShape,
  isTable,
  isToc,
  isTof,
  type Block,
  type DocumentModel,
} from "./documentModel";
import { type ListDefinition, type ListFormatting } from "./core/lists";
import {
  createDefinedListCounters,
  markerForDefinedItem,
  type DefinedListCounters,
} from "./core/listDefinitions";
import { clampListLevel, markerStyleOf, markerText } from "./core/lists";

function listOf(block: Block | undefined): ListFormatting | undefined {
  return block === undefined || isTable(block) || isImage(block) || isShape(block) || isChart(block) || isConnector(block) || isBreak(block) || isToc(block) || isTof(block)
    ? undefined
    : block.list;
}

/** Every block's displayed marker for one block list, in order — "" for
 * non-list paragraphs and for table blocks (a table interrupts numbering
 * exactly like a plain paragraph would, AC-11). One scan implements the
 * numbering semantics: consecutive items at a level count up from the
 * list's start value (AC-10, AC-22), any non-list block resets every
 * counter (AC-11), and returning to a shallower level resets the deeper
 * counters, so a later sub-list numbers from its start again. Shared by the
 * section canvas and each table cell's own block list (SPEC-022 — each cell
 * is its own numbering context).
 *
 * SPEC-032 widens the scan: an item referencing a list *definition*
 * (`list.defId`) is numbered from its (defId, listId) instance's counters
 * in `defined` instead — those counters survive interruptions (that is what
 * "Continue previous numbering" relies on) and are shared across a whole
 * numbering context via the optional `defined` parameter (lib/lists's
 * listMarkers threads one across all sections). A defined item still ends
 * any adjacent SPEC-018 run: it is a different list that happens to touch. */
export function markersForBlocks(
  blocks: readonly Block[],
  definitions: readonly ListDefinition[] = [],
  defined: DefinedListCounters = createDefinedListCounters(),
): string[] {
  // counters[level] is the last ordinal used at that level within the
  // current uninterrupted run of (ad-hoc) list paragraphs.
  let counters: Array<number | undefined> = [];
  return blocks.map((block) => {
    const list = listOf(block);
    if (!list) {
      counters = [];
      return "";
    }
    if (list.defId !== undefined) {
      counters = [];
      const def = definitions.find((entry) => entry.defId === list.defId);
      if (def) {
        return markerForDefinedItem(def, list, defined);
      }
      // A dangling defId (definition removed): fall through to the ad-hoc
      // rendering below rather than dropping the marker entirely.
    }
    const level = clampListLevel(list.level);
    for (let deeper = level + 1; deeper < counters.length; deeper++) {
      counters[deeper] = undefined;
    }
    const style = markerStyleOf(list);
    if (list.type === "bullet") {
      return markerText(style, 0);
    }
    const ordinal = (counters[level] ?? (list.startAt ?? 1) - 1) + 1;
    counters[level] = ordinal;
    return markerText(style, ordinal);
  });
}

/** Every top-level block's displayed marker, per section, in block order.
 * One defined-instance counter set spans all sections (SPEC-032): a defined
 * list interrupted by a section break continues its numbering. */
export function listMarkers(model: DocumentModel): string[][] {
  const defined = createDefinedListCounters();
  const definitions = model.listDefinitions ?? [];
  return model.sections.map((section) =>
    markersForBlocks(section.blocks, definitions, defined),
  );
}

/** The contiguous run of list paragraphs around `block` in its containing
 * block list — what AC-13 calls "the current list". Takes the container
 * (a section's blocks, or a table cell's, SPEC-022) rather than the model
 * so the run never leaks across a cell boundary. Returns the inclusive
 * [first, last] range, or null when the block is not a list item. */
export function listRunAround(
  blocks: readonly Block[] | undefined,
  block: number,
): { first: number; last: number } | null {
  if (!blocks || !listOf(blocks[block])) {
    return null;
  }
  let first = block;
  while (first > 0 && listOf(blocks[first - 1])) {
    first--;
  }
  let last = block;
  while (last + 1 < blocks.length && listOf(blocks[last + 1])) {
    last++;
  }
  return { first, last };
}
