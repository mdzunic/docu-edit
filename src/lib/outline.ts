// Document outline navigator (SPEC-037): builds the hierarchical outline of
// the document's Heading 1-6 paragraphs, resolves which heading the caret sits
// in, computes which body blocks a collapsed heading hides, and the
// drag-reorder command that moves a heading (with all its content) among its
// siblings. Framework-free like every lib module, and — like lib/toc.ts and
// lib/links.ts — a command module OUTSIDE commands.ts built on
// documentModel.ts's shared cloneContainers, so commands.undo.spec.ts's
// completeness guard (scoped to commands.ts's own exports) never has to
// account for moveHeading.
import { cloneContainers, type Block, type DocumentModel } from "./documentModel";
import { listHeadings, type HeadingTarget } from "./links";
import { compareCaret, type CaretPosition } from "./textEditing";
import { caretSelection, type CommandResult } from "./commands";

/** The empty-state message shown in the pane when the document has no
 * Heading 1-6 paragraphs (SPEC-037 AC-12). */
export const OUTLINE_EMPTY_TEXT = "Headings you add will appear here";

/** One heading in the outline tree: its derived anchor, visible text, 1-based
 * level, model address and nested child headings (deeper levels that follow it
 * before the next heading of the same-or-higher level). */
export interface OutlineNode {
  anchor: string;
  text: string;
  level: number;
  section: number;
  block: number;
  children: OutlineNode[];
}

/** Typed pass-through for OutlinePane's `outline` prop, which arrives as a
 * runtime-declared Array (espree bars the `Array as () => OutlineNode[]` cast
 * in `.vue` scripts — see lib/pageRender.asRenderEntries for the pattern). */
export function asOutlineNodes(value: unknown): OutlineNode[] {
  return value as OutlineNode[];
}

/** Typed pass-through for OutlinePane's `collapsed` prop (a runtime-declared
 * Object that is really a Set of anchor strings). */
export function asKeySet(value: unknown): ReadonlySet<string> {
  return value as ReadonlySet<string>;
}

/** A flattened, indented row the pane renders (SPEC-037 AC-1, AC-2): the node,
 * its depth in the tree (indent), and whether it has children (chevron). */
export interface OutlineRow {
  node: OutlineNode;
  depth: number;
  hasChildren: boolean;
  collapsed: boolean;
}

/** Only body headings drive the outline's structural operations: a heading
 * inside a table cell is not a top-level `section.blocks` entry, so it is
 * shown but never a drag source/target (moveHeading rejects it). */
function isTopLevelHeading(heading: HeadingTarget): boolean {
  return heading.address.cell === undefined;
}

/** The document's headings as a tree, nested by outline level (SPEC-037 AC-1).
 * Levels may skip (an H1 followed by an H3): each heading nests under the
 * nearest preceding heading of a strictly shallower level, regardless of the
 * gap, which matches how Word/Docs render a malformed outline. */
export function buildOutline(model: DocumentModel): OutlineNode[] {
  const roots: OutlineNode[] = [];
  const stack: OutlineNode[] = [];
  for (const heading of listHeadings(model)) {
    const node: OutlineNode = {
      anchor: heading.anchor,
      text: heading.text,
      level: heading.level,
      section: heading.address.section,
      block: heading.address.block,
      children: [],
    };
    while (stack.length > 0 && stack[stack.length - 1].level >= node.level) {
      stack.pop();
    }
    if (stack.length === 0) {
      roots.push(node);
    } else {
      stack[stack.length - 1].children.push(node);
    }
    stack.push(node);
  }
  return roots;
}

/** The visible rows of the tree given the set of collapsed anchors (SPEC-037
 * AC-7, AC-8): a collapsed node's descendants are omitted. Depth is the tree
 * depth (0 for a root), not the heading level, so indentation reflects nesting
 * even when levels skip. */
export function visibleOutlineRows(
  nodes: OutlineNode[],
  collapsed: ReadonlySet<string>,
  depth = 0,
): OutlineRow[] {
  const rows: OutlineRow[] = [];
  for (const node of nodes) {
    const hasChildren = node.children.length > 0;
    const isCollapsed = collapsed.has(node.anchor);
    rows.push({ node, depth, hasChildren, collapsed: hasChildren && isCollapsed });
    if (hasChildren && !isCollapsed) {
      rows.push(...visibleOutlineRows(node.children, collapsed, depth + 1));
    }
  }
  return rows;
}

/** The anchor of the heading that contains `caret` (SPEC-037 AC-3): the last
 * heading at or before the caret in document order. Null when the caret is
 * before the first heading (or there are none). */
export function activeHeadingAnchor(
  model: DocumentModel,
  caret: CaretPosition | null,
): string | null {
  if (!caret) {
    return null;
  }
  let active: string | null = null;
  for (const heading of listHeadings(model)) {
    const start: CaretPosition = {
      section: heading.address.section,
      block: heading.address.block,
      offset: 0,
      cell: heading.address.cell,
    };
    if (compareCaret(start, caret) <= 0) {
      active = heading.anchor;
    }
  }
  return active;
}

/** The half-open block range [start, end) a top-level heading owns within its
 * section: the heading itself plus every following block up to (not including)
 * the next paragraph that is a heading of the same-or-higher level. Tables,
 * images, breaks and deeper headings all count as content. */
function headingBlockRange(
  blocks: readonly Block[],
  startIndex: number,
  level: number,
): [number, number] {
  let end = startIndex + 1;
  while (end < blocks.length) {
    const block = blocks[end];
    if (block.type === "paragraph" && block.headingLevel !== undefined && block.headingLevel <= level) {
      break;
    }
    end += 1;
  }
  return [startIndex, end];
}

/** The `"section:block"` key of a top-level block address. */
export function blockKey(section: number, block: number): string {
  return `${section}:${block}`;
}

/** Typed empty `Set<string>` seed for the pane-collapsed / body-collapsed
 * anchor refs in `.vue` scripts, which espree bars from writing the
 * `new Set<string>()` generic themselves (see textEditing.noCaret). */
export function emptyKeySet(): Set<string> {
  return new Set();
}

/** How many content blocks follow a heading before the next heading of the
 * same-or-higher level — 0 when the heading owns nothing, so the body renderer
 * can hide the collapse chevron on a heading there is nothing to collapse. */
export function headingContentLength(
  model: DocumentModel,
  section: number,
  block: number,
): number {
  const blocks = model.sections[section]?.blocks;
  const heading = blocks?.[block];
  if (!blocks || !heading || heading.type !== "paragraph" || heading.headingLevel === undefined) {
    return 0;
  }
  const [start, end] = headingBlockRange(blocks, block, heading.headingLevel);
  return end - (start + 1);
}

/** The set of hidden block addresses ("section:block") when the given headings
 * are collapsed in the document body (SPEC-037 AC-16). `collapsed` holds the
 * `"section:block"` address of each collapsed heading; each contributes its
 * content blocks (its range minus the heading paragraph itself). A heading
 * nested inside a collapsed ancestor is itself hidden, so its own collapsed
 * state is irrelevant. This is view-only state — it is never written to the
 * model, so it never reaches the file (AC-17). Addresses that no longer point
 * at a heading (the document changed under a stale key) are ignored. */
export function hiddenBlockAddresses(
  model: DocumentModel,
  collapsed: ReadonlySet<string>,
): Set<string> {
  const hidden = new Set<string>();
  if (collapsed.size === 0) {
    return hidden;
  }
  for (const key of collapsed) {
    const [section, block] = key.split(":").map(Number);
    const blocks = model.sections[section]?.blocks;
    const heading = blocks?.[block];
    if (!blocks || !heading || heading.type !== "paragraph" || heading.headingLevel === undefined) {
      continue;
    }
    const [start, end] = headingBlockRange(blocks, block, heading.headingLevel);
    for (let index = start + 1; index < end; index += 1) {
      hidden.add(blockKey(section, index));
    }
  }
  return hidden;
}

/** Where a dragged heading is dropped relative to the target sibling. */
export type DropPosition = "before" | "after";

/** Typed `null` seeds for OutlinePane's anchor/drop refs, which espree bars
 * from writing the `string | null` / `DropPosition | null` annotations
 * themselves (see textEditing.noCaret for the same pattern). */
export function noAnchor(): string | null {
  return null;
}

export function noDropPosition(): DropPosition | null {
  return null;
}

// The nearest preceding heading of a strictly shallower level defines a
// heading's parent; two headings are siblings when they share a parent and
// level. Keyed by anchor over the whole document order.
function parentAnchors(headings: HeadingTarget[]): Map<string, string | null> {
  const parents = new Map<string, string | null>();
  const stack: HeadingTarget[] = [];
  for (const heading of headings) {
    while (stack.length > 0 && stack[stack.length - 1].level >= heading.level) {
      stack.pop();
    }
    parents.set(heading.anchor, stack.length > 0 ? stack[stack.length - 1].anchor : null);
    stack.push(heading);
  }
  return parents;
}

/** Whether dragging `sourceAnchor` onto `targetAnchor` at `position` is a
 * legal move (SPEC-037 AC-10, AC-11): both must be top-level body headings in
 * the same section, at the same level, sharing the same parent (siblings), and
 * the result must actually change the order. Dropping onto a heading of a
 * different level, onto a non-sibling, or onto the heading itself is rejected.
 * Shares its resolution with moveHeading so the pane's drag feedback and the
 * committed move never disagree. */
export function canMoveHeading(
  model: DocumentModel,
  sourceAnchor: string,
  targetAnchor: string,
  position: DropPosition,
): boolean {
  return resolveMove(model, sourceAnchor, targetAnchor, position) !== null;
}

interface ResolvedMove {
  section: number;
  sourceStart: number;
  sourceEnd: number;
  insertAt: number;
}

function resolveMove(
  model: DocumentModel,
  sourceAnchor: string,
  targetAnchor: string,
  position: DropPosition,
): ResolvedMove | null {
  if (sourceAnchor === targetAnchor) {
    return null;
  }
  const headings = listHeadings(model);
  const source = headings.find((heading) => heading.anchor === sourceAnchor);
  const target = headings.find((heading) => heading.anchor === targetAnchor);
  if (!source || !target || !isTopLevelHeading(source) || !isTopLevelHeading(target)) {
    return null;
  }
  if (source.address.section !== target.address.section) {
    return null;
  }
  if (source.level !== target.level) {
    return null;
  }
  const parents = parentAnchors(headings);
  if (parents.get(source.anchor) !== parents.get(target.anchor)) {
    return null;
  }
  const section = source.address.section;
  const blocks = model.sections[section].blocks;
  const [sourceStart, sourceEnd] = headingBlockRange(blocks, source.address.block, source.level);
  const [targetStart, targetEnd] = headingBlockRange(blocks, target.address.block, target.level);
  const insertAt = position === "before" ? targetStart : targetEnd;
  // A drop that lands inside or immediately adjacent to the source's own range
  // is a no-op — reject it so no empty undo step is recorded.
  if (insertAt >= sourceStart && insertAt <= sourceEnd) {
    return null;
  }
  return { section, sourceStart, sourceEnd, insertAt };
}

/** Moves a heading and all of its content — sub-headings included — to a new
 * position among its siblings, as one command (SPEC-037 AC-9, AC-10). Returns
 * null when the move is rejected (see canMoveHeading), so the caller records no
 * undo step. The caret ends at the moved heading's new start. Section-level
 * `opaqueBlocks` (preserved unmodelled XML, SPEC-008) are re-anchored to the
 * blocks they precede so round-trip fidelity survives the reorder. */
export function moveHeading(
  model: DocumentModel,
  sourceAnchor: string,
  targetAnchor: string,
  position: DropPosition,
): CommandResult | null {
  const resolved = resolveMove(model, sourceAnchor, targetAnchor, position);
  if (!resolved) {
    return null;
  }
  const { section, sourceStart, sourceEnd, insertAt } = resolved;
  const next = cloneContainers(model);
  const blocks = next.sections[section].blocks;
  // Re-anchor opaqueBlocks by the block they sit before, tracked by reference
  // so the reorder cannot misalign them; a trailing entry maps to the end.
  const opaque = next.sections[section].opaqueBlocks;
  const anchored = opaque?.map((entry) => ({
    entry,
    before: entry.blockIndex < blocks.length ? blocks[entry.blockIndex] : null,
  }));

  const moved = blocks.splice(sourceStart, sourceEnd - sourceStart);
  const destination = insertAt > sourceEnd ? insertAt - moved.length : insertAt;
  blocks.splice(destination, 0, ...moved);

  if (anchored) {
    next.sections[section] = {
      ...next.sections[section],
      opaqueBlocks: anchored.map(({ entry, before }) => ({
        ...entry,
        blockIndex: before === null ? blocks.length : blocks.indexOf(before),
      })),
    };
  }

  const caret: CaretPosition = { section, block: destination, offset: 0 };
  return { model: next, selection: caretSelection(caret) };
}
