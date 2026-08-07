// The pagination engine (SPEC-028): distributes the model's block flow onto
// pages, honouring manual breaks, section starts, per-section geometry and
// the paragraph pagination rules (keep with next, keep lines together, page
// break before, widow/orphan control). Framework-free and DOM-free: the
// caller supplies measured per-block metrics through a provider, so the
// engine is pure math over line boxes — unit-testable with fake metrics,
// and cheap enough (AC-34/35) to re-run in full on every edit while only
// *measurement* is incremental (cached per block object identity by the
// caller; untouched blocks keep their objects, see commands.ts).
import {
  isBreak,
  isImage,
  isTable,
  resolveColumnWidthsCm,
  sectionColumns,
  sectionContentWidthCm,
  sectionMargins,
  sectionPageSize,
  type Block,
  type DocumentModel,
  type Margins,
  type PageSize,
  type Paragraph,
} from "./documentModel";

/** CSS reference pixels per centimetre — the same 96 dpi the canvas's
 * cm-sized page CSS resolves to, so engine px and rendered px agree. */
export const PX_PER_CM = 96 / 2.54;

/** One rendered line of a paragraph, relative to the block's border-box
 * top, in px. */
export interface LineBox {
  top: number;
  height: number;
}

/** A block's measured (or estimated) geometry at a given content width. */
export interface BlockMetrics {
  /** The content width the block was measured at, px — stale when the
   * section's content width changes. */
  contentWidthPx: number;
  /** Border-box height, px; margins excluded. */
  heightPx: number;
  marginTopPx: number;
  marginBottomPx: number;
  /** Line boxes for a paragraph, in document order; empty for atoms
   * (tables, images, breaks) and for paragraphs that cannot be split. */
  lines: LineBox[];
  /** True when these metrics are an estimate rather than a DOM measurement
   * — the caller re-measures such blocks and re-runs the layout. */
  estimated?: boolean;
}

/** Supplies per-block metrics. The engine never caches — caching (and
 * invalidation by block object identity) is the provider's job. */
export interface MetricsProvider {
  metricsFor(
    sectionIndex: number,
    blockIndex: number,
    block: Block,
    contentWidthPx: number,
  ): BlockMetrics;
}

/** A block's share of one page. A paragraph split across pages appears as
 * several items with consecutive `[startLine, endLine)` windows; a whole
 * block has no line window. */
export interface PageItem {
  sectionIndex: number;
  blockIndex: number;
  /** First line rendered on this page (inclusive); absent = whole block. */
  startLine?: number;
  /** One past the last line rendered on this page; absent = to the end. */
  endLine?: number;
  /** Which column of a multi-column page this item flows into (SPEC-030);
   * absent for single-column pages, so single-column layouts are byte-for-
   * byte identical to pre-SPEC-030 output. */
  columnIndex?: number;
}

/** One resolved column of a multi-column page: its width and left offset (cm)
 * within the content area, so the renderer positions it with no model
 * lookups. */
export interface PageColumn {
  widthCm: number;
  offsetCm: number;
}

/** A page's multi-column geometry (SPEC-030); present only when the page has
 * more than one column. */
export interface PageColumns {
  count: number;
  gapCm: number;
  separator: boolean;
  columns: PageColumn[];
}

/** One laid-out page: the geometry it renders at (resolved from its opening
 * section, so the renderer needs no model lookups) plus its items. */
export interface PageLayout {
  sectionIndex: number;
  pageSize: PageSize;
  margins: Margins;
  items: PageItem[];
  /** Multi-column geometry (SPEC-030); absent = a single full-width column. */
  columns?: PageColumns;
}

/** A section's page/content geometry in px. */
export interface SectionGeometryPx {
  contentWidthPx: number;
  contentHeightPx: number;
}

export function sectionGeometryPx(
  model: DocumentModel,
  sectionIndex: number,
): SectionGeometryPx {
  const pageSize = sectionPageSize(model, sectionIndex);
  const margins = sectionMargins(model, sectionIndex);
  return {
    contentWidthPx: Math.max(1, (pageSize.widthCm - margins.leftCm - margins.rightCm) * PX_PER_CM),
    contentHeightPx: Math.max(1, (pageSize.heightCm - margins.topCm - margins.bottomCm) * PX_PER_CM),
  };
}

/** A section's resolved column geometry in px (SPEC-030): the per-column
 * widths/offsets the engine flows into and the renderer paints, plus the
 * single width every block in the section is measured at. */
export interface SectionColumnGeometryPx {
  count: number;
  gapCm: number;
  separator: boolean;
  contentHeightPx: number;
  /** The width blocks are measured at — the smallest column, so a block
   * measured once fits (never overflows) whichever column it flows into
   * (AC-15). For equal columns every column is this width, so the layout is
   * exact. */
  flowWidthPx: number;
  widthsCm: number[];
  offsetsCm: number[];
  columnsPx: { widthPx: number; offsetPx: number }[];
}

export function sectionColumnGeometryPx(
  model: DocumentModel,
  sectionIndex: number,
): SectionColumnGeometryPx {
  const columns = sectionColumns(model, sectionIndex);
  const contentWidthCm = sectionContentWidthCm(model, sectionIndex);
  const geometry = sectionGeometryPx(model, sectionIndex);
  const widthsCm = resolveColumnWidthsCm(columns, contentWidthCm);
  const gapCm = Math.max(0, columns.gapCm);
  const offsetsCm: number[] = [];
  const columnsPx: { widthPx: number; offsetPx: number }[] = [];
  let offsetCm = 0;
  for (const widthCm of widthsCm) {
    offsetsCm.push(offsetCm);
    columnsPx.push({ widthPx: Math.max(1, widthCm * PX_PER_CM), offsetPx: offsetCm * PX_PER_CM });
    offsetCm += widthCm + gapCm;
  }
  return {
    count: widthsCm.length,
    gapCm,
    separator: columns.separator === true,
    contentHeightPx: geometry.contentHeightPx,
    flowWidthPx: Math.max(1, Math.min(...widthsCm) * PX_PER_CM),
    widthsCm,
    offsetsCm,
    columnsPx,
  };
}

/** The width (px) every block in a section is measured/laid out at — the
 * section's (smallest) column width. The measurement controller and the
 * renderer's slice logic both read this so they agree with the engine. */
export function sectionColumnFlowWidthPx(model: DocumentModel, sectionIndex: number): number {
  return sectionColumnGeometryPx(model, sectionIndex).flowWidthPx;
}

// Sub-pixel measurement noise must never flip a fit decision into an
// off-by-one-line oscillation between layout passes.
const EPS = 0.5;

// Backtracking a keep-with-next chain stops rather than move more blocks
// than this — a runaway chain (every paragraph keepNext) must not empty
// pages or loop (AC-23's "honoured" never requires moving whole pages).
const MAX_KEEP_NEXT_CHAIN = 8;

// One column of one page — the unit content flows into (SPEC-030). A
// single-column page has exactly one frame, so the fill logic below is the
// same math it always was; multi-column pages have `columnCount` frames that
// fill left-to-right, and only the last one filling opens a new page.
interface Frame {
  page: PageLayout;
  /** The section that opened this frame's page (for keep-with-next scoping). */
  sectionIndex: number;
  columnIndex: number;
  columnCount: number;
  /** Width every block on this page is measured at (the section's smallest
   * column) — shared by all the page's frames so measurement stays cached
   * per block regardless of which column a block lands in. */
  flowWidthPx: number;
  /** Content height available in the column. */
  heightPx: number;
  /** Bottom edge (px, column-relative) of the last placed item. */
  cursor: number;
  /** The last item's collapsible bottom margin. */
  pendingMargin: number;
  /** Items placed in THIS column (page.items is shared across columns). */
  count: number;
}

function paragraphOf(block: Block): Paragraph | null {
  return isTable(block) || isImage(block) || isBreak(block) ? null : block;
}

function sliceHeight(metrics: BlockMetrics, startLine: number, endLine: number): number {
  const lines = metrics.lines;
  if (lines.length === 0) {
    return metrics.heightPx;
  }
  const top = lines[startLine]?.top ?? 0;
  if (endLine >= lines.length) {
    return metrics.heightPx - top;
  }
  const last = lines[endLine - 1];
  return last.top + last.height - top;
}

// Builds a page item, tagging its column only on multi-column pages so
// single-column layouts stay identical to pre-SPEC-030 output.
function frameItem(
  frame: Frame,
  sectionIndex: number,
  blockIndex: number,
  startLine?: number,
  endLine?: number,
): PageItem {
  const item: PageItem = { sectionIndex, blockIndex };
  if (startLine !== undefined) {
    item.startLine = startLine;
    item.endLine = endLine;
  }
  if (frame.columnCount > 1) {
    item.columnIndex = frame.columnIndex;
  }
  return item;
}

function placeWhole(frame: Frame, sectionIndex: number, blockIndex: number, m: BlockMetrics): void {
  const gap = frame.count === 0 ? m.marginTopPx : Math.max(frame.pendingMargin, m.marginTopPx);
  frame.page.items.push(frameItem(frame, sectionIndex, blockIndex));
  frame.cursor += gap + m.heightPx;
  frame.pendingMargin = m.marginBottomPx;
  frame.count += 1;
}

function placeSlice(
  frame: Frame,
  sectionIndex: number,
  blockIndex: number,
  m: BlockMetrics,
  startLine: number,
  endLine: number,
): void {
  // A continuation slice renders with its top margin stripped.
  const marginTop = startLine === 0 ? m.marginTopPx : 0;
  const gap = frame.count === 0 ? marginTop : Math.max(frame.pendingMargin, marginTop);
  frame.page.items.push(frameItem(frame, sectionIndex, blockIndex, startLine, endLine));
  frame.cursor += gap + sliceHeight(m, startLine, endLine);
  frame.pendingMargin = endLine >= m.lines.length ? m.marginBottomPx : 0;
  frame.count += 1;
}

// How many of the block's lines from `startLine` on fit within `available`
// px.
function linesFitting(m: BlockMetrics, startLine: number, available: number): number {
  let count = 0;
  const base = m.lines[startLine]?.top ?? 0;
  for (let i = startLine; i < m.lines.length; i += 1) {
    const line = m.lines[i];
    if (line.top + line.height - base <= available + EPS) {
      count = i - startLine + 1;
    } else {
      break;
    }
  }
  return count;
}

/** Lays the whole document out onto pages. */
export function paginate(model: DocumentModel, metrics: MetricsProvider): PageLayout[] {
  const pages: PageLayout[] = [];
  // Every column of every page, in fill order. `fi` is the frame being
  // filled; it only ever moves forward, so frames of earlier pages are never
  // revisited (their keep-with-next tail is still reachable at `fi - 1`).
  const frames: Frame[] = [];
  let fi = -1;

  function currentFrame(): Frame | null {
    return fi >= 0 ? frames[fi] : null;
  }

  function openPage(sectionIndex: number): Frame {
    const geometry = sectionColumnGeometryPx(model, sectionIndex);
    const layout: PageLayout = {
      sectionIndex,
      pageSize: { ...sectionPageSize(model, sectionIndex) },
      margins: { ...sectionMargins(model, sectionIndex) },
      items: [],
    };
    if (geometry.count > 1) {
      layout.columns = {
        count: geometry.count,
        gapCm: geometry.gapCm,
        separator: geometry.separator,
        columns: geometry.widthsCm.map((widthCm, index) => ({
          widthCm,
          offsetCm: geometry.offsetsCm[index],
        })),
      };
    }
    pages.push(layout);
    const base = frames.length;
    for (let c = 0; c < geometry.count; c += 1) {
      frames.push({
        page: layout,
        sectionIndex,
        columnIndex: c,
        columnCount: geometry.count,
        flowWidthPx: geometry.flowWidthPx,
        heightPx: geometry.contentHeightPx,
        cursor: 0,
        pendingMargin: 0,
        count: 0,
      });
    }
    fi = base;
    return frames[fi];
  }

  function breakPage(sectionIndex: number): Frame {
    return openPage(sectionIndex);
  }

  // Advances to the next column of the current page, or opens a new page when
  // the current column is the last (AC-9/AC-11: overflow past the last column
  // continues at the top of the first column of the next page).
  function advanceFrame(): Frame {
    const frame = frames[fi];
    if (frame.columnIndex + 1 < frame.columnCount) {
      fi += 1;
      return frames[fi];
    }
    return openPage(frame.sectionIndex);
  }

  // Moves a trailing chain of whole keep-with-next paragraphs from the frame
  // before `frame` onto it (AC-23). Called when a block naturally starts a
  // fresh frame (column or page); abandons over-long, cross-section or
  // frame-emptying chains. Because columns fill left-to-right, the previous
  // frame's items are always the tail of its page's shared item list.
  function pullKeepNextChain(frame: Frame, sectionIndex: number): void {
    const prev = frames[fi - 1];
    if (!prev || frame.count > 0) {
      return;
    }
    const prevItems = prev.page.items;
    const chain: PageItem[] = [];
    for (let i = prevItems.length - 1; i >= 0; i -= 1) {
      const item = prevItems[i];
      if (item.startLine !== undefined || item.endLine !== undefined) {
        break; // a split paragraph is already "with" its continuation
      }
      if (item.sectionIndex !== sectionIndex) {
        break; // a section break explicitly separates: never pull across it
      }
      const block = model.sections[item.sectionIndex].blocks[item.blockIndex];
      const paragraph = paragraphOf(block);
      if (!paragraph?.formatting?.keepNext) {
        break;
      }
      chain.unshift(item);
      if (chain.length > MAX_KEEP_NEXT_CHAIN) {
        return; // runaway chain: leave everything where it is
      }
    }
    if (chain.length === 0 || chain.length >= prev.count) {
      return; // nothing to move, or moving would empty the previous frame
    }
    let height = 0;
    for (const item of chain) {
      const block = model.sections[item.sectionIndex].blocks[item.blockIndex];
      const m = metrics.metricsFor(item.sectionIndex, item.blockIndex, block, frame.flowWidthPx);
      height += m.heightPx + m.marginTopPx;
    }
    if (height > frame.heightPx * 0.9) {
      return; // the chain alone would (nearly) fill the column: abandon
    }
    prevItems.splice(prevItems.length - chain.length, chain.length);
    prev.count -= chain.length;
    for (const item of chain) {
      const block = model.sections[item.sectionIndex].blocks[item.blockIndex];
      const m = metrics.metricsFor(item.sectionIndex, item.blockIndex, block, frame.flowWidthPx);
      placeWhole(frame, item.sectionIndex, item.blockIndex, m);
    }
  }

  function placeBlock(sectionIndex: number, blockIndex: number, block: Block): void {
    let frame = currentFrame() ?? openPage(sectionIndex);
    const m = metrics.metricsFor(sectionIndex, blockIndex, block, frame.flowWidthPx);
    const paragraph = paragraphOf(block);
    const formatting = paragraph?.formatting;
    const widow = formatting?.widowControl !== false;

    if (formatting?.pageBreakBefore && frame.count > 0) {
      frame = breakPage(sectionIndex);
    }
    // An explicit break before this block separates it from what precedes;
    // only natural frame starts may pull a keep-with-next tail after them —
    // and only once per block, so a chain can never chase a block that
    // keeps deferring (loop safety).
    let mayPull = formatting?.pageBreakBefore !== true;

    let startLine = 0;
    const lineCount = m.lines.length;
    for (;;) {
      if (frame.count === 0 && startLine === 0 && mayPull) {
        mayPull = false;
        pullKeepNextChain(frame, sectionIndex);
      }
      const empty = frame.count === 0;
      const marginTop = startLine === 0 ? m.marginTopPx : 0;
      const top = frame.cursor + (empty ? marginTop : Math.max(frame.pendingMargin, marginTop));
      const available = frame.heightPx - top;
      const remaining = startLine === 0 ? m.heightPx : sliceHeight(m, startLine, lineCount);

      if (remaining <= available + EPS) {
        if (startLine === 0) {
          placeWhole(frame, sectionIndex, blockIndex, m);
        } else {
          placeSlice(frame, sectionIndex, blockIndex, m, startLine, lineCount);
        }
        return;
      }

      // Does not fit. Decide how many lines stay in this column.
      const cannotSplit = lineCount <= 1 || formatting?.keepLines === true;
      let fit = cannotSplit ? 0 : linesFitting(m, startLine, available);
      if (!cannotSplit && widow) {
        if (lineCount - startLine - fit === 1 && fit > 0) {
          fit -= 1; // never strand a single last line in the next column (widow)
        }
        if (startLine === 0 && fit === 1) {
          fit = 0; // never leave a single first line behind (orphan)
        }
      }

      if (fit <= 0) {
        if (empty) {
          // Nothing fits even in an empty column: force progress — place at
          // least one line (split) or the whole unsplittable block
          // (overflow is clipped by the column).
          if (cannotSplit) {
            placeWhole(frame, sectionIndex, blockIndex, m);
            return;
          }
          fit = Math.max(1, linesFitting(m, startLine, available));
          placeSlice(frame, sectionIndex, blockIndex, m, startLine, startLine + fit);
          startLine += fit;
          frame = advanceFrame();
          continue;
        }
        // Defer the whole (remaining) block to a fresh column.
        frame = advanceFrame();
        continue;
      }

      placeSlice(frame, sectionIndex, blockIndex, m, startLine, startLine + fit);
      startLine += fit;
      frame = advanceFrame();
    }
  }

  for (let s = 0; s < model.sections.length; s += 1) {
    const section = model.sections[s];
    const current = currentFrame();
    if (current === null) {
      openPage(s);
    } else if (section.start !== "continuous") {
      breakPage(s);
    } else {
      // A continuous section keeps filling the current page only when both it
      // and the current page are single-column of the same width — blocks are
      // measured at one width, so mixed widths (or a column-count change)
      // cannot share a page and force a break instead.
      const geometry = sectionColumnGeometryPx(model, s);
      const sameSingleColumn =
        geometry.count === 1 &&
        current.columnCount === 1 &&
        Math.abs(geometry.flowWidthPx - current.flowWidthPx) <= EPS;
      if (!sameSingleColumn) {
        breakPage(s);
      }
    }
    for (let b = 0; b < section.blocks.length; b += 1) {
      const block = section.blocks[b];
      if (isBreak(block)) {
        // The break marker sits on the frame it ends (visible when formatting
        // marks are on). A page break then opens the next page; a column
        // break advances to the next column — but only in a multi-column
        // section, so in a single column it is inert (AC-12: a break with no
        // column to move to produces no visible break, page break or line
        // break).
        const frame = currentFrame() ?? openPage(s);
        const m = metrics.metricsFor(s, b, block, frame.flowWidthPx);
        placeWhole(frame, s, b, m);
        if (block.kind === "page") {
          breakPage(s);
        } else if (block.kind === "column" && frame.columnCount > 1) {
          advanceFrame();
        }
        continue;
      }
      placeBlock(s, b, block);
    }
  }
  if (pages.length === 0) {
    openPage(0);
  }
  return pages;
}

function sameItem(a: PageItem, b: PageItem): boolean {
  return (
    a.sectionIndex === b.sectionIndex &&
    a.blockIndex === b.blockIndex &&
    a.startLine === b.startLine &&
    a.endLine === b.endLine &&
    a.columnIndex === b.columnIndex
  );
}

function sameColumns(a: PageColumns | undefined, b: PageColumns | undefined): boolean {
  if (a === b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }
  return (
    a.count === b.count &&
    a.gapCm === b.gapCm &&
    a.separator === b.separator &&
    a.columns.length === b.columns.length &&
    a.columns.every(
      (col, index) =>
        col.widthCm === b.columns[index].widthCm && col.offsetCm === b.columns[index].offsetCm,
    )
  );
}

function samePage(a: PageLayout, b: PageLayout): boolean {
  return (
    a.sectionIndex === b.sectionIndex &&
    a.pageSize.widthCm === b.pageSize.widthCm &&
    a.pageSize.heightCm === b.pageSize.heightCm &&
    a.margins.topCm === b.margins.topCm &&
    a.margins.rightCm === b.margins.rightCm &&
    a.margins.bottomCm === b.margins.bottomCm &&
    a.margins.leftCm === b.margins.leftCm &&
    sameColumns(a.columns, b.columns) &&
    a.items.length === b.items.length &&
    a.items.every((item, index) => sameItem(item, b.items[index]))
  );
}

/** `next` with every page that is unchanged from `prev` replaced by the
 * previous page OBJECT — the renderer's v-memo keys on page identity, so an
 * edit near the end of the document re-renders only the pages it actually
 * changed (AC-35). */
export function reconcilePages(prev: PageLayout[], next: PageLayout[]): PageLayout[] {
  return next.map((page, index) => {
    const before = prev[index];
    return before && samePage(before, page) ? before : page;
  });
}

/** A whole-block page item — the typed seed for `.vue` scripts' default
 * parameters (espree bars the annotation; see textEditing.noCaret). */
export function wholePageItem(sectionIndex: number, blockIndex: number): PageItem {
  return { sectionIndex, blockIndex };
}

/** The page (0-based) a block starts on — the caret page indicator's coarse
 * lookup (AC-3); the caller refines split paragraphs by comparing the caret
 * offset against the rendered slices' character windows. */
export function pageOfBlock(pages: PageLayout[], sectionIndex: number, blockIndex: number): number {
  for (let p = 0; p < pages.length; p += 1) {
    for (const item of pages[p].items) {
      if (item.sectionIndex === sectionIndex && item.blockIndex === blockIndex) {
        return p;
      }
    }
  }
  return Math.max(0, pages.length - 1);
}

/** Rough metrics for a block that has not been measured yet — the
 * provisional layout's stand-in so the page count is approximately right
 * the instant a document opens, before rAF-batched measurement replaces
 * them (AC-34). Estimates use the document's default font geometry. */
export function estimateMetrics(
  block: Block,
  contentWidthPx: number,
  fontSizePt: number,
): BlockMetrics {
  const fontPx = (fontSizePt * 96) / 72;
  const lineHeight = fontPx * 1.4;
  if (isBreak(block)) {
    return { contentWidthPx, heightPx: 0, marginTopPx: 0, marginBottomPx: 0, lines: [], estimated: true };
  }
  if (isImage(block)) {
    return {
      contentWidthPx,
      heightPx: block.heightCm * PX_PER_CM,
      marginTopPx: 0,
      marginBottomPx: 0,
      lines: [],
      estimated: true,
    };
  }
  if (isTable(block)) {
    let rows = 0;
    for (const row of block.rows) {
      let tallest = 1;
      for (const cell of row.cells) {
        tallest = Math.max(tallest, cell.blocks.length);
      }
      rows += tallest;
    }
    return {
      contentWidthPx,
      heightPx: rows * lineHeight + 8,
      marginTopPx: 0,
      marginBottomPx: 0,
      lines: [],
      estimated: true,
    };
  }
  let length = 0;
  for (const run of block.runs) {
    length += run.text.length;
  }
  const charsPerLine = Math.max(8, Math.floor(contentWidthPx / (fontPx * 0.5)));
  const lineCount = Math.max(1, Math.ceil(length / charsPerLine));
  const lines: LineBox[] = [];
  for (let i = 0; i < lineCount; i += 1) {
    lines.push({ top: i * lineHeight, height: lineHeight });
  }
  const formatting = block.formatting;
  return {
    contentWidthPx,
    heightPx: lineCount * lineHeight,
    marginTopPx: ((formatting?.spaceBeforePt ?? 0) * 96) / 72,
    marginBottomPx: ((formatting?.spaceAfterPt ?? 0) * 96) / 72,
    lines,
    estimated: true,
  };
}
