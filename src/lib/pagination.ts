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
}

/** One laid-out page: the geometry it renders at (resolved from its opening
 * section, so the renderer needs no model lookups) plus its items. */
export interface PageLayout {
  sectionIndex: number;
  pageSize: PageSize;
  margins: Margins;
  items: PageItem[];
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

// Sub-pixel measurement noise must never flip a fit decision into an
// off-by-one-line oscillation between layout passes.
const EPS = 0.5;

// Backtracking a keep-with-next chain stops rather than move more blocks
// than this — a runaway chain (every paragraph keepNext) must not empty
// pages or loop (AC-23's "honoured" never requires moving whole pages).
const MAX_KEEP_NEXT_CHAIN = 8;

interface OpenPage {
  layout: PageLayout;
  /** Bottom edge (px, content-relative) of the last placed item. */
  cursor: number;
  /** The last item's collapsible bottom margin. */
  pendingMargin: number;
  geometry: SectionGeometryPx;
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

function placeWhole(page: OpenPage, sectionIndex: number, blockIndex: number, m: BlockMetrics): void {
  const gap = page.layout.items.length === 0 ? m.marginTopPx : Math.max(page.pendingMargin, m.marginTopPx);
  page.layout.items.push({ sectionIndex, blockIndex });
  page.cursor += gap + m.heightPx;
  page.pendingMargin = m.marginBottomPx;
}

function placeSlice(
  page: OpenPage,
  sectionIndex: number,
  blockIndex: number,
  m: BlockMetrics,
  startLine: number,
  endLine: number,
): void {
  // A continuation slice renders with its top margin stripped.
  const marginTop = startLine === 0 ? m.marginTopPx : 0;
  const gap = page.layout.items.length === 0 ? marginTop : Math.max(page.pendingMargin, marginTop);
  page.layout.items.push({ sectionIndex, blockIndex, startLine, endLine });
  page.cursor += gap + sliceHeight(m, startLine, endLine);
  page.pendingMargin = endLine >= m.lines.length ? m.marginBottomPx : 0;
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
  let current: OpenPage | null = null;

  function openPage(sectionIndex: number): OpenPage {
    const layout: PageLayout = {
      sectionIndex,
      pageSize: { ...sectionPageSize(model, sectionIndex) },
      margins: { ...sectionMargins(model, sectionIndex) },
      items: [],
    };
    pages.push(layout);
    current = {
      layout,
      cursor: 0,
      pendingMargin: 0,
      geometry: sectionGeometryPx(model, sectionIndex),
    };
    return current;
  }

  function breakPage(sectionIndex: number): OpenPage {
    return openPage(sectionIndex);
  }

  // Moves a trailing chain of whole keep-with-next paragraphs from the page
  // before `page` onto `page` (AC-23). Called when a block naturally starts
  // a fresh page; abandons over-long, cross-section or page-emptying chains.
  function pullKeepNextChain(page: OpenPage, sectionIndex: number): void {
    const prev = pages[pages.length - 2];
    if (!prev || page.layout.items.length > 0) {
      return;
    }
    const chain: PageItem[] = [];
    for (let i = prev.items.length - 1; i >= 0; i -= 1) {
      const item = prev.items[i];
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
    if (chain.length === 0 || chain.length >= prev.items.length) {
      return; // nothing to move, or moving would empty the previous page
    }
    let height = 0;
    for (const item of chain) {
      const block = model.sections[item.sectionIndex].blocks[item.blockIndex];
      const m = metrics.metricsFor(
        item.sectionIndex,
        item.blockIndex,
        block,
        page.geometry.contentWidthPx,
      );
      height += m.heightPx + m.marginTopPx;
    }
    if (height > page.geometry.contentHeightPx * 0.9) {
      return; // the chain alone would (nearly) fill the page: abandon
    }
    prev.items.splice(prev.items.length - chain.length, chain.length);
    for (const item of chain) {
      const block = model.sections[item.sectionIndex].blocks[item.blockIndex];
      const m = metrics.metricsFor(
        item.sectionIndex,
        item.blockIndex,
        block,
        page.geometry.contentWidthPx,
      );
      placeWhole(page, item.sectionIndex, item.blockIndex, m);
    }
  }

  function placeBlock(sectionIndex: number, blockIndex: number, block: Block): void {
    let page = current ?? openPage(sectionIndex);
    const m = metrics.metricsFor(sectionIndex, blockIndex, block, page.geometry.contentWidthPx);
    const paragraph = paragraphOf(block);
    const formatting = paragraph?.formatting;
    const widow = formatting?.widowControl !== false;

    if (formatting?.pageBreakBefore && page.layout.items.length > 0) {
      page = breakPage(sectionIndex);
    }
    // An explicit break before this block separates it from what precedes;
    // only natural page starts may pull a keep-with-next tail after them —
    // and only once per block, so a chain can never chase a block that
    // keeps deferring (loop safety).
    let mayPull = formatting?.pageBreakBefore !== true;

    let startLine = 0;
    const lineCount = m.lines.length;
    for (;;) {
      if (page.layout.items.length === 0 && startLine === 0 && mayPull) {
        mayPull = false;
        pullKeepNextChain(page, sectionIndex);
      }
      const empty = page.layout.items.length === 0;
      const marginTop = startLine === 0 ? m.marginTopPx : 0;
      const top = page.cursor + (empty ? marginTop : Math.max(page.pendingMargin, marginTop));
      const available = page.geometry.contentHeightPx - top;
      const remaining = startLine === 0 ? m.heightPx : sliceHeight(m, startLine, lineCount);

      if (remaining <= available + EPS) {
        if (startLine === 0) {
          placeWhole(page, sectionIndex, blockIndex, m);
        } else {
          placeSlice(page, sectionIndex, blockIndex, m, startLine, lineCount);
        }
        return;
      }

      // Does not fit. Decide how many lines stay on this page.
      const cannotSplit = lineCount <= 1 || formatting?.keepLines === true;
      let fit = cannotSplit ? 0 : linesFitting(m, startLine, available);
      if (!cannotSplit && widow) {
        if (lineCount - startLine - fit === 1 && fit > 0) {
          fit -= 1; // never strand a single last line on the next page (widow)
        }
        if (startLine === 0 && fit === 1) {
          fit = 0; // never leave a single first line behind (orphan)
        }
      }

      if (fit <= 0) {
        if (empty) {
          // Nothing fits even on an empty page: force progress — place at
          // least one line (split) or the whole unsplittable block
          // (overflow is clipped by the page).
          if (cannotSplit) {
            placeWhole(page, sectionIndex, blockIndex, m);
            return;
          }
          fit = Math.max(1, linesFitting(m, startLine, available));
          placeSlice(page, sectionIndex, blockIndex, m, startLine, startLine + fit);
          startLine += fit;
          page = breakPage(sectionIndex);
          continue;
        }
        // Defer the whole (remaining) block to a fresh page.
        page = breakPage(sectionIndex);
        continue;
      }

      placeSlice(page, sectionIndex, blockIndex, m, startLine, startLine + fit);
      startLine += fit;
      page = breakPage(sectionIndex);
    }
  }

  for (let s = 0; s < model.sections.length; s += 1) {
    const section = model.sections[s];
    if (current === null) {
      openPage(s);
    } else if (section.start !== "continuous") {
      breakPage(s);
    } else {
      // A continuous section keeps filling the current page — unless its
      // content width differs from the page's (blocks are measured at their
      // own section's width, so mixed widths on one page cannot render).
      const geometry = sectionGeometryPx(model, s);
      const active: OpenPage = current;
      if (Math.abs(geometry.contentWidthPx - active.geometry.contentWidthPx) > EPS) {
        breakPage(s);
      }
    }
    for (let b = 0; b < section.blocks.length; b += 1) {
      const block = section.blocks[b];
      if (isBreak(block)) {
        // The break block itself sits on the page it ends (its marker is
        // visible there when formatting marks are on); a page break then
        // opens the next page. Column breaks are inert until SPEC-030.
        const page: OpenPage = current ?? openPage(s);
        const m = metrics.metricsFor(s, b, block, page.geometry.contentWidthPx);
        placeWhole(page, s, b, m);
        if (block.kind === "page") {
          breakPage(s);
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
    a.endLine === b.endLine
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
