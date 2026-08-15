// SPEC-028 — the pagination engine, driven by fake metrics so every rule
// (breaks, sections, keeps, widow/orphan, splitting) is exercised as pure
// math, no DOM.
import { describe, expect, it } from "vitest";
import {
  estimateMetrics,
  paginate,
  pageOfBlock,
  reconcilePages,
  type MetricsProvider,
  type PageLayout,
} from "./pagination";
import { type BlockMetrics } from "./core/pageMetrics";
import { PX_PER_CM } from "./core/units";
import {
  createDocument,
  type Block,
  type DocumentModel,
  type Paragraph,
  type ParagraphFormatting,
} from "./documentModel";

const LINE = 10; // px per fake line
const PAGE_LINES = 10; // lines per page in the fixture geometry

function paragraph(lineCount: number, formatting?: ParagraphFormatting): Paragraph {
  const block: Paragraph = {
    type: "paragraph",
    // One character per line lets the fake provider derive the line count.
    runs: [{ text: "x".repeat(lineCount), formatting: {} }],
  };
  if (formatting) {
    block.formatting = formatting;
  }
  return block;
}

/** Fake metrics: a paragraph's line count is its text length, each line
 * LINE px tall; atoms are LINE px tall. */
const fakeMetrics: MetricsProvider = {
  metricsFor(_s, _b, block: Block, contentWidthPx): BlockMetrics {
    if (block.type !== "paragraph") {
      return { contentWidthPx, heightPx: block.type === "break" ? 0 : LINE, marginTopPx: 0, marginBottomPx: 0, lines: [] };
    }
    const count = Math.max(1, block.runs.reduce((n, run) => n + run.text.length, 0));
    return {
      contentWidthPx,
      heightPx: count * LINE,
      marginTopPx: 0,
      marginBottomPx: 0,
      lines: Array.from({ length: count }, (_, i) => ({ top: i * LINE, height: LINE })),
    };
  },
};

/** One-section fixture whose content area is exactly PAGE_LINES fake lines
 * tall (margins zero so px math stays exact). */
function fixture(...blocks: Block[]): DocumentModel {
  const model = createDocument("pagination-fixture");
  model.defaultStyles.pageSize = {
    widthCm: 100 / PX_PER_CM,
    heightCm: (PAGE_LINES * LINE) / PX_PER_CM,
  };
  model.defaultStyles.margins = { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 };
  model.sections = [{ blocks }];
  return model;
}

function itemsOf(page: PageLayout): Array<[number, number | undefined, number | undefined]> {
  return page.items.map((item) => [item.blockIndex, item.startLine, item.endLine]);
}

describe("paginate", () => {
  it("flows content across pages, breaking at the bottom margin (AC-1)", () => {
    const model = fixture(...Array.from({ length: 25 }, () => paragraph(1)));
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(3);
    expect(pages[0].items).toHaveLength(10);
    expect(pages[1].items).toHaveLength(10);
    expect(pages[2].items).toHaveLength(5);
  });

  it("renders an empty document as one page", () => {
    const pages = paginate(fixture(paragraph(0)), fakeMetrics);
    expect(pages).toHaveLength(1);
  });

  it("splits a long paragraph across pages at line boundaries", () => {
    const pages = paginate(fixture(paragraph(25)), fakeMetrics);
    expect(pages).toHaveLength(3);
    expect(itemsOf(pages[0])).toEqual([[0, 0, 10]]);
    expect(itemsOf(pages[1])).toEqual([[0, 10, 20]]);
    expect(itemsOf(pages[2])).toEqual([[0, 20, 25]]);
  });

  it("starts content after a manual page break on a new page (AC-5)", () => {
    const model = fixture(paragraph(2), { type: "break", kind: "page" }, paragraph(2));
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(2);
    // The break block sits on the page it ends.
    expect(pages[0].items.map((item) => item.blockIndex)).toEqual([0, 1]);
    expect(pages[1].items.map((item) => item.blockIndex)).toEqual([2]);
  });

  it("keeps a column break in the flow without breaking the page", () => {
    const model = fixture(paragraph(2), { type: "break", kind: "column" }, paragraph(2));
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(1);
  });

  it("honours page-break-before (AC-27)", () => {
    const model = fixture(paragraph(2), paragraph(2, { pageBreakBefore: true }));
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(2);
    expect(pages[1].items.map((item) => item.blockIndex)).toEqual([1]);
  });

  it("does not open a page before a page-break-before paragraph already at a page top", () => {
    const model = fixture(paragraph(2, { pageBreakBefore: true }), paragraph(1));
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(1);
  });

  it("honours keep-lines-together by moving the paragraph whole (AC-25)", () => {
    const split = paginate(fixture(paragraph(7), paragraph(5)), fakeMetrics);
    expect(itemsOf(split[0])).toEqual([
      [0, undefined, undefined],
      [1, 0, 3],
    ]);
    const kept = paginate(fixture(paragraph(7), paragraph(5, { keepLines: true })), fakeMetrics);
    expect(itemsOf(kept[0])).toEqual([[0, undefined, undefined]]);
    expect(itemsOf(kept[1])).toEqual([[1, undefined, undefined]]);
  });

  it("widow control never strands a single line on the next page (AC-29)", () => {
    // 8 + 3: only 2 of the 3 lines fit, which would leave a 1-line widow —
    // so one line less stays, splitting 1/2 — but that is a 1-line orphan,
    // so the paragraph moves whole.
    const pages = paginate(fixture(paragraph(8), paragraph(3)), fakeMetrics);
    expect(itemsOf(pages[0])).toEqual([[0, undefined, undefined]]);
    expect(itemsOf(pages[1])).toEqual([[1, undefined, undefined]]);
  });

  it("widow control off allows single stranded lines", () => {
    const pages = paginate(
      fixture(paragraph(8), paragraph(3, { widowControl: false })),
      fakeMetrics,
    );
    expect(itemsOf(pages[0])).toEqual([
      [0, undefined, undefined],
      [1, 0, 2],
    ]);
    expect(itemsOf(pages[1])).toEqual([[1, 2, 3]]);
  });

  it("widow control never leaves a single first line behind (orphan)", () => {
    // 9 + 4: one line of the second paragraph fits — an orphan, so the
    // paragraph moves whole.
    const pages = paginate(fixture(paragraph(9), paragraph(4)), fakeMetrics);
    expect(itemsOf(pages[0])).toEqual([[0, undefined, undefined]]);
    expect(itemsOf(pages[1])).toEqual([[1, undefined, undefined]]);
  });

  it("splits with at least two lines on each side when it can", () => {
    // 6 + 8: four of the 8 lines fit; 4/4 satisfies both minima.
    const pages = paginate(fixture(paragraph(6), paragraph(8)), fakeMetrics);
    expect(itemsOf(pages[0])).toEqual([
      [0, undefined, undefined],
      [1, 0, 4],
    ]);
    expect(itemsOf(pages[1])).toEqual([[1, 4, 8]]);
  });

  it("honours keep-with-next by pulling the paragraph to the next page (AC-23)", () => {
    // 9 + keepNext(1) fill page 1 exactly; the follower starts page 2 — the
    // keepNext paragraph must follow it.
    const model = fixture(paragraph(9), paragraph(1, { keepNext: true }), paragraph(2));
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(2);
    expect(pages[0].items.map((item) => item.blockIndex)).toEqual([0]);
    expect(pages[1].items.map((item) => item.blockIndex)).toEqual([1, 2]);
  });

  it("abandons a keep-with-next chain rather than emptying the page", () => {
    // Every block keepNext: nothing can move without emptying the page.
    const blocks = Array.from({ length: 12 }, () => paragraph(1, { keepNext: true }));
    const pages = paginate(fixture(...blocks), fakeMetrics);
    expect(pages).toHaveLength(2);
    expect(pages[0].items.length).toBeGreaterThan(0);
  });

  it("a section break (next page) starts a new page (AC-8)", () => {
    const model = fixture(paragraph(2));
    model.sections.push({ blocks: [paragraph(2)], start: "nextPage" });
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(2);
    expect(pages[1].sectionIndex).toBe(1);
  });

  it("a continuous section break keeps filling the same page (AC-9)", () => {
    const model = fixture(paragraph(2));
    model.sections.push({ blocks: [paragraph(2)], start: "continuous" });
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(1);
    expect(pages[0].items.map((item) => item.sectionIndex)).toEqual([0, 1]);
  });

  it("a continuous section with a different content width forces a page (measurement integrity)", () => {
    const model = fixture(paragraph(2));
    model.sections.push({
      blocks: [paragraph(2)],
      start: "continuous",
      pageSize: { widthCm: 50 / PX_PER_CM, heightCm: (PAGE_LINES * LINE) / PX_PER_CM },
    });
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(2);
  });

  it("pages carry their own section's page size and margins (AC-17..19)", () => {
    const model = fixture(paragraph(2));
    model.sections.push({
      blocks: [paragraph(2)],
      start: "nextPage",
      pageSize: { widthCm: 14.8, heightCm: 21 },
      margins: { topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 },
    });
    const pages = paginate(model, fakeMetrics);
    expect(pages[0].pageSize.widthCm).toBeCloseTo(100 / PX_PER_CM, 5);
    expect(pages[1].pageSize).toEqual({ widthCm: 14.8, heightCm: 21 });
    expect(pages[1].margins).toEqual({ topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 });
  });

  it("an unsplittable block taller than the page overflows rather than looping", () => {
    const model = fixture(paragraph(30, { keepLines: true }), paragraph(1));
    const pages = paginate(model, fakeMetrics);
    expect(pages).toHaveLength(2);
    expect(itemsOf(pages[0])).toEqual([[0, undefined, undefined]]);
  });

  it("lays out a 200-page document fast enough for interactive re-layout (AC-34/35)", () => {
    const blocks = Array.from({ length: 10 * 200 }, () => paragraph(1));
    const model = fixture(...blocks);
    const started = performance.now();
    const pages = paginate(model, fakeMetrics);
    const elapsed = performance.now() - started;
    expect(pages).toHaveLength(200);
    expect(elapsed).toBeLessThan(300);
  });
});

describe("reconcilePages", () => {
  it("keeps page object identity for unchanged pages (AC-35)", () => {
    const before = paginate(fixture(...Array.from({ length: 25 }, () => paragraph(1))), fakeMetrics);
    // Growing the last paragraph to 10 lines restructures the last page
    // (the paragraph now splits) while pages 1 and 2 stay identical.
    const model = fixture(...Array.from({ length: 24 }, () => paragraph(1)), paragraph(10));
    const after = reconcilePages(before, paginate(model, fakeMetrics));
    expect(after[0]).toBe(before[0]);
    expect(after[1]).toBe(before[1]);
    expect(after[2]).not.toBe(before[2]);
  });
});

describe("pageOfBlock", () => {
  it("finds the page a block starts on", () => {
    const pages = paginate(fixture(...Array.from({ length: 25 }, () => paragraph(1))), fakeMetrics);
    expect(pageOfBlock(pages, 0, 0)).toBe(0);
    expect(pageOfBlock(pages, 0, 10)).toBe(1);
    expect(pageOfBlock(pages, 0, 24)).toBe(2);
  });
});

// Multi-column fixture: two zero-gap equal columns over the 100px content
// width give two 50px columns, each PAGE_LINES fake lines tall.
function columnFixture(columnCount: number, ...blocks: Block[]): DocumentModel {
  const model = fixture(...blocks);
  model.sections[0].columns = {
    count: columnCount,
    equalWidth: true,
    gapCm: 0,
    separator: false,
  };
  return model;
}

function columnsOf(page: PageLayout): Array<number | undefined> {
  return page.items.map((item) => item.columnIndex);
}

describe("paginate — columns (SPEC-030)", () => {
  it("flows down the first column, then the next, then the next page (AC-9)", () => {
    const pages = paginate(
      columnFixture(2, ...Array.from({ length: 25 }, () => paragraph(1))),
      fakeMetrics,
    );
    expect(pages).toHaveLength(2);
    // Page 1: 10 in column 0, 10 in column 1.
    expect(pages[0].items).toHaveLength(20);
    expect(columnsOf(pages[0])).toEqual([...Array(10).fill(0), ...Array(10).fill(1)]);
    // Overflow continues at the top of page 2's first column.
    expect(pages[1].items).toHaveLength(5);
    expect(columnsOf(pages[1])).toEqual(Array(5).fill(0));
    expect(pages[1].items[0].columnIndex).toBe(0);
  });

  it("exposes each page's resolved column geometry", () => {
    const pages = paginate(columnFixture(2, paragraph(1)), fakeMetrics);
    expect(pages[0].columns?.count).toBe(2);
    expect(pages[0].columns?.columns).toHaveLength(2);
    expect(pages[0].columns?.columns[0].offsetCm).toBe(0);
    expect(pages[0].columns?.columns[1].offsetCm).toBeCloseTo(pages[0].columns!.columns[0].widthCm, 5);
  });

  it("keeps single-column pages free of column metadata (back-compat)", () => {
    const pages = paginate(fixture(paragraph(1)), fakeMetrics);
    expect(pages[0].columns).toBeUndefined();
    expect(pages[0].items[0].columnIndex).toBeUndefined();
  });

  it("moves following content to the next column at a column break (AC-10)", () => {
    const pages = paginate(
      columnFixture(2, paragraph(1), { type: "break", kind: "column" }, paragraph(1)),
      fakeMetrics,
    );
    expect(pages).toHaveLength(1);
    // paragraph 0 + the marker in column 0; paragraph 2 in column 1.
    const p0 = pages[0].items.find((i) => i.blockIndex === 0);
    const p2 = pages[0].items.find((i) => i.blockIndex === 2);
    expect(p0?.columnIndex).toBe(0);
    expect(p2?.columnIndex).toBe(1);
  });

  it("treats a column break as a no-op in a single-column section (AC-12)", () => {
    const pages = paginate(
      fixture(paragraph(1), { type: "break", kind: "column" }, paragraph(1)),
      fakeMetrics,
    );
    expect(pages).toHaveLength(1);
    // No page break, no new column: all three blocks flow on one page.
    expect(pages[0].items.map((i) => i.blockIndex)).toEqual([0, 1, 2]);
  });

  it("a column break in the last column moves to the next page (AC-11)", () => {
    const pages = paginate(
      columnFixture(2, paragraph(1), { type: "break", kind: "column" }, paragraph(1), {
        type: "break",
        kind: "column",
      }, paragraph(1)),
      fakeMetrics,
    );
    // block0 col0, block2 col1, block4 -> next page col0.
    expect(pages).toHaveLength(2);
    expect(pages[1].items.find((i) => i.blockIndex === 4)?.columnIndex).toBe(0);
  });

  it("restores column-break behaviour when the count grows again (AC-13)", () => {
    const single = paginate(
      fixture(paragraph(1), { type: "break", kind: "column" }, paragraph(1)),
      fakeMetrics,
    );
    expect(single).toHaveLength(1);
    expect(single[0].items).toHaveLength(3);
    const multi = paginate(
      columnFixture(2, paragraph(1), { type: "break", kind: "column" }, paragraph(1)),
      fakeMetrics,
    );
    expect(multi[0].items.find((i) => i.blockIndex === 2)?.columnIndex).toBe(1);
  });
});

describe("estimateMetrics", () => {
  it("estimates plausible paragraph lines from text length", () => {
    const m = estimateMetrics(paragraph(0), 500, 11);
    expect(m.estimated).toBe(true);
    expect(m.lines.length).toBe(1);
    const long: Paragraph = { type: "paragraph", runs: [{ text: "y".repeat(1000), formatting: {} }] };
    expect(estimateMetrics(long, 500, 11).lines.length).toBeGreaterThan(5);
  });

  it("estimates a break as zero-height", () => {
    expect(estimateMetrics({ type: "break", kind: "page" }, 500, 11).heightPx).toBe(0);
  });
});
