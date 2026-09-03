// SPEC-076 — a table's leading header run repeated at the top of every page
// it continues onto. Three surfaces meet here and all three are pure, so the
// whole rule is testable without a browser:
//
//   - tableSliceRows (lib/tables.ts)     which rows one page slice draws
//   - measureTableElement (core/pageMetrics.ts) one line box per MODEL row
//   - paginate (lib/pagination.ts)       the height the repeat costs a page
import { describe, expect, it } from "vitest";
import { insertTable, tableSliceRows } from "./tables";
import { caretSelection } from "./commands";
import { measureTableElement, type BlockMetrics } from "./core/pageMetrics";
import { paginate, type MetricsProvider } from "./pagination";
import { equalStructural } from "./equality";
import { documentBodyStats } from "./wordCount";
import { createDocument, type DocumentModel, type Table, type TableRow } from "./documentModel";
import { PX_PER_CM } from "./core/units";

const ROW = 10; // px per fake table row
const PAGE_ROWS = 10; // rows of content per page in the fixture geometry

function row(text: string, isHeader = false): TableRow {
  const built: TableRow = {
    cells: [
      { blocks: [{ type: "paragraph", runs: [{ text, formatting: {} }] }] },
      { blocks: [{ type: "paragraph", runs: [{ text: `${text} two`, formatting: {} }] }] },
    ],
  };
  if (isHeader) {
    built.isHeader = true;
  }
  return built;
}

function table(headerRows: number, bodyRows: number): Table {
  return {
    type: "table",
    columnWidthsCm: [4, 4],
    rows: [
      ...Array.from({ length: headerRows }, (_unused, index) => row(`head ${index}`, true)),
      ...Array.from({ length: bodyRows }, (_unused, index) => row(`body ${index}`)),
    ],
  };
}

/** Fake metrics in the shape measureTableElement produces: one ROW-tall line
 * box per model row, and nothing else on the page. */
const fakeMetrics: MetricsProvider = {
  metricsFor(_section, _block, block, contentWidthPx): BlockMetrics {
    const rows = block.type === "table" ? block.rows.length : 1;
    return {
      contentWidthPx,
      heightPx: rows * ROW,
      marginTopPx: 0,
      marginBottomPx: 0,
      lines: Array.from({ length: rows }, (_unused, index) => ({ top: index * ROW, height: ROW })),
    };
  },
};

function fixture(block: Table): DocumentModel {
  const model = createDocument("header-repeat-fixture");
  model.defaultStyles.pageSize = {
    widthCm: 400 / PX_PER_CM,
    heightCm: (PAGE_ROWS * ROW) / PX_PER_CM,
  };
  model.defaultStyles.margins = { topCm: 0, rightCm: 0, bottomCm: 0, leftCm: 0 };
  model.sections = [{ blocks: [block] }];
  return model;
}

describe("tableSliceRows", () => {
  it("draws every row once for a table that fits on one page (AC-6)", () => {
    const slice = tableSliceRows(table(1, 3));
    expect(slice.map((entry) => entry.rowIndex)).toEqual([0, 1, 2, 3]);
    expect(slice.some((entry) => entry.repeated)).toBe(false);
  });

  it("redraws the leading header run above a continuation's first body row (AC-1)", () => {
    const slice = tableSliceRows(table(1, 5), 3, 6);
    expect(slice.map((entry) => [entry.rowIndex, entry.repeated])).toEqual([
      [0, true],
      [3, false],
      [4, false],
      [5, false],
    ]);
  });

  it("repeats multiple contiguous header rows together, in model order (AC-4)", () => {
    const slice = tableSliceRows(table(3, 5), 5, 8);
    expect(slice.filter((entry) => entry.repeated).map((entry) => entry.rowIndex)).toEqual([0, 1, 2]);
    expect(slice[0].row.cells[0].blocks[0].runs[0].text).toBe("head 0");
  });

  it("repeats nothing when the first row is not a header, even if a later row is (AC-5)", () => {
    const late = table(0, 5);
    late.rows[2].isHeader = true;
    const slice = tableSliceRows(late, 3, 5);
    expect(slice.every((entry) => !entry.repeated)).toBe(true);
    expect(slice.map((entry) => entry.rowIndex)).toEqual([3, 4]);
  });

  it("renders a repeated row from the very same model row object (AC-2, AC-3)", () => {
    const source = table(1, 4);
    const slice = tableSliceRows(source, 3, 5);
    expect(slice[0].row).toBe(source.rows[0]);
  });
});

describe("measureTableElement", () => {
  function tableElement(heights: number[], repeatedCount: number): Element {
    const host = document.createElement("div");
    host.innerHTML = `<table><tbody>${heights
      .map((_unused, index) => `<tr${index < repeatedCount ? " data-repeated-header=\"true\"" : ""}></tr>`)
      .join("")}</tbody></table>`;
    const element = host.querySelector("table") as Element;
    // jsdom measures nothing, so stub the per-row rects the reader asks for.
    Array.from(element.querySelectorAll("tr")).forEach((tr, index) => {
      tr.getBoundingClientRect = () => ({ height: heights[index] }) as DOMRect;
    });
    return element;
  }

  it("measures one line box per model row (SPEC-076)", () => {
    const metrics = measureTableElement(tableElement([12, 20, 8], 0), 500);
    expect(metrics.lines).toEqual([
      { top: 0, height: 12 },
      { top: 12, height: 20 },
      { top: 32, height: 8 },
    ]);
    expect(metrics.heightPx).toBe(40);
  });

  it("skips repeated header rows so the run is measured once per table, not once per page (AC-11)", () => {
    const metrics = measureTableElement(tableElement([12, 20, 8], 1), 500);
    expect(metrics.lines).toEqual([
      { top: 0, height: 20 },
      { top: 20, height: 8 },
    ]);
    expect(metrics.heightPx).toBe(28);
  });
});

describe("paginate with a repeating header", () => {
  it("splits a tall table at row boundaries and charges each continuation the header's height (AC-1)", () => {
    // 1 header + 24 body rows, 10 rows of room per page.
    const pages = paginate(fixture(table(1, 24)), fakeMetrics);
    expect(pages.length).toBeGreaterThan(2);
    // Page 1 holds the header plus 9 body rows; every continuation loses one
    // row of room to the redrawn header.
    expect(pages[0].items[0].startLine).toBe(0);
    expect(pages[0].items[0].endLine).toBe(10);
    expect(pages[1].items[0].startLine).toBe(10);
    expect(pages[1].items[0].endLine).toBe(19);
    expect(pages[2].items[0].startLine).toBe(19);
  });

  it("never splits inside the header run itself (AC-1)", () => {
    const model = fixture(table(3, 20));
    const pages = paginate(model, fakeMetrics);
    for (const page of pages) {
      const item = page.items[0];
      // A slice either starts at the table's top or past the whole run.
      expect(item.startLine === 0 || (item.startLine ?? 0) >= 3).toBe(true);
      if (item.startLine === 0) {
        expect(item.endLine === undefined || item.endLine > 3).toBe(true);
      }
    }
  });

  it("leaves the model untouched — the repeat is a layout artefact (AC-3)", () => {
    const model = fixture(table(1, 24));
    const before = structuredClone(model);
    paginate(model, fakeMetrics);
    expect(equalStructural(model, before)).toBe(true);
  });

  it("counts the words of a four-page table exactly once (AC-10)", () => {
    const tall = fixture(table(1, 40));
    const flat = fixture(table(1, 40));
    // Lay one of them out across pages; the other never paginates.
    const pages = paginate(tall, fakeMetrics);
    expect(pages.length).toBeGreaterThanOrEqual(4);
    expect(documentBodyStats(tall).words).toBe(documentBodyStats(flat).words);
    // And the repeat adds nothing to the count itself: 41 rows x 2 cells,
    // "head 0" / "body n" (2 words) and "head 0 two" / "body n two" (3).
    expect(documentBodyStats(flat).words).toBe(41 * 5);
  });

  it("draws a table that fits on one page as a single unsplit item (AC-6)", () => {
    const pages = paginate(fixture(table(1, 5)), fakeMetrics);
    expect(pages).toHaveLength(1);
    expect(pages[0].items[0].startLine).toBeUndefined();
  });
});

describe("insertTable's header-row option (SPEC-076 AC-20, AC-21)", () => {
  const at = caretSelection({ section: 0, block: 0, offset: 0 });

  function emptyModel(): DocumentModel {
    const model = createDocument("insert-table-fixture");
    model.sections = [{ blocks: [{ type: "paragraph", runs: [] }] }];
    return model;
  }

  it("marks the first row isHeader when the dialog's checkbox is on (AC-20)", () => {
    const next = insertTable(emptyModel(), at, 3, 2, true).model;
    const table = next.sections[0].blocks.find((block) => block.type === "table") as Table;
    expect(table.rows[0].isHeader).toBe(true);
    expect(table.rows[1].isHeader).toBeUndefined();
  });

  it("creates no header row when it is off (AC-20)", () => {
    const next = insertTable(emptyModel(), at, 3, 2, false).model;
    const table = next.sections[0].blocks.find((block) => block.type === "table") as Table;
    expect(table.rows.every((row) => row.isHeader === undefined)).toBe(true);
  });

  it("creates no header row when the caller passes nothing — the Insert menu's grid quick-pick (AC-21)", () => {
    const next = insertTable(emptyModel(), at, 3, 2).model;
    const table = next.sections[0].blocks.find((block) => block.type === "table") as Table;
    expect(table.rows.every((row) => row.isHeader === undefined)).toBe(true);
  });
});
