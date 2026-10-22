// SPEC-076 — a table's leading header run repeated at the top of every page
// it continues onto. Three surfaces meet here and all three are pure, so the
// whole rule is testable without a browser:
//
//   - tableSliceRows (lib/tables.ts)     which rows one page slice draws
//   - measureTableElement (core/pageMetrics.ts) one line box per MODEL row
//   - paginate (lib/pagination.ts)       the height the repeat costs a page
//
// A table splitting at row boundaries at all is new here, so the rules about
// WHERE it may split — never inside the header run, never inside a vertical
// merge (core/tables.verticalMergeCrossings) — live here too.
import { describe, expect, it } from "vitest";
import { insertTable, tableSliceRows } from "./tables";
import { verticalMergeCrossings } from "./core/tables";
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

/** Metrics with an explicit height per model row — for the cases where the
 * split turns on one row being far taller than a page. */
function rowMetrics(heights: number[]): MetricsProvider {
  return {
    metricsFor(_section, _block, _block2, contentWidthPx): BlockMetrics {
      let top = 0;
      const lines = heights.map((height) => {
        const line = { top, height };
        top += height;
        return line;
      });
      return { contentWidthPx, heightPx: top, marginTopPx: 0, marginBottomPx: 0, lines };
    },
  };
}

/** Merges `rows` rows downward from `row` in column `col`, exactly as
 * SPEC-023's merge command leaves the model: the owner carries the rowSpan
 * and every slot below it becomes a covered placeholder. */
function mergeDown(built: Table, atRow: number, col: number, rows: number): Table {
  built.rows[atRow].cells[col] = { ...built.rows[atRow].cells[col], rowSpan: rows };
  for (let index = atRow + 1; index < atRow + rows; index += 1) {
    built.rows[index].cells[col] = { blocks: [], covered: true };
  }
  return built;
}

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

// A table splits at row boundaries now, so every row it draws must still fill
// the whole grid: the colgroup, the repeated header above it and the rows
// around it all assume a full complement of columns.
describe("verticalMergeCrossings", () => {
  it("crosses nothing when no cell is merged downward", () => {
    expect(verticalMergeCrossings(table(1, 3))).toEqual([false, false, false, false]);
  });

  it("marks every boundary a rowSpan reaches over, and not the one it starts at", () => {
    const merged = mergeDown(table(1, 4), 1, 0, 3); // rows 1..3 are one cell
    expect(verticalMergeCrossings(merged)).toEqual([false, false, true, true, false]);
  });

  it("ignores a purely horizontal merge", () => {
    const wide = table(1, 2);
    wide.rows[1].cells[0] = { ...wide.rows[1].cells[0], colSpan: 2 };
    wide.rows[1].cells[1] = { blocks: [], covered: true };
    expect(verticalMergeCrossings(wide)).toEqual([false, false, false]);
  });
});

describe("tableSliceRows cells", () => {
  it("gives every column an entry when nothing is merged", () => {
    const [first] = tableSliceRows(table(1, 1));
    expect(first.cells.map((entry) => [entry.colIndex, entry.colSpan, entry.rowSpan, entry.orphaned])).toEqual([
      [0, 1, 1, false],
      [1, 1, 1, false],
    ]);
  });

  it("draws a horizontally merged cell once, spanning its covered neighbour", () => {
    const wide = table(0, 1);
    wide.rows[0].cells[0] = { ...wide.rows[0].cells[0], colSpan: 2 };
    wide.rows[0].cells[1] = { blocks: [], covered: true };
    const [only] = tableSliceRows(wide);
    expect(only.cells).toHaveLength(1);
    expect(only.cells[0].colSpan).toBe(2);
  });

  it("clips a rowspan reaching past the last row this slice draws", () => {
    // Rows 1..3 are one cell; the slice stops after row 2, so it may draw
    // only two of the three rows — a taller rowspan would make the rendered
    // slice taller than measureTableElement reported.
    const merged = mergeDown(table(1, 4), 1, 0, 3);
    const slice = tableSliceRows(merged, 0, 3);
    expect(slice[1].cells[0].rowSpan).toBe(2);
    expect(slice[1].cells[0].orphaned).toBe(false);
  });

  it("stands in for a merge owner an earlier slice drew, keeping the columns aligned", () => {
    const merged = mergeDown(table(1, 4), 1, 0, 3);
    const slice = tableSliceRows(merged, 2, 5);
    const continuation = slice.find((entry) => entry.rowIndex === 2 && !entry.repeated);
    // Two entries, not one: without the stand-in the row's remaining cell
    // would lay out in column 0 and shift the whole row a column left.
    expect(continuation?.cells.map((entry) => [entry.colIndex, entry.orphaned])).toEqual([
      [0, true],
      [1, false],
    ]);
    // The stand-in is empty and its span stops where the slice does.
    expect(continuation?.cells[0].cell.blocks).toEqual([]);
    expect(continuation?.cells[0].rowSpan).toBe(2);
  });

  it("clips a header cell merged into the body at the repeated run, not the continuation", () => {
    const merged = mergeDown(table(2, 6), 0, 0, 3); // header rows 0-1 plus body row 2
    const slice = tableSliceRows(merged, 4, 7);
    const repeated = slice.filter((entry) => entry.repeated);
    expect(repeated).toHaveLength(2);
    expect(repeated[0].cells[0].rowSpan).toBe(2);
  });
});

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

describe("measureTableElement", () => {
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

  it("never splits inside a vertical merge", () => {
    // Merges sitting across where an unmerged table would have broken: the
    // first natural split is row 10, the second row 19.
    const merged = mergeDown(mergeDown(table(1, 24), 8, 0, 3), 18, 1, 3);
    const crossings = verticalMergeCrossings(merged);
    const pages = paginate(fixture(merged), fakeMetrics);
    expect(pages.length).toBeGreaterThan(2);
    let expectedStart = 0;
    for (const page of pages) {
      const item = page.items[0];
      const start = item.startLine ?? 0;
      const end = item.endLine ?? merged.rows.length;
      expect(crossings[start] ?? false).toBe(false);
      expect(crossings[end] ?? false).toBe(false);
      // And the slices still tile the table exactly once, in order.
      expect(start).toBe(expectedStart);
      expect(end).toBeGreaterThan(start);
      expectedStart = end;
    }
    expect(expectedStart).toBe(merged.rows.length);
  });

  it("places a table whose every interior boundary is merged as one whole block", () => {
    // One cell spanning every body row: there is nowhere legal to split, so
    // the table stays atomic and overflows rather than tearing the merge.
    const welded = mergeDown(table(1, 24), 1, 0, 24);
    const pages = paginate(fixture(welded), fakeMetrics);
    expect(pages).toHaveLength(1);
    expect(pages[0].items).toHaveLength(1);
    expect(pages[0].items[0].endLine ?? welded.rows.length).toBe(25);
  });

  it("draws no ghost continuation when the forced first slice swallows the table", () => {
    // A header row plus one body row taller than the whole page. The first
    // slice must clear the header run, which here means taking the entire
    // table: there is nothing left for a second page to draw.
    const pages = paginate(fixture(table(1, 1)), rowMetrics([ROW, PAGE_ROWS * ROW * 3]));
    expect(pages).toHaveLength(1);
    expect(pages[0].items).toHaveLength(1);
    expect(pages[0].items[0].startLine ?? 0).toBe(0);
    expect(pages[0].items[0].endLine ?? 2).toBe(2);
  });

  it("still continues onto a second page when a forced slice leaves rows behind", () => {
    // The same forced first slice, but with a row after it: the continuation
    // is real, and carries the repeated header above it.
    const pages = paginate(fixture(table(1, 2)), rowMetrics([ROW, PAGE_ROWS * ROW * 3, ROW]));
    expect(pages).toHaveLength(2);
    expect(pages[0].items[0].endLine).toBe(2);
    expect(pages[1].items[0].startLine).toBe(2);
    expect(pages[1].items[0].endLine).toBe(3);
  });
});

function emptyModel(): DocumentModel {
  const model = createDocument("insert-table-fixture");
  model.sections = [{ blocks: [{ type: "paragraph", runs: [] }] }];
  return model;
}

describe("insertTable's header-row option (SPEC-076 AC-20, AC-21)", () => {
  const at = caretSelection({ section: 0, block: 0, offset: 0 });

  it("marks the first row isHeader when the dialog's checkbox is on (AC-20)", () => {
    const next = insertTable(emptyModel(), at, 3, 2, true).model;
    const inserted = next.sections[0].blocks.find((block) => block.type === "table") as Table;
    expect(inserted.rows[0].isHeader).toBe(true);
    expect(inserted.rows[1].isHeader).toBeUndefined();
  });

  it("creates no header row when it is off (AC-20)", () => {
    const next = insertTable(emptyModel(), at, 3, 2, false).model;
    const inserted = next.sections[0].blocks.find((block) => block.type === "table") as Table;
    expect(inserted.rows.every((tableRow) => tableRow.isHeader === undefined)).toBe(true);
  });

  it("creates no header row when the caller passes nothing — the Insert menu's grid quick-pick (AC-21)", () => {
    const next = insertTable(emptyModel(), at, 3, 2).model;
    const inserted = next.sections[0].blocks.find((block) => block.type === "table") as Table;
    expect(inserted.rows.every((tableRow) => tableRow.isHeader === undefined)).toBe(true);
  });
});
