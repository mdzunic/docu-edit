// SPEC-117 "Geometry": the maximum-digit-width table (never a measured
// font), the xlsx character-width ↔ px conversions, and GridMetrics — integer
// px, hidden rows and columns zero-sized, O(log n) per query over sorted runs
// of custom sizes, never an array of MAX_ROWS entries. GridMetrics is checked
// against a naive per-index reference on random sheets, and the 1,048,576-row
// performance budget is asserted here in the unit suite.
import { describe, expect, it } from "vitest";
import {
  baseColWidthToPx,
  buildGridMetrics,
  CELL_PADDING_PX,
  defaultColumnPx,
  defaultRowPx,
  hiddenColumnRuns,
  hiddenRowRuns,
  INDENT_PX_PER_LEVEL,
  maxDigitWidthPx,
  ptToPx,
  pxToXlsxWidth,
  xlsxWidthToPx,
  type GridMetrics,
} from "./geometry";
import { MAX_COLS, MAX_ROWS, type ColumnSpan, type Row, type Sheet } from "./model";

function sheetWith(partial: Partial<Sheet> = {}): Sheet {
  return {
    conditionalFormats: [],
    id: "sheet-1",
    name: "Sheet1",
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    ...partial,
  };
}

/** A small deterministic PRNG (mulberry32) so a failure reproduces. */
function seeded(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomInt(rand: () => number, lo: number, hi: number): number {
  return lo + Math.floor(rand() * (hi - lo + 1));
}

describe("SPEC-117 Geometry: constants", () => {
  it("pads 2 px each side (+ the 1 px gridline = Excel's 5 px) and indents 9 px per level", () => {
    expect(CELL_PADDING_PX).toBe(2);
    expect(INDENT_PX_PER_LEVEL).toBe(9);
  });
});

describe("SPEC-117 Geometry: the MDW table", () => {
  // The spec's test table, row by row.
  const TABLE: Array<[string, number, number]> = [
    ["Calibri", 10, 7],
    ["Calibri", 11, 7],
    ["Calibri", 12, 8],
    ["Aptos Narrow", 11, 7],
    ["Arial", 10, 7],
    ["Arial", 11, 8],
    ["Liberation Sans", 10, 7],
    ["Helvetica", 12, 9],
    ["Times New Roman", 11, 7],
    ["Times New Roman", 12, 8],
    ["Cambria", 11, 8],
    ["Courier New", 10, 8],
    ["Foo", 11, 7],
  ];

  it.each(TABLE)("%s %d pt → %d px", (family, sizePt, mdw) => {
    expect(maxDigitWidthPx(family, sizePt)).toBe(mdw);
  });

  it("matches family names case-insensitively", () => {
    expect(maxDigitWidthPx("CALIBRI", 12)).toBe(8);
    expect(maxDigitWidthPx("calibri", 12)).toBe(8);
    expect(maxDigitWidthPx("aPtOs nArRoW", 11)).toBe(7);
    expect(maxDigitWidthPx("HELVETICA", 12)).toBe(9);
    expect(maxDigitWidthPx("times new roman", 12)).toBe(8);
    expect(maxDigitWidthPx("COURIER NEW", 10)).toBe(8);
  });

  it("gives a metric twin its original's row", () => {
    for (const size of [8, 10, 11, 12, 14, 18, 24]) {
      expect(maxDigitWidthPx("Carlito", size)).toBe(maxDigitWidthPx("Calibri", size));
      expect(maxDigitWidthPx("Arimo", size)).toBe(maxDigitWidthPx("Arial", size));
      expect(maxDigitWidthPx("Liberation Serif", size)).toBe(maxDigitWidthPx("Times New Roman", size));
      expect(maxDigitWidthPx("Tinos", size)).toBe(maxDigitWidthPx("Times New Roman", size));
      expect(maxDigitWidthPx("Caladea", size)).toBe(maxDigitWidthPx("Cambria", size));
      expect(maxDigitWidthPx("Liberation Mono", size)).toBe(maxDigitWidthPx("Courier New", size));
      expect(maxDigitWidthPx("Cousine", size)).toBe(maxDigitWidthPx("Courier New", size));
    }
  });

  it("is round(ratio × sizePt × 96 / 72) and at least 1", () => {
    expect(maxDigitWidthPx("Arial", 20)).toBe(Math.round((0.5562 * 20 * 96) / 72));
    expect(maxDigitWidthPx("Unknown Sans", 30)).toBe(Math.round((0.5068 * 30 * 96) / 72));
    expect(maxDigitWidthPx("Calibri", 1)).toBe(1);
    expect(maxDigitWidthPx("Calibri", 0.5)).toBe(1);
    expect(maxDigitWidthPx("Calibri", 0)).toBe(1);
  });
});

describe("SPEC-117 Geometry: width conversions", () => {
  it("xlsxWidthToPx(9.140625, 7) === 64 and pxToXlsxWidth(64, 7) === 9.140625", () => {
    expect(xlsxWidthToPx(9.140625, 7)).toBe(64);
    expect(pxToXlsxWidth(64, 7)).toBe(9.140625);
  });

  it("applies Excel's formulas exactly", () => {
    // trunc(((256·w + trunc(128/mdw)) / 256) · mdw)
    expect(xlsxWidthToPx(8.43, 7)).toBe(Math.trunc(((256 * 8.43 + 18) / 256) * 7));
    expect(xlsxWidthToPx(0, 7)).toBe(0);
    expect(xlsxWidthToPx(20, 8)).toBe(Math.trunc(((256 * 20 + 16) / 256) * 8));
    // trunc((trunc((px − 5)/mdw · 100 + 0.5)/100 · mdw + 5)/mdw · 256)/256
    expect(pxToXlsxWidth(100, 7)).toBe(14.28125);
    expect(pxToXlsxWidth(20, 7)).toBe(2.8515625);
  });

  it("round-trips every px width through the character width", () => {
    for (const mdw of [6, 7, 8, 9, 11]) {
      for (let px = 1; px <= 1000; px += 1) {
        expect(xlsxWidthToPx(pxToXlsxWidth(px, mdw), mdw)).toBe(px);
      }
    }
  });

  it("a baseColWidth-only sheet defaults to 64 px for Calibri 11", () => {
    const mdw = maxDigitWidthPx("Calibri", 11);
    expect(baseColWidthToPx(8, mdw)).toBe(64);
    // ceil8(trunc(((256·b + trunc(128/mdw)) / 256) · mdw) + 5): always a multiple of 8.
    for (const base of [1, 4, 8, 10, 12, 20]) {
      for (const width of [6, 7, 8, 9]) {
        const px = baseColWidthToPx(base, width);
        expect(px % 8).toBe(0);
        expect(px).toBeGreaterThanOrEqual(Math.trunc(((256 * base + Math.trunc(128 / width)) / 256) * width) + 5);
        expect(px - 8).toBeLessThan(Math.trunc(((256 * base + Math.trunc(128 / width)) / 256) * width) + 5);
      }
    }
  });
});

describe("SPEC-117 Geometry: integer px", () => {
  it("converts heights pt → px rounded (15 pt → 20 px)", () => {
    expect(ptToPx(15)).toBe(20);
    expect(ptToPx(12.75)).toBe(17);
    expect(ptToPx(12.8)).toBe(17);
    expect(ptToPx(0)).toBe(0);
    expect(ptToPx(409.5)).toBe(546);
  });

  it("answers integer px: colWidth = round(widthPx), rowHeight = round(heightPt × 4/3)", () => {
    const metrics = buildGridMetrics(
      sheetWith({
        defaultColWidthPx: 85.33,
        defaultRowHeightPt: 12.8,
        columns: [{ min: 2, max: 2, widthPx: 100.6 }],
        rows: [
          { index: 1, cells: [], heightPt: 15 },
          { index: 3, cells: [], heightPt: 21.75 },
        ],
      }),
    );
    expect(metrics.colWidth(0)).toBe(85);
    expect(metrics.colWidth(2)).toBe(101);
    expect(metrics.colLeft(3)).toBe(85 + 85 + 101);
    expect(metrics.rowHeight(0)).toBe(17);
    expect(metrics.rowHeight(1)).toBe(20);
    expect(metrics.rowHeight(3)).toBe(29);
    expect(metrics.rowTop(4)).toBe(17 + 20 + 17 + 29);
    for (const value of [metrics.colLeft(5), metrics.rowTop(7), metrics.colAt(1000.5), metrics.rowAt(33.3)]) {
      expect(Number.isInteger(value)).toBe(true);
    }
  });

  it("defaults: 64 × 20 px for Calibri 11 / 15 pt; a span without a width uses the sheet default", () => {
    const sheet = sheetWith({ columns: [{ min: 0, max: 3, styleId: 2 }] });
    expect(defaultColumnPx(sheet)).toBe(64);
    expect(defaultRowPx(sheet)).toBe(20);
    const metrics = buildGridMetrics(sheet);
    expect(metrics.colWidth(0)).toBe(64);
    expect(metrics.colLeft(4)).toBe(256);
    expect(metrics.rowHeight(0)).toBe(20);
  });

  it("sizes hidden rows and columns zero, so their neighbours touch", () => {
    const metrics = buildGridMetrics(
      sheetWith({
        columns: [
          { min: 2, max: 3, widthPx: 120, hidden: true },
          { min: 6, max: 6, hidden: true },
        ],
        rows: [
          { index: 3, cells: [], hidden: true },
          { index: 4, cells: [], hidden: true, heightPt: 30 },
          { index: 5, cells: [], hidden: true },
        ],
      }),
    );
    expect(metrics.colWidth(2)).toBe(0);
    expect(metrics.colWidth(3)).toBe(0);
    expect(metrics.colWidth(6)).toBe(0);
    expect(metrics.colLeft(2)).toBe(128);
    expect(metrics.colLeft(4)).toBe(128);
    expect(metrics.colLeft(7)).toBe(6 * 64 - 2 * 64);
    expect(metrics.colAt(128)).toBe(4); // never a zero-sized column
    expect(metrics.rowHeight(4)).toBe(0);
    expect(metrics.rowTop(3)).toBe(60);
    expect(metrics.rowTop(6)).toBe(60);
    expect(metrics.rowAt(60)).toBe(6);
    expect(metrics.rowAt(59)).toBe(2);
  });

  it("clamps hit tests to the grid", () => {
    const metrics = buildGridMetrics(sheetWith());
    expect(metrics.colAt(-50)).toBe(0);
    expect(metrics.rowAt(-1)).toBe(0);
    expect(metrics.colAt(metrics.totalWidth(MAX_COLS) + 10)).toBe(MAX_COLS - 1);
    expect(metrics.rowAt(metrics.totalHeight(MAX_ROWS))).toBe(MAX_ROWS - 1);
  });

  it("totalWidth/totalHeight are the extent of the first n columns/rows", () => {
    const metrics = buildGridMetrics(
      sheetWith({ columns: [{ min: 1, max: 1, widthPx: 100 }], rows: [{ index: 0, cells: [], heightPt: 30 }] }),
    );
    expect(metrics.totalWidth(0)).toBe(0);
    expect(metrics.totalWidth(1)).toBe(64);
    expect(metrics.totalWidth(2)).toBe(164);
    expect(metrics.totalWidth(MAX_COLS)).toBe(MAX_COLS * 64 + 36);
    expect(metrics.totalHeight(1)).toBe(40);
    expect(metrics.totalHeight(MAX_ROWS)).toBe(MAX_ROWS * 20 + 20);
    expect(buildGridMetrics(sheetWith()).totalHeight(MAX_ROWS)).toBe(20_971_520);
  });
});

// --- a naive reference -------------------------------------------------------

/** One axis as per-index arrays: every item's size and the prefix sums. */
interface NaiveAxis {
  sizes: Float64Array;
  offsets: Float64Array; // offsets[i] = Σ sizes[0..i−1]; length count + 1
}

function naiveAxis(count: number, defaultSize: number, assign: (sizes: Float64Array) => void): NaiveAxis {
  const sizes = new Float64Array(count).fill(defaultSize);
  assign(sizes);
  const offsets = new Float64Array(count + 1);
  for (let i = 0; i < count; i += 1) {
    offsets[i + 1] = offsets[i]! + sizes[i]!;
  }
  return { sizes, offsets };
}

/** The item holding `pos`: the last index whose leading edge is at or
 * before it (so never a zero-sized one), clamped to the axis. */
function naiveAt(axis: NaiveAxis, pos: number): number {
  const count = axis.sizes.length;
  const at = Math.max(0, pos);
  let lo = 0;
  let hi = count;
  while (lo < hi) {
    const mid = (lo + hi + 1) >>> 1;
    if (axis.offsets[mid]! <= at) {
      lo = mid;
    } else {
      hi = mid - 1;
    }
  }
  return Math.min(lo, count - 1);
}

/** The spec's sizing rules, index by index. */
function naiveMetrics(sheet: Sheet): { cols: NaiveAxis; rows: NaiveAxis } {
  const colDefault = Math.round(sheet.defaultColWidthPx);
  const rowDefault = Math.round((sheet.defaultRowHeightPt * 4) / 3);
  const cols = naiveAxis(MAX_COLS, colDefault, (sizes) => {
    for (const span of sheet.columns) {
      for (let c = span.min; c <= span.max; c += 1) {
        sizes[c] = span.hidden ? 0 : Math.round(span.widthPx ?? sheet.defaultColWidthPx);
      }
    }
  });
  const rows = naiveAxis(MAX_ROWS, rowDefault, (sizes) => {
    for (const row of sheet.rows) {
      if (row.hidden) {
        sizes[row.index] = 0;
      } else if (row.heightPt !== undefined) {
        sizes[row.index] = Math.round((row.heightPt * 4) / 3);
      }
    }
  });
  return { cols, rows };
}

/** A random sheet: column spans with widths, hidden runs and style-only
 * spans; rows in clusters (adjacent equal sizes merge into one run), hidden
 * runs, rows with cells only, rows at the grid's edges. */
function randomSheet(seed: number): Sheet {
  const rand = seeded(seed);
  const pick = <T,>(items: readonly T[]): T => items[Math.floor(rand() * items.length)]!;
  const defaultColWidthPx = pick([64, 72, 85.33, 48.6, 100]);
  const defaultRowHeightPt = pick([15, 12.8, 20.25, 14.4]);
  const columns: ColumnSpan[] = [];
  let c = randomInt(rand, 0, 3);
  while (c < MAX_COLS) {
    const length = rand() < 0.1 ? randomInt(rand, 200, 3000) : randomInt(rand, 1, 12);
    const max = Math.min(MAX_COLS - 1, c + length - 1);
    const roll = rand();
    const span: ColumnSpan = { min: c, max };
    if (roll < 0.25) {
      span.hidden = true;
      if (rand() < 0.5) {
        span.widthPx = randomInt(rand, 10, 200);
      }
    } else if (roll < 0.4) {
      span.styleId = randomInt(rand, 1, 9);
    } else if (roll < 0.5) {
      span.widthPx = defaultColWidthPx; // equal to the default: no difference
    } else {
      span.widthPx = randomInt(rand, 1, 4000) / 10;
    }
    columns.push(span);
    c = max + 1 + (rand() < 0.3 ? 0 : randomInt(rand, 1, 400));
  }
  const rows: Row[] = [];
  let r = rand() < 0.5 ? 0 : randomInt(rand, 1, 50);
  while (r < MAX_ROWS) {
    const cluster = rand() < 0.2 ? randomInt(rand, 20, 400) : randomInt(rand, 1, 4);
    const roll = rand();
    const heightPt = randomInt(rand, 0, 1640) / 4;
    for (let k = 0; k < cluster && r < MAX_ROWS; k += 1, r += 1) {
      const row: Row = { index: r, cells: [] };
      if (roll < 0.2) {
        row.hidden = true;
        if (rand() < 0.3) {
          row.heightPt = heightPt;
        }
      } else if (roll < 0.3) {
        row.cells = [{ col: 0, value: r }]; // no size property
      } else if (roll < 0.4) {
        row.heightPt = defaultRowHeightPt;
      } else {
        row.heightPt = rand() < 0.7 ? heightPt : randomInt(rand, 0, 1640) / 4;
      }
      rows.push(row);
    }
    r += rand() < 0.2 ? 0 : randomInt(rand, 1, 20_000);
  }
  if (seed % 2 === 0 && rows.at(-1)!.index < MAX_ROWS - 1) {
    rows.push({ index: MAX_ROWS - 1, cells: [], hidden: true });
  }
  return sheetWith({ defaultColWidthPx, defaultRowHeightPt, columns, rows });
}

/** Every leading edge ±1 and the inside of every item: the positions where
 * an off-by-one in a hit test shows. */
function probePositions(axis: NaiveAxis, index: number): number[] {
  const start = axis.offsets[index]!;
  const size = axis.sizes[index]!;
  return [start - 1, start, start + 1, start + size - 1, start + size, start + Math.floor(size / 2)];
}

function expectAxisMatches(
  naive: NaiveAxis,
  start: (i: number) => number,
  size: (i: number) => number,
  at: (pos: number) => number,
  total: (n: number) => number,
  indices: Iterable<number>,
): void {
  const count = naive.sizes.length;
  for (const i of indices) {
    if (start(i) !== naive.offsets[i]) {
      expect.fail(`start(${i}) = ${start(i)}, reference ${naive.offsets[i]}`);
    }
    if (total(i) !== naive.offsets[i]) {
      expect.fail(`total(${i}) = ${total(i)}, reference ${naive.offsets[i]}`);
    }
    if (i >= count) {
      continue;
    }
    if (size(i) !== naive.sizes[i]) {
      expect.fail(`size(${i}) = ${size(i)}, reference ${naive.sizes[i]}`);
    }
    for (const pos of probePositions(naive, i)) {
      if (at(pos) !== naiveAt(naive, pos)) {
        expect.fail(`at(${pos}) = ${at(pos)}, reference ${naiveAt(naive, pos)} (near item ${i})`);
      }
    }
  }
}

describe("SPEC-117 Geometry: GridMetrics against a naive per-index reference", () => {
  for (const seed of [1, 2, 3, 4, 5, 6]) {
    it(`random sheet #${seed}: every column and a broad sample of rows agree`, () => {
      const sheet = randomSheet(seed);
      const metrics: GridMetrics = buildGridMetrics(sheet);
      const naive = naiveMetrics(sheet);
      const rand = seeded(seed * 7919);

      // Every column (and the edge one past the last).
      const colIndices = Array.from({ length: MAX_COLS + 1 }, (_, i) => i);
      expectAxisMatches(
        naive.cols,
        (c) => metrics.colLeft(c),
        (c) => metrics.colWidth(c),
        (x) => metrics.colAt(x),
        (n) => metrics.totalWidth(n),
        colIndices,
      );

      // Rows: around every stored row, the grid's edges and random rows.
      const rowIndices = new Set<number>([0, 1, MAX_ROWS - 2, MAX_ROWS - 1, MAX_ROWS]);
      for (const row of sheet.rows) {
        for (const i of [row.index - 1, row.index, row.index + 1]) {
          if (i >= 0 && i <= MAX_ROWS) {
            rowIndices.add(i);
          }
        }
      }
      for (let k = 0; k < 3000; k += 1) {
        rowIndices.add(Math.floor(rand() * MAX_ROWS));
      }
      expectAxisMatches(
        naive.rows,
        (r) => metrics.rowTop(r),
        (r) => metrics.rowHeight(r),
        (y) => metrics.rowAt(y),
        (n) => metrics.totalHeight(n),
        rowIndices,
      );

      // Random hit tests anywhere, including past either end.
      const width = naive.cols.offsets[MAX_COLS]!;
      const height = naive.rows.offsets[MAX_ROWS]!;
      for (let k = 0; k < 3000; k += 1) {
        const x = Math.floor(rand() * (width + 200)) - 100;
        const y = Math.floor(rand() * (height + 200)) - 100;
        expect(metrics.colAt(x)).toBe(naiveAt(naive.cols, x));
        expect(metrics.rowAt(y)).toBe(naiveAt(naive.rows, y));
      }
      expect(metrics.totalWidth(MAX_COLS)).toBe(width);
      expect(metrics.totalHeight(MAX_ROWS)).toBe(height);
    });
  }

  it("lands a hit test inside a zero-size run on the first visible item after it", () => {
    const sheet = sheetWith({
      columns: [
        { min: 0, max: 1, hidden: true },
        { min: 2, max: 2, widthPx: 30 },
        { min: 5, max: 9, hidden: true },
        { min: 10, max: 10, widthPx: 50 },
      ],
      rows: [
        { index: 0, cells: [], hidden: true },
        { index: 1, cells: [], hidden: true },
        { index: 2, cells: [], heightPt: 30 },
        { index: 4, cells: [], hidden: true },
        { index: 5, cells: [], hidden: true },
      ],
    });
    const metrics = buildGridMetrics(sheet);
    expect(metrics.colAt(0)).toBe(2);
    expect(metrics.colLeft(5)).toBe(30 + 64 + 64);
    expect(metrics.colAt(metrics.colLeft(5))).toBe(10);
    expect(metrics.colAt(metrics.colLeft(5) - 1)).toBe(4);
    expect(metrics.rowAt(0)).toBe(2);
    expect(metrics.rowAt(metrics.rowTop(4))).toBe(6);
    expect(metrics.rowAt(metrics.rowTop(4) - 1)).toBe(3);
  });
});

describe("SPEC-117 Geometry: the O(log n) budget", () => {
  it("a 1,048,576-row sheet with 10,000 custom heights answers 100,000 random rowAt/rowTop queries in under 50 ms", () => {
    const rand = seeded(117);
    const indices = new Set<number>();
    while (indices.size < 10_000) {
      indices.add(Math.floor(rand() * MAX_ROWS));
    }
    const rows: Row[] = [...indices]
      .toSorted((a, b) => a - b)
      .map((index) => (rand() < 0.1 ? { index, cells: [], hidden: true } : { index, cells: [], heightPt: randomInt(rand, 4, 160) / 4 }));
    const metrics = buildGridMetrics(sheetWith({ rows }));
    const total = metrics.totalHeight(MAX_ROWS);
    const QUERIES = 50_000; // of each kind: 100,000 in all
    const ys = Float64Array.from({ length: QUERIES }, () => Math.floor(rand() * total));
    const rs = Int32Array.from({ length: QUERIES }, () => Math.floor(rand() * MAX_ROWS));

    let checksum = 0;
    const started = performance.now();
    for (let i = 0; i < QUERIES; i += 1) {
      checksum += metrics.rowAt(ys[i]!);
      checksum += metrics.rowTop(rs[i]!);
    }
    const elapsed = performance.now() - started;

    expect(checksum).toBeGreaterThan(0);
    expect(elapsed).toBeLessThan(50);
    // …and the answers are right: a sample against the naive reference.
    const naive = naiveMetrics(sheetWith({ rows }));
    for (let i = 0; i < 2000; i += 1) {
      expect(metrics.rowAt(ys[i]!)).toBe(naiveAt(naive.rows, ys[i]!));
      expect(metrics.rowTop(rs[i]!)).toBe(naive.rows.offsets[rs[i]!]);
    }
  });

  it("never allocates per row: the build is O(custom runs), however tall the grid", () => {
    // A build that allocated (or zeroed) an array of MAX_ROWS entries would
    // cost milliseconds each — a thousand would take seconds, not this.
    const sparse = sheetWith({
      rows: [
        { index: 5, cells: [], heightPt: 30 },
        { index: 900_000, cells: [], hidden: true },
        { index: MAX_ROWS - 1, cells: [], heightPt: 45 },
      ],
      columns: [{ min: MAX_COLS - 1, max: MAX_COLS - 1, widthPx: 10 }],
    });
    let started = performance.now();
    for (let i = 0; i < 1000; i += 1) {
      buildGridMetrics(sparse);
    }
    expect(performance.now() - started).toBeLessThan(150);

    // Rows with no size property are skipped in the one pass.
    const valuesOnly: Row[] = Array.from({ length: 100_000 }, (_, index) => ({ index, cells: [{ col: 0, value: index }] }));
    started = performance.now();
    const metrics = buildGridMetrics(sheetWith({ rows: valuesOnly }));
    expect(performance.now() - started).toBeLessThan(150);
    expect(metrics.rowTop(MAX_ROWS - 1)).toBe((MAX_ROWS - 1) * 20);

    const edge = buildGridMetrics(sparse);
    expect(edge.rowHeight(MAX_ROWS - 1)).toBe(60);
    expect(edge.rowTop(MAX_ROWS - 1)).toBe((MAX_ROWS - 1) * 20 + 20 - 20);
    expect(edge.colWidth(MAX_COLS - 1)).toBe(10);
  });
});

describe("SPEC-117 Frozen panes and hidden rows and columns: hidden runs for the header markers", () => {
  it("hiddenColumnRuns: the runs touching [first − 1, last + 1], adjacent spans joined", () => {
    const sheet = sheetWith({
      columns: [
        { min: 2, max: 3, hidden: true }, // C–D
        { min: 4, max: 4, widthPx: 100 },
        { min: 6, max: 6, hidden: true }, // G
        { min: 10, max: 11, hidden: true }, // K–L
        { min: 12, max: 13, hidden: true }, // M–N: joins K–L
        { min: 40, max: 40, hidden: true },
      ],
    });
    expect(hiddenColumnRuns(sheet, 0, 20)).toEqual([
      { start: 2, end: 3 },
      { start: 6, end: 6 },
      { start: 10, end: 13 },
    ]);
    // A run ending just left of the window still marks its left boundary.
    expect(hiddenColumnRuns(sheet, 4, 9)).toEqual([
      { start: 2, end: 3 },
      { start: 6, end: 6 },
      { start: 10, end: 13 },
    ]);
    expect(hiddenColumnRuns(sheet, 5, 8)).toEqual([{ start: 6, end: 6 }]);
    // A joined run is reported whole from either side: the window touches
    // only M–N (or only K–L), the marker still names K–N.
    expect(hiddenColumnRuns(sheet, 14, 20)).toEqual([{ start: 10, end: 13 }]);
    expect(hiddenColumnRuns(sheet, 7, 9)).toEqual([
      { start: 6, end: 6 },
      { start: 10, end: 13 },
    ]);
    expect(hiddenColumnRuns(sheet, 15, 39)).toEqual([{ start: 40, end: 40 }]);
    expect(hiddenColumnRuns(sheet, 15, 30)).toEqual([]);
    expect(hiddenColumnRuns(sheetWith(), 0, 100)).toEqual([]);
  });

  it("hiddenRowRuns: runs reaching in from above and continuing below the window are whole", () => {
    const rows: Row[] = [
      { index: 1, cells: [{ col: 0, value: 1 }] },
      { index: 3, cells: [], hidden: true },
      { index: 4, cells: [], hidden: true },
      { index: 5, cells: [], hidden: true },
      { index: 6, cells: [], heightPt: 30 },
      { index: 10, cells: [], hidden: true },
      { index: 12, cells: [], hidden: true }, // not adjacent to 10
      ...Array.from({ length: 8 }, (_, k): Row => ({ index: 18 + k, cells: [], hidden: true })), // 18–25
    ];
    const sheet = sheetWith({ rows });
    expect(hiddenRowRuns(sheet, 0, 30)).toEqual([
      { start: 3, end: 5 },
      { start: 10, end: 10 },
      { start: 12, end: 12 },
      { start: 18, end: 25 },
    ]);
    // Window starting inside the 3–5 run: the run is reported from its start.
    expect(hiddenRowRuns(sheet, 5, 8)).toEqual([{ start: 3, end: 5 }]);
    // Window ending inside the 18–25 run: the run is reported to its end.
    expect(hiddenRowRuns(sheet, 14, 19)).toEqual([{ start: 18, end: 25 }]);
    expect(hiddenRowRuns(sheet, 7, 8)).toEqual([]);
    // A run ending just above the window (at first − 1) still marks its top
    // boundary, reported whole; one ending further up does not.
    expect(hiddenRowRuns(sheet, 26, 40)).toEqual([{ start: 18, end: 25 }]);
    expect(hiddenRowRuns(sheet, 27, 40)).toEqual([]);
    expect(hiddenRowRuns(sheetWith(), 0, 100)).toEqual([]);
  });
});
