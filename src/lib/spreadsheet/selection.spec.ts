// SPEC-117 "Selection" and "Keyboard map": the pure selection model. Every
// rule is pinned here — `active` inside its range, visible and a merge's
// top-left; ranges merge-snapped; moving collapses, leaves a merge from its
// far edge and lands on a merge's top-left; hidden rows and columns are
// skipped; the grid's edges clamp; extending moves the corner farthest from
// the anchor; Mod+arrow finds the data-region edge; Mod+A takes the current
// region, then the whole sheet; Tab/Enter cycle inside a multi-cell
// selection — together with the spec's worked examples verbatim, the whole
// keyboard map (null for keys it does not own), the pointer helpers, the
// validation of a stored selection, and a sparse-walk performance bound on a
// 1,000,000-row sheet.
import { describe, expect, it } from "vitest";
import { formatA1, formatRange, parseA1, parseRange, type CellRef, type GridRange } from "./address";
import { buildGridMetrics } from "./geometry";
import { buildMergeIndex } from "./merges";
import { MAX_COLS, MAX_ROWS, type Cell, type CellValue, type ColumnSpan, type Row, type Sheet } from "./model";
import {
  addRange,
  currentRegion,
  cycleWithin,
  edgeInDirection,
  extendBy,
  extendTo,
  extendToLine,
  headerSelection,
  isCellSelected,
  moveActive,
  navigateByKey,
  normaliseSelection,
  rangeContains,
  selectAll,
  selectCell,
  selectColumns,
  selectRows,
  singleCell,
  snapToMerges,
  wholeColumns,
  wholeRows,
  type NavContext,
  type NavKey,
  type SheetSelection,
} from "./selection";

const LAST_ROW = MAX_ROWS - 1;
const LAST_COL = MAX_COLS - 1;

function ref(a1: string): CellRef {
  const parsed = parseA1(a1);
  if (!parsed) {
    throw new Error(`bad reference ${a1}`);
  }
  return { row: parsed.row, col: parsed.col };
}

function range(text: string): GridRange {
  const parsed = parseRange(text);
  if (!parsed) {
    throw new Error(`bad range ${text}`);
  }
  return parsed;
}

interface Layout {
  cells?: Record<string, CellValue>;
  merges?: string[];
  /** 1-based row numbers. */
  hiddenRows?: number[];
  /** Column letters. */
  hiddenCols?: string[];
}

/** A grid sheet: the cells, merges and hidden rows/columns given. */
function sheetOf(layout: Layout = {}): Sheet {
  const byRow = new Map<number, Cell[]>();
  for (const [a1, value] of Object.entries(layout.cells ?? {})) {
    const at = ref(a1);
    byRow.set(at.row, [...(byRow.get(at.row) ?? []), { col: at.col, value }]);
  }
  const hiddenRows = new Set((layout.hiddenRows ?? []).map((n) => n - 1));
  for (const index of hiddenRows) {
    byRow.set(index, byRow.get(index) ?? []);
  }
  const rows: Row[] = [...byRow.entries()]
    .toSorted(([a], [b]) => a - b)
    .map(([index, cells]): Row => ({ index, cells: cells.toSorted((a, b) => a.col - b.col), ...(hiddenRows.has(index) ? { hidden: true } : {}) }));
  const columns: ColumnSpan[] = (layout.hiddenCols ?? [])
    .map((letters): ColumnSpan => {
      const col = ref(`${letters}1`).col;
      return { min: col, max: col, hidden: true };
    })
    .toSorted((a, b) => a.min - b.min);
  return {
    id: "sheet-1",
    name: "Sheet1",
    visibility: "visible",
    kind: "grid",
    rows,
    columns,
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: (layout.merges ?? []).map(range),
  };
}

/** The NavContext the grid builds: hidden = zero-sized in GridMetrics. */
function navOf(sheet: Sheet, page = { rows: 10, cols: 5 }): NavContext {
  const metrics = buildGridMetrics(sheet);
  return {
    sheet,
    merges: buildMergeIndex(sheet.merges),
    isRowHidden: (r) => metrics.rowHeight(r) === 0,
    isColHidden: (c) => metrics.colWidth(c) === 0,
    pageRows: page.rows,
    pageCols: page.cols,
  };
}

/** Numbers in every cell of the given ranges ("A1:C3", "E1"). */
function fill(...ranges: string[]): Record<string, CellValue> {
  const cells: Record<string, CellValue> = {};
  for (const text of ranges) {
    const r = range(text);
    for (let row = r.top; row <= r.bottom; row += 1) {
      for (let col = r.left; col <= r.right; col += 1) {
        cells[formatA1({ row, col })] = row * 100 + col;
      }
    }
  }
  return cells;
}

/** A selection from A1 text: `active`, its ranges, the active range's index. */
function sel(active: string, ranges: string[] = [active], activeRangeIndex = 0): SheetSelection {
  return { active: ref(active), ranges: ranges.map(range), activeRangeIndex };
}

/** "B6 [B5:E6]": the active cell, then every range with the active one
 * bracketed. */
function show(selection: SheetSelection | null): string {
  if (!selection) {
    return "null";
  }
  const ranges = selection.ranges.map((r, i) => (i === selection.activeRangeIndex ? `[${formatRange(r)}]` : formatRange(r)));
  return `${formatA1(selection.active)} ${ranges.join(" ")}`;
}

function key(name: string, mods: Partial<Omit<NavKey, "key">> = {}): NavKey {
  return { key: name, shift: false, mod: false, alt: false, ctrl: false, ...mods };
}

/** Windows/Linux: Ctrl is the primary modifier, so both flags are set. */
const MOD = { mod: true, ctrl: true } as const;

/** The best of three timed runs of `walk` (after one untimed run) — immune
 * to a GC pause or a busy machine. */
function fastest(walk: () => CellRef): { result: CellRef; ms: number } {
  let best = { result: walk(), ms: Number.POSITIVE_INFINITY };
  for (let run = 0; run < 3; run += 1) {
    const started = performance.now();
    const result = walk();
    best = { result, ms: Math.min(best.ms, performance.now() - started) };
  }
  return best;
}

/** Presses keys in turn, failing on a key the map does not handle. */
function keys(start: SheetSelection, nav: NavContext, ...presses: NavKey[]): SheetSelection {
  let current = start;
  for (const press of presses) {
    const next = navigateByKey(current, press, nav);
    if (!next) {
      throw new Error(`unhandled ${press.key}`);
    }
    current = next;
  }
  return current;
}

describe("the spec's worked examples", () => {
  it("in {A1..C3 filled, E1 filled}, Ctrl+Right from A1 → C1, again → E1, again → XFD1", () => {
    const nav = navOf(sheetOf({ cells: fill("A1:C3", "E1") }));
    const right = key("ArrowRight", MOD);
    const first = keys(selectCell(ref("A1"), nav), nav, right);
    expect(show(first)).toBe("C1 [C1]");
    const second = keys(first, nav, right);
    expect(show(second)).toBe("E1 [E1]");
    expect(show(keys(second, nav, right))).toBe("XFD1 [XFD1]");
    expect(edgeInDirection(ref("A1"), "right", nav)).toEqual(ref("C1"));
    expect(edgeInDirection(ref("C1"), "right", nav)).toEqual(ref("E1"));
    expect(edgeInDirection(ref("E1"), "right", nav)).toEqual(ref("XFD1"));
  });

  it("Ctrl+A from B2 → A1:C3, again → A1:XFD1048576", () => {
    const nav = navOf(sheetOf({ cells: fill("A1:C3", "E1") }));
    const once = keys(selectCell(ref("B2"), nav), nav, key("a", MOD));
    expect(show(once)).toBe("B2 [A1:C3]");
    expect(currentRegion(ref("B2"), nav)).toEqual(range("A1:C3"));
    const twice = keys(once, nav, key("a", MOD));
    expect(show(twice)).toBe("A1 [A1:XFD1048576]");
    expect(twice.ranges).toEqual([{ top: 0, left: 0, bottom: LAST_ROW, right: LAST_COL }]);
  });

  it("with C5:E6 merged: Right from B5 → C5, Right again → F5; Shift+Right from B6 → B6:C6 snapped to B5:E6", () => {
    const nav = navOf(sheetOf({ merges: ["C5:E6"] }));
    const into = moveActive(selectCell(ref("B5"), nav), "right", nav);
    expect(show(into)).toBe("C5 [C5:E6]");
    expect(show(moveActive(into, "right", nav))).toBe("F5 [F5]");
    const extended = extendBy(selectCell(ref("B6"), nav), "right", nav);
    expect(show(extended)).toBe("B6 [B5:E6]");
    expect(snapToMerges(range("B6:C6"), nav.merges)).toEqual(range("B5:E6"));
  });
});

describe("snapToMerges", () => {
  it("expands a range until no merge crosses its edge, to a fixed point", () => {
    const merges = buildMergeIndex([range("A1:B2"), range("B3:C4")]);
    // A2:A3 takes in A1:B2 (→ A1:B3), which now crosses B3:C4 (→ A1:C4).
    expect(snapToMerges(range("A2:A3"), merges)).toEqual(range("A1:C4"));
    expect(snapToMerges(range("D1:D9"), merges)).toEqual(range("D1:D9"));
    expect(snapToMerges(range("A1:C4"), merges)).toEqual(range("A1:C4"));
  });

  it("returns a fresh object and never touches its input or the merges", () => {
    const merges = [range("C5:E6")];
    const index = buildMergeIndex(merges);
    const input = range("B6:C6");
    const snapped = snapToMerges(input, index);
    expect(input).toEqual(range("B6:C6"));
    expect(snapped).not.toBe(input);
    expect(snapToMerges(range("C5:E6"), index)).not.toBe(merges[0]);
    snapped.top = 99;
    expect(merges[0]).toEqual(range("C5:E6"));
  });
});

describe("singleCell, selectCell and selectAll", () => {
  it("singleCell is one cell, clamped to the grid", () => {
    expect(show(singleCell(ref("B3")))).toBe("B3 [B3]");
    expect(singleCell({ row: 2_000_000, col: -4 })).toEqual({ active: { row: LAST_ROW, col: 0 }, ranges: [{ top: LAST_ROW, left: 0, bottom: LAST_ROW, right: 0 }], activeRangeIndex: 0 });
  });

  it("selectCell snaps a press inside a merge to the merge, its top-left active", () => {
    const nav = navOf(sheetOf({ merges: ["C5:E6"] }));
    expect(show(selectCell(ref("D6"), nav))).toBe("C5 [C5:E6]");
    expect(show(selectCell({ row: 10_000_000, col: 99_999 }, nav))).toBe("XFD1048576 [XFD1048576]");
  });

  it("selectCell moves a hidden cell to the nearest visible one", () => {
    const nav = navOf(sheetOf({ hiddenRows: [4, 5, 6], hiddenCols: ["B"] }));
    expect(show(selectCell(ref("A4"), nav))).toBe("A3 [A3]");
    expect(show(selectCell(ref("A6"), nav))).toBe("A7 [A7]");
    expect(show(selectCell(ref("B1"), nav))).toBe("C1 [C1]");
  });

  it("selectAll is the whole sheet with A1 active", () => {
    expect(selectAll()).toEqual({ active: { row: 0, col: 0 }, ranges: [{ top: 0, left: 0, bottom: LAST_ROW, right: LAST_COL }], activeRangeIndex: 0 });
    expect(selectAll()).not.toBe(selectAll());
  });
});

describe("moveActive", () => {
  const nav = navOf(sheetOf());

  it("moves one cell in each direction and collapses the selection to it", () => {
    expect(show(moveActive(sel("B2"), "down", nav))).toBe("B3 [B3]");
    expect(show(moveActive(sel("B2"), "up", nav))).toBe("B1 [B1]");
    expect(show(moveActive(sel("B2"), "left", nav))).toBe("A2 [A2]");
    expect(show(moveActive(sel("B2"), "right", nav))).toBe("C2 [C2]");
    expect(show(moveActive(sel("B2", ["A1:C3", "E5"], 0), "down", nav))).toBe("B3 [B3]");
  });

  it("clamps at the grid's edges — row 1 and 1048576, columns A and XFD", () => {
    expect(show(moveActive(sel("A1"), "up", nav))).toBe("A1 [A1]");
    expect(show(moveActive(sel("A1"), "left", nav))).toBe("A1 [A1]");
    expect(show(moveActive(sel("B1048576"), "down", nav))).toBe("B1048576 [B1048576]");
    expect(show(moveActive(sel("XFD5"), "right", nav))).toBe("XFD5 [XFD5]");
    expect(show(moveActive(sel("XFD1048576", ["A1:XFD1048576"]), "down", nav))).toBe("XFD1048576 [XFD1048576]");
  });

  it("skips hidden rows and columns", () => {
    const hidden = navOf(sheetOf({ hiddenRows: [3, 4], hiddenCols: ["C", "D"] }));
    expect(show(moveActive(sel("B2"), "down", hidden))).toBe("B5 [B5]");
    expect(show(moveActive(sel("B5"), "up", hidden))).toBe("B2 [B2]");
    expect(show(moveActive(sel("B2"), "right", hidden))).toBe("E2 [E2]");
    expect(show(moveActive(sel("E2"), "left", hidden))).toBe("B2 [B2]");
  });

  it("stops at the last visible row or column when the grid's edge is hidden", () => {
    const hidden = navOf(sheetOf({ hiddenRows: [1_048_575, 1_048_576], hiddenCols: ["A"] }));
    expect(show(moveActive(sel("C1048574"), "down", hidden))).toBe("C1048574 [C1048574]");
    expect(show(moveActive(sel("B3"), "left", hidden))).toBe("B3 [B3]");
  });

  it("leaves a merge from its far edge and enters one at its top-left", () => {
    const merged = navOf(sheetOf({ merges: ["C5:E6"] }));
    const merge = selectCell(ref("C5"), merged);
    expect(show(moveActive(merge, "down", merged))).toBe("C7 [C7]");
    expect(show(moveActive(merge, "up", merged))).toBe("C4 [C4]");
    expect(show(moveActive(merge, "left", merged))).toBe("B5 [B5]");
    expect(show(moveActive(merge, "right", merged))).toBe("F5 [F5]");
    expect(show(moveActive(sel("D7"), "up", merged))).toBe("C5 [C5:E6]");
    expect(show(moveActive(sel("F6"), "left", merged))).toBe("C5 [C5:E6]");
  });

  it("takes `steps` visible steps (a merge is one), stopping at the edge", () => {
    const merged = navOf(sheetOf({ merges: ["C5:E6"], hiddenRows: [3] }));
    expect(show(moveActive(sel("A1"), "down", merged, 5))).toBe("A7 [A7]");
    expect(show(moveActive(sel("B5"), "right", merged, 2))).toBe("F5 [F5]");
    expect(show(moveActive(sel("A4"), "up", merged, 10))).toBe("A1 [A1]");
  });
});

describe("extendBy", () => {
  const nav = navOf(sheetOf());

  it("keeps the anchor and moves the far edge, growing and shrinking through it", () => {
    let s = extendBy(sel("B2"), "down", nav, 2);
    expect(show(s)).toBe("B2 [B2:B4]");
    s = extendBy(s, "up", nav);
    expect(show(s)).toBe("B2 [B2:B3]");
    s = extendBy(s, "up", nav, 2);
    expect(show(s)).toBe("B2 [B1:B2]");
    expect(show(extendBy(sel("B2"), "left", nav))).toBe("B2 [A2:B2]");
  });

  it("moves the corner farthest from the anchor on each axis", () => {
    // Anchor at the bottom-right: Right and Down move the left and top edges.
    expect(show(extendBy(sel("C3", ["A1:C3"]), "right", nav))).toBe("C3 [B1:C3]");
    expect(show(extendBy(sel("C3", ["A1:C3"]), "down", nav))).toBe("C3 [A2:C3]");
    expect(show(extendBy(sel("C3", ["B2:C3"]), "up", nav))).toBe("C3 [B1:C3]");
    // At the grid's top edge the moving (top) edge cannot go further.
    expect(show(extendBy(sel("C3", ["A1:C3"]), "up", nav))).toBe("C3 [A1:C3]");
  });

  it("changes only the active range", () => {
    const s = extendBy(sel("B2", ["E5", "B2"], 1), "right", nav);
    expect(show(s)).toBe("B2 E5 [B2:C2]");
  });

  it("skips hidden rows and columns and clamps at the grid's edges", () => {
    const hidden = navOf(sheetOf({ hiddenRows: [3], hiddenCols: ["C"] }));
    expect(show(extendBy(sel("B2"), "down", hidden))).toBe("B2 [B2:B4]");
    expect(show(extendBy(sel("B2"), "right", hidden))).toBe("B2 [B2:D2]");
    expect(show(extendBy(sel("XFD1"), "right", nav))).toBe("XFD1 [XFD1]");
    expect(show(extendBy(sel("A1048576"), "down", nav))).toBe("A1048576 [A1048576]");
    expect(show(extendBy(sel("A1", ["A1:XFD1048576"]), "down", nav))).toBe("A1 [A1:XFD1048576]");
  });

  it("snaps to merges and never stalls against one", () => {
    const merged = navOf(sheetOf({ merges: ["C5:E6"] }));
    const snapped = extendBy(selectCell(ref("B6"), merged), "right", merged);
    expect(show(snapped)).toBe("B6 [B5:E6]");
    // The snap moved the top edge; Shift+Down still grows the selection.
    expect(show(extendBy(snapped, "down", merged))).toBe("B6 [B5:E7]");
    expect(show(extendBy(snapped, "up", merged))).toBe("B6 [B4:E6]");
    // From the merge itself, Shift+Left takes in column B.
    const merge = selectCell(ref("C5"), merged);
    expect(show(extendBy(merge, "left", merged))).toBe("C5 [B5:E6]");
    expect(show(extendBy(merge, "right", merged))).toBe("C5 [C5:F6]");
    expect(show(extendBy(extendBy(merge, "left", merged), "right", merged))).toBe("C5 [C5:E6]");
  });
});

describe("extendTo", () => {
  const nav = navOf(sheetOf({ merges: ["C5:E6"] }));

  it("spans the anchor and the target, snapped, keeping the active cell and other ranges", () => {
    expect(show(extendTo(sel("B2"), ref("D4"), nav))).toBe("B2 [B2:D4]");
    expect(show(extendTo(sel("B2"), ref("A1"), nav))).toBe("B2 [A1:B2]");
    expect(show(extendTo(sel("B6"), ref("C6"), nav))).toBe("B6 [B5:E6]");
    expect(show(extendTo(selectCell(ref("C5"), nav), ref("B4"), nav))).toBe("C5 [B4:E6]");
    expect(show(extendTo(sel("H1", ["A1", "H1"], 1), ref("H3"), nav))).toBe("H1 A1 [H1:H3]");
    expect(show(extendTo(sel("B2"), { row: 9_999_999, col: 1 }, nav))).toBe("B2 [B2:B1048576]");
  });
});

describe("edgeInDirection", () => {
  it("runs to the end of a filled block, jumps gaps, and stops at the grid's edge", () => {
    const nav = navOf(sheetOf({ cells: fill("B3:B12", "B15") }));
    expect(edgeInDirection(ref("B3"), "down", nav)).toEqual(ref("B12"));
    expect(edgeInDirection(ref("B12"), "down", nav)).toEqual(ref("B15"));
    expect(edgeInDirection(ref("B15"), "down", nav)).toEqual(ref("B1048576"));
    expect(edgeInDirection(ref("B15"), "up", nav)).toEqual(ref("B12"));
    expect(edgeInDirection(ref("B12"), "up", nav)).toEqual(ref("B3"));
    expect(edgeInDirection(ref("B3"), "up", nav)).toEqual(ref("B1"));
    expect(edgeInDirection(ref("B13"), "down", nav)).toEqual(ref("B15"));
    expect(edgeInDirection(ref("B1"), "down", nav)).toEqual(ref("B3"));
    expect(edgeInDirection(ref("B1"), "up", nav)).toEqual(ref("B1"));
    expect(edgeInDirection(ref("B5"), "left", nav)).toEqual(ref("A5"));
    expect(edgeInDirection(ref("B5"), "right", nav)).toEqual(ref("XFD5"));
  });

  it("walks left through {A1..C3, E1}", () => {
    const nav = navOf(sheetOf({ cells: fill("A1:C3", "E1") }));
    expect(edgeInDirection(ref("XFD1"), "left", nav)).toEqual(ref("E1"));
    expect(edgeInDirection(ref("E1"), "left", nav)).toEqual(ref("C1"));
    expect(edgeInDirection(ref("C1"), "left", nav)).toEqual(ref("A1"));
    expect(edgeInDirection(ref("A1"), "left", nav)).toEqual(ref("A1"));
  });

  it("skips hidden rows and columns — a hidden value neither stops nor ends a run", () => {
    const run = navOf(sheetOf({ cells: fill("B3", "B5"), hiddenRows: [4] }));
    expect(edgeInDirection(ref("B3"), "down", run)).toEqual(ref("B5"));
    const jump = navOf(sheetOf({ cells: fill("B3", "B5", "B8"), hiddenRows: [5] }));
    expect(edgeInDirection(ref("B3"), "down", jump)).toEqual(ref("B8"));
    const across = navOf(sheetOf({ cells: fill("A1", "C1", "E1"), hiddenCols: ["B"] }));
    expect(edgeInDirection(ref("A1"), "right", across)).toEqual(ref("C1"));
  });

  it("ends at the last visible row or column when the grid's edge is hidden", () => {
    const nav = navOf(sheetOf({ cells: fill("B3"), hiddenRows: [1_048_576], hiddenCols: ["XFD", "XFC"] }));
    expect(edgeInDirection(ref("B3"), "down", nav)).toEqual(ref("B1048575"));
    expect(edgeInDirection(ref("B3"), "right", nav)).toEqual(ref("XFB3"));
  });

  it("treats a style-only cell (value null) as empty and \"\" as a value", () => {
    const sheet = sheetOf({ cells: { B2: 1, B3: null, B4: null, B6: "" } });
    sheet.rows = sheet.rows.map((row) => ({ ...row, cells: row.cells.map((cell) => (cell.value === null ? { ...cell, styleId: 3 } : cell)) }));
    const nav = navOf(sheet);
    expect(edgeInDirection(ref("B2"), "down", nav)).toEqual(ref("B6"));
  });

  it("treats a merge as one cell holding its top-left's value", () => {
    const nav = navOf(sheetOf({ cells: { A5: "x", B5: "y", C5: "Title", F5: "z" }, merges: ["C5:E5"] }));
    expect(edgeInDirection(ref("A5"), "right", nav)).toEqual(ref("F5"));
    expect(edgeInDirection(ref("G5"), "left", nav)).toEqual(ref("F5"));
    expect(edgeInDirection(ref("F5"), "left", nav)).toEqual(ref("A5"));
    // Down column D meets the merge at D5 and returns its top-left.
    expect(edgeInDirection(ref("D1"), "down", nav)).toEqual(ref("C5"));
    // From the merge, the walk leaves from its far edge.
    expect(edgeInDirection(ref("C5"), "right", nav)).toEqual(ref("F5"));
    expect(edgeInDirection(ref("C5"), "down", nav)).toEqual(ref("C1048576"));
  });

  it("is a sparse walk: an empty column of a 1,000,000-row sheet in under 50 ms", () => {
    const shared: Cell[] = [{ col: 0, value: 1 }];
    const big = navOf({ ...sheetOf(), rows: Array.from({ length: 1_000_000 }, (_, index): Row => ({ index, cells: shared })) });
    const small = navOf(sheetOf({ cells: fill("A1:A2000") }));
    edgeInDirection(ref("B1"), "down", small); // warm the JIT on a small sheet
    edgeInDirection(ref("B2000"), "up", small);
    // Probing each of the 1,048,576 row indices would take several times longer.
    const down = fastest(() => edgeInDirection(ref("B1"), "down", big));
    const up = fastest(() => edgeInDirection({ row: 999_999, col: 1 }, "up", big));
    expect(down.result).toEqual({ row: LAST_ROW, col: 1 });
    expect(up.result).toEqual({ row: 0, col: 1 });
    expect(down.ms).toBeLessThan(50);
    expect(up.ms).toBeLessThan(50);
    // A filled column runs to its last value.
    expect(edgeInDirection(ref("A1"), "down", big)).toEqual({ row: 999_999, col: 0 });
  });
});

describe("currentRegion", () => {
  it("grows across diagonal neighbours and stops at empty rows and columns", () => {
    const nav = navOf(sheetOf({ cells: fill("A1:C3", "E1") }));
    expect(currentRegion(ref("C3"), nav)).toEqual(range("A1:C3"));
    expect(currentRegion(ref("E1"), nav)).toEqual(range("E1"));
    expect(currentRegion(ref("A4"), nav)).toEqual(range("A1:C4"));
    expect(currentRegion(ref("H9"), nav)).toEqual(range("H9"));
    const diagonal = navOf(sheetOf({ cells: fill("A1", "B2", "C3") }));
    expect(currentRegion(ref("A1"), diagonal)).toEqual(range("A1:C3"));
  });

  it("includes merges whole, growing again from what they reach", () => {
    const nav = navOf(sheetOf({ cells: { ...fill("A2:C4"), A1: "Title" }, merges: ["A1:E1"] }));
    expect(currentRegion(ref("B3"), nav)).toEqual(range("A1:E4"));
  });

  it("is the cell itself on an empty sheet", () => {
    expect(currentRegion(ref("C3"), navOf(sheetOf()))).toEqual(range("C3"));
  });
});

describe("cycleWithin", () => {
  const nav = navOf(sheetOf());

  it("with one cell selected is moveActive right/down (left/up backwards)", () => {
    expect(show(cycleWithin(sel("B2"), "row", false, nav))).toBe("C2 [C2]");
    expect(show(cycleWithin(sel("B2"), "row", true, nav))).toBe("A2 [A2]");
    expect(show(cycleWithin(sel("B2"), "column", false, nav))).toBe("B3 [B3]");
    expect(show(cycleWithin(sel("B2"), "column", true, nav))).toBe("B1 [B1]");
    const merged = navOf(sheetOf({ merges: ["C5:E6"] }));
    expect(show(cycleWithin(selectCell(ref("D5"), merged), "row", false, merged))).toBe("F5 [F5]");
  });

  it("walks a range in row order (Tab) or column order (Enter), wrapping, keeping the ranges", () => {
    const tab = (s: SheetSelection): SheetSelection => cycleWithin(s, "row", false, nav);
    const enter = (s: SheetSelection): SheetSelection => cycleWithin(s, "column", false, nav);
    let s = sel("A1", ["A1:B2"]);
    const order: string[] = [];
    for (let i = 0; i < 4; i += 1) {
      s = tab(s);
      order.push(formatA1(s.active));
    }
    expect(order).toEqual(["B1", "A2", "B2", "A1"]);
    expect(show(s)).toBe("A1 [A1:B2]");
    order.length = 0;
    for (let i = 0; i < 4; i += 1) {
      s = enter(s);
      order.push(formatA1(s.active));
    }
    expect(order).toEqual(["A2", "B1", "B2", "A1"]);
    expect(show(cycleWithin(sel("A1", ["A1:B2"]), "row", true, nav))).toBe("B2 [A1:B2]");
    expect(show(cycleWithin(sel("A1", ["A1:B2"]), "column", true, nav))).toBe("B2 [A1:B2]");
  });

  it("wraps to the next range, and from the last range back to the first", () => {
    const start = sel("B1", ["A1:B1", "D3"], 0);
    const next = cycleWithin(start, "row", false, nav);
    expect(show(next)).toBe("D3 A1:B1 [D3]");
    expect(show(cycleWithin(next, "row", false, nav))).toBe("A1 [A1:B1] D3");
    expect(show(cycleWithin(sel("A1", ["A1:B1", "D3"], 0), "row", true, nav))).toBe("D3 A1:B1 [D3]");
    expect(show(cycleWithin(sel("D3", ["A1:B1", "D3"], 1), "column", false, nav))).toBe("A1 [A1:B1] D3");
  });

  it("skips hidden cells", () => {
    const hidden = navOf(sheetOf({ hiddenCols: ["B"], hiddenRows: [2] }));
    expect(show(cycleWithin(sel("A1", ["A1:C3"]), "row", false, hidden))).toBe("C1 [A1:C3]");
    expect(show(cycleWithin(sel("C1", ["A1:C3"]), "row", false, hidden))).toBe("A3 [A1:C3]");
    expect(show(cycleWithin(sel("A1", ["A1:C3"]), "column", false, hidden))).toBe("A3 [A1:C3]");
  });

  it("skips merge members, meeting a merge at its top-left — backwards in exact reverse", () => {
    const merged = navOf(sheetOf({ merges: ["B1:C2"] }));
    const walk = (order: "row" | "column", back: boolean): string[] => {
      let s = sel("A1", ["A1:D2"]);
      const seen: string[] = [];
      for (let i = 0; i < 5; i += 1) {
        s = cycleWithin(s, order, back, merged);
        seen.push(formatA1(s.active));
      }
      return seen;
    };
    expect(walk("row", false)).toEqual(["B1", "D1", "A2", "D2", "A1"]);
    expect(walk("row", true)).toEqual(["D2", "A2", "D1", "B1", "A1"]);
    expect(walk("column", false)).toEqual(["A2", "B1", "D1", "D2", "A1"]);
    expect(walk("column", true)).toEqual(["D2", "D1", "B1", "A2", "A1"]);
  });
});

describe("addRange, selectRows, selectColumns and headerSelection", () => {
  const nav = navOf(sheetOf({ merges: ["C5:E6", "A9:A10"] }));

  it("addRange appends the pressed cell (its merge) as the new active range", () => {
    const s = addRange(sel("A1", ["A1:B2"]), ref("D6"), nav);
    expect(show(s)).toBe("C5 A1:B2 [C5:E6]");
    expect(show(addRange(s, ref("H8"), nav))).toBe("H8 A1:B2 C5:E6 [H8]");
  });

  it("Shift+Space selects the whole rows of the active range; Ctrl+Space its whole columns", () => {
    expect(show(selectRows(sel("B2", ["B2:C3"]), nav))).toBe("B2 [A2:XFD3]");
    expect(show(selectColumns(sel("G2", ["G2:H3"]), nav))).toBe("G2 [G1:H1048576]");
    // Snapped: columns B:C are crossed by C5:E6.
    expect(show(selectColumns(sel("B2", ["B2:C3"]), nav))).toBe("B2 [B1:E1048576]");
    // Snapped: row 5 alone is crossed by C5:E6.
    expect(show(selectRows(sel("B5"), nav))).toBe("B5 [A5:XFD6]");
    expect(show(selectRows(sel("H1", ["B2", "H1"], 1), nav))).toBe("H1 B2 [A1:XFD1]");
  });

  it("headerSelection selects whole rows or columns, the pressed line's first visible cell active", () => {
    expect(show(headerSelection("row", 3, 3, nav))).toBe("A4 [A4:XFD4]");
    expect(show(headerSelection("row", 3, 1, nav))).toBe("A4 [A2:XFD4]");
    expect(show(headerSelection("column", 6, 6, nav))).toBe("G1 [G1:G1048576]");
    expect(show(headerSelection("column", 6, 5, nav))).toBe("G1 [F1:G1048576]");
    expect(show(headerSelection("column", 2, 2, nav))).toBe("C1 [C1:E1048576]");
    // Snapped to merges; the active cell lands on a merge's top-left.
    expect(show(headerSelection("row", 9, 9, nav))).toBe("A9 [A9:XFD10]");
    const hidden = navOf(sheetOf({ hiddenCols: ["A"], hiddenRows: [1] }));
    expect(show(headerSelection("row", 3, 3, hidden))).toBe("B4 [A4:XFD4]");
    expect(show(headerSelection("column", 0, 2, hidden))).toBe("B2 [A1:C1048576]");
  });

  it("headerSelection with a base (Mod+press) appends the band as the active range", () => {
    const base = sel("B2", ["B2:C3"]);
    expect(show(headerSelection("column", 6, 6, nav, base))).toBe("G1 B2:C3 [G1:G1048576]");
    // Rows 8:9 are crossed by A9:A10.
    expect(show(headerSelection("row", 7, 8, nav, base))).toBe("A8 B2:C3 [A8:XFD10]");
    expect(base).toEqual(sel("B2", ["B2:C3"]));
  });

  it("extendToLine (Shift+header press) extends the active range to whole rows or columns, as Shift does for a cell", () => {
    // From B3 to row 6, and up to row 1: the active cell stays B3.
    expect(show(extendToLine(sel("B3"), "row", 5, nav))).toBe("B3 [A3:XFD6]");
    expect(show(extendToLine(sel("B3"), "row", 0, nav))).toBe("B3 [A1:XFD3]");
    expect(show(extendToLine(sel("B3"), "column", 6, nav))).toBe("B3 [B1:G1048576]");
    // The other ranges are kept; only the active one changes.
    const two = sel("H8", ["B2:C3", "H8"], 1);
    expect(show(extendToLine(two, "column", 9, nav))).toBe("H8 B2:C3 [H1:J1048576]");
    expect(two).toEqual(sel("H8", ["B2:C3", "H8"], 1));
    // From inside a merge the band covers the whole merge; snapped.
    expect(show(extendToLine(sel("C5", ["C5:E6"]), "row", 6, nav))).toBe("C5 [A5:XFD7]");
    expect(show(extendToLine(sel("B2"), "row", 8, nav))).toBe("B2 [A2:XFD10]");
  });
});

describe("normaliseSelection", () => {
  const nav = navOf(sheetOf({ merges: ["C5:E6"] }));

  it("selects the first visible cell for no selection", () => {
    expect(show(normaliseSelection(undefined, nav))).toBe("A1 [A1]");
    expect(show(normaliseSelection(undefined, navOf(sheetOf({ hiddenRows: [1], hiddenCols: ["A"] }))))).toBe("B2 [B2]");
  });

  it("keeps a valid selection as it is", () => {
    const valid = sel("B2", ["A1:C3", "G5:H6"], 0);
    expect(normaliseSelection(valid, nav)).toEqual(valid);
    expect(normaliseSelection(valid, nav)).not.toBe(valid);
  });

  it("normalises, clamps and snaps ranges, dropping unreadable ones and ones outside the grid", () => {
    const stored = {
      active: { row: 5, col: 3 },
      ranges: [{ top: 9, left: 9, bottom: 2, right: 2 }, "B2", { top: 2_000_000, left: 0, bottom: 2_000_001, right: 0 }, { top: 5, left: 3, bottom: 5, right: 3 }, { top: 0, left: 0, bottom: 2e9, right: Number.NaN }],
      activeRangeIndex: 3,
    } as unknown as SheetSelection;
    // D6 is a member of C5:E6: the active cell lands on C5, the range snaps.
    expect(show(normaliseSelection(stored, nav))).toBe("C5 C3:J10 [C5:E6]");
  });

  it("repairs the active range index and an active cell outside its range", () => {
    expect(show(normaliseSelection({ ...sel("H8", ["A1:B2", "H8"]), activeRangeIndex: 7 }, nav))).toBe("H8 A1:B2 [H8]");
    expect(show(normaliseSelection({ ...sel("H8", ["A1:B2", "C3"]), activeRangeIndex: 1 }, nav))).toBe("C3 A1:B2 [C3]");
    expect(show(normaliseSelection({ ...sel("H8", ["A1:B2", "C3"]), activeRangeIndex: -1 }, nav))).toBe("A1 [A1:B2] C3");
  });

  it("moves a hidden active cell to the first visible cell of its range", () => {
    const hidden = navOf(sheetOf({ hiddenRows: [2] }));
    expect(show(normaliseSelection(sel("B2", ["B2:C3"]), hidden))).toBe("B3 [B2:C3]");
    expect(show(normaliseSelection(sel("B2", ["B2"]), hidden))).toBe("B3 [B3]");
  });

  it("falls back to the stored active cell, else A1, when no range is usable", () => {
    expect(show(normaliseSelection({ active: ref("D4"), ranges: [], activeRangeIndex: 0 }, nav))).toBe("D4 [D4]");
    expect(show(normaliseSelection({ ranges: "nope" } as unknown as SheetSelection, nav))).toBe("A1 [A1]");
    expect(show(normaliseSelection({ active: { row: 4, col: 3 }, ranges: [{ top: -5, left: -5, bottom: -1, right: -1 }], activeRangeIndex: 0 }, nav))).toBe("C5 [C5:E6]");
  });
});

describe("navigateByKey", () => {
  const nav = navOf(sheetOf({ cells: { ...fill("B3:B12"), B2: "Header" } }), { rows: 10, cols: 5 });

  it("arrows move, Shift+arrows extend, Mod+arrows jump to the data edge", () => {
    expect(show(keys(sel("B3"), nav, key("ArrowDown")))).toBe("B4 [B4]");
    expect(show(keys(sel("B3"), nav, key("ArrowRight", { shift: true })))).toBe("B3 [B3:C3]");
    expect(show(keys(sel("B3"), nav, key("ArrowDown", MOD)))).toBe("B12 [B12]");
    expect(show(keys(sel("B3"), nav, key("ArrowUp", MOD)))).toBe("B2 [B2]");
    // macOS: Cmd is the primary modifier, Control is not.
    expect(show(keys(sel("B3"), nav, key("ArrowDown", { mod: true })))).toBe("B12 [B12]");
    expect(navigateByKey(sel("B3"), key("ArrowDown", { ctrl: true }), nav)).toBeNull();
  });

  it("Mod+Shift+arrows extend the moving corner to its data edge", () => {
    const down = keys(sel("B3"), nav, key("ArrowDown", { ...MOD, shift: true }));
    expect(show(down)).toBe("B3 [B3:B12]");
    // From the moving corner B12 the block runs up to B2 ("Header").
    expect(show(keys(down, nav, key("ArrowUp", { ...MOD, shift: true })))).toBe("B3 [B2:B3]");
    expect(show(keys(sel("B12"), nav, key("ArrowUp", { ...MOD, shift: true })))).toBe("B12 [B2:B12]");
    // The moving corner is the one farthest from the anchor: B3 anchors B3:C3,
    // so Mod+Shift+Down searches down column C (empty) to the grid's edge.
    expect(show(keys(sel("B3", ["B3:C3"]), nav, key("ArrowDown", { ...MOD, shift: true })))).toBe("B3 [B3:C1048576]");
  });

  it("Home goes to column A of the row, Mod+Home to A1, Mod+End to the used range's bottom-right", () => {
    expect(show(keys(sel("C7"), nav, key("Home")))).toBe("A7 [A7]");
    expect(show(keys(sel("C7"), nav, key("Home", { shift: true })))).toBe("C7 [A7:C7]");
    expect(show(keys(sel("C7"), nav, key("Home", MOD)))).toBe("A1 [A1]");
    expect(show(keys(sel("C7"), nav, key("Home", { ...MOD, shift: true })))).toBe("C7 [A1:C7]");
    expect(show(keys(sel("A1"), nav, key("End", MOD)))).toBe("B12 [B12]");
    expect(show(keys(sel("A1"), nav, key("End", { ...MOD, shift: true })))).toBe("A1 [A1:B12]");
    expect(show(keys(sel("C7"), navOf(sheetOf()), key("End", MOD)))).toBe("A1 [A1]");
    const hidden = navOf(sheetOf({ cells: fill("C3", "D9"), hiddenRows: [9, 1], hiddenCols: ["A", "D"] }));
    expect(show(keys(sel("E5"), hidden, key("Home")))).toBe("B5 [B5]");
    expect(show(keys(sel("E5"), hidden, key("Home", MOD)))).toBe("B2 [B2]");
    expect(show(keys(sel("E5"), hidden, key("End", MOD)))).toBe("C8 [C8]");
  });

  it("PageDown/PageUp move ± pageRows rows; Alt moves ± pageCols columns; Shift extends", () => {
    expect(show(keys(sel("C3"), nav, key("PageDown")))).toBe("C13 [C13]");
    expect(show(keys(sel("C13"), nav, key("PageUp")))).toBe("C3 [C3]");
    expect(show(keys(sel("C3"), nav, key("PageUp")))).toBe("C1 [C1]");
    expect(show(keys(sel("C3"), nav, key("PageDown", { shift: true })))).toBe("C3 [C3:C13]");
    expect(show(keys(sel("C3"), nav, key("PageDown", { alt: true })))).toBe("H3 [H3]");
    expect(show(keys(sel("H3"), nav, key("PageUp", { alt: true })))).toBe("C3 [C3]");
    expect(show(keys(sel("C3"), nav, key("PageDown", { alt: true, shift: true })))).toBe("C3 [C3:H3]");
    expect(show(keys(sel("C1048570"), nav, key("PageDown")))).toBe("C1048576 [C1048576]");
  });

  it("Mod+A selects the current region, then the whole sheet; an isolated cell goes straight to the whole sheet", () => {
    const region = keys(sel("B5"), nav, key("a", MOD));
    expect(show(region)).toBe("B5 [B2:B12]");
    expect(show(keys(region, nav, key("A", MOD)))).toBe("A1 [A1:XFD1048576]");
    expect(show(keys(sel("H20"), nav, key("a", MOD)))).toBe("A1 [A1:XFD1048576]");
    expect(show(keys(selectAll(), nav, key("a", MOD)))).toBe("A1 [A1:XFD1048576]");
    expect(navigateByKey(sel("B5"), key("a", { ctrl: true }), nav)).toBeNull();
  });

  it("Shift+Space selects rows, Ctrl+Space columns — literal Ctrl on every platform", () => {
    expect(show(keys(sel("B3", ["B3:C4"]), nav, key(" ", { shift: true })))).toBe("B3 [A3:XFD4]");
    expect(show(keys(sel("B3", ["B3:C4"]), nav, key(" ", MOD)))).toBe("B3 [B1:C1048576]");
    expect(show(keys(sel("B3", ["B3:C4"]), nav, key(" ", { ctrl: true })))).toBe("B3 [B1:C1048576]");
    expect(navigateByKey(sel("B3"), key(" ", { mod: true }), nav)).toBeNull();
    expect(navigateByKey(sel("B3"), key(" "), nav)).toBeNull();
    expect(navigateByKey(sel("B3"), key(" ", { ...MOD, shift: true }), nav)).toBeNull();
  });

  it("Tab/Shift+Tab and Enter/Shift+Enter cycle inside a multi-cell selection", () => {
    const block = sel("A1", ["A1:B2"]);
    expect(show(keys(block, nav, key("Tab")))).toBe("B1 [A1:B2]");
    expect(show(keys(block, nav, key("Tab", { shift: true })))).toBe("B2 [A1:B2]");
    expect(show(keys(block, nav, key("Enter")))).toBe("A2 [A1:B2]");
    expect(show(keys(block, nav, key("Enter", { shift: true })))).toBe("B2 [A1:B2]");
    expect(show(keys(sel("C3"), nav, key("Tab")))).toBe("D3 [D3]");
    expect(show(keys(sel("C3"), nav, key("Enter")))).toBe("C4 [C4]");
  });

  it("returns null for keys it does not handle — including Ctrl+PageUp/PageDown", () => {
    const s = sel("C3");
    for (const press of [
      key("x"),
      key("F2"),
      key("Escape"),
      key("Delete"),
      key("End"),
      key("PageUp", MOD),
      key("PageDown", MOD),
      key("PageUp", { ctrl: true }),
      key("PageDown", { ctrl: true }),
      key("PageDown", { mod: true }),
      key("ArrowDown", { alt: true }),
      key("Home", { alt: true }),
      key("Tab", MOD),
      key("Tab", { alt: true }),
      key("Enter", MOD),
      key("Enter", { alt: true }),
      key("a", { ...MOD, shift: true }),
      key("a"),
      key("constructor"),
      key("toString", MOD),
    ]) {
      expect(navigateByKey(s, press, nav), JSON.stringify(press)).toBeNull();
    }
  });
});

describe("a merge whose top-left is hidden", () => {
  // Column C is hidden, so C5:E5 shows as D5:E5; row 2 is hidden, so B2:C3
  // shows as its third row. Each is still one visible cell, its top-left active.
  const nav = navOf(sheetOf({ merges: ["C5:E5", "B2:C3"], hiddenCols: ["C"], hiddenRows: [2] }));

  it("is entered at its top-left and left from its far edge", () => {
    const into = moveActive(sel("B5"), "right", nav);
    expect(show(into)).toBe("C5 [C5:E5]");
    expect(show(moveActive(into, "right", nav))).toBe("F5 [F5]");
    expect(show(moveActive(into, "down", nav))).toBe("D6 [D6]");
    expect(show(moveActive(sel("B1"), "down", nav))).toBe("B2 [B2:C3]");
    expect(show(selectCell(ref("D5"), nav))).toBe("C5 [C5:E5]");
  });

  it("is kept as the active cell and met by Tab at its top-left", () => {
    expect(show(normaliseSelection(selectCell(ref("E5"), nav), nav))).toBe("C5 [C5:E5]");
    const row = sel("B5", ["A5:F5"]);
    const first = cycleWithin(row, "row", false, nav);
    expect(show(first)).toBe("C5 [A5:F5]");
    expect(show(cycleWithin(first, "row", false, nav))).toBe("F5 [A5:F5]");
    // After B1 (C1 is hidden), row 2 is hidden but holds B2:C3's top-left.
    expect(show(cycleWithin(sel("B1", ["A1:C3"]), "row", false, nav))).toBe("B2 [A1:C3]");
    expect(show(cycleWithin(sel("B2", ["A1:C3"]), "row", false, nav))).toBe("A3 [A1:C3]");
    expect(show(cycleWithin(sel("A3", ["A1:C3"]), "row", true, nav))).toBe("B2 [A1:C3]");
  });
});

describe("selections are fresh plain data", () => {
  it("never share an object with the input or the sheet's merges", () => {
    const sheet = sheetOf({ cells: fill("A1:C3"), merges: ["C5:E6"] });
    const nav = navOf(sheet);
    const input = sel("C5", ["A1:B2", "C5:E6"], 1);
    const snapshot = JSON.parse(JSON.stringify(input)) as SheetSelection;
    const results = [
      moveActive(input, "down", nav),
      extendBy(input, "right", nav),
      extendTo(input, ref("F7"), nav),
      cycleWithin(input, "row", false, nav),
      addRange(input, ref("D5"), nav),
      selectRows(input, nav),
      selectColumns(input, nav),
      headerSelection("row", 4, 4, nav, input),
      normaliseSelection(input, nav),
      selectCell(ref("D6"), nav),
      navigateByKey(input, key("a", MOD), nav)!,
    ];
    for (const result of results) {
      expect(JSON.parse(JSON.stringify(result))).toEqual(result);
      expect(result.active).not.toBe(input.active);
      for (const r of result.ranges) {
        expect(input.ranges).not.toContain(r);
        expect(sheet.merges).not.toContain(r);
        r.top = -1;
        r.left = -1;
      }
      result.active.row = -1;
    }
    expect(input).toEqual(snapshot);
    expect(sheet.merges).toEqual([range("C5:E6")]);
  });
});

describe("range helpers", () => {
  it("rangeContains and isCellSelected test inclusive bounds over every range", () => {
    expect(rangeContains(range("B2:C3"), 1, 1)).toBe(true);
    expect(rangeContains(range("B2:C3"), 2, 2)).toBe(true);
    expect(rangeContains(range("B2:C3"), 3, 2)).toBe(false);
    expect(rangeContains(range("B2:C3"), 1, 0)).toBe(false);
    const s = sel("A1", ["A1", "D4:E5"]);
    expect(isCellSelected(s, 0, 0)).toBe(true);
    expect(isCellSelected(s, 4, 4)).toBe(true);
    expect(isCellSelected(s, 1, 1)).toBe(false);
  });

  it("wholeRows and wholeColumns span the grid, in either order, clamped", () => {
    expect(wholeRows(4, 2)).toEqual({ top: 2, left: 0, bottom: 4, right: LAST_COL });
    expect(wholeColumns(3, 1)).toEqual({ top: 0, left: 1, bottom: LAST_ROW, right: 3 });
    expect(wholeRows(-3, 5_000_000)).toEqual({ top: 0, left: 0, bottom: LAST_ROW, right: LAST_COL });
    expect(wholeColumns(20_000, 20_000)).toEqual({ top: 0, left: LAST_COL, bottom: LAST_ROW, right: LAST_COL });
  });
});
