// SPEC-117 "Merges": the merge index every renderer, the selection model and
// the overflow pass query — merges sorted by top (then left) with a running
// maximum of their bottoms, so mergeAt/mergesIn are a binary search plus a
// walk back that stops as soon as the running max-bottom falls above the
// query (proven by counting index reads on 1,000 merges) — paneMerges, the
// cell layer's query (which merges a pane draws, hosted by the first row of
// each that the pane shows), and acceptMerges, the readers' rule that only
// ranges of two or more cells overlapping no earlier merge are modelled, in
// source order.
import { describe, expect, it } from "vitest";
import { parseRange, type GridRange } from "./address";
import { acceptMerges, buildMergeIndex, mergeAt, mergesIn, paneMerges, rangesIntersect, type MergeIndex, type PaneMerge } from "./merges";

function range(text: string): GridRange {
  const parsed = parseRange(text);
  if (!parsed) {
    throw new Error(`bad range ${text}`);
  }
  return parsed;
}

/** `index` with its arrays behind Proxies counting numeric-index reads. */
function counted(index: MergeIndex): { index: MergeIndex; reads: () => number } {
  let reads = 0;
  const watch = <T>(items: T[]): T[] =>
    new Proxy(items, {
      get(target, property, receiver) {
        if (typeof property === "string" && /^\d+$/.test(property)) {
          reads += 1;
        }
        return Reflect.get(target, property, receiver);
      },
    });
  return { index: { ranges: watch(index.ranges), maxBottom: watch(index.maxBottom) }, reads: () => reads };
}

describe("buildMergeIndex", () => {
  it("sorts merges by top, then left, with a running maximum of their bottoms", () => {
    const merges = [range("C5:E6"), range("A1:A100"), range("D2:E2"), range("A102:B103"), range("B2:C2")];
    const index = buildMergeIndex(merges);
    expect(index.ranges).toEqual([range("A1:A100"), range("B2:C2"), range("D2:E2"), range("C5:E6"), range("A102:B103")]);
    expect(index.maxBottom).toEqual([99, 99, 99, 99, 102]);
    // The input keeps its order.
    expect(merges[0]).toEqual(range("C5:E6"));
  });

  it("is empty for no merges", () => {
    expect(buildMergeIndex(undefined)).toEqual({ ranges: [], maxBottom: [] });
    expect(buildMergeIndex([])).toEqual({ ranges: [], maxBottom: [] });
    expect(mergeAt(buildMergeIndex([]), 0, 0)).toBeUndefined();
    expect(mergesIn(buildMergeIndex(undefined), range("A1:XFD1048576"))).toEqual([]);
  });
});

describe("rangesIntersect", () => {
  it("is true exactly when the ranges share a cell", () => {
    expect(rangesIntersect(range("A1:B2"), range("B2:C3"))).toBe(true);
    expect(rangesIntersect(range("A1:C3"), range("B2"))).toBe(true);
    expect(rangesIntersect(range("A1:B2"), range("C1:D2"))).toBe(false);
    expect(rangesIntersect(range("A1:B2"), range("A3:B4"))).toBe(false);
  });
});

describe("mergeAt", () => {
  const index = buildMergeIndex([range("A1:A100"), range("C2:D2"), range("C4:D5")]);

  it("finds the merge covering a cell, including its edges", () => {
    expect(mergeAt(index, 0, 0)).toEqual(range("A1:A100"));
    expect(mergeAt(index, 99, 0)).toEqual(range("A1:A100"));
    expect(mergeAt(index, 1, 3)).toEqual(range("C2:D2"));
    expect(mergeAt(index, 4, 2)).toEqual(range("C4:D5"));
    expect(mergeAt(index, 3, 3)).toEqual(range("C4:D5"));
  });

  it("returns undefined outside every merge", () => {
    expect(mergeAt(index, 100, 0)).toBeUndefined();
    expect(mergeAt(index, 2, 2)).toBeUndefined();
    expect(mergeAt(index, 1, 1)).toBeUndefined();
    expect(mergeAt(index, 5, 2)).toBeUndefined();
  });

  it("finds a tall merge that starts above later, shorter ones", () => {
    // A1:A100 is index 0; the walk back from row 49 passes C4:D5 and C2:D2.
    expect(mergeAt(index, 49, 0)).toEqual(range("A1:A100"));
  });
});

describe("mergesIn", () => {
  it("returns every merge sharing a cell with the window, in index order", () => {
    const index = buildMergeIndex([range("A1:B5"), range("D2:E3"), range("A7:C8"), range("G1:H1"), range("B10:C11")]);
    // A1:B5's top is above the window and its bottom inside it.
    expect(mergesIn(index, range("A3:E8"))).toEqual([range("A1:B5"), range("D2:E3"), range("A7:C8")]);
    expect(mergesIn(index, range("F2:F9"))).toEqual([]);
    expect(mergesIn(index, range("H1"))).toEqual([range("G1:H1")]);
    expect(mergesIn(index, range("A11:XFD1048576"))).toEqual([range("B10:C11")]);
  });

  it("stops walking back once the running max-bottom falls above the window", () => {
    const merges = Array.from({ length: 1000 }, (_, i): GridRange => ({ top: 2 * i, left: 0, bottom: 2 * i + 1, right: 1 }));
    const { index, reads } = counted(buildMergeIndex(merges));
    expect(mergesIn(index, { top: 1990, left: 0, bottom: 1995, right: 5 })).toEqual(merges.slice(995, 998));
    expect(reads()).toBeLessThan(40);
  });

  it("still finds a tall early merge reaching into the window", () => {
    const merges = [{ top: 0, left: 5, bottom: 5000, right: 6 }, ...Array.from({ length: 100 }, (_, i): GridRange => ({ top: 2 * i, left: 0, bottom: 2 * i + 1, right: 1 }))];
    const index = buildMergeIndex(merges);
    expect(mergesIn(index, { top: 4000, left: 0, bottom: 4001, right: 10 })).toEqual([merges[0]]);
    expect(mergeAt(index, 4000, 6)).toEqual(merges[0]);
  });
});

/** A pane's rendered rows: `first..last` minus the hidden ones. */
function rendered(first: number, last: number, hidden: readonly number[] = []): Array<{ index: number }> {
  return Array.from({ length: last - first + 1 }, (_, i) => ({ index: first + i })).filter((row) => !hidden.includes(row.index));
}

/** 64 px columns, `hidden` ones zero-width. */
function colLeft(hidden: readonly number[] = []): (col: number) => number {
  return (col) => 64 * col - 64 * hidden.filter((c) => c < col).length;
}

/** Each drawn merge with the index of the row hosting it. */
function hosts(found: PaneMerge<{ index: number }>[]): Array<[GridRange, number]> {
  return found.map(({ range: r, host }) => [r, host.index]);
}

describe("paneMerges", () => {
  it("hosts each merge in the window by its anchor row", () => {
    const index = buildMergeIndex([range("C5:E5"), range("B2:B3"), range("H20:I21")]);
    expect(hosts(paneMerges(index, range("A1:F10"), rendered(0, 9), colLeft()))).toEqual([
      [range("B2:B3"), 1],
      [range("C5:E5"), 4],
    ]);
  });

  it("hands a merge whose anchor row is hidden to the first row of it that shows", () => {
    // B4:D6 with row 4 hidden: rows 5–6 show it, hosted by row 5.
    const index = buildMergeIndex([range("B4:D6")]);
    expect(hosts(paneMerges(index, range("A1:F10"), rendered(0, 9, [3]), colLeft()))).toEqual([[range("B4:D6"), 4]]);
    expect(hosts(paneMerges(index, range("A1:F10"), rendered(0, 9, [3, 4]), colLeft()))).toEqual([[range("B4:D6"), 5]]);
  });

  it("hosts a merge whose anchor row is above the window by the window's first row", () => {
    // Rows 1–2 frozen: the main pane's window starts at row 3 and still
    // draws A2:A4; scrolled further, the first rendered row hosts B2:C50.
    const index = buildMergeIndex([range("A2:A4"), range("B2:C50")]);
    expect(hosts(paneMerges(index, range("A3:F20"), rendered(2, 19), colLeft()))).toEqual([
      [range("A2:A4"), 2],
      [range("B2:C50"), 2],
    ]);
    expect(hosts(paneMerges(index, range("A23:F60"), rendered(22, 59), colLeft()))).toEqual([[range("B2:C50"), 22]]);
  });

  it("draws nothing for a merge whose rows or columns in the window are all hidden", () => {
    const index = buildMergeIndex([range("B4:D5"), range("G1:H1"), range("J1:K1")]);
    // Rows 4–5 hidden; columns G–H hidden; J1:K1 has J hidden but K shows.
    expect(hosts(paneMerges(index, range("A1:M10"), rendered(0, 9, [3, 4]), colLeft([6, 7, 9])))).toEqual([[range("J1:K1"), 0]]);
    // Only the columns inside the window count: C–D of A1:D1 are hidden,
    // and the window holds just those two.
    const wide = buildMergeIndex([range("A1:D1")]);
    expect(hosts(paneMerges(wide, range("C1:F5"), rendered(0, 4), colLeft([2, 3])))).toEqual([]);
    expect(hosts(paneMerges(wide, range("A1:B5"), rendered(0, 4), colLeft([2, 3])))).toEqual([[range("A1:D1"), 0]]);
  });

  it("draws a merge crossing the pane's edge in each pane it reaches", () => {
    // Column A frozen: A1:H1 shows in the frozen pane (A) and the scrolled one (B on).
    const index = buildMergeIndex([range("A1:H1")]);
    expect(hosts(paneMerges(index, range("A1:A30"), rendered(0, 29), colLeft()))).toEqual([[range("A1:H1"), 0]]);
    expect(hosts(paneMerges(index, range("B1:T30"), rendered(0, 29), colLeft()))).toEqual([[range("A1:H1"), 0]]);
    expect(hosts(paneMerges(index, range("K1:Z30"), rendered(0, 29), colLeft()))).toEqual([]);
  });
});

describe("acceptMerges", () => {
  it("keeps ranges of two or more cells that overlap no earlier kept merge, in source order", () => {
    const candidates = [range("C5:E6"), range("A1:B2"), range("D6:F7"), range("H1"), range("B2:B3"), range("H2:H3"), range("A10:B10")];
    expect(acceptMerges(candidates)).toEqual([range("C5:E6"), range("A1:B2"), range("H2:H3"), range("A10:B10")]);
  });

  it("checks overlap against kept merges only", () => {
    // B1:B2 overlaps the rejected single cell B1 only, so it is kept.
    expect(acceptMerges([range("B1"), range("B1:B2")])).toEqual([range("B1:B2")]);
    // A later merge overlapping a rejected (overlapping) one but no kept one is kept.
    expect(acceptMerges([range("A1:B2"), range("B2:C3"), range("C3:D4")])).toEqual([range("A1:B2"), range("C3:D4")]);
    expect(acceptMerges([])).toEqual([]);
  });
});
