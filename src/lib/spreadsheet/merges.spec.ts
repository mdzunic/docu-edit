// SPEC-117 "Merges": the merge index every renderer, the selection model and
// the overflow pass query — merges sorted by top (then left) with a running
// maximum of their bottoms, so mergeAt/mergesIn are a binary search plus a
// walk back that stops as soon as the running max-bottom falls above the
// query (proven by counting index reads on 1,000 merges) — and acceptMerges,
// the readers' rule that only ranges of two or more cells overlapping no
// earlier merge are modelled, in source order.
import { describe, expect, it } from "vitest";
import { parseRange, type GridRange } from "./address";
import { acceptMerges, buildMergeIndex, mergeAt, mergesIn, rangesIntersect, type MergeIndex } from "./merges";

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
