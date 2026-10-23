// SPEC-109: the assembly layout engine. Every layout against 1, 2, 3, 5 and
// 12 items on a wide and a tall sheet with exact rectangles; "auto" column
// selection and its tie-breaks; gap arithmetic at the content box's edges;
// fit and align within each cell; overlay transforms and z-order; the empty
// tray; and purity — rectangles exact to 3 decimals, inputs never touched.
import { describe, expect, it } from "vitest";
import {
  ASSEMBLY_PLAN_DEFAULTS,
  autoColumns,
  DEFAULT_GAP_CM,
  DEFAULT_OVERLAY_TRANSFORM,
  LAYOUT_KINDS,
  MAX_ASSEMBLY_ITEMS,
  naturalAssemblyPx,
  normaliseTransform,
  planAssembly,
  type AssemblyItem,
  type AssemblyPlan,
  type LayoutKind,
  type PlacedItem,
} from "./assemblyLayout";
import type { SheetGeometry } from "./sheetLayout";

/** 32 × 22 cm, 1 cm margins: a 30 × 20 cm content box. */
const WIDE: SheetGeometry = { widthCm: 32, heightCm: 22, content: { xCm: 1, yCm: 1, widthCm: 30, heightCm: 20 } };
/** 22 × 32 cm, 1 cm margins: a 20 × 30 cm content box. */
const TALL: SheetGeometry = { widthCm: 22, heightCm: 32, content: { xCm: 1, yCm: 1, widthCm: 20, heightCm: 30 } };

/** Items cycle through three shapes: 2:1, 1:1 and 1:2. */
const SHAPES = [
  [200, 100],
  [100, 100],
  [100, 200],
];

function items(count: number): AssemblyItem[] {
  return Array.from({ length: count }, (_, i) => ({ id: `item-${i}`, widthPx: SHAPES[i % 3][0], heightPx: SHAPES[i % 3][1] }));
}

function plan(overrides: Partial<AssemblyPlan> = {}): AssemblyPlan {
  return { ...ASSEMBLY_PLAN_DEFAULTS, ...overrides };
}

/** [xCm, yCm, widthCm, heightCm] of each placed item. */
function rects(placed: PlacedItem[]): number[][] {
  return placed.map((item) => [item.xCm, item.yCm, item.widthCm, item.heightCm]);
}

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object") {
    for (const inner of Object.values(value)) {
      deepFreeze(inner);
    }
    Object.freeze(value);
  }
  return value;
}

/** Contain, centred, gapCm 1, "auto" columns, the default transforms: the
 * visible placed rectangle of every item. */
const EXPECTED: Record<LayoutKind, Record<"wide" | "tall", Record<number, number[][]>>> = {
  grid: {
    wide: {
      1: [[1, 3.5, 30, 15]],
      2: [[1, 7.375, 14.5, 7.25], [16.5, 3.75, 14.5, 14.5]],
      3: [[1, 2.125, 14.5, 7.25], [19, 1, 9.5, 9.5], [5.875, 11.5, 4.75, 9.5]],
      5: [[1, 3.417, 9.333, 4.667], [11.333, 1.083, 9.333, 9.333], [23.958, 1, 4.75, 9.5], [1, 13.917, 9.333, 4.667], [11.333, 11.583, 9.333, 9.333]],
      12: [[1, 2.313, 6.75, 3.375], [9.125, 1, 6, 6], [18.375, 1, 3, 6], [24.25, 2.313, 6.75, 3.375], [1.375, 8, 6, 6], [10.625, 8, 3, 6], [16.5, 9.313, 6.75, 3.375], [24.625, 8, 6, 6], [2.875, 15, 3, 6], [8.75, 16.313, 6.75, 3.375], [16.875, 15, 6, 6], [26.125, 15, 3, 6]],
    },
    tall: {
      1: [[1, 11, 20, 10]],
      2: [[1, 3.25, 20, 10], [3.75, 16.5, 14.5, 14.5]],
      3: [[1, 5.875, 9.5, 4.75], [11.5, 3.5, 9.5, 9.5], [2.125, 16.5, 7.25, 14.5]],
      5: [[1, 3.292, 9.5, 4.75], [11.583, 1, 9.333, 9.333], [3.417, 11.333, 4.667, 9.333], [11.5, 13.625, 9.5, 4.75], [1.083, 21.667, 9.333, 9.333]],
      12: [[1, 2.875, 6, 3], [8, 1.375, 6, 6], [16.313, 1, 3.375, 6.75], [1, 10.625, 6, 3], [8, 9.125, 6, 6], [16.313, 8.75, 3.375, 6.75], [1, 18.375, 6, 3], [8, 16.875, 6, 6], [16.313, 16.5, 3.375, 6.75], [1, 26.125, 6, 3], [8, 24.625, 6, 6], [16.313, 24.25, 3.375, 6.75]],
    },
  },
  horizontal: {
    wide: {
      1: [[1, 3.5, 30, 15]],
      2: [[1, 6.167, 19.333, 9.667], [21.333, 6.167, 9.667, 9.667]],
      3: [[1, 7, 16, 8], [18, 7, 8, 8], [27, 7, 4, 8]],
      5: [[1, 9, 8, 4], [10, 9, 4, 4], [15, 9, 2, 4], [18, 9, 8, 4], [27, 9, 4, 4]],
      12: [[1, 10.321, 2.714, 1.357], [4.714, 10.321, 1.357, 1.357], [7.071, 10.321, 0.679, 1.357], [8.75, 10.321, 2.714, 1.357], [12.464, 10.321, 1.357, 1.357], [14.821, 10.321, 0.679, 1.357], [16.5, 10.321, 2.714, 1.357], [20.214, 10.321, 1.357, 1.357], [22.571, 10.321, 0.679, 1.357], [24.25, 10.321, 2.714, 1.357], [27.964, 10.321, 1.357, 1.357], [30.321, 10.321, 0.679, 1.357]],
    },
    tall: {
      1: [[1, 11, 20, 10]],
      2: [[1, 12.833, 12.667, 6.333], [14.667, 12.833, 6.333, 6.333]],
      3: [[1, 13.429, 10.286, 5.143], [12.286, 13.429, 5.143, 5.143], [18.429, 13.429, 2.571, 5.143]],
      5: [[1, 14.769, 4.923, 2.462], [6.923, 14.769, 2.462, 2.462], [10.385, 14.769, 1.231, 2.462], [12.615, 14.769, 4.923, 2.462], [18.538, 14.769, 2.462, 2.462]],
      12: [[1, 15.679, 1.286, 0.643], [3.286, 15.679, 0.643, 0.643], [4.929, 15.679, 0.321, 0.643], [6.25, 15.679, 1.286, 0.643], [8.536, 15.679, 0.643, 0.643], [10.179, 15.679, 0.321, 0.643], [11.5, 15.679, 1.286, 0.643], [13.786, 15.679, 0.643, 0.643], [15.429, 15.679, 0.321, 0.643], [16.75, 15.679, 1.286, 0.643], [19.036, 15.679, 0.643, 0.643], [20.679, 15.679, 0.321, 0.643]],
    },
  },
  vertical: {
    wide: {
      1: [[1, 3.5, 30, 15]],
      2: [[9.667, 1, 12.667, 6.333], [9.667, 8.333, 12.667, 12.667]],
      3: [[13.429, 1, 5.143, 2.571], [13.429, 4.571, 5.143, 5.143], [13.429, 10.714, 5.143, 10.286]],
      5: [[14.4, 1, 3.2, 1.6], [14.4, 3.6, 3.2, 3.2], [14.4, 7.8, 3.2, 6.4], [14.4, 15.2, 3.2, 1.6], [14.4, 17.8, 3.2, 3.2]],
      12: [[15.679, 1, 0.643, 0.321], [15.679, 2.321, 0.643, 0.643], [15.679, 3.964, 0.643, 1.286], [15.679, 6.25, 0.643, 0.321], [15.679, 7.571, 0.643, 0.643], [15.679, 9.214, 0.643, 1.286], [15.679, 11.5, 0.643, 0.321], [15.679, 12.821, 0.643, 0.643], [15.679, 14.464, 0.643, 1.286], [15.679, 16.75, 0.643, 0.321], [15.679, 18.071, 0.643, 0.643], [15.679, 19.714, 0.643, 1.286]],
    },
    tall: {
      1: [[1, 11, 20, 10]],
      2: [[1.333, 1, 19.333, 9.667], [1.333, 11.667, 19.333, 19.333]],
      3: [[7, 1, 8, 4], [7, 6, 8, 8], [7, 15, 8, 16]],
      5: [[8.4, 1, 5.2, 2.6], [8.4, 4.6, 5.2, 5.2], [8.4, 10.8, 5.2, 10.4], [8.4, 22.2, 5.2, 2.6], [8.4, 25.8, 5.2, 5.2]],
      12: [[10.321, 1, 1.357, 0.679], [10.321, 2.679, 1.357, 1.357], [10.321, 5.036, 1.357, 2.714], [10.321, 8.75, 1.357, 0.679], [10.321, 10.429, 1.357, 1.357], [10.321, 12.786, 1.357, 2.714], [10.321, 16.5, 1.357, 0.679], [10.321, 18.179, 1.357, 1.357], [10.321, 20.536, 1.357, 2.714], [10.321, 24.25, 1.357, 0.679], [10.321, 25.929, 1.357, 1.357], [10.321, 28.286, 1.357, 2.714]],
    },
  },
  masonry: {
    wide: {
      1: [[1, 1, 30, 15]],
      2: [[1, 1, 14.5, 7.25], [16.5, 1, 14.5, 14.5]],
      3: [[7.9, 1, 7.6, 3.8], [16.5, 1, 7.6, 7.6], [7.9, 5.8, 7.6, 15.2]],
      5: [[1, 1, 9.333, 4.667], [11.333, 1, 9.333, 9.333], [21.667, 1, 9.333, 18.667], [1, 6.667, 9.333, 4.667], [11.333, 11.333, 9.333, 9.333]],
      12: [[6.5, 1, 4, 2], [11.5, 1, 4, 4], [16.5, 1, 4, 8], [21.5, 1, 4, 2], [6.5, 4, 4, 4], [21.5, 4, 4, 8], [11.5, 6, 4, 2], [6.5, 9, 4, 4], [11.5, 9, 4, 8], [16.5, 10, 4, 2], [16.5, 13, 4, 4], [21.5, 13, 4, 8]],
    },
    tall: {
      1: [[1, 1, 20, 10]],
      2: [[1.333, 1, 19.333, 9.667], [1.333, 11.667, 19.333, 19.333]],
      3: [[1, 1, 9.5, 4.75], [11.5, 1, 9.5, 9.5], [1, 6.75, 9.5, 19]],
      5: [[1, 1, 9.5, 4.75], [11.5, 1, 9.5, 9.5], [1, 6.75, 9.5, 19], [11.5, 11.5, 9.5, 4.75], [11.5, 17.25, 9.5, 9.5]],
      12: [[2.909, 1, 4.727, 2.364], [8.636, 1, 4.727, 4.727], [14.364, 1, 4.727, 9.455], [2.909, 4.364, 4.727, 2.364], [8.636, 6.727, 4.727, 4.727], [2.909, 7.727, 4.727, 9.455], [14.364, 11.455, 4.727, 2.364], [8.636, 12.455, 4.727, 4.727], [14.364, 14.818, 4.727, 9.455], [2.909, 18.182, 4.727, 2.364], [8.636, 18.182, 4.727, 4.727], [2.909, 21.545, 4.727, 9.455]],
    },
  },
  overlay: {
    wide: {
      1: [[1, 3.5, 30, 15]],
      2: [[1, 3.5, 30, 15], [6, 1, 20, 20]],
      3: [[1, 3.5, 30, 15], [6, 1, 20, 20], [11, 1, 10, 20]],
      5: [[1, 3.5, 30, 15], [6, 1, 20, 20], [11, 1, 10, 20], [1, 3.5, 30, 15], [6, 1, 20, 20]],
      12: [[1, 3.5, 30, 15], [6, 1, 20, 20], [11, 1, 10, 20], [1, 3.5, 30, 15], [6, 1, 20, 20], [11, 1, 10, 20], [1, 3.5, 30, 15], [6, 1, 20, 20], [11, 1, 10, 20], [1, 3.5, 30, 15], [6, 1, 20, 20], [11, 1, 10, 20]],
    },
    tall: {
      1: [[1, 11, 20, 10]],
      2: [[1, 11, 20, 10], [1, 6, 20, 20]],
      3: [[1, 11, 20, 10], [1, 6, 20, 20], [3.5, 1, 15, 30]],
      5: [[1, 11, 20, 10], [1, 6, 20, 20], [3.5, 1, 15, 30], [1, 11, 20, 10], [1, 6, 20, 20]],
      12: [[1, 11, 20, 10], [1, 6, 20, 20], [3.5, 1, 15, 30], [1, 11, 20, 10], [1, 6, 20, 20], [3.5, 1, 15, 30], [1, 11, 20, 10], [1, 6, 20, 20], [3.5, 1, 15, 30], [1, 11, 20, 10], [1, 6, 20, 20], [3.5, 1, 15, 30]],
    },
  },
};

describe("planAssembly — every layout, 1/2/3/5/12 items, wide and tall", () => {
  for (const layout of LAYOUT_KINDS) {
    for (const [name, sheet] of [["wide", WIDE], ["tall", TALL]] as const) {
      for (const count of [1, 2, 3, 5, 12]) {
        it(`${layout}, ${count} item${count === 1 ? "" : "s"}, ${name} sheet`, () => {
          const placed = planAssembly(items(count), plan({ layout, gapCm: 1 }), sheet);
          expect(rects(placed)).toEqual(EXPECTED[layout][name][count]);
          expect(placed.map((item) => item.index)).toEqual([...Array(count).keys()]);
        });
      }
    }
  }

  it("keeps every rectangle inside the content box, exact to 3 decimals", () => {
    for (const layout of LAYOUT_KINDS) {
      for (const sheet of [WIDE, TALL]) {
        for (const item of planAssembly(items(12), plan({ layout }), sheet)) {
          const values = [item.xCm, item.yCm, item.widthCm, item.heightCm, ...Object.values(item.cell), ...Object.values(item.sourcePx)];
          for (const value of values) {
            expect(Math.round(value * 1000) / 1000).toBe(value);
          }
          const box = sheet.content;
          expect(item.xCm).toBeGreaterThanOrEqual(box.xCm);
          expect(item.yCm).toBeGreaterThanOrEqual(box.yCm);
          expect(item.xCm + item.widthCm).toBeLessThanOrEqual(box.xCm + box.widthCm + 0.001);
          expect(item.yCm + item.heightCm).toBeLessThanOrEqual(box.yCm + box.heightCm + 0.001);
        }
      }
    }
  });

  it("horizontal scales every item to one common height; vertical to one common width", () => {
    for (const sheet of [WIDE, TALL]) {
      const row = planAssembly(items(12), plan({ layout: "horizontal" }), sheet);
      expect(new Set(row.map((item) => item.heightCm)).size).toBe(1);
      expect(new Set(row.map((item) => item.yCm)).size).toBe(1);
      const column = planAssembly(items(12), plan({ layout: "vertical" }), sheet);
      expect(new Set(column.map((item) => item.widthCm)).size).toBe(1);
      expect(new Set(column.map((item) => item.xCm)).size).toBe(1);
    }
  });

  it("masonry puts each item in the currently shortest column, ties to the leftmost, aspect kept", () => {
    // 2 columns of 14.5 cm: 2:1 (7.25 tall) → col 0; 1:1 (14.5) → col 1;
    // 1:2 → col 0 (7.25 < 14.5); 1:1 → col 1 (14.5 + 1 + … vs col 0's
    // 7.25 + 1 + 29 = 37.25) — then the stack is scaled to fit 20 cm.
    const placed = planAssembly(
      [
        { id: "a", widthPx: 200, heightPx: 100 },
        { id: "b", widthPx: 100, heightPx: 100 },
        { id: "c", widthPx: 100, heightPx: 200 },
        { id: "d", widthPx: 100, heightPx: 100 },
      ],
      plan({ layout: "masonry", columns: 2, gapCm: 1 }),
      { ...WIDE, content: { xCm: 1, yCm: 1, widthCm: 30, heightCm: 100 } },
    );
    expect(rects(placed)).toEqual([
      [1, 1, 14.5, 7.25],
      [16.5, 1, 14.5, 14.5],
      [1, 9.25, 14.5, 29],
      [16.5, 16.5, 14.5, 14.5],
    ]);
    for (const [i, item] of placed.entries()) {
      const source = [200, 100, 100, 100][i] / [100, 100, 200, 100][i];
      expect(item.widthCm / item.heightCm).toBeCloseTo(source, 3);
    }
    // A tie (two empty columns) goes left.
    const tie = planAssembly([{ id: "a", widthPx: 100, heightPx: 100 }], plan({ layout: "masonry", columns: 3 }), WIDE);
    expect(tie[0].xCm).toBe(1);
    expect(placed.map((item) => item.row)).toEqual([0, 0, 1, 1]);
  });

  it("masonry scales a stack taller than the box down uniformly, gaps unchanged, centred across", () => {
    // Three 1:2 items in one column of 10 cm would be 3 × 20 + 2 gaps.
    const placed = planAssembly(
      Array.from({ length: 3 }, (_, i) => ({ id: `t${i}`, widthPx: 100, heightPx: 200 })),
      plan({ layout: "masonry", columns: 1, gapCm: 1 }),
      { widthCm: 12, heightCm: 32, content: { xCm: 1, yCm: 1, widthCm: 10, heightCm: 30 } },
    );
    // (30 − 2) / 60 of each: 4.667 × 9.333, centred in the 10 cm box.
    expect(rects(placed)).toEqual([
      [3.667, 1, 4.667, 9.333],
      [3.667, 11.333, 4.667, 9.333],
      [3.667, 21.667, 4.667, 9.333],
    ]);
    expect(placed[2].yCm + placed[2].heightCm).toBe(31);
  });
});

describe("auto columns", () => {
  it("chooses by item count and sheet aspect", () => {
    const counts = [1, 2, 3, 5, 12];
    expect(counts.map((count) => autoColumns(count, WIDE.content))).toEqual([1, 2, 2, 3, 4]);
    expect(counts.map((count) => autoColumns(count, TALL.content))).toEqual([1, 1, 2, 2, 3]);
  });

  it("minimises the wasted area of square cells: 5 items on 30 × 20 → 3 columns of 10 cm", () => {
    // c = 1: side min(30, 4) = 4 → 600 − 80 waste; c = 2: min(15, 6.667);
    // c = 3: min(10, 10) = 10 → 600 − 500 = 100, the least; c = 4: 7.5;
    // c = 5: min(6, 20) = 6.
    expect(autoColumns(5, { widthCm: 30, heightCm: 20 })).toBe(3);
  });

  it("breaks a waste tie towards the more square grid", () => {
    // 3 items on 30 × 20: 2 × 2 and 3 × 1 both give 10 cm cells (waste 300);
    // 2 columns × 2 rows is the squarer grid.
    expect(autoColumns(3, { widthCm: 30, heightCm: 20 })).toBe(2);
    // 3 items on 20 × 30: 1 × 3 and 2 × 2 both give 10 cm cells.
    expect(autoColumns(3, { widthCm: 20, heightCm: 30 })).toBe(2);
  });

  it("then towards fewer columns", () => {
    // 2 items on a square box: 1 × 2 and 2 × 1 tie on waste and on squareness.
    expect(autoColumns(2, { widthCm: 10, heightCm: 10 })).toBe(1);
    // 6 items on a square box: 2 × 3 and 3 × 2 tie the same way.
    expect(autoColumns(6, { widthCm: 10, heightCm: 10 })).toBe(2);
  });

  it("is what an 'auto' grid and masonry use, and a number is used as given", () => {
    const auto = planAssembly(items(5), plan({ columns: "auto", gapCm: 0, fit: "stretch" }), WIDE);
    expect(new Set(auto.map((item) => item.xCm)).size).toBe(3);
    const fixed = planAssembly(items(5), plan({ columns: 5, gapCm: 0, fit: "stretch" }), WIDE);
    expect(rects(fixed).map(([x, y, w, h]) => [x, y, w, h])).toEqual([0, 1, 2, 3, 4].map((i) => [1 + i * 6, 1, 6, 20]));
    // More columns than items keeps the column width.
    const sparse = planAssembly(items(2), plan({ columns: 4, gapCm: 0, fit: "stretch" }), WIDE);
    expect(rects(sparse)).toEqual([[1, 1, 7.5, 20], [8.5, 1, 7.5, 20]]);
    const masonry = planAssembly(items(5), plan({ layout: "masonry", columns: "auto", gapCm: 0 }), WIDE);
    expect(new Set(masonry.map((item) => item.widthCm))).toEqual(new Set([10]));
  });
});

describe("gaps", () => {
  it("defaults to 0.2 cm", () => {
    expect(DEFAULT_GAP_CM).toBe(0.2);
    expect(ASSEMBLY_PLAN_DEFAULTS.gapCm).toBe(0.2);
    expect(ASSEMBLY_PLAN_DEFAULTS.cellBackground).toBe("transparent");
  });

  it("spaces grid cells by the gap and never at the content box's edge", () => {
    const placed = planAssembly(items(4), plan({ columns: 2, fit: "stretch" }), WIDE);
    // (30 − 0.2) / 2 = 14.9 across, (20 − 0.2) / 2 = 9.9 down.
    expect(rects(placed)).toEqual([
      [1, 1, 14.9, 9.9],
      [16.1, 1, 14.9, 9.9],
      [1, 11.1, 14.9, 9.9],
      [16.1, 11.1, 14.9, 9.9],
    ]);
    expect(placed[1].xCm - (placed[0].xCm + placed[0].widthCm)).toBeCloseTo(0.2, 9);
    expect(placed[3].xCm + placed[3].widthCm).toBe(31);
    expect(placed[3].yCm + placed[3].heightCm).toBe(21);
  });

  it("is never added around a single cell", () => {
    for (const layout of ["grid", "horizontal", "vertical", "masonry"] as const) {
      const [only] = planAssembly(items(1), plan({ layout, gapCm: 3, fit: "stretch" }), WIDE);
      expect(only.cell).toEqual(layout === "masonry" ? { xCm: 1, yCm: 1, widthCm: 30, heightCm: 15 } : WIDE.content);
    }
  });

  it("spaces the horizontal row and the vertical column edge to edge", () => {
    const row = planAssembly(items(3), plan({ layout: "horizontal", gapCm: 0.5, fit: "stretch" }), WIDE);
    // 29 cm shared 2 : 1 : 0.5.
    expect(rects(row)).toEqual([
      [1, 1, 16.571, 20],
      [18.071, 1, 8.286, 20],
      [26.857, 1, 4.143, 20],
    ]);
    expect(row[2].xCm + row[2].widthCm).toBe(31);
    const column = planAssembly(items(3), plan({ layout: "vertical", gapCm: 0.5, fit: "stretch" }), TALL);
    // 29 cm shared 0.5 : 1 : 2.
    expect(rects(column)).toEqual([
      [1, 1, 20, 4.143],
      [1, 5.643, 20, 8.286],
      [1, 14.429, 20, 16.571],
    ]);
    expect(column[2].yCm + column[2].heightCm).toBe(31);
  });

  it("spaces masonry columns and the items within a column, not above the first", () => {
    const placed = planAssembly(
      Array.from({ length: 4 }, (_, i) => ({ id: `s${i}`, widthPx: 100, heightPx: 100 })),
      plan({ layout: "masonry", columns: 2, gapCm: 0.4 }),
      { ...WIDE, content: { xCm: 1, yCm: 1, widthCm: 30.4, heightCm: 40 } },
    );
    expect(rects(placed)).toEqual([
      [1, 1, 15, 15],
      [16.4, 1, 15, 15],
      [1, 16.4, 15, 15],
      [16.4, 16.4, 15, 15],
    ]);
  });

  it("does not space overlay items: every one is on the whole content box", () => {
    const placed = planAssembly(items(2), plan({ layout: "overlay", gapCm: 5, fit: "stretch" }), WIDE);
    expect(rects(placed)).toEqual([[1, 1, 30, 20], [1, 1, 30, 20]]);
  });
});

describe("fit, align and the cell background, within each cell", () => {
  it("aligns a contained item inside its own cell", () => {
    const placed = planAssembly(items(2), plan({ columns: 2, gapCm: 0, align: "top-left" }), WIDE);
    // Cells 15 × 20; 2:1 → 15 × 7.5 at the cell's top-left; 1:1 → 15 × 15.
    expect(rects(placed)).toEqual([
      [1, 1, 15, 7.5],
      [16, 1, 15, 15],
    ]);
    const bottomRight = planAssembly(items(2), plan({ columns: 2, gapCm: 0, align: "bottom-right" }), WIDE);
    expect(rects(bottomRight)).toEqual([
      [1, 13.5, 15, 7.5],
      [16, 6, 15, 15],
    ]);
    expect(placed.map((item) => item.cell)).toEqual([
      { xCm: 1, yCm: 1, widthCm: 15, heightCm: 20 },
      { xCm: 16, yCm: 1, widthCm: 15, heightCm: 20 },
    ]);
  });

  it("crops a cover fit to the cell, the source rectangle naming the surviving pixels", () => {
    const [wide] = planAssembly([{ id: "w", widthPx: 200, heightPx: 100 }], plan({ columns: 2, gapCm: 0, fit: "cover" }), WIDE);
    // Cell 15 × 20: scale max(15/200, 20/100) = 0.2 → 40 × 20, centred:
    // the middle 75 px of the 200 survive.
    expect(rects([wide])).toEqual([[1, 1, 15, 20]]);
    expect(wide.sourcePx).toEqual({ x: 62.5, y: 0, width: 75, height: 100 });
    const [left] = planAssembly([{ id: "w", widthPx: 200, heightPx: 100 }], plan({ columns: 2, gapCm: 0, fit: "cover", align: "left" }), WIDE);
    expect(left.sourcePx).toEqual({ x: 0, y: 0, width: 75, height: 100 });
  });

  it("places an actual-size item at one pixel per dot of the given dpi", () => {
    const [item] = planAssembly([{ id: "a", widthPx: 150, heightPx: 75 }], plan({ fit: "actual", align: "top-left" }), WIDE, 150);
    expect(rects([item])).toEqual([[1, 1, 2.54, 1.27]]);
    expect(item.sourcePx).toEqual({ x: 0, y: 0, width: 150, height: 75 });
  });

  it("carries the cell background on every placed item, with full opacity and normal blending", () => {
    const placed = planAssembly(items(3), plan({ cellBackground: "#112233" }), WIDE);
    expect(placed.map((item) => [item.cellBackground, item.opacity, item.blend])).toEqual([
      ["#112233", 1, "normal"],
      ["#112233", 1, "normal"],
      ["#112233", 1, "normal"],
    ]);
  });

  it("ignores overlay transforms outside the overlay layout", () => {
    const moved = items(2).map((item) => ({ ...item, transform: { ...DEFAULT_OVERLAY_TRANSFORM, xPercent: 0, opacity: 10, blend: "screen" as const } }));
    expect(planAssembly(moved, plan(), WIDE)).toEqual(planAssembly(items(2), plan(), WIDE));
  });

  it("chunks drawing by row: grid rows, one horizontal row, one per vertical item", () => {
    expect(planAssembly(items(5), plan({ columns: 2 }), WIDE).map((item) => item.row)).toEqual([0, 0, 1, 1, 2]);
    expect(planAssembly(items(3), plan({ layout: "horizontal" }), WIDE).map((item) => item.row)).toEqual([0, 0, 0]);
    expect(planAssembly(items(3), plan({ layout: "vertical" }), WIDE).map((item) => item.row)).toEqual([0, 1, 2]);
  });
});

describe("overlay", () => {
  const square = { id: "sq", widthPx: 100, heightPx: 100 };

  it("places every item on the whole content box, in tray order back to front", () => {
    const placed = planAssembly(items(3), plan({ layout: "overlay" }), WIDE);
    expect(placed.map((item) => [item.index, item.row])).toEqual([
      [0, 0],
      [1, 1],
      [2, 2],
    ]);
    expect(rects(placed)).toEqual(EXPECTED.overlay.wide[3]);
  });

  it("centres the item at xPercent/yPercent of the content box, scaled by scalePercent", () => {
    const [item] = planAssembly(
      [{ ...square, transform: { ...DEFAULT_OVERLAY_TRANSFORM, xPercent: 25, yPercent: 75, scalePercent: 50 } }],
      plan({ layout: "overlay" }),
      WIDE,
    );
    // Contained: 20 × 20, halved: 10 × 10, centred at (1 + 7.5, 1 + 15).
    expect(rects([item])).toEqual([[3.5, 11, 10, 10]]);
    expect(item.sourcePx).toEqual({ x: 0, y: 0, width: 100, height: 100 });
  });

  it("clips an item pushed past the content box, cropping its source", () => {
    const [item] = planAssembly(
      [{ ...square, transform: { ...DEFAULT_OVERLAY_TRANSFORM, xPercent: 100, yPercent: 0 } }],
      plan({ layout: "overlay" }),
      WIDE,
    );
    // 20 × 20 centred on the top-right corner (31, 1): its bottom-left quarter shows.
    expect(rects([item])).toEqual([[21, 1, 10, 10]]);
    expect(item.sourcePx).toEqual({ x: 0, y: 50, width: 50, height: 50 });
    expect(item.cell).toEqual({ xCm: 21, yCm: 1, widthCm: 10, heightCm: 10 });
  });

  it("carries opacity (0–100 → 0–1) and blend per item", () => {
    const placed = planAssembly(
      [
        { ...square, transform: { ...DEFAULT_OVERLAY_TRANSFORM, opacity: 40, blend: "multiply" } },
        { ...square, transform: { ...DEFAULT_OVERLAY_TRANSFORM, opacity: 0, blend: "screen" } },
        { ...square, transform: { ...DEFAULT_OVERLAY_TRANSFORM, blend: "overlay" } },
        square,
      ],
      plan({ layout: "overlay" }),
      WIDE,
    );
    expect(placed.map((item) => [item.opacity, item.blend])).toEqual([
      [0.4, "multiply"],
      [0, "screen"],
      [1, "overlay"],
      [1, "normal"],
    ]);
  });

  it("builds the base size from the plan's fit", () => {
    const transform = { ...DEFAULT_OVERLAY_TRANSFORM, scalePercent: 50 };
    const [stretched] = planAssembly([{ ...square, transform }], plan({ layout: "overlay", fit: "stretch" }), WIDE);
    expect(rects([stretched])).toEqual([[8.5, 6, 15, 10]]);
    const [covered] = planAssembly([{ ...square, transform }], plan({ layout: "overlay", fit: "cover" }), WIDE);
    expect(rects([covered])).toEqual([[8.5, 3.5, 15, 15]]);
  });

  it("normalises a transform into its ranges", () => {
    expect(normaliseTransform(undefined)).toEqual(DEFAULT_OVERLAY_TRANSFORM);
    expect(
      normaliseTransform({ xPercent: -5, yPercent: 140, scalePercent: 0, opacity: 250, blend: "lighten" as never }),
    ).toEqual({ xPercent: 0, yPercent: 100, scalePercent: 1, opacity: 100, blend: "normal" });
    expect(normaliseTransform({ xPercent: Number.NaN, scalePercent: 5000 })).toMatchObject({ xPercent: 50, scalePercent: 1000 });
  });
});

describe("the empty tray, bounds and purity", () => {
  it("places nothing for an empty tray, in every layout, without an error", () => {
    for (const layout of LAYOUT_KINDS) {
      expect(planAssembly([], plan({ layout }), WIDE)).toEqual([]);
      expect(planAssembly([], plan({ layout, columns: 3 }), TALL)).toEqual([]);
    }
    expect(naturalAssemblyPx([], plan())).toBeNull();
  });

  it("holds at most 100 items per tray", () => {
    expect(MAX_ASSEMBLY_ITEMS).toBe(100);
    expect(planAssembly(items(MAX_ASSEMBLY_ITEMS), plan(), WIDE)).toHaveLength(100);
  });

  it("is pure: deep-frozen inputs, the same answer twice", () => {
    const input = deepFreeze(items(12).map((item, i) => ({ ...item, transform: { ...DEFAULT_OVERLAY_TRANSFORM, xPercent: i * 8 } })));
    for (const layout of LAYOUT_KINDS) {
      const frozenPlan = deepFreeze(plan({ layout, columns: layout === "masonry" ? 3 : "auto" }));
      const sheet = deepFreeze(structuredClone(WIDE));
      const first = planAssembly(input, frozenPlan, sheet);
      expect(planAssembly(input, frozenPlan, sheet)).toEqual(first);
    }
  });

  it("reads only each item's pixel size", () => {
    const withExtras = items(3).map((item) => ({ ...item, bytes: new Uint8Array([1, 2, 3]), title: "x" }));
    expect(planAssembly(withExtras, plan(), WIDE)).toEqual(planAssembly(items(3), plan(), WIDE));
  });
});

describe("naturalAssemblyPx (Fit to content)", () => {
  const three = items(3);
  it("sizes the arrangement from the items' own pixels", () => {
    expect(naturalAssemblyPx(three, plan({ columns: 2 }))).toEqual({ widthPx: 400, heightPx: 400 });
    expect(naturalAssemblyPx(three, plan({ layout: "horizontal" }))).toEqual({ widthPx: 700, heightPx: 200 });
    expect(naturalAssemblyPx(three, plan({ layout: "vertical" }))).toEqual({ widthPx: 200, heightPx: 700 });
    expect(naturalAssemblyPx(three, plan({ layout: "masonry", columns: 2 }))).toEqual({ widthPx: 400, heightPx: 500 });
    expect(naturalAssemblyPx(three, plan({ layout: "overlay" }))).toEqual({ widthPx: 200, heightPx: 200 });
  });
});
