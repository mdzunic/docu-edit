// SPEC-131 "Anchors follow cells", "Tests": the drawing layer's geometry
// (every anchor shape, hidden rows, oversize offsets), behaviours, every
// command (copy-on-write identity, minting, clamping), the `drawings`
// reference adjuster row by row — with "a row insert above moves it" per
// behaviour — the deletion rules, a cell move leaving drawings alone, and
// the recomputed points within and beyond one pixel.
import { describe, expect, it } from "vitest";
import { base64ToBytes } from "../core/packageParts";
import {
  anchorBehaviour,
  anchorFromRectPx,
  applyDrawingResult,
  deleteDrawing,
  drawingAfterEdit,
  drawingNoun,
  drawingRectPx,
  EMU_PER_PX,
  insertedImageSizePx,
  insertImage,
  insertShape,
  isDrawingShown,
  nextDrawingId,
  nextDrawingNumber,
  nextMediaId,
  pointFromPx,
  recomputedPoint,
  reorderDrawing,
  setDrawingAltText,
  setDrawingRect,
  visibleDrawings,
  withDerivedSize,
  writtenPoints,
} from "./drawings";
import { buildGridMetrics } from "./geometry";
import type { AnchorPoint, DrawingAnchor, Sheet, SheetDrawing, SpreadsheetModel } from "./model";
import { moveRange } from "./move";
import { applyStructuralEdit, type StructuralEdit } from "./structure";

const E = EMU_PER_PX;

function sheet(fields: Partial<Sheet> = {}): Sheet {
  return {
    id: "s1",
    name: "Sheet1",
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15, // 20 px rows
    merges: [],
    conditionalFormats: [],
    validations: [],
    ...fields,
  };
}

function book(sheets: Sheet[], format: SpreadsheetModel["format"] = "xlsx"): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: "m",
    title: "T",
    format,
    readerVersion: 12,
    sheets,
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
    dxfs: [],
  };
}

const point = (col: number, row: number, colOffPx = 0, rowOffPx = 0): AnchorPoint => ({ col, colOffEmu: colOffPx * E, row, rowOffEmu: rowOffPx * E });

function drawing(id: string, anchor: DrawingAnchor, extra: Partial<SheetDrawing> = {}): SheetDrawing {
  return { id, anchor, name: `Shape ${id.slice(2)}`, content: { kind: "shape", visual: { geometry: "rect" } }, ...extra };
}

const metrics = (s: Sheet = sheet()) => buildGridMetrics(s);
const from = (d: SheetDrawing | null) => d?.anchor.from;
const to = (d: SheetDrawing | null) => d?.anchor.to;

// --- the five anchor shapes ---------------------------------------------------------------

const MOVE_AND_SIZE: DrawingAnchor = { kind: "twoCell", editAs: "twoCell", from: point(1, 2, 10, 5), to: point(3, 6, 0, 0) };
const TWO_CELL_DEFAULT: DrawingAnchor = { kind: "twoCell", from: point(1, 2, 10, 5), to: point(3, 6) };
const TWO_CELL_ONE_CELL: DrawingAnchor = { kind: "twoCell", editAs: "oneCell", from: point(1, 2), to: point(3, 6), extEmu: { cx: 100 * E, cy: 50 * E } };
const TWO_CELL_ABSOLUTE: DrawingAnchor = { kind: "twoCell", editAs: "absolute", from: point(1, 2), to: point(3, 6), posEmu: { x: 70 * E, y: 45 * E }, extEmu: { cx: 30 * E, cy: 40 * E } };
const ONE_CELL: DrawingAnchor = { kind: "oneCell", from: point(2, 1, 4, 3), extEmu: { cx: 64 * E, cy: 20 * E } };
const ABSOLUTE: DrawingAnchor = { kind: "absolute", posEmu: { x: 10 * E, y: 12 * E }, extEmu: { cx: 5 * E, cy: 6 * E } };

describe("anchorBehaviour", () => {
  it("twoCell with editAs absent or twoCell → moveAndSize; oneCell or editAs oneCell → move; absolute or editAs absolute → fixed", () => {
    expect(anchorBehaviour(TWO_CELL_DEFAULT)).toBe("moveAndSize");
    expect(anchorBehaviour(MOVE_AND_SIZE)).toBe("moveAndSize");
    expect(anchorBehaviour(TWO_CELL_ONE_CELL)).toBe("move");
    expect(anchorBehaviour(ONE_CELL)).toBe("move");
    expect(anchorBehaviour(TWO_CELL_ABSOLUTE)).toBe("fixed");
    expect(anchorBehaviour(ABSOLUTE)).toBe("fixed");
  });
});

describe("drawingRectPx", () => {
  it("renders a moveAndSize anchor from its from point to its to point", () => {
    expect(drawingRectPx(MOVE_AND_SIZE, metrics())).toEqual({ left: 64 + 10, top: 40 + 5, width: 192 - 74, height: 120 - 45 });
    expect(drawingRectPx(TWO_CELL_DEFAULT, metrics())).toEqual(drawingRectPx(MOVE_AND_SIZE, metrics()));
  });

  it("renders a move anchor from its from point plus extEmu, whatever its to says", () => {
    expect(drawingRectPx(TWO_CELL_ONE_CELL, metrics())).toEqual({ left: 64, top: 40, width: 100, height: 50 });
    expect(drawingRectPx(ONE_CELL, metrics())).toEqual({ left: 132, top: 23, width: 64, height: 20 });
  });

  it("renders a fixed anchor at posEmu plus extEmu", () => {
    expect(drawingRectPx(TWO_CELL_ABSOLUTE, metrics())).toEqual({ left: 70, top: 45, width: 30, height: 40 });
    expect(drawingRectPx(ABSOLUTE, metrics())).toEqual({ left: 10, top: 12, width: 5, height: 6 });
  });

  it("clamps an offset larger than its cell for rendering only — the anchor keeps it", () => {
    const oversize: DrawingAnchor = { kind: "oneCell", from: point(0, 0, 500, 90), extEmu: { cx: E, cy: E } };
    expect(drawingRectPx(oversize, metrics())).toEqual({ left: 64, top: 20, width: 1, height: 1 });
    expect(oversize.from!.colOffEmu).toBe(500 * E);
  });

  it("a moveAndSize drawing over a hidden row shrinks by its height; move and fixed ones keep their size", () => {
    const hidden = sheet({ rows: [{ index: 3, cells: [], hidden: true }] });
    expect(drawingRectPx(MOVE_AND_SIZE, metrics(hidden)).height).toBe(drawingRectPx(MOVE_AND_SIZE, metrics()).height - 20);
    expect(drawingRectPx(TWO_CELL_ONE_CELL, metrics(hidden)).height).toBe(50);
    expect(drawingRectPx(ONE_CELL, metrics(hidden)).height).toBe(20);
    expect(drawingRectPx(TWO_CELL_ABSOLUTE, metrics(hidden)).height).toBe(40);
  });

  it("a moveAndSize drawing whose whole row span is hidden is not shown; a filter-hidden span counts", () => {
    const anchor: DrawingAnchor = { kind: "twoCell", editAs: "twoCell", from: point(0, 2), to: point(1, 4) };
    const rowsHidden = sheet({ rows: [2, 3].map((index) => ({ index, cells: [], hidden: true as const })) });
    const filtered = sheet({ rows: [2, 3].map((index) => ({ index, cells: [], filterHidden: true as const })) });
    const partly = sheet({ rows: [{ index: 2, cells: [], hidden: true }] });
    expect(isDrawingShown(drawing("dr1", anchor), metrics(rowsHidden))).toBe(false);
    expect(isDrawingShown(drawing("dr1", anchor), metrics(filtered))).toBe(false);
    expect(isDrawingShown(drawing("dr1", anchor), metrics(partly))).toBe(true);
    // A move drawing keeps its size, so it is shown.
    expect(isDrawingShown(drawing("dr2", { kind: "oneCell", from: point(0, 2), extEmu: { cx: E, cy: E } }), metrics(rowsHidden))).toBe(true);
    expect(isDrawingShown(drawing("dr3", anchor, { hidden: true }), metrics())).toBe(false);
    const s = sheet({ rows: rowsHidden.rows, drawings: [drawing("dr1", anchor), drawing("dr2", ONE_CELL)] });
    expect(visibleDrawings(s, metrics(s)).map((each) => each.id)).toEqual(["dr2"]);
  });

  it("a horizontal line (zero height) inside one row is shown", () => {
    const line: DrawingAnchor = { kind: "twoCell", editAs: "twoCell", from: point(0, 2, 0, 10), to: point(2, 2, 0, 10) };
    expect(isDrawingShown(drawing("dr1", line), metrics())).toBe(true);
    expect(drawingRectPx(line, metrics()).height).toBe(0);
  });
});

describe("anchorFromRectPx", () => {
  it("is drawingRectPx's inverse for each behaviour, keeping a like anchor's kind and editAs", () => {
    const m = metrics(sheet({ columns: [{ min: 1, max: 1, widthPx: 100 }], rows: [{ index: 2, cells: [], heightPt: 30 }] }));
    const rect = { left: 70, top: 47, width: 123, height: 61 };
    for (const like of [MOVE_AND_SIZE, TWO_CELL_ONE_CELL, TWO_CELL_ABSOLUTE, ONE_CELL, ABSOLUTE]) {
      const anchor = anchorFromRectPx(rect, anchorBehaviour(like), m, like);
      expect(anchor.kind).toBe(like.kind);
      expect(anchor.editAs).toBe(like.editAs);
      const back = drawingRectPx(anchor, m);
      expect(back.left).toBeCloseTo(rect.left, 3);
      expect(back.top).toBeCloseTo(rect.top, 3);
      expect(back.width).toBeCloseTo(rect.width, 3);
      expect(back.height).toBeCloseTo(rect.height, 3);
    }
  });

  it("without a like anchor: moveAndSize → twoCell \"twoCell\", move → twoCell \"oneCell\" with extEmu, fixed → absolute", () => {
    const rect = { left: 10, top: 10, width: 20, height: 20 };
    expect(anchorFromRectPx(rect, "moveAndSize", metrics())).toMatchObject({ kind: "twoCell", editAs: "twoCell" });
    expect(anchorFromRectPx(rect, "move", metrics())).toMatchObject({ kind: "twoCell", editAs: "oneCell", extEmu: { cx: 20 * E, cy: 20 * E } });
    expect(anchorFromRectPx(rect, "fixed", metrics())).toEqual({ kind: "absolute", posEmu: { x: 10 * E, y: 10 * E }, extEmu: { cx: 20 * E, cy: 20 * E } });
  });

  it("always produces offsets smaller than their cell — fractional edges, hidden neighbours and cell boundaries included", () => {
    const s = sheet({ columns: [{ min: 2, max: 3, hidden: true }, { min: 5, max: 5, widthPx: 7 }], rows: [{ index: 1, cells: [], hidden: true }, { index: 4, cells: [], heightPt: 3 }] });
    const m = metrics(s);
    let seed = 7;
    const next = (): number => {
      seed = (seed * 1103515245 + 12345) % 2147483648;
      return seed / 2147483648;
    };
    const edges = [0, 63.99999, 64, 127.5, 128, 135.9999999, 200];
    const xs = [...edges, ...Array.from({ length: 200 }, () => next() * 600)];
    for (const x of xs) {
      for (const y of [0, 19.999999, 20, 40, 47.9999, 60, next() * 400]) {
        const p = pointFromPx(x, y, m);
        expect(p.colOffEmu).toBeGreaterThanOrEqual(0);
        expect(p.colOffEmu).toBeLessThan(m.colWidth(p.col) * E);
        expect(p.rowOffEmu).toBeLessThan(m.rowHeight(p.row) * E);
        expect(Number.isInteger(p.colOffEmu) && Number.isInteger(p.rowOffEmu)).toBe(true);
        const anchor = anchorFromRectPx({ left: x, top: y, width: next() * 300, height: next() * 300 }, "moveAndSize", m);
        for (const corner of [anchor.from!, anchor.to!]) {
          expect(corner.colOffEmu).toBeLessThan(m.colWidth(corner.col) * E);
          expect(corner.rowOffEmu).toBeLessThan(m.rowHeight(corner.row) * E);
        }
      }
    }
  });

  it("withDerivedSize gives a move twoCell anchor extEmu and a fixed one posEmu + extEmu from the read-time geometry", () => {
    const read: DrawingAnchor = { kind: "twoCell", editAs: "oneCell", from: point(1, 2, 4, 0), to: point(3, 6) };
    expect(withDerivedSize(read, metrics())).toEqual({ ...read, extEmu: { cx: (192 - 68) * E, cy: 80 * E } });
    const fixed: DrawingAnchor = { kind: "twoCell", editAs: "absolute", from: point(1, 2), to: point(2, 3) };
    expect(withDerivedSize(fixed, metrics())).toEqual({ ...fixed, posEmu: { x: 64 * E, y: 40 * E }, extEmu: { cx: 64 * E, cy: 20 * E } });
    expect(withDerivedSize(MOVE_AND_SIZE, metrics())).toBe(MOVE_AND_SIZE);
  });
});

describe("recomputed points", () => {
  const m = metrics();
  it("keeps the source point while the derived one is within one pixel on both axes", () => {
    const source = point(1, 2, 10, 5);
    expect(recomputedPoint({ ...source, colOffEmu: source.colOffEmu + E }, source, m)).toBe(source);
    expect(recomputedPoint({ ...source, rowOffEmu: source.rowOffEmu - E }, source, m)).toBe(source);
    // The same position spelled at the next cell's corner.
    const corner = point(2, 2, 0, 0);
    expect(recomputedPoint(corner, point(1, 2, 64, 0), m)).toEqual(point(1, 2, 64, 0));
  });

  it("takes the derived point when it is more than one pixel away on either axis", () => {
    const source = point(1, 2, 10, 5);
    const beyond = { ...source, colOffEmu: source.colOffEmu + E + 1 };
    expect(recomputedPoint(beyond, source, m)).toBe(beyond);
    const below = point(1, 4, 10, 5);
    expect(recomputedPoint(below, source, m)).toBe(below);
    expect(recomputedPoint(below, undefined, m)).toBe(below);
  });

  it("writtenPoints: a move anchor's to from extEmu, a fixed one's both from posEmu — the source's when within a pixel", () => {
    const move: DrawingAnchor = { kind: "twoCell", editAs: "oneCell", from: point(1, 2), to: point(3, 4), extEmu: { cx: 128 * E, cy: 40 * E } };
    expect(writtenPoints(move, m, { to: point(3, 4) })).toEqual({ from: point(1, 2), to: point(3, 4) });
    // A taller row 2 since read: the derived to moves up a row.
    const taller = metrics(sheet({ rows: [{ index: 2, cells: [], heightPt: 30 }] }));
    expect(writtenPoints(move, taller, { to: point(3, 4) }).to).toEqual(pointFromPx(64 + 128, 40 + 40, taller));
    const fixed: DrawingAnchor = { kind: "twoCell", editAs: "absolute", from: point(1, 2), to: point(2, 3), posEmu: { x: 64 * E, y: 40 * E }, extEmu: { cx: 64 * E, cy: 20 * E } };
    expect(writtenPoints(fixed, m, { from: point(1, 2), to: point(2, 3) })).toEqual({ from: point(1, 2), to: point(2, 3) });
    expect(writtenPoints(MOVE_AND_SIZE, m)).toEqual({ from: MOVE_AND_SIZE.from, to: MOVE_AND_SIZE.to });
  });
});

// --- the adjuster -------------------------------------------------------------------------

const rows = (kind: "insertRows" | "deleteRows", at: number, count: number): StructuralEdit => ({ kind, sheetId: "s1", at, count });
const cols = (kind: "insertCols" | "deleteCols", at: number, count: number): StructuralEdit => ({ kind, sheetId: "s1", at, count });

describe("the drawings adjuster — the anchor-follow table", () => {
  const mas = drawing("dr1", { kind: "twoCell", editAs: "twoCell", from: point(1, 4, 3, 2), to: point(3, 8, 5, 6) });
  const mv = drawing("dr2", { kind: "oneCell", from: point(1, 4, 3, 2), extEmu: { cx: E, cy: E } });
  const mv2 = drawing("dr3", { kind: "twoCell", editAs: "oneCell", from: point(1, 4, 3, 2), to: point(3, 8), extEmu: { cx: E, cy: E } });
  const fx = drawing("dr4", { kind: "absolute", posEmu: { x: 0, y: 99 * E }, extEmu: { cx: E, cy: E } });

  it("insert at or before from: moveAndSize from and to += count; move from += count; fixed untouched", () => {
    expect(from(drawingAfterEdit(mas, rows("insertRows", 4, 2)))).toEqual(point(1, 6, 3, 2));
    expect(to(drawingAfterEdit(mas, rows("insertRows", 4, 2)))).toEqual(point(3, 10, 5, 6));
    expect(from(drawingAfterEdit(mv, rows("insertRows", 0, 3)))).toEqual(point(1, 7, 3, 2));
    expect(from(drawingAfterEdit(mv2, rows("insertRows", 0, 3)))).toEqual(point(1, 7, 3, 2));
    expect(to(drawingAfterEdit(mv2, rows("insertRows", 0, 3)))).toEqual(point(3, 8));
    expect(drawingAfterEdit(fx, rows("insertRows", 0, 3))).toBe(fx);
  });

  it("insert inside the span (from < at ≤ to): moveAndSize grows; move does not move", () => {
    const grown = drawingAfterEdit(mas, rows("insertRows", 8, 1));
    expect(from(grown)).toEqual(mas.anchor.from);
    expect(to(grown)).toEqual(point(3, 9, 5, 6));
    expect(drawingAfterEdit(mv, rows("insertRows", 5, 1))).toBe(mv);
  });

  it("insert below every point leaves the drawing's identity", () => {
    expect(drawingAfterEdit(mas, rows("insertRows", 9, 1))).toBe(mas);
    expect(drawingAfterEdit(mas, cols("insertCols", 4, 1))).toBe(mas);
  });

  it("delete a band containing neither point: points below shift up", () => {
    expect(from(drawingAfterEdit(mas, rows("deleteRows", 0, 2)))).toEqual(point(1, 2, 3, 2));
    expect(to(drawingAfterEdit(mas, rows("deleteRows", 0, 2)))).toEqual(point(3, 6, 5, 6));
    // A band inside the span shrinks it from below.
    const shrunk = drawingAfterEdit(mas, rows("deleteRows", 5, 2));
    expect(from(shrunk)).toEqual(mas.anchor.from);
    expect(to(shrunk)).toEqual(point(3, 6, 5, 6));
    expect(from(drawingAfterEdit(mv, rows("deleteRows", 1, 2)))).toEqual(point(1, 2, 3, 2));
    expect(drawingAfterEdit(mas, rows("deleteRows", 9, 3))).toBe(mas);
  });

  it("delete a band containing to only: to = (at, 0)", () => {
    const cut = drawingAfterEdit(mas, rows("deleteRows", 7, 5));
    expect(from(cut)).toEqual(mas.anchor.from);
    expect(to(cut)).toEqual({ ...mas.anchor.to!, row: 7, rowOffEmu: 0 });
    expect(drawingAfterEdit(mv, rows("deleteRows", 7, 5))).toBe(mv);
  });

  it("delete a band containing from only: from = (at, 0), to shifts up; move: from = (at, 0)", () => {
    const cut = drawingAfterEdit(mas, rows("deleteRows", 3, 2));
    expect(from(cut)).toEqual({ ...mas.anchor.from!, row: 3, rowOffEmu: 0 });
    expect(to(cut)).toEqual(point(3, 6, 5, 6));
    expect(from(drawingAfterEdit(mv, rows("deleteRows", 3, 2)))).toEqual({ ...mv.anchor.from!, row: 3, rowOffEmu: 0 });
  });

  it("delete a band containing both: moveAndSize is deleted; move and fixed never are", () => {
    expect(drawingAfterEdit(mas, rows("deleteRows", 2, 10))).toBeNull();
    expect(from(drawingAfterEdit(mv, rows("deleteRows", 2, 10)))).toEqual({ ...mv.anchor.from!, row: 2, rowOffEmu: 0 });
    expect(drawingAfterEdit(mv2, rows("deleteRows", 2, 10))).not.toBeNull();
    expect(drawingAfterEdit(fx, rows("deleteRows", 0, 200))).toBe(fx);
  });

  it("columns are symmetric", () => {
    expect(from(drawingAfterEdit(mas, cols("insertCols", 0, 2)))).toEqual(point(3, 4, 3, 2));
    expect(to(drawingAfterEdit(mas, cols("insertCols", 2, 1)))).toEqual(point(4, 8, 5, 6));
    expect(to(drawingAfterEdit(mas, cols("deleteCols", 3, 1)))).toEqual({ ...mas.anchor.to!, col: 3, colOffEmu: 0 });
    expect(drawingAfterEdit(mas, cols("deleteCols", 1, 3))).toBeNull();
  });

  it("a row insert above each of the three kinds moves the first two by exactly the inserted height and leaves the third", () => {
    const three = [
      drawing("dr1", { kind: "twoCell", editAs: "twoCell", from: point(1, 5, 3, 4), to: point(2, 7, 10, 2) }),
      drawing("dr2", { kind: "twoCell", editAs: "oneCell", from: point(1, 5, 3, 4), to: point(2, 7), extEmu: { cx: 50 * E, cy: 30 * E } }),
      drawing("dr3", { kind: "absolute", posEmu: { x: 64 * E, y: 100 * E }, extEmu: { cx: 50 * E, cy: 30 * E } }),
    ];
    const m = book([sheet({ drawings: three })]);
    const before = three.map((each) => drawingRectPx(each.anchor, metrics(m.sheets[0])));
    expect(applyStructuralEdit(m, rows("insertRows", 2, 2))).toEqual({ ok: true });
    const s = m.sheets[0]!;
    const after = s.drawings!.map((each) => drawingRectPx(each.anchor, metrics(s)));
    expect(after[0]).toEqual({ ...before[0], top: before[0]!.top + 40 });
    expect(after[1]).toEqual({ ...before[1], top: before[1]!.top + 40 });
    expect(after[2]).toEqual(before[2]);
    expect(s.drawings![2]).toBe(three[2]);
    expect(s.drawingsDirty).toBe(true);
    expect(s.dirty).toBe(true);
  });

  it("a row deleted under a moveAndSize drawing removes it; an ods sheet is marked dirty, not drawingsDirty", () => {
    const m = book([sheet({ drawings: [drawing("dr1", { kind: "twoCell", editAs: "twoCell", from: point(0, 3), to: point(1, 4) })] })], "ods");
    applyStructuralEdit(m, rows("deleteRows", 3, 2));
    expect(m.sheets[0]!.drawings).toEqual([]);
    expect(m.sheets[0]!.drawingsDirty).toBeUndefined();
    expect(m.sheets[0]!.dirty).toBe(true);
  });

  it("an edit on another sheet leaves the drawings' sheet by identity", () => {
    const s1 = sheet({ drawings: [drawing("dr1", MOVE_AND_SIZE)] });
    const s2 = sheet({ id: "s2", name: "Sheet2" });
    const m = book([s1, s2]);
    applyStructuralEdit(m, { kind: "insertRows", sheetId: "s2", at: 0, count: 1 });
    expect(m.sheets[0]).toBe(s1);
  });
});

describe("a move of cells (SPEC-125) leaves every drawing where it is", () => {
  it("moveRange keeps Sheet.drawings identical", () => {
    const drawings = [drawing("dr1", MOVE_AND_SIZE), drawing("dr2", ONE_CELL)];
    const m = book([sheet({ drawings, rows: [{ index: 2, cells: [{ col: 1, value: 5 }] }] })]);
    expect(moveRange(m, { fromSheetId: "s1", from: { top: 1, left: 1, bottom: 6, right: 3 }, toSheetId: "s1", toTop: 20, toLeft: 5 })).toEqual({ ok: true });
    expect(m.sheets[0]!.drawings).toBe(drawings);
  });
});

// --- commands ----------------------------------------------------------------------------------

describe("commands", () => {
  const a = drawing("dr1", MOVE_AND_SIZE);
  const b = drawing("dr2", ONE_CELL);
  const c = drawing("dr3", ABSOLUTE);
  const base = (format: SpreadsheetModel["format"] = "xlsx") => book([sheet({ drawings: [a, b, c] }), sheet({ id: "s2", name: "Other" })], format);

  it("setDrawingRect re-anchors one drawing keeping its behaviour, clamped to the grid's top-left; the others keep identity", () => {
    const m = base();
    const result = setDrawingRect(m, "s1", "dr2", { left: -20, top: 30, width: 64, height: 20 })!;
    expect(result.drawingId).toBe("dr2");
    expect(result.sheet.drawings).not.toBe(m.sheets[0]!.drawings);
    expect(result.sheet.drawings![0]).toBe(a);
    expect(result.sheet.drawings![2]).toBe(c);
    expect(result.sheet.drawings![1]!.anchor).toEqual({ kind: "oneCell", from: point(0, 1, 0, 10), extEmu: { cx: 64 * E, cy: 20 * E } });
    expect(result.sheet.drawingsDirty).toBe(true);
    expect(m.sheets[0]!.drawings![1]).toBe(b); // nothing mutated
    applyDrawingResult(m, result);
    expect(m.sheets[0]).toBe(result.sheet);
    expect(setDrawingRect(m, "s1", "nope", { left: 0, top: 0, width: 1, height: 1 })).toBeNull();
  });

  it("an ods command sets dirty instead of drawingsDirty", () => {
    const result = setDrawingRect(base("ods"), "s1", "dr1", { left: 5, top: 5, width: 10, height: 10 })!;
    expect(result.sheet.dirty).toBe(true);
    expect(result.sheet.drawingsDirty).toBeUndefined();
    expect(result.sheet.drawings![1]).toBe(b);
  });

  it("deleteDrawing removes one drawing", () => {
    const result = deleteDrawing(base(), "s1", "dr1")!;
    expect(result.sheet.drawings).toEqual([b, c]);
    expect(result.sheet.drawings![0]).toBe(b);
    expect(deleteDrawing(base(), "s2", "dr1")).toBeNull();
  });

  it("reorderDrawing: front, forward, backward, back — null when already there", () => {
    expect(reorderDrawing(base(), "s1", "dr1", "front")!.sheet.drawings!.map((d) => d.id)).toEqual(["dr2", "dr3", "dr1"]);
    expect(reorderDrawing(base(), "s1", "dr1", "forward")!.sheet.drawings!.map((d) => d.id)).toEqual(["dr2", "dr1", "dr3"]);
    expect(reorderDrawing(base(), "s1", "dr3", "backward")!.sheet.drawings!.map((d) => d.id)).toEqual(["dr1", "dr3", "dr2"]);
    expect(reorderDrawing(base(), "s1", "dr3", "back")!.sheet.drawings!.map((d) => d.id)).toEqual(["dr3", "dr1", "dr2"]);
    expect(reorderDrawing(base(), "s1", "dr3", "front")).toBeNull();
    expect(reorderDrawing(base(), "s1", "dr1", "back")).toBeNull();
    const moved = reorderDrawing(base(), "s1", "dr1", "front")!;
    expect(moved.sheet.drawings![2]).toBe(a);
  });

  it("setDrawingAltText sets it, and an empty field removes the attribute", () => {
    const set = setDrawingAltText(base(), "s1", "dr2", "A chart of sales")!;
    expect(set.sheet.drawings![1]!.altText).toBe("A chart of sales");
    const m = base();
    applyDrawingResult(m, set);
    const removed = setDrawingAltText(m, "s1", "dr2", "  ")!;
    expect("altText" in removed.sheet.drawings![1]!).toBe(false);
    expect(setDrawingAltText(base(), "s1", "dr2", "")).toBeNull();
  });

  it("setDrawingAltText keeps line breaks and tabs but drops what XML cannot carry", () => {
    const set = setDrawingAltText(base(), "s1", "dr2", "Sales\u0001 by\tregion\nQ3")!;
    expect(set.sheet.drawings![1]!.altText).toBe("Sales by\tregion\nQ3");
    expect(setDrawingAltText(base(), "s1", "dr2", "\u0002\u0003")).toBeNull();
  });

  it("insertImage lands top-left on the active cell, at natural size scaled to fit 0.8 of the visible grid, twoCell editAs oneCell, with a new media part", () => {
    const m = base();
    const bytes = new Uint8Array([1, 2, 3, 4]);
    const result = insertImage(m, "s1", { bytes, contentType: "image/png", naturalPx: { width: 1000, height: 500 }, at: { row: 3, col: 2 }, visiblePx: { width: 800, height: 600 } })!;
    const added = result.sheet.drawings!.at(-1)!;
    expect(result.drawingId).toBe(added.id);
    expect(added.id).toBe("dr4");
    expect(added.name).toBe("Picture 4"); // after the created "Shape 1".."Shape 3"
    expect(added.content).toEqual({ kind: "image", mediaId: "xl/media/image1.png" });
    expect(added.anchor.kind).toBe("twoCell");
    expect(added.anchor.editAs).toBe("oneCell");
    expect(drawingRectPx(added.anchor, metrics(result.sheet))).toEqual({ left: 128, top: 60, width: 640, height: 320 });
    expect(result.media).toHaveLength(1);
    expect(base64ToBytes(result.media![0]!.data)).toEqual(bytes);
    expect(result.media![0]!.mediaId).toBe("xl/media/image1.png");
    expect(m.media).toBeUndefined();
    expect(result.sheet.drawings!.slice(0, 3)).toEqual([a, b, c]);
  });

  it("insertedImageSizePx keeps a small image at its natural size and scales a tall one by height", () => {
    expect(insertedImageSizePx({ width: 100, height: 50 }, { width: 800, height: 600 })).toEqual({ width: 100, height: 50 });
    expect(insertedImageSizePx({ width: 100, height: 1000 }, { width: 800, height: 600 })).toEqual({ width: 48, height: 480 });
  });

  it("insertShape lands twoCell editAs twoCell, styled with the new-shape defaults, named after its definition", () => {
    const result = insertShape(base(), "s1", "rect", { left: 70, top: 30, width: 100, height: 40 })!;
    const added = result.sheet.drawings!.at(-1)!;
    expect(added.name).toBe("Rectangle 4");
    expect(added.anchor).toMatchObject({ kind: "twoCell", editAs: "twoCell" });
    expect(added.content).toEqual({ kind: "shape", visual: { geometry: "rect", fill: { kind: "solid", color: "#4472c4" }, outline: { color: "#2e5395", widthPt: 1, dash: "solid" } } });
    expect(drawingRectPx(added.anchor, metrics())).toEqual({ left: 70, top: 30, width: 100, height: 40 });
    const line = insertShape(base(), "s1", "line", { left: 0, top: 0, width: 50, height: 0 })!.sheet.drawings!.at(-1)!;
    expect(line.content.kind === "shape" && line.content.visual.fill).toBeUndefined();
  });

  it("minting: drawing ids are unique per sheet; names count past every source cNvPr id and created drawing", () => {
    const sourced = drawing("dr7", ONE_CELL, { sourceXml: `<xdr:oneCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="12" name="Shape 11"/></xdr:nvSpPr></xdr:sp></xdr:oneCellAnchor>` });
    const created = drawing("dr2", ONE_CELL, { name: "Picture 14" });
    expect(nextDrawingId(sheet({ drawings: [sourced, created] }))).toBe("dr8");
    expect(nextDrawingNumber(sheet({ drawings: [sourced] }))).toBe(12);
    expect(nextDrawingNumber(sheet({ drawings: [sourced, created] }))).toBe(15);
    expect(nextDrawingNumber(sheet())).toBe(1);
  });

  it("nextMediaId: the package path past the highest image suffix across opaque and media", () => {
    const m = { ...base(), opaque: [{ name: "xl/media/image3.jpeg", data: "" }, { name: "xl/media/other9.png", data: "" }], media: [{ mediaId: "xl/media/image4.png", contentType: "image/png", data: "" }] };
    expect(nextMediaId(m, "image/png")).toBe("xl/media/image5.png");
    expect(nextMediaId({ ...base("ods"), opaque: [{ name: "Pictures/image2.svg", data: "" }] }, "image/svg+xml")).toBe("Pictures/image3.svg");
    expect(nextMediaId(base(), "image/jpeg")).toBe("xl/media/image1.jpeg");
  });

  it("drawingNoun names each content kind", () => {
    expect(drawingNoun(a)).toBe("shape");
    expect(drawingNoun({ ...a, content: { kind: "image", mediaId: "x" } })).toBe("image");
    expect(drawingNoun({ ...a, content: { kind: "group", members: [] } })).toBe("group");
    expect(drawingNoun({ ...a, content: { kind: "chart" } })).toBe("chart");
    expect(drawingNoun({ ...a, content: { kind: "foreign", label: "Slicer" } })).toBe("object");
  });

  it("the model stays plain JSON", () => {
    const m = base();
    applyDrawingResult(m, insertImage(m, "s1", { bytes: new Uint8Array([9]), contentType: "image/gif", naturalPx: { width: 2, height: 2 }, at: { row: 0, col: 0 }, visiblePx: { width: 10, height: 10 } })!);
    expect(JSON.parse(JSON.stringify(m))).toEqual(m);
  });
});
