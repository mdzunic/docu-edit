import { describe, expect, it } from "vitest";
import type { Row } from "./model";
import { chunkKeySegment, chunkOf, chunkRows, firstRowAtOrAfter, rowsInChunk, SHEET_ROWS_CHUNK } from "./rowChunks";

const row = (index: number): Row => ({ index, cells: [{ col: 0, value: index }] });

describe("row chunks (SPEC-137)", () => {
  it("bands row indices in SHEET_ROWS_CHUNK = 512", () => {
    expect(SHEET_ROWS_CHUNK).toBe(512);
    expect(chunkOf(0)).toBe(0);
    expect(chunkOf(511)).toBe(0);
    expect(chunkOf(512)).toBe(1);
    expect(chunkOf(1_048_575)).toBe(2047);
  });

  it("pads a chunk to four digits for its record key", () => {
    expect(chunkKeySegment(0)).toBe("0000");
    expect(chunkKeySegment(48)).toBe("0048");
    expect(chunkKeySegment(2047)).toBe("2047");
  });

  it("finds a band's rows by index, keeping the row objects", () => {
    const rows = [row(0), row(3), row(511), row(512), row(2000), row(5000)];
    expect(firstRowAtOrAfter(rows, 512)).toBe(3);
    expect(firstRowAtOrAfter(rows, 6000)).toBe(6);
    expect(rowsInChunk(rows, 0)).toEqual([rows[0], rows[1], rows[2]]);
    expect(rowsInChunk(rows, 0)[2]).toBe(rows[2]);
    expect(rowsInChunk(rows, 1)).toEqual([rows[3]]);
    expect(rowsInChunk(rows, 2)).toEqual([]);
  });

  it("cuts sparse rows into their non-empty bands, in order", () => {
    const rows = [row(0), row(3), row(512), row(2000), row(5000)];
    expect(chunkRows(rows).map(([chunk, band]) => [chunk, band.map((each) => each.index)])).toEqual([
      [0, [0, 3]],
      [1, [512]],
      [3, [2000]],
      [9, [5000]],
    ]);
    expect(chunkRows([])).toEqual([]);
  });
});
