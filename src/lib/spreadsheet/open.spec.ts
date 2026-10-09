// SPEC-116 "Reader versions": SPREADSHEET_READER_VERSION, the (empty)
// upgrade list, assembly, and upgradeSpreadsheetModel — a current record
// comes back untouched; an older one with no dirty sheet is re-read from its
// retained package keeping its id, title and every sheet's id (by position);
// a dirty sheet instead runs every pending step, in order.
import { afterEach, describe, expect, it } from "vitest";
import { getCell } from "./address";
import { buildMiniBookOdsBytes, buildMiniBookXlsxBytes } from "./fixtures/build";
import type { Sheet, SpreadsheetModel } from "./model";
import {
  assembleSpreadsheet,
  READER_UPGRADES,
  readSpreadsheetBytes,
  SPREADSHEET_READER_VERSION,
  upgradeSpreadsheetModel,
  type ReaderUpgrade,
} from "./open";

/** A record as an older reader would have stored it: version 0, the same
 * package retained, but nothing extracted from it yet (empty rows). */
async function versionZero(build: () => Promise<Uint8Array>, format: "xlsx" | "ods"): Promise<SpreadsheetModel> {
  const model = await readSpreadsheetBytes(await build(), format, "Renamed by the user");
  return {
    ...model,
    readerVersion: 0,
    sheets: model.sheets.map((sheet, index) => ({ ...sheet, id: `stored-sheet-${index}`, rows: [] })),
  };
}

afterEach(() => {
  READER_UPGRADES.length = 0;
});

describe("reader versions", () => {
  it("starts at version 1 with no upgrades", () => {
    expect(SPREADSHEET_READER_VERSION).toBe(1);
    expect(READER_UPGRADES).toEqual([]);
  });

  it("stamps every fresh read with the current version and a card preview", async () => {
    const model = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
    expect(model.kind).toBe("spreadsheet");
    expect(model.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(model.title).toBe("mini-book");
    expect(model.preview?.sheetName).toBe("Summary");
    expect(model.preview?.cells[1]?.[1]).toBe("Revenue");
  });

  it("omits empty optional fields when assembling", () => {
    const model = assembleSpreadsheet("t", "xlsx", {
      sheets: [{ id: "s", name: "S", visibility: "visible", kind: "grid", rows: [] }],
      activeSheetIndex: 0,
      dateSystem: "1900",
      opaque: [],
      properties: {},
    });
    expect(Object.keys(model).toSorted()).toEqual(
      ["activeSheetIndex", "dateSystem", "format", "id", "kind", "opaque", "preview", "readerVersion", "sheets", "title"].toSorted(),
    );
  });

  it("returns a current record as the same object", async () => {
    const model = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
    expect(await upgradeSpreadsheetModel(model)).toBe(model);
  });

  for (const [format, build] of [
    ["xlsx", buildMiniBookXlsxBytes],
    ["ods", buildMiniBookOdsBytes],
  ] as const) {
    it(`re-reads a version-0 ${format} record from opaque, keeping id, title and sheet ids`, async () => {
      const stored = await versionZero(build, format);
      const upgraded = await upgradeSpreadsheetModel(stored);
      expect(upgraded).not.toBe(stored);
      expect(upgraded.readerVersion).toBe(SPREADSHEET_READER_VERSION);
      expect(upgraded.id).toBe(stored.id);
      expect(upgraded.title).toBe("Renamed by the user");
      expect(upgraded.sheets.map((sheet) => sheet.id)).toEqual(stored.sheets.map((sheet) => sheet.id));
      expect(upgraded.opaque).toEqual(stored.opaque);
      // What the retained source holds is extracted again.
      expect(upgraded.sheets.some((sheet) => sheet.rows.length > 0)).toBe(true);
      expect(upgraded.preview?.cells.length).toBeGreaterThan(0);
    });
  }

  it("runs every pending step, in order, on a dirty sheet and re-reads the clean ones", async () => {
    const stored = await versionZero(buildMiniBookXlsxBytes, "xlsx");
    const dirty = stored.sheets[1] as Sheet & { dirty?: boolean };
    dirty.dirty = true;
    dirty.rows = [{ index: 0, cells: [{ col: 0, value: "edited" }] }];
    const calls: string[] = [];
    const step = (from: number): ReaderUpgrade => ({
      from,
      upgradeDirtySheet(model, sheet, sourceParts) {
        calls.push(`${from}:${sheet.id}:${sourceParts.has(sheet.partName!)}:${model.id}`);
      },
    });
    READER_UPGRADES.push(step(0));
    const upgraded = await upgradeSpreadsheetModel(stored);
    expect(calls).toEqual([`0:${dirty.id}:true:${stored.id}`]);
    // The dirty sheet keeps its edits; the clean ones are re-read.
    expect(upgraded.sheets[1]).toBe(dirty);
    expect(getCell(upgraded.sheets[1]!, 0, 0)?.value).toBe("edited");
    expect(getCell(upgraded.sheets[0]!, 1, 1)?.value).toBe("Revenue");
    expect(upgraded.sheets[0]!.id).toBe(stored.sheets[0]!.id);
  });
});
