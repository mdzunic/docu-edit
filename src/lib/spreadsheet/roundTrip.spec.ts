// SPEC-116's round-trip promise (kept by SPEC-117's and SPEC-118's readers,
// whose styles and formats fixtures join the list), over every fixture (the ones built here and
// the Numbers export a person committed): downloading an untouched workbook
// re-emits every source part byte-identical, in source order — the ZIP
// container itself is rebuilt, so the relation is part-level equalBytes, as
// the presentation suites assert — `.ods` keeps `mimetype` first and STOREd,
// a macro-enabled workbook keeps its VBA project and content types, the file
// name and MIME type follow format + package variant, `.ods` splice equality
// holds at read, and the model survives the store's JSON trip unchanged.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { getCell, parseA1 } from "./address";
import { spreadsheetDownload, spreadsheetFileName } from "./download";
import {
  buildMiniBookOdsBytes,
  buildMiniBookXlsxBytes,
  buildSpec1904XlsxBytes,
  buildSpecMacrosXlsmBytes,
  buildSpecTemplateXltxBytes,
  buildSpecWideOdsBytes,
  buildSpec117StylesOdsBytes,
  buildSpec117StylesXlsxBytes,
  buildSpec118FormatsOdsBytes,
  buildSpec118FormatsXlsxBytes,
  buildSpec118GeneratedXlsxBytes,
  buildSpec119FormulasOdsBytes,
  buildSpec119FormulasXlsxBytes,
  SPEC116_VBA_PROJECT,
} from "./fixtures/build";
import type { PackageVariant, Sheet, SpreadsheetFormat, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { workbookForCompare } from "./fixtures/workbookForCompare";
import { addSheet, createBlankWorkbook, duplicateSheet, moveSheet, renameSheet, setSheetHidden, setTabColor } from "./sheets";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const NUMBERS_XLSX = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec116-numbers.xlsx")));

interface Fixture {
  file: string;
  format: SpreadsheetFormat;
  variant?: PackageVariant;
  extension: string;
  mimeType: string;
  bytes: () => Promise<Uint8Array>;
}

const FIXTURES: Fixture[] = [
  {
    file: "mini-book.xlsx",
    format: "xlsx",
    extension: "xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    bytes: buildMiniBookXlsxBytes,
  },
  {
    file: "mini-book.ods",
    format: "ods",
    extension: "ods",
    mimeType: "application/vnd.oasis.opendocument.spreadsheet",
    bytes: buildMiniBookOdsBytes,
  },
  {
    file: "spec116-1904.xlsx",
    format: "xlsx",
    extension: "xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    bytes: buildSpec1904XlsxBytes,
  },
  {
    file: "spec116-macros.xlsm",
    format: "xlsx",
    variant: "macroEnabled",
    extension: "xlsm",
    mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12",
    bytes: buildSpecMacrosXlsmBytes,
  },
  {
    file: "spec116-template.xltx",
    format: "xlsx",
    variant: "template",
    extension: "xltx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    bytes: buildSpecTemplateXltxBytes,
  },
  {
    file: "spec116-wide.ods",
    format: "ods",
    extension: "ods",
    mimeType: "application/vnd.oasis.opendocument.spreadsheet",
    bytes: buildSpecWideOdsBytes,
  },
  {
    file: "spec116-numbers.xlsx",
    format: "xlsx",
    extension: "xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    bytes: async () => NUMBERS_XLSX,
  },
  // SPEC-117: reading every style, size, merge and view setting writes
  // nothing — both styles fixtures still download byte-identical.
  {
    file: "spec117-styles.xlsx",
    format: "xlsx",
    extension: "xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    bytes: buildSpec117StylesXlsxBytes,
  },
  {
    file: "spec117-styles.ods",
    format: "ods",
    extension: "ods",
    mimeType: "application/vnd.oasis.opendocument.spreadsheet",
    bytes: buildSpec117StylesOdsBytes,
  },
  // SPEC-118: reading number formats — and adding the implied-format styles
  // an unformatted date or percentage gets — writes nothing either.
  {
    file: "spec118-formats.xlsx",
    format: "xlsx",
    extension: "xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    bytes: buildSpec118FormatsXlsxBytes,
  },
  {
    file: "spec118-generated.xlsx",
    format: "xlsx",
    extension: "xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    bytes: buildSpec118GeneratedXlsxBytes,
  },
  {
    file: "spec118-formats.ods",
    format: "ods",
    extension: "ods",
    mimeType: "application/vnd.oasis.opendocument.spreadsheet",
    bytes: buildSpec118FormatsOdsBytes,
  },
  // SPEC-119: reading formulas — shared, array, data-table, dynamic-array and
  // external ones, the three ODS dialects — changes the model, never a byte.
  {
    file: "spec119-formulas.xlsx",
    format: "xlsx",
    extension: "xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    bytes: buildSpec119FormulasXlsxBytes,
  },
  {
    file: "spec119-formulas.ods",
    format: "ods",
    extension: "ods",
    mimeType: "application/vnd.oasis.opendocument.spreadsheet",
    bytes: buildSpec119FormulasOdsBytes,
  },
];

/** The 0-based stored cell at an A1 address of a sheet (undefined = empty). */
function at(sheet: Sheet, a1: string): unknown {
  const ref = parseA1(a1)!;
  return getCell(sheet, ref.row, ref.col)?.value;
}

/** The STORE flag of a ZIP's first local entry, read off its header. */
function firstEntry(bytes: Uint8Array): { name: string; stored: boolean } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const method = view.getUint16(8, true);
  const nameLength = view.getUint16(26, true);
  return { name: new TextDecoder().decode(bytes.subarray(30, 30 + nameLength)), stored: method === 0 };
}

describe("every fixture downloads part for part", () => {
  for (const fixture of FIXTURES) {
    it(`${fixture.file}: every source part byte-identical and in order; .${fixture.extension}, ${fixture.mimeType}`, async () => {
      const source = await fixture.bytes();
      const model = await readSpreadsheetBytes(source, fixture.format, "fixture");
      expect(model.format).toBe(fixture.format);
      expect(model.packageVariant).toBe(fixture.variant);

      const download = await spreadsheetDownload(model);
      expect(download.fileName).toBe(`fixture.${fixture.extension}`);
      expect(download.mimeType).toBe(fixture.mimeType);
      const out = await unzip(download.bytes);
      const original = await unzip(source);
      expect([...out.keys()]).toEqual([...original.keys()]);
      for (const [name, bytes] of original) {
        expect(equalBytes(out.get(name)!, bytes), name).toBe(true);
      }

      // And after the store's JSON trip — the model is plain data.
      const stored = JSON.parse(JSON.stringify(model)) as SpreadsheetModel;
      expect(stored).toEqual(model);
      const again = await unzip((await spreadsheetDownload(stored)).bytes);
      for (const [name, bytes] of original) {
        expect(equalBytes(again.get(name)!, bytes), name).toBe(true);
      }
    });
  }
});

describe("package rules", () => {
  it(".ods keeps mimetype first and STOREd", async () => {
    for (const build of [buildMiniBookOdsBytes, buildSpecWideOdsBytes]) {
      const download = await spreadsheetDownload(await readSpreadsheetBytes(await build(), "ods", "book"));
      expect(firstEntry(download.bytes)).toEqual({ name: "mimetype", stored: true });
    }
  });

  it("a macro-enabled workbook keeps xl/vbaProject.bin, its relationship and its content type", async () => {
    const model = await readSpreadsheetBytes(await buildSpecMacrosXlsmBytes(), "xlsx", "Macros");
    const parts = await unzip((await spreadsheetDownload(model)).bytes);
    expect(equalBytes(parts.get("xl/vbaProject.bin")!, SPEC116_VBA_PROJECT)).toBe(true);
    const decode = (name: string) => new TextDecoder().decode(parts.get(name));
    expect(decode("[Content_Types].xml")).toContain(`ContentType="application/vnd.ms-office.vbaProject"`);
    expect(decode("[Content_Types].xml")).toContain("application/vnd.ms-excel.sheet.macroEnabled.main+xml");
    expect(decode("xl/_rels/workbook.xml.rels")).toContain("relationships/vbaProject");
  });

  it("names the download by format + variant (xlsx/xlsm/xltx/xltm/ods/ots), titled or not", async () => {
    const model = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "Budget 2026");
    const cases: Array<[SpreadsheetFormat, PackageVariant | undefined, string]> = [
      ["xlsx", undefined, "xlsx"],
      ["xlsx", "macroEnabled", "xlsm"],
      ["xlsx", "template", "xltx"],
      ["xlsx", "macroEnabledTemplate", "xltm"],
      ["ods", undefined, "ods"],
      ["ods", "template", "ots"],
    ];
    for (const [format, packageVariant, extension] of cases) {
      expect(spreadsheetFileName({ ...model, format, packageVariant })).toBe(`Budget 2026.${extension}`);
    }
    expect(spreadsheetFileName({ ...model, title: "  " })).toBe("Untitled spreadsheet.xlsx");
  });
});

describe(".ods splice equality at read", () => {
  for (const build of [buildMiniBookOdsBytes, buildSpecWideOdsBytes, buildSpec119FormulasOdsBytes]) {
    it(`${build.name}: prefix + every sheet's sourceXml + suffix is content.xml exactly`, async () => {
      const bytes = await build();
      const model = await readSpreadsheetBytes(bytes, "ods", "book");
      const content = new TextDecoder().decode((await unzip(bytes)).get("content.xml"));
      const envelope = model.odsContentEnvelope!;
      expect(envelope.prefix + model.sheets.map((sheet) => sheet.sourceXml).join("") + envelope.suffix).toBe(content);
    });
  }
});

describe("the Numbers export reads (e2e/fixtures/spec116-numbers.md)", () => {
  it("reads Numbers' three worksheets and the values its note records", async () => {
    const model = await readSpreadsheetBytes(NUMBERS_XLSX, "xlsx", "spec116-numbers");
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(["Export Summary", "Summary - Figures", "Second - Notes"]);
    expect(model.sheets.every((sheet) => sheet.kind === "grid" && sheet.visibility === "visible")).toBe(true);
    expect(model.activeSheetIndex).toBe(0);
    expect(model.dateSystem).toBe("1900");

    const figures = model.sheets[1]!;
    // The table name sits in the merged title row; the data lands at A2..C6.
    expect(at(figures, "A1")).toBe("Figures");
    expect(at(figures, "B2")).toBe("Amount");
    expect(at(figures, "A3")).toBe("Revenue");
    expect(at(figures, "B3")).toBe(1250.5);
    expect(at(figures, "B4")).toBe(400);
    // 2024-03-15 01:00 (the CET offset Numbers wrote), not local midnight.
    expect(at(figures, "C3")).toBe(45366.041666666664);
    // The dynamic-array formula contributes its cached value.
    expect(at(figures, "B5")).toBe(850.5);
    expect(at(figures, "B6")).toBe("Plain text, with a comma");
    // Cells with a style and no value are stored as style-only cells
    // (SPEC-117; SPEC-116's reader dropped them).
    const styleOnly = getCell(figures, 3, 2);
    expect(styleOnly?.value).toBeNull();
    expect(styleOnly?.styleId).toEqual(expect.any(Number));

    const notes = model.sheets[2]!;
    expect(at(notes, "B3")).toBe(42);
    // Numbers exports the boolean true as the number 1.
    expect(at(notes, "B4")).toBe(1);

    const summary = model.sheets[0]!;
    expect(at(summary, "B3")).toMatch(/^This document was exported from Numbers/);
    expect(at(summary, "D10")).toBe("Summary - Figures");
  });
});

// --- SPEC-121 ("Tests": roundTrip.spec.ts) -----------------------------------------------

/** Puts every model field back as `snapshot` had it — what undoing every
 * command restores (structureDirty and each sheet's dirty included). */
function restore(model: SpreadsheetModel, snapshot: SpreadsheetModel): void {
  for (const key of Object.keys(model) as (keyof SpreadsheetModel)[]) {
    if (!(key in snapshot)) {
      delete model[key];
    }
  }
  Object.assign(model, structuredClone(snapshot));
}

describe("SPEC-121: from-scratch write → read", () => {
  for (const format of ["xlsx", "ods"] as const) {
    it(`.${format}: a new workbook and an edited one read back to the same workbook (ids masked)`, async () => {
      const blank: SpreadsheetModel = { ...createBlankWorkbook(), format };
      const edited: SpreadsheetModel = { ...createBlankWorkbook(), format };
      addSheet(edited, 0);
      addSheet(edited, 1);
      renameSheet(edited, 1, "Budget");
      setTabColor(edited, 1, { theme: "accent5", lumMod: 0.75 });
      setSheetHidden(edited, [0], true);
      for (const model of [blank, edited]) {
        const { bytes } = await spreadsheetDownload(model);
        expect(await packageIntegrity(bytes)).toEqual([]);
        expect(workbookForCompare(await readSpreadsheetBytes(bytes, format, "copy"))).toEqual(workbookForCompare(model));
      }
    });
  }
});

describe("SPEC-121: no-op and undo-to-origin saves are equalBytes per part", () => {
  for (const fixture of FIXTURES) {
    it(`${fixture.file}`, async () => {
      const source = await fixture.bytes();
      const original = await unzip(source);
      const model = await readSpreadsheetBytes(source, fixture.format, "fixture");
      const expectVerbatim = async (): Promise<void> => {
        const out = await unzip((await spreadsheetDownload(model)).bytes);
        expect([...out.keys()]).toEqual([...original.keys()]);
        for (const [name, bytes] of original) {
          expect(equalBytes(out.get(name)!, bytes), name).toBe(true);
        }
      };
      await expectVerbatim();
      const snapshot = structuredClone(model);
      addSheet(model, model.activeSheetIndex);
      renameSheet(model, 0, "Renamed by a test");
      duplicateSheet(model, model.activeSheetIndex);
      moveSheet(model, model.activeSheetIndex, 0);
      setTabColor(model, 0, "#336699");
      setSheetHidden(model, [0], true);
      expect(model.structureDirty).toBe(true);
      restore(model, snapshot);
      await expectVerbatim();
    });
  }
});

describe("SPEC-121: untouched parts survive structural commands", () => {
  for (const fixture of FIXTURES.filter((each) => each.format === "xlsx")) {
    it(`${fixture.file}: every worksheet part the commands did not touch is equalBytes its source`, async () => {
      const source = await fixture.bytes();
      const original = await unzip(source);
      const model = await readSpreadsheetBytes(source, "xlsx", "fixture");
      const active = model.sheets[model.activeSheetIndex]!;
      addSheet(model, model.sheets.length - 1);
      const others = model.sheets.flatMap((sheet, index) => (sheet !== active && sheet.visibility === "visible" && sheet.partName !== undefined ? [index] : []));
      if (others.length > 1) {
        setSheetHidden(model, [others[0]!], true);
      }
      model.activeSheetIndex = model.sheets.indexOf(active); // view state: the same sheet stays active
      const bytes = (await spreadsheetDownload(model)).bytes;
      expect(await packageIntegrity(bytes)).toEqual([]);
      const out = await unzip(bytes);
      for (const sheet of model.sheets.filter((each) => each.partName !== undefined)) {
        const before = new TextDecoder().decode(original.get(sheet.partName!));
        const selected = /^(1|true)$/.test(/<(?:\w+:)?sheetView\b[^>]*\btabSelected="([^"]*)"/.exec(before)?.[1] ?? "");
        if (selected === (sheet === active)) {
          expect(equalBytes(out.get(sheet.partName!)!, original.get(sheet.partName!)!), sheet.partName).toBe(true);
        }
      }
    });
  }

  for (const fixture of FIXTURES.filter((each) => each.format === "ods")) {
    it(`${fixture.file}: content.xml is the prefix (+ generated styles) + every untouched table's source + the new table + the suffix`, async () => {
      const model = await readSpreadsheetBytes(await fixture.bytes(), "ods", "fixture");
      const envelope = model.odsContentEnvelope!;
      const tables = model.sheets.map((sheet) => sheet.sourceXml!).join("");
      const active = model.activeSheetIndex;
      addSheet(model, model.sheets.length - 1);
      model.activeSheetIndex = active;
      const bytes = (await spreadsheetDownload(model)).bytes;
      expect(await packageIntegrity(bytes)).toEqual([]);
      const content = new TextDecoder().decode((await unzip(bytes)).get("content.xml"));
      const start = content.indexOf(tables);
      expect(start).toBeGreaterThan(0);
      expect(content.slice(0, start).replace(/<style:style style:name="de-[a-z]+\d+"[\s\S]*?<\/style:style>/g, "")).toBe(envelope.prefix);
      const rest = content.slice(start + tables.length);
      expect(rest.endsWith(envelope.suffix)).toBe(true);
      expect(rest.slice(0, rest.length - envelope.suffix.length)).toMatch(/^<table:table table:name="Sheet\d+" table:style-name="de-ta\d+">.*<\/table:table>$/);
    });
  }
});
