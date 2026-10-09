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
    expect(spreadsheetFileName({ ...model, title: "  " })).toBe("Untitled workbook.xlsx");
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
