// SPEC-116's round-trip promise (kept by SPEC-117's and SPEC-118's readers,
// whose styles and formats fixtures join the list), over every fixture (the ones built here and
// the Numbers export a person committed): downloading an untouched workbook
// re-emits every source part byte-identical, in source order — the ZIP
// container itself is rebuilt, so the relation is part-level equalBytes, as
// the presentation suites assert — `.ods` keeps `mimetype` first and STOREd,
// a macro-enabled workbook keeps its VBA project and content types, the file
// name and MIME type follow format + package variant, `.ods` splice equality
// holds at read, and the model survives the store's JSON trip unchanged.
// SPEC-122: over every fixture an edit → write → read gives the edited model
// (cells, formulas, number formats and resolved styles, and
// workbookForCompare), the other sheets' parts and tables keep their bytes,
// xl/styles.xml and xl/sharedStrings.xml only grow (appended entries and
// their counts), write → read → write is equalBytes-stable, forcing
// every .ods sheet dirty without an edit reads back to the same model, and a
// copy in the other format re-opens showing the same displayed values,
// formulas and number formats. SPEC-129: the conditional-format fixtures
// download byte-identical, a dirty sheet re-emits unedited rules verbatim,
// rules edited, saved and reopened come back with their dxfs, and a copy in
// the other format carries every modelled rule. SPEC-130: the validation
// fixtures download byte-identical, set → save → reopen reproduces the rules
// in both formats, and a copy in the other format maps every type through
// the model (an x14 rule into the ODS form; a sort-ascending list into an
// ordinary xlsx list, the one property that does not cross).
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { formatA1, formatRange, getCell, parseA1, parseRange, usedRange } from "./address";
import { setCellEntry } from "./cells";
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
  buildSpec120CalcOdsBytes,
  buildSpec120CalcXlsxBytes,
  buildSpec121LinksOdsBytes,
  buildSpec121LinksXlsxBytes,
  buildSpec122ShapesOdsBytes,
  buildSpec122ShapesXlsxBytes,
  buildSpec128FilterOdsBytes,
  buildSpec128FilterXlsxBytes,
  buildSpec129CfOdsBytes,
  buildSpec129CfXlsxBytes,
  buildSpec129LegacyOdsBytes,
  buildSpec130ValidationOdsBytes,
  buildSpec130ValidationXlsxBytes,
  SPEC116_VBA_PROJECT,
  SPEC128_FILTER_XLSX,
  SPEC129_CF_ODS,
  SPEC129_CF_XLSX,
} from "./fixtures/build";
import {
  addConditionalFormat,
  deleteRule,
  modelledDxf,
  moveRule,
  rulesByPriority,
  setFormatRanges,
  updateRule,
  type ConditionalFormat,
} from "./conditionalFormats";
import { setFilterCriteria } from "./filter";
import type { PackageVariant, Sheet, SpreadsheetFormat, SpreadsheetModel } from "./model";
import { formatCodeOf } from "./numberFormat/builtin";
import { parseFormatCode } from "./numberFormat/parse";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { formatValue, plainText, workbookFormatContext } from "./numberFormat/render";
import { readSpreadsheetBytes } from "./open";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { workbookForCompare } from "./fixtures/workbookForCompare";
import { addSheet, createBlankWorkbook, duplicateSheet, moveSheet, renameSheet, setSheetHidden, setTabColor, stripRetention } from "./sheets";
import { cellStyle, effectiveStyleId, internDxf, styleKey } from "./styles";
import { clearValidation, defaultValidationSettings, formatListLiteral, setValidation } from "./validation";

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
  // SPEC-128: reading an autofilter, its criteria, the sort records and
  // which hidden rows the filter hid writes nothing.
  {
    file: "spec128-filter.xlsx",
    format: "xlsx",
    extension: "xlsx",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    bytes: buildSpec128FilterXlsxBytes,
  },
  {
    file: "spec128-filter.ods",
    format: "ods",
    extension: "ods",
    mimeType: "application/vnd.oasis.opendocument.spreadsheet",
    bytes: buildSpec128FilterOdsBytes,
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

// --- SPEC-122 ("Tests": roundTrip.spec.ts) -----------------------------------------------

const EN_US = SPREADSHEET_LOCALES["en-US"];

/** Every fixture SPEC-122's round trips run over. */
const EDIT_FIXTURES: Pick<Fixture, "file" | "format" | "bytes">[] = [
  ...FIXTURES,
  { file: "spec120-calc.xlsx", format: "xlsx", bytes: buildSpec120CalcXlsxBytes },
  { file: "spec120-calc.ods", format: "ods", bytes: buildSpec120CalcOdsBytes },
  { file: "spec121-links.xlsx", format: "xlsx", bytes: buildSpec121LinksXlsxBytes },
  { file: "spec121-links.ods", format: "ods", bytes: buildSpec121LinksOdsBytes },
  { file: "spec122-shapes.xlsx", format: "xlsx", bytes: buildSpec122ShapesXlsxBytes },
  { file: "spec122-shapes.ods", format: "ods", bytes: buildSpec122ShapesOdsBytes },
];

/** A number as the comparisons hold it: an .ods date-time is exact to the
 * millisecond, not to the serial's last binary digit. */
function comparableNumber(n: number): number {
  return Number(n.toPrecision(12));
}

/** A sheet's content as SPEC-122's round trips compare it: per stored cell
 * its address, value, formula and array extent, the number format code its
 * effective style shows in en-US, and the rest of that style's look
 * (styleKey — where a style was read from aside); the merges; per stored
 * row its height and visibility. */
function sheetForCompare(model: SpreadsheetModel, sheet: Sheet) {
  return {
    name: sheet.name,
    cells: sheet.rows.flatMap((row) =>
      row.cells.map((cell) => {
        const style = cellStyle(model, effectiveStyleId(sheet, row.index, cell.col));
        const { numberFormat: _code, builtinFormatId: _builtin, ...look } = style;
        return {
          ref: formatA1({ row: row.index, col: cell.col }),
          value: typeof cell.value === "number" ? comparableNumber(cell.value) : cell.value,
          formula: cell.formula ?? null,
          arrayRange: cell.arrayRange ?? null,
          code: formatCodeOf(style, EN_US) ?? "General",
          look: styleKey(look),
        };
      }),
    ),
    rows: sheet.rows.flatMap((row) => (row.heightPt !== undefined || row.hidden ? [{ index: row.index, heightPt: row.heightPt ?? null, hidden: row.hidden === true }] : [])),
    merges: sheet.merges,
  };
}

/** Edits the first grid sheet through SPEC-122's cell command: a number with
 * a format, a two-line wrapped string and a formula in an empty row below
 * the data, and the first plain value cell (outside merges and arrays)
 * replaced — the sheet marked dirty, as the edit pipeline would. Returns
 * the sheet's index. */
function editFirstGridSheet(model: SpreadsheetModel): number {
  const index = model.sheets.findIndex((sheet) => sheet.kind === "grid");
  const sheet = model.sheets[index]!;
  const below = (usedRange(sheet)?.bottom ?? -1) + 2;
  const inArray = (row: number, col: number): boolean =>
    sheet.rows.some((each) =>
      each.cells.some((cell) => cell.arrayRange !== undefined && row >= cell.arrayRange.top && row <= cell.arrayRange.bottom && col >= cell.arrayRange.left && col <= cell.arrayRange.right),
    );
  const inMerge = (row: number, col: number): boolean => sheet.merges.some((merge) => row >= merge.top && row <= merge.bottom && col >= merge.left && col <= merge.right);
  const plain = sheet.rows
    .flatMap((row) => row.cells.map((cell) => ({ row: row.index, cell })))
    .find(({ row, cell }) => cell.value !== null && cell.formula === undefined && cell.dataTableXml === undefined && !inArray(row, cell.col) && !inMerge(row, cell.col));
  if (plain) {
    setCellEntry(model, index, { row: plain.row, col: plain.cell.col }, { value: "changed by SPEC-122" });
  }
  setCellEntry(model, index, { row: below, col: 0 }, { value: 1234.5, numberFormat: "#,##0.00" });
  setCellEntry(model, index, { row: below, col: 1 }, { value: "edited\ntwo lines", wrap: true });
  setCellEntry(model, index, { row: below, col: 2 }, { value: 2469, formula: `A${below + 1}*2` });
  model.sheets = model.sheets.map((each, position) => (position === index ? { ...each, dirty: true } : each));
  return index;
}

/** A part's text. */
function partText(parts: Map<string, Uint8Array>, name: string): string {
  return new TextDecoder().decode(parts.get(name));
}

/** A part's tags and texts in order, its count attributes left out. */
function xmlTokens(xml: string): string[] {
  return xml.replace(/ (?:count|uniqueCount)="\d+"/g, "").match(/<[^>]*>|[^<]+/g) ?? [];
}

/** Whether `written` is `source` with whole elements inserted and nothing
 * else changed but `count`/`uniqueCount` values (or a `count` removed) —
 * SPEC-122's append-only rule for xl/styles.xml and xl/sharedStrings.xml:
 * compared token by token (tags and text), every source token appears in
 * order, and a token the source does not have is part of an inserted run. */
function appendOnly(source: string, written: string): boolean {
  const from = xmlTokens(source);
  const to = xmlTokens(written);
  let j = 0;
  for (const token of from) {
    const found = to.indexOf(token, j);
    if (found < 0) {
      return false;
    }
    // What lies between is inserted: whole elements, never text.
    const run = to.slice(j, found).join("");
    if (run !== "" && !/^<[^/!?][\s\S]*>$/.test(run)) {
      return false;
    }
    j = found + 1;
  }
  return j === to.length;
}

describe("SPEC-122: edit → write → read gives the edited model, and write → read → write is stable", () => {
  it("(the append-only comparison accepts inserted elements and new counts, nothing else)", () => {
    const source = `<sst count="2" uniqueCount="2"><si><t>a</t></si><si><t>b</t></si></sst>`;
    expect(appendOnly(source, source)).toBe(true);
    expect(appendOnly(source, `<sst uniqueCount="3"><si><t>a</t></si><si><t>b</t></si><si><t>c</t></si></sst>`)).toBe(true);
    expect(appendOnly(`<fonts count="1"><font/></fonts><fills/>`, `<numFmts count="1"><numFmt/></numFmts><fonts count="2"><font/><font><b/></font></fonts><fills/>`)).toBe(true);
    expect(appendOnly(source, `<sst uniqueCount="2"><si><t>a</t></si><si><t>B</t></si></sst>`)).toBe(false);
    expect(appendOnly(source, `<sst uniqueCount="1"><si><t>a</t></si></sst>`)).toBe(false);
    expect(appendOnly(source, `<sst uniqueCount="2"><si><t>a!</t></si><si><t>b</t></si></sst>`)).toBe(false);
  });

  for (const fixture of EDIT_FIXTURES) {
    it(`${fixture.file}`, async () => {
      const sourceBytes = await fixture.bytes();
      const model = await readSpreadsheetBytes(sourceBytes, fixture.format, "fixture");
      const index = editFirstGridSheet(model);
      const written = (await spreadsheetDownload(model)).bytes;
      expect(await packageIntegrity(written)).toEqual([]);
      const back = await readSpreadsheetBytes(written, fixture.format, "fixture");
      expect(workbookForCompare(back)).toEqual(workbookForCompare(model));
      model.sheets.forEach((sheet, position) => {
        expect(sheetForCompare(back, back.sheets[position]!), sheet.name).toEqual(sheetForCompare(model, sheet));
      });

      // Only the edited sheet was rebuilt: every other sheet's part (xlsx) or
      // table (ods) keeps its source bytes, and the workbook-wide string and
      // style tables only grew.
      const source = await unzip(sourceBytes);
      const first = await unzip(written);
      if (fixture.format === "xlsx") {
        model.sheets.forEach((sheet, position) => {
          if (position !== index && sheet.partName !== undefined && source.has(sheet.partName)) {
            expect(equalBytes(first.get(sheet.partName)!, source.get(sheet.partName)!), sheet.partName).toBe(true);
          }
        });
        for (const name of ["xl/styles.xml", "xl/sharedStrings.xml"].filter((each) => source.has(each))) {
          expect(appendOnly(partText(source, name), partText(first, name)), name).toBe(true);
        }
      } else {
        const content = partText(first, "content.xml");
        model.sheets.forEach((sheet, position) => {
          if (position !== index && sheet.sourceXml !== undefined) {
            expect(content.includes(sheet.sourceXml), sheet.name).toBe(true);
          }
        });
      }

      // The same sheet rebuilt again from what was read: every part as before.
      back.sheets = back.sheets.map((sheet, position) => (position === index ? { ...sheet, dirty: true } : sheet));
      const second = await unzip((await spreadsheetDownload(back)).bytes);
      expect([...second.keys()]).toEqual([...first.keys()]);
      for (const [name, bytes] of first) {
        expect(equalBytes(second.get(name)!, bytes), name).toBe(true);
      }
    });
  }
});

describe("SPEC-122: forcing every .ods sheet dirty without an edit reads back to the same model", () => {
  for (const fixture of EDIT_FIXTURES.filter((each) => each.format === "ods")) {
    it(`${fixture.file}`, async () => {
      const model = await readSpreadsheetBytes(await fixture.bytes(), "ods", "fixture");
      const forced: SpreadsheetModel = { ...model, sheets: model.sheets.map((sheet) => ({ ...sheet, dirty: true as const })) };
      const back = await readSpreadsheetBytes((await spreadsheetDownload(forced)).bytes, "ods", "fixture");
      expect(workbookForCompare(back)).toEqual(workbookForCompare(model));
      model.sheets.forEach((sheet, position) => {
        expect(sheetForCompare(back, back.sheets[position]!), sheet.name).toEqual(sheetForCompare(model, sheet));
      });
    });
  }
});

/** What a user sees of a sheet: per stored cell with content, its displayed
 * text in en-US, its formula and its number format — compared by meaning
 * (SPEC-118's parseFormatCode), since a copy in the other format carries a
 * format's meaning, not its spelling: an .ods number style writes Excel's
 * `[Red]\-#,##0.00` as text "-" and reads back `[Red]-#,##0.00`. */
function displayedCells(model: SpreadsheetModel, sheet: Sheet) {
  const ctx = workbookFormatContext(model, EN_US);
  return sheet.rows.flatMap((row) =>
    row.cells
      .filter((cell) => cell.value !== null || cell.formula !== undefined)
      .map((cell) => {
        const code = formatCodeOf(cellStyle(model, effectiveStyleId(sheet, row.index, cell.col)), EN_US);
        return {
          ref: formatA1({ row: row.index, col: cell.col }),
          text: plainText(formatValue(cell.value, code, ctx)),
          formula: cell.formula ?? null,
          format: parseFormatCode(code ?? "General"),
        };
      }),
  );
}

describe("SPEC-122: a copy in the other format re-opens with the same displayed values, formulas and number formats", () => {
  for (const fixture of EDIT_FIXTURES.filter((each) => /^(mini-book|spec122-shapes)\./.test(each.file))) {
    it(`${fixture.file}`, async () => {
      const model = await readSpreadsheetBytes(await fixture.bytes(), fixture.format, "book");
      const target: SpreadsheetFormat = fixture.format === "xlsx" ? "ods" : "xlsx";
      const { bytes } = await spreadsheetDownload(model, target);
      expect(await packageIntegrity(bytes)).toEqual([]);
      const back = await readSpreadsheetBytes(bytes, target, "copy");
      const expected = stripRetention({ ...model, format: target });
      expect(back.sheets.map((sheet) => sheet.name)).toEqual(expected.sheets.map((sheet) => sheet.name));
      expected.sheets.forEach((sheet, position) => {
        expect(displayedCells(back, back.sheets[position]!), sheet.name).toEqual(displayedCells(expected, sheet));
      });
    });
  }
});

/** A modelled element without its verbatim source text. */
function strip<T extends { sourceXml?: string }>(value: T | undefined): Omit<T, "sourceXml"> | undefined {
  if (value === undefined) {
    return undefined;
  }
  const { sourceXml: _source, ...rest } = value;
  return rest;
}

/** What SPEC-128 round-trips of a sheet: its autofilter and sort record
 * (their verbatim source text is retention, not model) and every row's two
 * hidden flags. */
function filterForCompare(sheet: Sheet) {
  const autoFilter = sheet.autoFilter && { ...strip(sheet.autoFilter)!, sort: strip(sheet.autoFilter.sort) };
  if (autoFilter && autoFilter.sort === undefined) {
    delete autoFilter.sort;
  }
  return {
    autoFilter,
    sortState: strip(sheet.sortState),
    rows: sheet.rows.filter((row) => row.hidden || row.filterHidden).map((row) => ({ index: row.index, hidden: row.hidden === true, filterHidden: row.filterHidden === true })),
  };
}

describe("SPEC-128: filter → save → reopen reproduces the autofilter, the sort records and every row's hidden flags", () => {
  const ctx = workbookFormatContext({ dateSystem: "1900", format: "xlsx" }, EN_US);
  for (const [file, format, bytes] of [
    ["spec128-filter.xlsx", "xlsx", buildSpec128FilterXlsxBytes],
    ["spec128-filter.ods", "ods", buildSpec128FilterOdsBytes],
  ] as const) {
    it(`${file}: Region without West`, async () => {
      const model = await readSpreadsheetBytes(await bytes(), format, "book");
      const sales = model.sheets[0]!;
      expect(sales.autoFilter?.range).toEqual(SPEC128_FILTER_XLSX.sales.autoFilter.range);
      const result = setFilterCriteria(model, sales.id, 0, { kind: "values", values: ["East", "North"], blanks: true, dateGroups: [] }, { ...ctx, dateQuirks: format === "ods" ? "proleptic" : "excel" });
      expect(result.ok).toBe(true);
      const back = await readSpreadsheetBytes((await spreadsheetDownload(model)).bytes, format, "book");
      expect(back.sheets.map(filterForCompare)).toEqual(model.sheets.map(filterForCompare));
      expect(filterForCompare(back.sheets[0]!).rows.filter((row) => row.filterHidden).map((row) => row.index + 1)).toEqual([3, 4, 5, 8, 9, 10]);
    });
  }

  it("spec128-filter.xlsx: a filter whose source text survives is re-emitted verbatim inside a rebuilt sheet", async () => {
    const model = await readSpreadsheetBytes(await buildSpec128FilterXlsxBytes(), "xlsx", "book");
    setCellEntry(model, 0, { row: 20, col: 5 }, { value: "note" });
    model.sheets = model.sheets.map((each, position) => (position === 0 ? { ...each, dirty: true } : each));
    const parts = await unzip((await spreadsheetDownload(model)).bytes);
    const sheetXml = new TextDecoder().decode(parts.get(SPEC128_FILTER_XLSX.sales.partName)!);
    expect(sheetXml).toContain(SPEC128_FILTER_XLSX.sales.autoFilter.xml);
    const back = await readSpreadsheetBytes(await spreadsheetDownload(model).then((d) => d.bytes), "xlsx", "book");
    expect(back.sheets.map(filterForCompare)).toEqual(model.sheets.map(filterForCompare));
  });

  it("spec128-filter.xlsx: turning the filter off writes no autoFilter, no filterMode and no _xlnm._FilterDatabase for the sheet", async () => {
    const model = await readSpreadsheetBytes(await buildSpec128FilterXlsxBytes(), "xlsx", "book");
    const { toggleAutoFilter } = await import("./filter");
    expect(toggleAutoFilter(model, model.sheets[0]!.id, null, ctx).ok).toBe(true);
    const parts = await unzip((await spreadsheetDownload(model)).bytes);
    const sheetXml = new TextDecoder().decode(parts.get(SPEC128_FILTER_XLSX.sales.partName)!);
    expect(sheetXml).not.toContain("<autoFilter");
    expect(sheetXml).not.toContain("filterMode");
    const workbookXml = new TextDecoder().decode(parts.get("xl/workbook.xml")!);
    const filterDatabases = [...workbookXml.matchAll(/<definedName\b[^>]*name="_xlnm\._FilterDatabase"[^>]*>/g)].map((match) => /localSheetId="(\d+)"/.exec(match[0])?.[1]);
    expect(filterDatabases).toEqual(["1"]);
    const back = await readSpreadsheetBytes((await spreadsheetDownload(model)).bytes, "xlsx", "book");
    expect(back.sheets[0]!.autoFilter).toBeUndefined();
    expect(filterForCompare(back.sheets[0]!).rows).toEqual([{ index: SPEC128_FILTER_XLSX.sales.userHiddenIndex, hidden: true, filterHidden: false }]);
  });
});

// ============================================================================
// SPEC-129 (xlsx) — conditional formatting over spec129-cf.xlsx: an untouched
// download is equalBytes per part; a dirty sheet whose rules were not edited
// re-emits each block (and the extLst) verbatim; add/edit → save → reopen
// reproduces the conditional formats and the dxfs they reference, minus the
// minted ids, the x14 GUIDs and the source text, in the same priority order.
// ============================================================================

/** A model's conditional formats as the SPEC-129 round trips compare them:
 * no minted ids, GUIDs or source text, each dxfId replaced by the dxf it
 * names, a preserved rule's XML without its wire `priority` (the writer
 * patches it to the model's); formats by their first rule's priority (xlsx
 * keeps base and x14 formats in two lists, so their interleaving in the
 * array is not carried — precedence is the priority). */
function spec129Formats(model: SpreadsheetModel): unknown[] {
  return model.sheets.map((sheet) =>
    sheet.conditionalFormats
      .toSorted((a, b) => (a.rules[0]?.priority ?? 0) - (b.rules[0]?.priority ?? 0))
      .map((format) => ({
        ranges: format.ranges,
        pivot: format.pivot,
        rules: format.rules.map(({ id: _id, x14Id: _guid, dxfId, ...rule }) => ({
          ...rule,
          ...(rule.kind === "preserved" ? { xml: rule.xml.replace(/^(<[^>]*?)\s+priority="\d+"/, "$1") } : {}),
          dxf: dxfId === undefined ? undefined : modelledDxf(model.dxfs[dxfId]!),
        })),
      })),
  );
}

/** A part's text. */
function spec129Decode(bytes: Uint8Array | undefined): string {
  return new TextDecoder().decode(bytes);
}

/** The first sheet's rules in priority order, as the manager lists them. */
function spec129Order(model: SpreadsheetModel): string[] {
  return rulesByPriority(model.sheets[0]!.conditionalFormats).map(({ format, rule }) => `${format.ranges.map(formatRange).join(" ")} ${rule.kind}`);
}

describe("SPEC-129 (xlsx): spec129-cf.xlsx round trips", () => {
  const F = SPEC129_CF_XLSX;

  it("an untouched download is equalBytes per part, in order — after the store's JSON trip too", async () => {
    const source = await buildSpec129CfXlsxBytes();
    const model = await readSpreadsheetBytes(source, "xlsx", "book");
    const original = await unzip(source);
    for (const candidate of [model, JSON.parse(JSON.stringify(model)) as SpreadsheetModel]) {
      const out = await unzip((await spreadsheetDownload(candidate)).bytes);
      expect([...out.keys()]).toEqual([...original.keys()]);
      for (const [name, bytes] of original) {
        expect(equalBytes(out.get(name)!, bytes), name).toBe(true);
      }
    }
  });

  it("a dirty sheet whose rules were not edited re-emits each block, the conditional-formatting ext and the sparkline ext verbatim", async () => {
    const source = await buildSpec129CfXlsxBytes();
    const model = await readSpreadsheetBytes(source, "xlsx", "book");
    setCellEntry(model, 0, { row: 20, col: 9 }, { value: "note" });
    model.sheets = model.sheets.map((each, position) => (position === 0 ? { ...each, dirty: true } : each));
    const parts = await unzip((await spreadsheetDownload(model)).bytes);
    const original = await unzip(source);
    const scores = spec129Decode(parts.get(F.scores.partName));
    expect(scores).not.toBe(spec129Decode(original.get(F.scores.partName)));
    expect(scores).toContain(F.scores.blocks.join(""));
    expect(scores).toContain(`<extLst>${F.scores.cfExt}${F.scores.sparklineExt}</extLst>`);
    for (const name of [F.more.partName, "xl/styles.xml"]) {
      expect(equalBytes(parts.get(name)!, original.get(name)!), name).toBe(true);
    }
    const back = await readSpreadsheetBytes(await spreadsheetDownload(model).then((download) => download.bytes), "xlsx", "book");
    expect(spec129Formats(back)).toEqual(spec129Formats(model));
  });

  it("add/edit → save → reopen reproduces the conditional formats and the dxfs they reference, in the same priority order", async () => {
    const model = await readSpreadsheetBytes(await buildSpec129CfXlsxBytes(), "xlsx", "book");
    const scores = model.sheets[0]!;
    // The e2e's flow: "> 120" in green over B2:B13 (first), moved down one.
    const green = internDxf(model, { font: { color: "#006100" }, fill: { kind: "solid", color: "#c6efce" } });
    expect(green).toBe(F.dxfs.length);
    const { ruleId } = addConditionalFormat(model, scores.id, [parseRange("B2:B13")!], { kind: "cellIs", operator: "greaterThan", formulas: ["120"], dxfId: green });
    moveRule(model, scores.id, ruleId, "down");
    // The data bar recoloured (its GUID kept), More's duplicates rule deleted.
    const bar = model.sheets[0]!.conditionalFormats.flatMap((format) => format.rules).find((rule) => rule.kind === "dataBar")!;
    if (bar.kind !== "dataBar") {
      throw new Error("no data bar");
    }
    updateRule(model, scores.id, bar.id, { kind: "dataBar", bar: { ...bar.bar, color: "#ff8800", gradient: true } });
    const more = model.sheets[1]!;
    deleteRule(model, more.id, more.conditionalFormats.flatMap((format) => format.rules).find((rule) => rule.kind === "duplicates")!.id);

    const parts = await unzip((await spreadsheetDownload(model)).bytes);
    const styles = spec129Decode(parts.get("xl/styles.xml"));
    expect(styles).toContain(`<dxfs count="${F.dxfs.length + 1}">${F.dxfXml.join("")}<dxf><font><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf></dxfs>`);
    expect(spec129Decode(parts.get(F.scores.partName))).toContain(F.scores.sparklineExt);
    const back = await readSpreadsheetBytes(await spreadsheetDownload(model).then((download) => download.bytes), "xlsx", "book");
    expect(spec129Formats(back)).toEqual(spec129Formats(model));
    // The manager lists the same rules in the same order.
    expect(spec129Order(back)).toEqual(spec129Order(model));
    expect(spec129Order(back).slice(0, 2)).toEqual(["B2:B13 cellIs", "B2:B13 cellIs"]);
  });
});
// ============================================================================
// end of SPEC-129 (xlsx)
// ============================================================================

// ============================================================================
// SPEC-129 (ODS): conditional formats round trip — spec129-cf.ods and
// spec129-legacy.ods download byte-identical when untouched; add/edit → save
// → reopen reproduces each sheet's conditional formats and the dxfs their
// rules apply under toEqual, minus the minted ids, `sourceXml` and a dxf's
// `sourceStyleName` (a new dxf comes back as its generated `de-cf` style).
// ODS numbers priorities by document order, so formats compare in priority
// order, and dxfs by content (the re-read table is in first-use order).
// ============================================================================

/** An edit of a sheet's cells, the sheet marked dirty as the editor does. */
function editSpec129Sheet(model: SpreadsheetModel, index: number, row: number, col: number, value: string | number): void {
  setCellEntry(model, index, { row, col }, { value });
  model.sheets = model.sheets.map((each, position) => (position === index ? { ...each, dirty: true } : each));
}

function spec129FirstPriority(format: ConditionalFormat): number {
  return Math.min(...format.rules.map((rule) => rule.priority));
}

/** A sheet's formats as an ODS save round-trips them (see above). */
function odsConditionalFormatsForCompare(model: SpreadsheetModel, sheet: Sheet) {
  return sheet.conditionalFormats
    .toSorted((a, b) => spec129FirstPriority(a) - spec129FirstPriority(b))
    .map((format) => ({
      ranges: format.ranges,
      rules: format.rules.map(({ id: _id, dxfId, ...rule }) => {
        if (dxfId === undefined) {
          return rule;
        }
        const { sourceXml: _xml, sourceStyleName: _name, ...dxf } = model.dxfs[dxfId]!;
        return { ...rule, dxf };
      }),
    }));
}

describe("SPEC-129 (ODS): conditional formats round trip", () => {
  for (const [file, build] of [
    ["spec129-cf.ods", buildSpec129CfOdsBytes],
    ["spec129-legacy.ods", buildSpec129LegacyOdsBytes],
  ] as const) {
    it(`${file}: an untouched download is every part byte-identical and in order`, async () => {
      const source = await build();
      const model = await readSpreadsheetBytes(source, "ods", "fixture");
      const out = await unzip((await spreadsheetDownload(model)).bytes);
      const original = await unzip(source);
      expect([...out.keys()]).toEqual([...original.keys()]);
      for (const [name, bytes] of original) {
        expect(equalBytes(out.get(name)!, bytes), name).toBe(true);
      }
    });
  }

  it("spec129-cf.ods: a dirty table whose rules were not edited re-emits each format's source verbatim", async () => {
    const model = await readSpreadsheetBytes(await buildSpec129CfOdsBytes(), "ods", "book");
    editSpec129Sheet(model, 0, 20, 1, "note");
    const content = new TextDecoder().decode((await unzip((await spreadsheetDownload(model)).bytes)).get("content.xml"));
    expect(content).toContain(`<calcext:conditional-formats>${SPEC129_CF_ODS.scoresFormats.join("")}</calcext:conditional-formats>`);
  });

  it("spec129-cf.ods: add (Greater than 120, green), move it down, edit, re-range and delete rules → save → reopen reproduces them", async () => {
    const model = await readSpreadsheetBytes(await buildSpec129CfOdsBytes(), "ods", "book");
    const [scores, more] = model.sheets;
    const green = internDxf(model, { font: { color: "#006100" }, fill: { kind: "solid", color: "#c6efce" } });
    const { ruleId } = addConditionalFormat(model, scores!.id, [parseRange("B2:B13")!], { kind: "cellIs", operator: "greaterThan", formulas: ["120"], dxfId: green });
    moveRule(model, scores!.id, ruleId, "down");
    const sheetFormats = (index: number) => model.sheets[index]!.conditionalFormats;
    const stars = sheetFormats(0).find((format) => format.rules[0]!.kind === "iconSet" && format.rules[0]!.icons.set === "3Stars")!;
    deleteRule(model, scores!.id, stars.rules[0]!.id);
    const bar = sheetFormats(0).find((format) => format.rules[0]!.kind === "dataBar")!;
    setFormatRanges(model, scores!.id, bar.id, [parseRange("C2:C20")!, parseRange("J2:J5")!]);
    const text = sheetFormats(1).flatMap((format) => format.rules).find((rule) => rule.kind === "text")!;
    const bold = internDxf(model, { font: { bold: true, underline: "single" }, numberFormat: "0.00" });
    updateRule(model, more!.id, text.id, { kind: "text", operator: "beginsWith", text: `say "hi"`, dxfId: bold });

    const back = await readSpreadsheetBytes((await spreadsheetDownload(model)).bytes, "ods", "book");
    for (const index of [0, 1]) {
      expect(odsConditionalFormatsForCompare(back, back.sheets[index]!)).toEqual(odsConditionalFormatsForCompare(model, model.sheets[index]!));
    }
    // The new rule ranks second, under Bad's "> 100".
    const ranked = rulesByPriority(back.sheets[0]!.conditionalFormats).slice(0, 2);
    expect(ranked.map(({ rule }) => rule.kind === "cellIs" && rule.formulas[0])).toEqual(["100", "120"]);
    // And a second save of the re-read model is stable.
    const again = await readSpreadsheetBytes((await spreadsheetDownload(back)).bytes, "ods", "book");
    expect(odsConditionalFormatsForCompare(again, again.sheets[0]!)).toEqual(odsConditionalFormatsForCompare(back, back.sheets[0]!));
  });

  it("spec129-legacy.ods: an edited table writes its style:map rules as calcext ones and reopens to the same formats", async () => {
    const model = await readSpreadsheetBytes(await buildSpec129LegacyOdsBytes(), "ods", "book");
    editSpec129Sheet(model, 0, 2, 2, 120);
    const back = await readSpreadsheetBytes((await spreadsheetDownload(model)).bytes, "ods", "book");
    expect(odsConditionalFormatsForCompare(back, back.sheets[0]!)).toEqual(odsConditionalFormatsForCompare(model, model.sheets[0]!));
  });
});

// ============================================================================
// SPEC-129: "Download a copy as…" carries every modelled rule across formats
// per the wire tables — in priority order, over the same ranges; a preserved
// rule stays behind (UNCONVERTIBLE_CONTENT's "some conditional formatting").
// A kind with no wire form of its own in the target travels as the
// expression the tables give it: blanks and a std-dev average as ODS
// `formula-is`, LibreOffice's year periods as xlsx expressions.
// ============================================================================

/** Each sheet's rules in priority order as "<ranges> <kind>[ <icon set>]". */
function carriedRules(model: SpreadsheetModel, target: SpreadsheetFormat): string[][] {
  return model.sheets.map((sheet) =>
    rulesByPriority(sheet.conditionalFormats)
      .filter(({ rule }) => rule.kind !== "preserved")
      .map(({ format, rule }) => {
        const asExpression =
          target === "ods"
            ? rule.kind === "blanks" || (rule.kind === "aboveAverage" && rule.stdDev !== undefined)
            : rule.kind === "timePeriod" && rule.period.endsWith("Year");
        const kind = asExpression ? "expression" : rule.kind;
        return `${format.ranges.map(formatRange).join(" ")} ${kind}${rule.kind === "iconSet" ? ` ${rule.icons.set}` : ""}`;
      }),
  );
}

describe("SPEC-129: a copy in the other format carries every modelled rule", () => {
  for (const [file, format, build] of [
    ["spec129-cf.xlsx", "xlsx", buildSpec129CfXlsxBytes],
    ["spec129-cf.ods", "ods", buildSpec129CfOdsBytes],
  ] as const) {
    it(file, async () => {
      const model = await readSpreadsheetBytes(await build(), format, "book");
      const target: SpreadsheetFormat = format === "xlsx" ? "ods" : "xlsx";
      const { bytes } = await spreadsheetDownload(model, target);
      expect(await packageIntegrity(bytes)).toEqual([]);
      const back = await readSpreadsheetBytes(bytes, target, "copy");
      expect(carriedRules(back, target)).toEqual(carriedRules(model, target));
      expect(carriedRules(back, target).flat().length).toBeGreaterThan(10);
    });
  }
});

// --- SPEC-130 ------------------------------------------------------------------------------------

/** A sheet's rules as a round trip compares them: no minted ids, no source
 * text, no minted ODS name (wire identities, re-made on every write). */
function rulesForCompare(sheet: Sheet) {
  return sheet.validations.map(({ id: _id, sourceXml: _xml, odsName: _name, ...rule }) => rule);
}

/** What crosses formats: the rules without their placement in the source
 * format (x14) and, into an xlsx, without LibreOffice's sort order — as a
 * set: an xlsx file lists its x14 ext's rules after the base block's, so a
 * cross-sheet rule's place among the others cannot cross. */
function crossedRules(sheet: Sheet, target: SpreadsheetFormat) {
  return rulesForCompare(sheet)
    .map(({ x14: _x14, odsListOrder, ...rule }) => (target === "ods" && odsListOrder !== undefined ? { ...rule, odsListOrder } : rule))
    .map((rule) => JSON.stringify(rule))
    .toSorted();
}

describe("SPEC-130: data validations round trip", () => {
  for (const [file, format, build] of [
    ["spec130-validation.xlsx", "xlsx", buildSpec130ValidationXlsxBytes],
    ["spec130-validation.ods", "ods", buildSpec130ValidationOdsBytes],
  ] as const) {
    it(`${file}: an untouched download is equalBytes per part, in order — after the store's JSON trip too`, async () => {
      const source = await build();
      const model = await readSpreadsheetBytes(source, format, "book");
      expect(model.sheets[0]!.validations.length).toBeGreaterThan(5);
      const original = await unzip(source);
      for (const candidate of [model, JSON.parse(JSON.stringify(model)) as SpreadsheetModel]) {
        const out = await unzip((await spreadsheetDownload(candidate)).bytes);
        expect([...out.keys()]).toEqual([...original.keys()]);
        for (const [name, bytes] of original) {
          expect(equalBytes(out.get(name)!, bytes), name).toBe(true);
        }
      }
    });

    it(`${file}: set → save → reopen reproduces the validations (minus minted ids and source text)`, async () => {
      const model = await readSpreadsheetBytes(await build(), format, "book");
      const form = model.sheets[0]!;
      // The e2e's D2 list, a custom rule over a block, and B4's rule cleared.
      setValidation(model, form.id, [parseRange("D2")!], { ...defaultValidationSettings(), type: "list", formula1: formatListLiteral(["Yes", "No"]) });
      setValidation(model, form.id, [parseRange("G2:H5")!], {
        ...defaultValidationSettings(),
        type: "custom",
        formula1: "G2>$B$3",
        input: { show: true, title: "Above B3", text: "Greater than\nB3" },
        error: { show: true, style: "warning", title: "Too small", text: "Pick more" },
      });
      clearValidation(model, form.id, [parseRange("B4")!]);
      const back = await readSpreadsheetBytes((await spreadsheetDownload(model)).bytes, format, "book");
      expect(rulesForCompare(back.sheets[0]!)).toEqual(rulesForCompare(model.sheets[0]!));
      expect(back.sheets[0]!.validations.some((rule) => rule.formula1 === '"Yes,No"')).toBe(true);
      expect(back.odsOrphanValidations).toEqual(model.odsOrphanValidations);
    });

    it(`${file}: a copy in the other format maps every type through the model`, async () => {
      const model = await readSpreadsheetBytes(await build(), format, "book");
      const target: SpreadsheetFormat = format === "xlsx" ? "ods" : "xlsx";
      const { bytes } = await spreadsheetDownload(model, target);
      expect(await packageIntegrity(bytes)).toEqual([]);
      const back = await readSpreadsheetBytes(bytes, target, "copy");
      for (const [index, sheet] of model.sheets.entries()) {
        expect(crossedRules(back.sheets[index]!, target), sheet.name).toEqual(crossedRules(sheet, target));
      }
      const types = new Set(back.sheets[0]!.validations.map((rule) => rule.type));
      for (const type of ["list", "whole", "decimal", "date", "textLength", "custom"] as const) {
        expect(types.has(type), type).toBe(true);
      }
      if (target === "ods") {
        // The x14 cross-sheet list took the ODS form.
        expect(model.sheets[0]!.validations.some((rule) => rule.x14)).toBe(true);
        expect(back.sheets[0]!.validations.some((rule) => rule.formula1 === "Lists!$A$1:$A$4")).toBe(true);
      } else {
        // The sort-ascending list became an ordinary list; the cross-sheet one went to the x14 ext.
        expect(model.sheets[0]!.validations.some((rule) => rule.odsListOrder === "sort-ascending")).toBe(true);
        expect(back.sheets[0]!.validations.some((rule) => rule.odsListOrder !== undefined)).toBe(false);
        expect(back.sheets[0]!.validations.find((rule) => rule.formula1?.startsWith("Lists!"))?.x14).toBe(true);
      }
    });
  }
});
