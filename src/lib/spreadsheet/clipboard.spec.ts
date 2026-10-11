// SPEC-125 "Flavours", "The payload", "Validation", "External sources",
// "Capture fixtures": the payload's build/parse round trip, every validation
// bound, the security cases (oversize, malformed, prototype-polluting,
// out-of-shape and script-bearing payloads; SPEC-016's sanitisation list for
// the external importer), TSV both ways, HTML out, each producer's markers
// as inline snippets, and each real capture under
// e2e/fixtures/spec125-clipboard/ to the expected cells of "Capture
// fixtures" — skipped with "fixture not yet captured: <name>" while the file
// is absent.
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import type { GridRange } from "./address";
import {
  buildClipboardCopy,
  buildClipboardPayload,
  cellCss,
  clipboardFlavours,
  copyArea,
  copyFlavours,
  copyRefusal,
  encodePayload,
  MAX_COPY_CELLS,
  MULTIPLE_SELECTIONS_MESSAGE,
  parseClipboardPayload,
  parseTsv,
  PAYLOAD_ATTRIBUTE,
  TOO_MANY_TO_CHANGE_MESSAGE,
  TOO_MANY_TO_COPY_MESSAGE,
  tsvField,
  tsvText,
} from "./clipboard";
import { cssColor, cssDeclarations, decodeCssString, importClipboardHtml, importedStyle } from "./clipboardImport";
import { MAX_PAYLOAD_ATTRIBUTE_LENGTH } from "./clipboardPayload";
import { bytesToBase64 } from "../core/packageParts";
import { DEFAULT_THEME } from "../core/theme";
import { r1c1ToA1 } from "./formula/r1c1";
import { MAX_COLS, MAX_ROWS, type Cell, type CellStyle, type Row, type Sheet, type SpreadsheetModel } from "./model";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { blockFromPayload, pasteBlock, type PasteBlock } from "./paste";
import { dateKindOf } from "./series";
import { dateToSerial } from "./serial";

const EN_US = SPREADSHEET_LOCALES["en-US"];
const CAPTURES = join(__dirname, "../../../e2e/fixtures/spec125-clipboard");

function sheet(id: string, name: string, fields: Partial<Sheet> = {}): Sheet {
  return { conditionalFormats: [], id, name, visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], ...fields };
}

function model(sheets: Sheet[], fields: Partial<SpreadsheetModel> = {}): SpreadsheetModel {
  return {
    dxfs: [],
    kind: "spreadsheet",
    id: "book",
    title: "T",
    format: "xlsx",
    readerVersion: 8,
    sheets,
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
    ...fields,
  };
}

const value = (col: number, v: Cell["value"], extra: Partial<Cell> = {}): Cell => ({ col, value: v, ...extra });
const row = (index: number, cells: Cell[], extra: Partial<Row> = {}): Row => ({ index, cells, ...extra });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });

/** A workbook with B2:C3 = 10, 20 / =B2*2 (20), "a\tb" — C3 bold, B2:C2's
 * fill yellow, D1:E1 merged. */
function sample(): SpreadsheetModel {
  const bold: CellStyle = { font: { bold: true } };
  const yellow: CellStyle = { fill: { kind: "solid", color: "#ffff00" }, sourceXf: 3 };
  return model(
    [
      sheet("s1", "Sheet1", {
        rows: [
          row(0, [value(3, "Merged")]),
          row(1, [value(1, 10, { styleId: 2 }), value(2, 20, { styleId: 2 })]),
          row(2, [value(1, 20, { formula: "B2*2" }), value(2, "a\tb", { styleId: 1 })]),
        ],
        merges: [range(0, 3, 0, 4)],
      }),
    ],
    { styles: [{}, bold, yellow] },
  );
}

const encode = (data: unknown): string => bytesToBase64(new TextEncoder().encode(JSON.stringify(data)));

function minimal(fields: Record<string, unknown> = {}): Record<string, unknown> {
  return { kind: "docu-edit/cells", version: 1, workbookId: "w", sheetName: "S", dateSystem: "1900", origin: { row: 0, col: 0 }, rows: 2, cols: 2, cells: [], styles: [], merges: [], ...fields };
}

describe("the copied rectangle", () => {
  it("refuses a multi-range selection and an area over MAX_COPY_CELLS", () => {
    const s = sample().sheets[0]!;
    expect(copyRefusal(s, [range(0, 0, 0, 0), range(2, 2, 2, 2)])).toBe(MULTIPLE_SELECTIONS_MESSAGE);
    const big = sheet("s", "S", { rows: [row(999, [value(999, 1)])] });
    expect(copyRefusal(big, [range(0, 0, 999, 999)])).toBeNull(); // exactly 1,000,000
    expect(copyRefusal(big, [range(0, 0, 5000, 999)])).toBeNull(); // clipped back to the used range
    const bigger = sheet("s", "S", { rows: [row(2000, [value(999, 1)])] });
    expect(copyRefusal(bigger, [range(0, 0, 1000, 999)])).toBe(TOO_MANY_TO_COPY_MESSAGE);
    expect(MAX_COPY_CELLS).toBe(1_000_000);
  });

  it("clips to the used range: copying column A copies A1 to its last used row", () => {
    const s = sheet("s", "S", { rows: [row(4, [value(0, 1)]), row(9, [value(2, 2)])] });
    expect(copyArea(s, range(0, 0, MAX_ROWS - 1, 0))).toEqual({ range: range(0, 0, 9, 0), wholeRows: false, wholeCols: true });
    expect(copyArea(s, range(3, 0, 3, MAX_COLS - 1))).toEqual({ range: range(3, 0, 3, 2), wholeRows: true, wholeCols: false });
    // Never above the top-left, and grown back over a merge the clip would cut.
    expect(copyArea(s, range(20, 5, 30, 8)).range).toEqual(range(20, 5, 20, 5));
    const merged = sheet("s", "S", { rows: [row(0, [value(0, 1)])], merges: [range(0, 0, 0, 3)] });
    expect(copyArea(merged, range(0, 0, 0, MAX_COLS - 1)).range).toEqual(range(0, 0, 0, 3));
  });
});

describe("the payload", () => {
  it("carries values, canonical formulas, styles without sourceXf, relative merges and the source facts", () => {
    const m = sample();
    const { payload, styleIds } = buildClipboardCopy(m, "s1", range(0, 1, 2, 4), "cut-1");
    expect(payload).toMatchObject({ kind: "docu-edit/cells", version: 1, workbookId: "book", sheetName: "Sheet1", dateSystem: "1900", origin: { row: 0, col: 1 }, rows: 3, cols: 4, cut: "cut-1" });
    expect(payload.cells).toEqual([
      { r: 0, c: 2, v: "Merged" },
      { r: 1, c: 0, v: 10, s: 0 },
      { r: 1, c: 1, v: 20, s: 0 },
      { r: 2, c: 0, v: 20, f: "B2*2" },
      { r: 2, c: 1, v: "a\tb", s: 1 },
    ]);
    expect(payload.styles).toEqual([{ fill: { kind: "solid", color: "#ffff00" } }, { font: { bold: true } }]);
    expect(styleIds).toEqual([2, 1]);
    expect(payload.merges).toEqual([range(0, 2, 0, 3)]);
  });

  it("round-trips through the attribute: parse(encode(p)) equals p", () => {
    const payload = buildClipboardPayload(sample(), "s1", range(0, 0, 2, 4));
    expect(parseClipboardPayload(encodePayload(payload))).toEqual(payload);
  });

  it("carries whole-row properties and rich text, and an array only when its whole range was copied", () => {
    const s = sheet("s1", "S", {
      rows: [
        row(0, [value(0, 6, { formula: "SUM(B1:B2)", arrayRange: range(0, 0, 1, 0) }), value(1, "ab", { runs: [{ text: "a", font: { bold: true } }, { text: "b" }] })], { heightPt: 30, customHeight: true }),
        row(1, [value(0, 6)]),
      ],
    });
    const m = model([s], { styles: [{}, { font: { italic: true } }] });
    const whole = buildClipboardPayload(m, "s1", range(0, 0, 1, MAX_COLS - 1));
    expect(whole.wholeRows).toBe(true);
    expect(whole.rowProps).toEqual([{ r: 0, heightPt: 30, customHeight: true }]);
    expect(whole.cells[0]).toEqual({ r: 0, c: 0, v: 6, f: "SUM(B1:B2)", a: range(0, 0, 1, 0) });
    expect(whole.cells[1]!.t).toEqual([{ text: "a", font: { bold: true } }, { text: "b" }]);
    // The anchor alone: the formula, no array range.
    expect(buildClipboardPayload(m, "s1", range(0, 0, 0, 0)).cells[0]).toEqual({ r: 0, c: 0, v: 6, f: "SUM(B1:B2)" });
    expect(parseClipboardPayload(encodePayload(whole))).toEqual(whole);
  });
});

describe("parseClipboardPayload — every validation bound", () => {
  const ok = (data: unknown): boolean => parseClipboardPayload(encode(data)) !== null;

  it("accepts the minimal payload and ignores unknown keys", () => {
    expect(ok(minimal())).toBe(true);
    const parsed = parseClipboardPayload(encode(minimal({ extra: { anything: 1 } })))!;
    expect(Object.keys(parsed)).not.toContain("extra");
  });

  it("checks the attribute's size before decoding, and decodes strict UTF-8 JSON", () => {
    expect(parseClipboardPayload("A".repeat(MAX_PAYLOAD_ATTRIBUTE_LENGTH + 1))).toBeNull();
    expect(parseClipboardPayload("")).toBeNull();
    expect(parseClipboardPayload("!!!not base64!!!")).toBeNull();
    expect(parseClipboardPayload(btoa("\xff\xfe{}"))).toBeNull();
    expect(parseClipboardPayload(btoa("{not json"))).toBeNull();
    expect(parseClipboardPayload(btoa("[1,2]"))).toBeNull();
    expect(parseClipboardPayload(btoa("null"))).toBeNull();
  });

  it("checks kind and version", () => {
    expect(ok(minimal({ kind: "docu-edit/other" }))).toBe(false);
    expect(ok(minimal({ version: 2 }))).toBe(false);
    expect(ok(minimal({ version: "1" }))).toBe(false);
  });

  it("bounds rows, cols and origin", () => {
    expect(ok(minimal({ rows: 0 }))).toBe(false);
    expect(ok(minimal({ cols: 1.5 }))).toBe(false);
    expect(ok(minimal({ rows: 1001, cols: 1000 }))).toBe(false);
    expect(ok(minimal({ rows: 1000, cols: 1000 }))).toBe(true);
    expect(ok(minimal({ origin: { row: -1, col: 0 } }))).toBe(false);
    expect(ok(minimal({ origin: { row: MAX_ROWS - 1, col: 0 } }))).toBe(false); // 2 rows from the last row
    expect(ok(minimal({ origin: { row: 0, col: MAX_COLS - 2 } }))).toBe(true);
  });

  it("bounds cells: inside the shape, strictly increasing row-major, at most rows × cols", () => {
    expect(ok(minimal({ cells: [{ r: 2, c: 0, v: 1 }] }))).toBe(false);
    expect(ok(minimal({ cells: [{ r: 0, c: -1, v: 1 }] }))).toBe(false);
    expect(ok(minimal({ cells: [{ r: 0, c: 1, v: 1 }, { r: 0, c: 0, v: 1 }] }))).toBe(false);
    expect(ok(minimal({ cells: [{ r: 0, c: 0, v: 1 }, { r: 0, c: 0, v: 2 }] }))).toBe(false);
    expect(ok(minimal({ cells: [{ r: 0, c: 0 }] }))).toBe(false);
    expect(ok(minimal({ cells: Array.from({ length: 5 }, (_, i) => ({ r: 0, c: i, v: 1 })) }))).toBe(false);
  });

  it("bounds values: finite numbers, strings ≤ 32,767, booleans, null, known errors", () => {
    for (const v of [1.5, "x", true, null, { error: "#N/A" }, "x".repeat(32_767)]) {
      expect(ok(minimal({ cells: [{ r: 0, c: 0, v }] })), JSON.stringify(v).slice(0, 20)).toBe(true);
    }
    for (const v of ["x".repeat(32_768), { error: "#BOGUS!" }, { error: 1 }, [1], {}]) {
      expect(ok(minimal({ cells: [{ r: 0, c: 0, v }] })), JSON.stringify(v).slice(0, 20)).toBe(false);
    }
    // Infinity cannot be written as JSON at all; a huge exponent reads as it.
    expect(parseClipboardPayload(btoa(JSON.stringify(minimal({ cells: [{ r: 0, c: 0, v: 0 }] })).replace('"v":0', '"v":1e999')))).toBeNull();
  });

  it("drops a formula parseFormula rejects (or over 8,192 characters), keeping the value", () => {
    const parsed = (f: string) => parseClipboardPayload(encode(minimal({ cells: [{ r: 0, c: 0, v: 3, f }] })))!.cells[0];
    expect(parsed("A1+1")).toEqual({ r: 0, c: 0, v: 3, f: "A1+1" });
    expect(parsed("SUM(")).toEqual({ r: 0, c: 0, v: 3 });
    const long = `${"A1+".repeat(2730)}A1`;
    expect(long).toHaveLength(8192);
    expect(parsed(long)?.f).toBe(long);
    expect(parsed(`${long}+1`)).toEqual({ r: 0, c: 0, v: 3 });
  });

  it("drops an array range outside the shape or not anchored at its cell", () => {
    const parsed = (a: unknown) => parseClipboardPayload(encode(minimal({ cells: [{ r: 0, c: 0, v: 1, f: "A1", a }] })))!.cells[0]!.a;
    expect(parsed(range(0, 0, 1, 1))).toEqual(range(0, 0, 1, 1));
    expect(parsed(range(0, 0, 2, 1))).toBeUndefined();
    expect(parsed(range(1, 0, 1, 1))).toBeUndefined();
  });

  it("bounds style indices and every style field", () => {
    expect(ok(minimal({ styles: [{}], cells: [{ r: 0, c: 0, v: 1, s: 0 }] }))).toBe(true);
    expect(ok(minimal({ styles: [{}], cells: [{ r: 0, c: 0, v: 1, s: 1 }] }))).toBe(false);
    const style = (s: unknown): boolean => ok(minimal({ styles: [s] }));
    expect(style({ font: { family: "Aptos", sizePt: 11, bold: true, color: "#AABBCC" }, numberFormat: "0.00%", namedStyle: "Good" })).toBe(true);
    expect(style({ font: { family: { theme: "minor" } } })).toBe(true);
    expect(style({ font: { family: "x".repeat(256) } })).toBe(false);
    expect(style({ font: { family: "a\u0001b" } })).toBe(false);
    expect(style({ font: { sizePt: 0.5 } })).toBe(false);
    expect(style({ font: { sizePt: 410 } })).toBe(false);
    expect(style({ font: { bold: false } })).toBe(false);
    expect(style({ font: { color: "red" } })).toBe(false);
    expect(style({ font: { color: { theme: "accent1", tint: 0.4, lumMod: 1 } } })).toBe(true);
    expect(style({ font: { color: { theme: "accent9" } } })).toBe(false);
    expect(style({ font: { color: { theme: "accent1", shade: 2 } } })).toBe(false);
    expect(style({ borders: { top: { style: "thin" }, diagonal: { style: "double", up: true } } })).toBe(true);
    expect(style({ borders: { top: { style: "groovy" } } })).toBe(false);
    expect(style({ alignment: { horizontal: "center", vertical: "top", wrap: true, indent: 250, textRotation: 255 } })).toBe(true);
    expect(style({ alignment: { indent: 251 } })).toBe(false);
    expect(style({ alignment: { textRotation: 200 } })).toBe(false);
    expect(style({ alignment: { horizontal: "middle" } })).toBe(false);
    expect(style({ numberFormat: "x".repeat(256) })).toBe(false);
    expect(style({ numberFormat: '0.00"' })).toBe(false);
    expect(style({ builtinFormatId: 14 })).toBe(true);
    expect(style({ builtinFormatId: 99 })).toBe(false);
    expect(style({ namedStyle: "n".repeat(256) })).toBe(false);
    expect(style({ fill: { kind: "pattern", pattern: "gray125", fg: "#000000" } })).toBe(true);
    expect(style({ fill: { kind: "pattern", pattern: "url(x)" } })).toBe(false);
    expect(style({ protection: { locked: false, hidden: true } })).toBe(true);
    expect(style({ protection: { locked: true } })).toBe(false);
    expect(ok(minimal({ styles: Array.from({ length: 64_001 }, () => ({})) }))).toBe(false);
  });

  it("never keeps sourceXf or sourceStyleName, nor any unknown style key", () => {
    const parsed = parseClipboardPayload(encode(minimal({ styles: [{ font: { bold: true }, sourceXf: 4, sourceStyleName: "ce1", junk: 1 }] })))!;
    expect(parsed.styles).toEqual([{ font: { bold: true } }]);
  });

  it("bounds rich text: 32,767 characters in total, fonts as styles'; runs not spelling the value dropped", () => {
    const cell = (t: unknown, v: unknown = "ab") => parseClipboardPayload(encode(minimal({ cells: [{ r: 0, c: 0, v, t }] })));
    expect(cell([{ text: "a", font: { bold: true } }, { text: "b" }])!.cells[0]!.t).toHaveLength(2);
    expect(cell([{ text: "a" }, { text: "x" }])!.cells[0]!.t).toBeUndefined();
    expect(cell([{ text: "a", font: { sizePt: 999 } }])).toBeNull();
    expect(cell([{ text: "x".repeat(20_000) }, { text: "x".repeat(20_000) }], "x".repeat(32_767))).toBeNull();
  });

  it("bounds merges: inside the shape, at least two cells, pairwise disjoint", () => {
    expect(ok(minimal({ merges: [range(0, 0, 0, 1)] }))).toBe(true);
    expect(ok(minimal({ merges: [range(0, 0, 2, 0)] }))).toBe(false);
    expect(ok(minimal({ merges: [range(0, 0, 0, 0)] }))).toBe(false);
    expect(ok(minimal({ merges: [range(0, 0, 1, 0), range(1, 0, 1, 1)] }))).toBe(false);
    expect(ok(minimal({ merges: [range(0, 0, 1, 0), range(0, 1, 1, 1)] }))).toBe(true);
    expect(ok(minimal({ merges: undefined }))).toBe(false);
  });

  it("bounds row and column properties: only with their flag, inside the shape, ascending, sizes in SPEC-124's bounds", () => {
    expect(ok(minimal({ wholeRows: true, rowProps: [{ r: 0, heightPt: 409, customHeight: true, hidden: true }] }))).toBe(true);
    expect(ok(minimal({ rowProps: [{ r: 0, heightPt: 20 }] }))).toBe(false);
    expect(ok(minimal({ wholeRows: true, rowProps: [{ r: 0, heightPt: 410 }] }))).toBe(false);
    expect(ok(minimal({ wholeRows: true, rowProps: [{ r: 2 }] }))).toBe(false);
    expect(ok(minimal({ wholeRows: true, rowProps: [{ r: 1 }, { r: 0 }] }))).toBe(false);
    expect(ok(minimal({ wholeCols: true, colProps: [{ c: 1, widthPx: 120, hidden: true }] }))).toBe(true);
    expect(ok(minimal({ wholeCols: true, colProps: [{ c: 1, widthPx: -1 }] }))).toBe(false);
    expect(ok(minimal({ wholeCols: true, colProps: [{ c: 1, widthPx: 100_000 }] }))).toBe(false);
    expect(ok(minimal({ wholeRows: false }))).toBe(false);
  });

  it("bounds workbookId, sheetName and cut", () => {
    expect(ok(minimal({ workbookId: "w".repeat(256) }))).toBe(false);
    expect(ok(minimal({ sheetName: "" }))).toBe(false);
    expect(ok(minimal({ sheetName: "s".repeat(32) }))).toBe(false);
    expect(ok(minimal({ cut: "t".repeat(255) }))).toBe(true);
    expect(ok(minimal({ cut: "t".repeat(256) }))).toBe(false);
    expect(ok(minimal({ dateSystem: "1905" }))).toBe(false);
  });
});

describe("parseClipboardPayload — the security test", () => {
  const before = Object.getOwnPropertyNames(Object.prototype).toSorted();

  /** Pastes whatever the hostile attribute yields into a fresh workbook. */
  function pasteHostile(attribute: string): SpreadsheetModel {
    const m = model([sheet("s1", "S")]);
    const payload = parseClipboardPayload(attribute);
    if (payload) {
      pasteBlock(m, "s1", { ranges: [range(0, 0, 0, 0)], active: { row: 0, col: 0 } }, blockFromPayload(payload), { what: "all" });
    }
    return m;
  }

  it("refuses oversize, malformed and out-of-shape payloads — nothing reaches the model", () => {
    const hostile = [
      "A".repeat(MAX_PAYLOAD_ATTRIBUTE_LENGTH + 4),
      "%%%%",
      btoa("{"),
      encode(minimal({ rows: 1e9, cols: 1e9 })),
      encode(minimal({ cells: [{ r: 5, c: 5, v: "out of shape" }] })),
      encode(minimal({ merges: [range(0, 0, 9, 9)] })),
      encode(minimal({ cells: [{ r: 0, c: 0, v: 1, s: 7 }] })),
    ];
    for (const attribute of hostile) {
      expect(parseClipboardPayload(attribute)).toBeNull();
      expect(pasteHostile(attribute).sheets[0]!.rows).toEqual([]);
    }
  });

  it("is not polluted through __proto__, and rebuilds only whitelisted fields", () => {
    const json =
      '{"kind":"docu-edit/cells","version":1,"workbookId":"w","sheetName":"S","dateSystem":"1900","origin":{"row":0,"col":0},"rows":1,"cols":1,' +
      '"__proto__":{"polluted":true},"constructor":{"prototype":{"polluted":true}},' +
      '"cells":[{"r":0,"c":0,"v":"x","__proto__":{"polluted":true},"onclick":"alert(1)"}],' +
      '"styles":[{"__proto__":{"polluted":true},"font":{"bold":true,"__proto__":{"x":1}}}],"merges":[]}';
    const attribute = btoa(json);
    const parsed = parseClipboardPayload(attribute)!;
    expect(parsed).not.toBeNull();
    expect(Object.getPrototypeOf(parsed)).toBe(Object.prototype);
    expect(Object.keys(parsed.cells[0]!).toSorted()).toEqual(["c", "r", "v"]);
    expect(JSON.stringify(parsed.styles)).toBe('[{"font":{"bold":true}}]');
    const m = pasteHostile(attribute);
    expect(JSON.stringify(m.sheets[0]!.rows)).not.toContain("onclick");
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
    expect(Object.getOwnPropertyNames(Object.prototype).toSorted()).toEqual(before);
  });

  it("keeps script-bearing text inert: a value is text, a hostile formula or format is dropped or refused", () => {
    const attribute = encode(
      minimal({
        sheetName: "<script>x</script>",
        cells: [
          { r: 0, c: 0, v: "<script>alert(1)</script>" },
          { r: 0, c: 1, v: 0, f: "<script>alert(1)</script>" },
        ],
      }),
    );
    const parsed = parseClipboardPayload(attribute)!;
    expect(parsed.cells[0]!.v).toBe("<script>alert(1)</script>");
    expect(parsed.cells[1]!.f).toBeUndefined();
    expect(parseClipboardPayload(encode(minimal({ styles: [{ numberFormat: "<script>" }] })))).not.toBeNull(); // a literal code SPEC-118 accepts — inert text
    expect(parseClipboardPayload(encode(minimal({ styles: [{ font: { color: "javascript:alert(1)" } }] })))).toBeNull();
    const m = pasteHostile(attribute);
    expect(m.sheets[0]!.rows[0]!.cells.map((cell) => cell.formula)).toEqual([undefined, undefined]);
  });
});

describe("TSV both ways", () => {
  it("quotes a field holding a tab, a line break or a quote, doubling quotes", () => {
    expect(tsvField("plain")).toBe("plain");
    expect(tsvField("a\tb")).toBe('"a\tb"');
    expect(tsvField("a\nb")).toBe('"a\nb"');
    expect(tsvField('say "hi"')).toBe('"say ""hi"""');
    expect(tsvText([["a", "b"], ["c", ""]])).toBe("a\tb\r\nc\t");
  });

  it("parses quoted fields, every row separator, and ignores a trailing empty line", () => {
    expect(parseTsv("a\tb\r\nc\td\n")).toEqual([["a", "b"], ["c", "d"]]);
    expect(parseTsv("a\rb\nc")).toEqual([["a"], ["b"], ["c"]]);
    expect(parseTsv('"x\ty"\t"line\nbreak"\t"q ""t"""')).toEqual([["x\ty", "line\nbreak", 'q "t"']]);
    expect(parseTsv("a\t\tb")).toEqual([["a", "", "b"]]);
    expect(parseTsv("a\n\n")).toEqual([["a"], [""]]);
    expect(parseTsv("")).toEqual([]);
    expect(parseTsv(tsvText([["1,234.50", 'say "hi"\tthere'], ["x\r\ny", ""]]))).toEqual([["1,234.50", 'say "hi"\tthere'], ["x\r\ny", ""]]);
  });
});

describe("the flavours", () => {
  it("text/plain is each position's display text at full width, merged-away positions empty", () => {
    const m = sample();
    m.sheets[0]!.rows[1]!.cells[0]!.value = 123456789012;
    m.styles = [...m.styles, { numberFormat: "#,##0.00" }];
    m.sheets[0]!.rows[1]!.cells[0]!.styleId = 3;
    m.sheets[0]!.columns = [{ min: 1, max: 1, widthPx: 10 }];
    const { text } = clipboardFlavours(m, "s1", range(0, 1, 2, 4));
    expect(text).toBe('\t\tMerged\t\r\n123,456,789,012.00\t20\t\t\r\n20\t"a\tb"\t\t');
    expect(text).not.toContain("#");
  });

  it("text/html is a meta charset and a table carrying the payload, one td per position, merges as spans, whitelisted inline CSS", () => {
    const m = sample();
    const { html } = clipboardFlavours(m, "s1", range(0, 1, 2, 4));
    expect(html.startsWith('<meta charset="utf-8"><table data-docu-edit-cells="')).toBe(true);
    const doc = new DOMParser().parseFromString(html, "text/html");
    const table = doc.querySelector(`table[${PAYLOAD_ATTRIBUTE}]`)!;
    expect(parseClipboardPayload(table.getAttribute(PAYLOAD_ATTRIBUTE)!)).toEqual(buildClipboardPayload(m, "s1", range(0, 1, 2, 4)));
    const rows = Array.from(table.querySelectorAll("tr"));
    expect(rows.map((tr) => tr.querySelectorAll("td").length)).toEqual([3, 4, 4]);
    expect(rows[0]!.querySelectorAll("td")[2]!.getAttribute("colspan")).toBe("2");
    expect(rows[1]!.querySelector("td")!.getAttribute("style")).toBe("background-color:#ffff00");
    expect(rows[2]!.querySelectorAll("td")[1]!.getAttribute("style")).toBe("font-weight:700");
    expect(html.replace(/data-docu-edit-cells="[^"]*"/, "")).not.toMatch(/class=|<script|\son\w+=/i);
  });

  it("escapes text and resolves theme colours to hex in the CSS whitelist", () => {
    const m = model([sheet("s1", "S", { rows: [row(0, [value(0, '<b>&"x"</b>', { styleId: 1 })])] })], {
      styles: [
        {},
        {
          font: { family: "Aptos's Narrow", sizePt: 12, italic: true, underline: "single", strike: true, color: { theme: "accent1" } },
          alignment: { horizontal: "center", vertical: "center", wrap: true },
          borders: { top: { style: "medium", color: "#ff0000" }, bottom: { style: "double" } },
        },
      ],
    });
    const { html } = clipboardFlavours(m, "s1", range(0, 0, 0, 0));
    expect(html).toContain("&lt;b&gt;&amp;&quot;x&quot;&lt;/b&gt;</td>");
    expect(cellCss(m.styles[1]!, DEFAULT_THEME)).toBe(
      `font-family:'Aptoss Narrow';font-size:12pt;font-style:italic;text-decoration:underline line-through;color:${DEFAULT_THEME.colors.accent1};` +
        "text-align:center;vertical-align:middle;white-space:pre-wrap;border-top:2px solid #ff0000;border-bottom:3px double #000000",
    );
  });
});

// --- external sources ---------------------------------------------------------------------

const CTX = { locale: EN_US, dateSystem: "1900" as const };

function imported(html: string, ctx: { locale: typeof EN_US; dateSystem: "1900" | "1904" } = CTX): PasteBlock {
  const block = importClipboardHtml(html, ctx);
  if (block === null || typeof block === "string") {
    throw new Error(`no block: ${String(block)}`);
  }
  return block;
}

/** A block's cell at A1-style offset, with its style. */
function at(block: PasteBlock, r: number, c: number): { v: unknown; f?: string; r1c1?: string; style?: CellStyle } | undefined {
  const cell = block.cells.find((each) => each.r === r && each.c === c);
  if (!cell) {
    return undefined;
  }
  return { v: cell.v, ...(cell.f ? { f: cell.f } : {}), ...(cell.r1c1 ? { r1c1: cell.r1c1 } : {}), ...(cell.s !== undefined ? { style: block.styles[cell.s] } : {}) };
}

describe("SPEC-128: Copy over filter-hidden rows takes the visible rows only", () => {
  /** A1:C6 — Region | Amount over West, East, West, North, West, filtered to
   * West (rows 3 and 5 filter-hidden); C4 = B4*2, merged C5:C6 hidden-crossing. */
  function filtered(): SpreadsheetModel {
    return model([
      sheet("s1", "Sheet1", {
        rows: [
          row(0, [value(0, "Region"), value(1, "Amount")]),
          row(1, [value(0, "West"), value(1, 1)]),
          row(2, [value(0, "East"), value(1, 2)], { filterHidden: true }),
          row(3, [value(0, "West"), value(1, 3), value(2, 6, { formula: "B4*2" })]),
          row(4, [value(0, "North"), value(1, 4), value(2, "x")], { filterHidden: true }),
          row(5, [value(0, "West"), value(1, 5)]),
        ],
        merges: [range(4, 2, 5, 2)],
        autoFilter: { range: range(0, 0, 5, 1), columns: [{ colId: 0, criteria: { kind: "values", values: ["West"], blanks: false, dateGroups: [] } }] },
      }),
    ]);
  }

  it("packs the visible rows' spans together in the payload, a packed formula kept as written with its source row", () => {
    const copy = buildClipboardCopy(filtered(), "s1", range(0, 0, 5, 2), undefined, { visibleRows: true });
    expect(copy.spans).toEqual([range(0, 0, 1, 2), range(3, 0, 3, 2), range(5, 0, 5, 2)]);
    expect(copy.area.range).toEqual(range(0, 0, 5, 2));
    expect(copy.payload.rows).toBe(4);
    expect(copy.payload.cells.map((cell) => [cell.r, cell.c, cell.v])).toEqual([
      [0, 0, "Region"],
      [0, 1, "Amount"],
      [1, 0, "West"],
      [1, 1, 1],
      [2, 0, "West"],
      [2, 1, 3],
      [2, 2, 6],
      [3, 0, "West"],
      [3, 1, 5],
    ]);
    expect(copy.payload.cells[6]).toMatchObject({ f: "B4*2", fr: 3 });
    expect(copy.payload.cells.filter((cell) => cell.fr !== undefined)).toHaveLength(1);
    // The merge crosses a filter-hidden row: it does not travel.
    expect(copy.payload.merges).toEqual([]);
  });

  it("the flavours hold the visible rows only", () => {
    const m = filtered();
    const copy = buildClipboardCopy(m, "s1", range(0, 0, 5, 2), undefined, { visibleRows: true });
    const { text, html } = copyFlavours(m, "s1", copy);
    expect(text).toBe(tsvText([["Region", "Amount", ""], ["West", "1", ""], ["West", "3", "6"], ["West", "5", ""]]));
    const table = new DOMParser().parseFromString(html, "text/html").querySelector(`table[${PAYLOAD_ATTRIBUTE}]`)!;
    expect(table.querySelectorAll("tr").length).toBe(4);
    expect(parseClipboardPayload(table.getAttribute(PAYLOAD_ATTRIBUTE)!)).toEqual(copy.payload);
  });

  it("a paste of it moves each relative reference by the distance from its own source row (Excel)", () => {
    const m = filtered();
    const { payload } = buildClipboardCopy(m, "s1", range(0, 0, 5, 2), undefined, { visibleRows: true });
    pasteBlock(m, "s1", { ranges: [range(9, 4, 9, 4)], active: { row: 9, col: 4 } }, blockFromPayload(payload), { what: "all" });
    const pasted = m.sheets[0]!.rows.find((each) => each.index === 11)!;
    expect(pasted.cells.map((cell) => [cell.col, cell.value, cell.formula])).toEqual([
      [4, "West", undefined],
      [5, 3, undefined],
      [6, 6, "F12*2"],
    ]);
  });

  it("never passes a relative reference through a row above row 1: a formula far below what it names pastes as Excel's would", () => {
    const m = model([
      sheet("s1", "Sheet1", {
        rows: [
          row(0, [value(0, "Region")]),
          row(1, [value(0, "East")], { filterHidden: true }),
          row(2, [value(0, "East")], { filterHidden: true }),
          row(3, [value(0, "West"), value(2, "Region", { formula: "A1" })]),
        ],
        autoFilter: { range: range(0, 0, 3, 0), columns: [{ colId: 0, criteria: { kind: "values", values: ["West"], blanks: false, dateGroups: [] } }] },
      }),
    ]);
    const { payload } = buildClipboardCopy(m, "s1", range(0, 0, 3, 2), undefined, { visibleRows: true });
    expect(payload.cells.at(-1)).toMatchObject({ r: 1, c: 2, f: "A1", fr: 3 });
    expect(parseClipboardPayload(encodePayload(payload))).toEqual(payload);
    pasteBlock(m, "s1", { ranges: [range(9, 4, 9, 4)], active: { row: 9, col: 4 } }, blockFromPayload(payload), { what: "all" });
    expect(m.sheets[0]!.rows.find((each) => each.index === 10)!.cells.find((cell) => cell.col === 6)?.formula).toBe("E8");
  });

  it("takes every row without filter-hidden rows inside, for a cut, and unless asked (a drag)", () => {
    const m = filtered();
    expect(buildClipboardCopy(m, "s1", range(0, 0, 1, 2), undefined, { visibleRows: true }).spans).toBeUndefined();
    const cut = buildClipboardCopy(m, "s1", range(0, 0, 5, 2), "cut-token", { visibleRows: true });
    expect([cut.spans, cut.payload.rows]).toEqual([undefined, 6]);
    const drag = buildClipboardCopy(m, "s1", range(0, 0, 5, 2));
    expect([drag.spans, drag.payload.rows, drag.payload.merges]).toEqual([undefined, 6, [range(4, 2, 5, 2)]]);
  });
});

describe("external sources — each producer's markers (inline snippets)", () => {
  it("Excel: xmlns:x, x:num/x:bool/x:err/x:str, mso-number-format from class rules with CSS escapes, class rules read as data", () => {
    const html =
      '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><style><!--td {color:black; font-size:11.0pt; font-family:"Aptos Narrow", sans-serif; white-space:nowrap; mso-number-format:General;} ' +
      '.xl65 {mso-number-format:"\\0022$\\0022\\#\\,\\#\\#0\\.00"; font-weight:700;} .xl66 {mso-number-format:"Short Date";} .xl67 {white-space:normal; background:yellow; border-top:.5pt solid windowtext; border-bottom:2.0pt double red;}--></style></head><body><table>' +
      '<tr><td class="xl65" x:num="1234.5">$1,234.50</td><td x:num>1,234</td><td x:bool="TRUE">TRUE</td><td x:err="#DIV/0!">#DIV/0!</td><td x:str>007</td></tr>' +
      '<tr><td class="xl66" x:num="46096">3/15/2026</td><td class="xl67" style="mso-number-format:\'0%\'" x:num="0.25">25%</td><td x:fmla="=A1">ignored formula</td></tr></table></body></html>';
    const block = imported(html);
    expect(block.source).toBe("excel");
    expect(at(block, 0, 0)).toMatchObject({ v: 1234.5, style: { numberFormat: '"$"#,##0.00', font: { bold: true, family: "Aptos Narrow", sizePt: 11, color: "#000000" } } });
    expect(at(block, 0, 1)!.v).toBe(1234);
    expect(at(block, 0, 2)!.v).toBe(true);
    expect(at(block, 0, 3)!.v).toEqual({ error: "#DIV/0!" });
    expect(at(block, 0, 4)!.v).toBe("007");
    expect(at(block, 1, 0)!.style).toMatchObject({ numberFormat: "m/d/yyyy", builtinFormatId: 14 });
    expect(at(block, 1, 1)!.style).toMatchObject({
      numberFormat: "0%",
      alignment: { wrap: true },
      fill: { kind: "solid", color: "#ffff00" },
      borders: { top: { style: "thin" }, bottom: { style: "double", color: "#ff0000" } },
    });
    expect(at(block, 1, 2)!.f).toBeUndefined();
  });

  it("Google Sheets: data-sheets-value types, data-sheets-numberformat and R1C1 formulas", () => {
    const html =
      '<google-sheets-html-origin><table data-sheets-root="1"><tr>' +
      '<td data-sheets-value=\'{"1":2,"2":"Widget"}\'>Widget</td>' +
      '<td data-sheets-value=\'{"1":3,"3":1234567}\' data-sheets-numberformat=\'{"1":2,"2":"#,##0","3":1}\'>1,234,567</td>' +
      '<td data-sheets-value=\'{"1":4,"4":1}\'>TRUE</td>' +
      '<td data-sheets-value=\'{"1":3,"3":1235801.5}\' data-sheets-formula="=R[0]C[-2]+R[0]C[-1]" style="font-weight:bold;">1235801.5</td>' +
      "<td data-sheets-value='not json'>12</td></tr></table></google-sheets-html-origin>";
    const block = imported(html);
    expect(block.source).toBe("sheets");
    expect(at(block, 0, 0)).toEqual({ v: "Widget" });
    expect(at(block, 0, 1)).toEqual({ v: 1234567, style: { numberFormat: "#,##0" } });
    expect(at(block, 0, 2)!.v).toBe(true);
    expect(at(block, 0, 3)).toEqual({ v: 1235801.5, r1c1: "=R[0]C[-2]+R[0]C[-1]", style: { font: { bold: true } } });
    expect(at(block, 0, 4)!.v).toBe(12);
    expect(r1c1ToA1("=R[0]C[-2]+R[0]C[-1]", { row: 0, col: 6 })).toBe("E1+F1");
    expect(r1c1ToA1("=SUM(R[-3]C:R[-1]C)", { row: 9, col: 1 })).toBe("SUM(B7:B9)");
    expect(r1c1ToA1("='My Sheet'!R2C3&\"R1C1\"", { row: 0, col: 0 })).toBe("'My Sheet'!$C$2&\"R1C1\"");
    expect(r1c1ToA1("=R[-1]C", { row: 0, col: 0 })).toBeNull();
    expect(r1c1ToA1("=SUM(C[1])", { row: 0, col: 0 })).toBeNull();
    expect(r1c1ToA1("=ROUND(R1C1,2)", { row: 4, col: 4 })).toBe("ROUND($A$1,2)");
    // Transposed: a reference relative on both axes has its offset transposed;
    // absolute and mixed parts convert at the cell as ever.
    expect(r1c1ToA1("=R[-1]C[2]+R2C[1]+R[1]C3+RC", { row: 5, col: 5 })).toBe("H5+G$2+$C7+F6");
    expect(r1c1ToA1("=R[-1]C[2]+R2C[1]+R[1]C3+RC", { row: 5, col: 5 }, true)).toBe("E8+G$2+$C7+F6");
  });

  it("LibreOffice: sdval, sdnum codes, a BOOLEAN code, inline CSS and font/b/i/u tags", () => {
    const html =
      '<table><tr><td sdval="1234567" sdnum="1033;0;#,##0"><font face="Liberation Sans" size=2 color="#000080">1,234,567</font></td>' +
      '<td sdval="1" sdnum="1033;0;BOOLEAN">TRUE</td><td sdval="0.25" sdnum="1033;0;0%"><b><i>25%</i></b></td>' +
      '<td sdnum="1033;0;@"><u>Widget</u></td><td sdval="46096" sdnum="1033;1033;MM/DD/YY">03/15/26</td></tr></table>';
    const block = imported(html);
    expect(block.source).toBe("libreoffice");
    expect(at(block, 0, 0)).toEqual({ v: 1234567, style: { numberFormat: "#,##0", font: { family: "Liberation Sans", sizePt: 10, color: "#000080" } } });
    expect(at(block, 0, 1)!.v).toBe(true);
    expect(at(block, 0, 2)).toEqual({ v: 0.25, style: { numberFormat: "0%", font: { bold: true, italic: true } } });
    expect(at(block, 0, 3)).toEqual({ v: "Widget", style: { numberFormat: "@", font: { underline: "single" } } });
    expect(dateKindOf(at(block, 0, 4)!.style!.numberFormat)).toBe("date");
  });

  it("any other table: parseInput of the text, inline CSS and b/i/u/s tags, colspan/rowspan as merges", () => {
    const html =
      '<table><tr><td style="color:#ff0000;font-size:16px;text-align:center">1,234.50</td><td><b>bold</b></td><td><s>gone</s> <i>mixed</i></td><td colspan="2">=1+1</td></tr>' +
      '<tr><td rowspan="2">10/8/2026</td><td>TRUE</td><td>x<br>y</td><td>$5</td><td>12%</td></tr><tr><td>after</td></tr></table>';
    const block = imported(html);
    expect(block.source).toBe("html");
    expect(at(block, 0, 0)).toEqual({ v: 1234.5, style: { numberFormat: "#,##0.00", font: { color: "#ff0000", sizePt: 12 }, alignment: { horizontal: "center" } } });
    expect(at(block, 0, 1)!.style).toEqual({ font: { bold: true } });
    expect(at(block, 0, 2)!.style).toBeUndefined();
    expect(at(block, 0, 3)).toEqual({ v: null, f: "1+1" });
    expect(at(block, 1, 0)!.v).toBe(dateToSerial(2026, 10, 8, "1900"));
    expect(at(block, 1, 1)!.v).toBe(true);
    expect(at(block, 1, 2)!.v).toBe("x\ny");
    expect(at(block, 1, 3)!.v).toBe(5);
    expect(at(block, 2, 1)!.v).toBe("after"); // placed past the rowspan
    expect(block.merges).toEqual([range(0, 3, 0, 4), range(1, 0, 2, 0)]);
    expect([block.rows, block.cols]).toEqual([3, 5]);
  });

  it("moves attribute serials with a date format by 1462 days into a 1904 workbook", () => {
    const html = '<table><tr><td sdval="46096" sdnum="1033;0;YYYY-MM-DD">2026-03-15</td><td sdval="46096" sdnum="1033;0;0">46096</td></tr></table>';
    const block = imported(html, { locale: EN_US, dateSystem: "1904" });
    expect(at(block, 0, 0)!.v).toBe(46096 - 1462);
    expect(at(block, 0, 1)!.v).toBe(46096);
    expect(block.dateSystem).toBe("1904");
  });

  it("reads CSS as data: declarations, escapes, colours", () => {
    expect([...cssDeclarations('a:1; b:"x;y" ; C : url(a;b) !important;')]).toEqual([["a", "1"], ["b", '"x;y"'], ["c", "url(a;b)"]]);
    expect(decodeCssString('"\\0022$\\0022\\#\\,\\#\\#0\\.00"')).toBe('"$"#,##0.00');
    expect(cssColor("windowtext")).toBeUndefined();
    expect(cssColor("Navy")).toBe("#000080");
    expect(cssColor("rgb(1, 2, 3)")).toBe("#010203");
    expect(cssColor("url(https://evil.example/x)")).toBeUndefined();
    expect(importedStyle(new Map([["font-weight", "600"], ["text-decoration", "underline"], ["text-underline-style", "double"], ["mso-rotate", "-45"]]))).toEqual({
      font: { bold: true, underline: "double" },
      alignment: { textRotation: 135 },
    });
  });

  it("decodes Excel's escapes as Excel writes them: exactly four hex digits, no terminating space", () => {
    expect(decodeCssString('"\\0022Dec\\0022"')).toBe('"Dec"');
    expect(decodeCssString('"0\\0022 kg\\0022"')).toBe('0" kg"');
    expect(decodeCssString('"\\0022Q\\00220"')).toBe('"Q"0');
    expect(decodeCssString('"\\[$-409\\]mmm\\\\ d\\\\\\,\\\\ yyyy\\;\\@"')).toBe("[$-409]mmm\\ d\\,\\ yyyy;@");
    const html =
      '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><style>.xl70 {mso-number-format:"0\\.0\\0022 kg\\0022";} .xl71 {mso-number-format:"\\0022Dec \\00220";}</style></head>' +
      '<body><table><tr><td class="xl70" x:num="12.5">12.5 kg</td><td class="xl71" x:num="3">Dec 3</td></tr></table></body></html>';
    const block = imported(html);
    expect(at(block, 0, 0)).toEqual({ v: 12.5, style: { numberFormat: '0.0" kg"' } });
    expect(at(block, 0, 1)).toEqual({ v: 3, style: { numberFormat: '"Dec "0' } });
  });

  it("refuses input over 64 MiB or over MAX_COPY_CELLS cells before parsing; HTML without a table is no block", () => {
    expect(importClipboardHtml(`<table><tr><td>x</td></tr></table>${" ".repeat(64 * 1024 * 1024)}`, CTX)).toBe("That's too many cells to change at once");
    expect(importClipboardHtml(`<table><tr>${"<td>".repeat(MAX_COPY_CELLS + 1)}</tr></table>`, CTX)).toBe(TOO_MANY_TO_CHANGE_MESSAGE);
    expect(importClipboardHtml("<p>no table</p>", CTX)).toBeNull();
  });

  it("reads hostile CSS in linear time: brace-less style text, long digit runs (quadratic regexes took minutes)", () => {
    const digits = "1".repeat(100_000);
    const html =
      `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><style>${"a".repeat(400_000)}} @page {margin:1in} .xl65 {font-weight:700}</style></head>` +
      `<body><table><tr><td class="xl65" style="font-size:${digits}x;border-top:${digits}x solid red">x</td></tr></table></body></html>`;
    expect(at(imported(html), 0, 0)).toEqual({ v: "x", style: { font: { bold: true }, borders: { top: { style: "thin", color: "#ff0000" } } } });
  });

  it("refuses a table whose spans reach past MAX_COPY_CELLS as it is laid out, never walking them", () => {
    // 20,000 × 1,000 positions: past V8's Set limit, were they ever marked.
    expect(importClipboardHtml('<table><tr><td rowspan="20000" colspan="1000">x</td></tr></table>', CTX)).toBe(TOO_MANY_TO_CHANGE_MESSAGE);
    expect(importClipboardHtml('<table><tr><td colspan="1000">a</td></tr><tr><td rowspan="1001">b</td></tr></table>', CTX)).toBe(TOO_MANY_TO_CHANGE_MESSAGE);
    expect(imported('<table><tr><td rowspan="2" colspan="3">x</td><td>y</td></tr></table>')).toMatchObject({ rows: 2, cols: 4, merges: [range(0, 0, 1, 2)] });
  });

  it("refuses spans that overlap past MAX_COPY_CELLS positions walked, though the table's area stays inside it", () => {
    // Each row's 500 × 500 span starts one column left of the last, inside a
    // 999 × 1,000 area: walked in full, 125 million positions.
    const staircase = Array.from({ length: 500 }, (_, r) => `<tr><td colspan="${500 - r}">f</td><td rowspan="500" colspan="500">x</td></tr>`).join("");
    expect(importClipboardHtml(`<table>${staircase}</table>`, CTX)).toBe(TOO_MANY_TO_CHANGE_MESSAGE);
    // Spans that do not overlap walk at most the area: exactly the cap lands.
    expect(imported('<table><tr><td rowspan="1000" colspan="1000">x</td></tr></table>')).toMatchObject({ rows: 1000, cols: 1000, merges: [range(0, 0, 999, 999)] });
  });

  it("holds what attributes give to the payload's bounds: XML-legal text of at most 32,767 characters, format codes of at most 255, a Sheets formula of at most 8,192", () => {
    const excel = imported(
      `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table><tr><td x:str="a\u0001b">t</td><td x:str="${"x".repeat(40_000)}">t</td>` +
        `<td style="mso-number-format:'${"0".repeat(256)}'" x:num="1">1</td><td style="mso-number-format:'${"0".repeat(255)}'" x:num="1">1</td></tr></table></body></html>`,
    );
    expect(at(excel, 0, 0)!.v).toBe("ab");
    expect(at(excel, 0, 1)!.v).toBe("x".repeat(32_767));
    expect(at(excel, 0, 2)).toEqual({ v: 1 });
    expect(at(excel, 0, 3)).toEqual({ v: 1, style: { numberFormat: "0".repeat(255) } });
    const sheets = imported(
      `<google-sheets-html-origin><table data-sheets-root="1"><tr><td data-sheets-value='{"1":2,"2":"a\\u0001b"}'>ab</td>` +
        `<td data-sheets-value='{"1":3,"3":1}' data-sheets-formula="=${"1+".repeat(4096)}1">1</td>` +
        `<td data-sheets-value='{"1":3,"3":1}' data-sheets-formula="=${"1+".repeat(4095)}1">1</td></tr></table></google-sheets-html-origin>`,
    );
    expect(at(sheets, 0, 0)!.v).toBe("ab");
    expect(at(sheets, 0, 1)).toEqual({ v: 1 });
    expect(at(sheets, 0, 2)!.r1c1).toHaveLength(8192);
  });

  it("passes SPEC-016's sanitisation list: script, style with url(), on*, javascript:, remote <img>, iframe", () => {
    const html =
      '<html><head><style>td { background: url(https://evil.example/bg) }</style><script>window.__xss = 1</script></head><body>' +
      '<table onmouseover="window.__xss = 2"><tr>' +
      '<td style="background:url(https://evil.example/x);color:expression(alert(1));font-family:url(evil)" onclick="window.__xss = 3">safe ' +
      '<a href="javascript:window.__xss = 4">link</a><img src="https://evil.example/tracker.png" onerror="window.__xss = 5"></td>' +
      '<td><iframe src="https://evil.example/frame"></iframe><script>window.__xss = 6</script>after</td></tr></table></body></html>';
    const block = imported(html);
    expect(block.cells.map((cell) => cell.v)).toEqual(["safe link", "after"]);
    expect(block.styles).toEqual([]);
    const serialised = JSON.stringify(block);
    for (const hostile of ["script", "__xss", "iframe", "evil.example", "javascript:", "url(", "onclick", "expression"]) {
      expect(serialised).not.toContain(hostile);
    }
    expect(Reflect.get(window, "__xss")).toBeUndefined();
  });
});

// --- the captures ------------------------------------------------------------------------------

interface CaptureExpectation {
  /** H1:I1 merged. */
  mergedPair: boolean;
  /** G1 is the formula =B1+C1 (Google Sheets), else its value. */
  formula: boolean;
}

const CAPTURE_EXPECTATIONS: Record<string, CaptureExpectation & { source: PasteBlock["source"] }> = {
  excel: { source: "excel", mergedPair: true, formula: false },
  sheets: { source: "sheets", mergedPair: true, formula: true },
  libreoffice: { source: "libreoffice", mergedPair: true, formula: false },
  numbers: { source: "html", mergedPair: true, formula: false },
};

describe("the real captures (e2e/fixtures/spec125-clipboard) to the expected cells of \"Capture fixtures\"", () => {
  for (const [name, expectation] of Object.entries(CAPTURE_EXPECTATIONS)) {
    const path = join(CAPTURES, `${name}.html`);
    it(`${name}.html: A1:I1 — text, #,##0, $ currency, 0%, a date, TRUE, =B1+C1, a merged pair`, (ctx) => {
      ctx.skip(!existsSync(path), `fixture not yet captured: ${name}.html`);
      const html = readFileSync(path, "utf8");
      const block = imported(html);
      expect(block.source).toBe(expectation.source);
      const row0 = (c: number) => at(block, 0, c);
      expect(row0(0)!.v).toBe("Widget");
      expect(row0(1)!.v).toBe(1234567);
      expect(row0(1)!.style?.numberFormat).toMatch(/#,##0/);
      expect(row0(2)!.v).toBe(1234.5);
      expect(row0(2)!.style?.numberFormat).toMatch(/\$/);
      expect(row0(3)!.v).toBe(0.25);
      expect(row0(3)!.style?.numberFormat).toMatch(/%/);
      expect(row0(4)!.v).toBe(dateToSerial(2026, 3, 15, "1900"));
      expect(dateKindOf(row0(4)!.style?.builtinFormatId === 14 ? "m/d/yyyy" : row0(4)!.style?.numberFormat)).toBe("date");
      expect(row0(5)!.v).toBe(true);
      expect(row0(6)!.v).toBe(1235801.5);
      if (expectation.formula) {
        expect(r1c1ToA1(row0(6)!.r1c1!, { row: 0, col: 6 })).toBe("B1+C1");
      }
      expect(row0(7)!.v).toBe("Merged");
      if (expectation.mergedPair) {
        expect(block.merges).toContainEqual(range(0, 7, 0, 8));
      }
    });
  }
});
