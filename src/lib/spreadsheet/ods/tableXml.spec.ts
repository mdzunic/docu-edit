// SPEC-117 "Reading ODS sheet layout" (with SPEC-116's cell rules it
// extends): column widths, visibility and default cell styles with the
// last-run default; row heights, custom heights, visibility and default
// styles; spans → merges; covered-cell storage; text:span fonts; the styled
// repeat-run rules (the 16380-column run that stores no cell) and their cell
// budget — over synthetic content and the LibreOffice-shaped fixture.
import { describe, expect, it } from "vitest";
import { getCell, parseA1, parseRange, type GridRange } from "../address";
import { WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { buildSpec117StylesOdsBytes, SPEC117_STYLES_ODS } from "../fixtures/build";
import { acceptMerges } from "../merges";
import type { Sheet } from "../model";
import { effectiveStyleId } from "../styles";
import { parseOds } from "./reader";
import { acceptMergesByBand } from "./tableXml";

const encoder = new TextEncoder();

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ` +
  `xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"`;

/** Automatic styles most tests share: column widths, row heights, cell styles. */
const STYLES =
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="0.889in"/></style:style>` +
  `<style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="2.54cm"/></style:style>` +
  `<style:style style:name="co3" style:family="table-column"><style:table-column-properties fo:break-before="auto"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="30pt" style:use-optimal-row-height="false"/></style:style>` +
  `<style:style style:name="ro3" style:family="table-row"><style:table-row-properties style:row-height="20pt" style:use-optimal-row-height="true"/></style:style>` +
  `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold"/></style:style>` +
  `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ffff00"/></style:style>` +
  `<style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-style="italic"/></style:style>` +
  `<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>` +
  `<style:style style:name="T2" style:family="text"><style:text-properties fo:color="#ff0000"/></style:style>`;

function workbook(body: string, options: { automatic?: string; common?: string } = {}): Map<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>();
  parts.set("mimetype", encoder.encode("application/vnd.oasis.opendocument.spreadsheet"));
  parts.set(
    "content.xml",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS}>` +
        `<office:automatic-styles>${options.automatic ?? STYLES}</office:automatic-styles>` +
        `<office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`,
    ),
  );
  parts.set(
    "styles.xml",
    encoder.encode(
      `<office:document-styles ${NS}><office:styles>${options.common ?? `<style:style style:name="Default" style:family="table-cell"/>`}` +
        `</office:styles></office:document-styles>`,
    ),
  );
  return parts;
}

const table = (columns: string, rows: string): string => `<table:table table:name="T">${columns}${rows}</table:table>`;
const column = (attrs: string): string => `<table:table-column ${attrs}/>`;
const row = (cells: string, attrs = ""): string => `<table:table-row${attrs ? ` ${attrs}` : ""}>${cells}</table:table-row>`;
const float = (value: number, attrs = ""): string =>
  `<table:table-cell${attrs ? ` ${attrs}` : ""} office:value-type="float" office:value="${value}"><text:p>${value}</text:p></table:table-cell>`;
const empty = (repeat: number, style?: string): string =>
  `<table:table-cell${style ? ` table:style-name="${style}"` : ""}${repeat > 1 ? ` table:number-columns-repeated="${repeat}"` : ""}/>`;
const string = (inner: string, attrs = ""): string =>
  `<table:table-cell${attrs ? ` ${attrs}` : ""} office:value-type="string"><text:p>${inner}</text:p></table:table-cell>`;

async function read(body: string, options?: { automatic?: string; common?: string }) {
  const parsed = await parseOds(new Uint8Array(0), workbook(body, options));
  const sheet = parsed.sheets[0]!;
  const id = (name: string): number => parsed.styles.findIndex((style) => style.sourceStyleName === name);
  return { parsed, sheet, id };
}

function styleAt(sheet: Sheet, a1: string): number {
  const ref = parseA1(a1)!;
  return effectiveStyleId(sheet, ref.row, ref.col);
}

describe("columns", () => {
  it("reads widths in any ODF length as unrounded px, the last run's width being the default", async () => {
    const { sheet } = await read(
      table(column(`table:style-name="co2"`) + column(`table:style-name="co1" table:number-columns-repeated="16383"`), row(float(1))),
    );
    expect(sheet.defaultColWidthPx).toBe(85.344);
    expect(sheet.columns).toEqual([{ min: 0, max: 0, widthPx: 96, sourceStyleName: "co2" }]);
  });

  it("falls back to the table-column default style's width, then to 85.33 px", async () => {
    const withDefault = `<style:default-style style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:default-style>`;
    const { sheet } = await read(table(column(`table:style-name="co3" table:number-columns-repeated="16384"`), row(float(1))), {
      common: withDefault,
    });
    expect(sheet.defaultColWidthPx).toBe(96);
    expect(sheet.columns).toEqual([]);
    const bare = await read(table(column(`table:style-name="co3"`), row(float(1))));
    expect(bare.sheet.defaultColWidthPx).toBe(85.33);
    const none = await read(table("", row(float(1))));
    expect(none.sheet.defaultColWidthPx).toBe(85.33);
  });

  it("stores a run only for a width of its own, hidden or a default cell style; collapse and filter hide", async () => {
    const { sheet, id } = await read(
      table(
        column(`table:style-name="co1" table:default-cell-style-name="Default"`) +
          column(`table:style-name="co1" table:visibility="collapse" table:number-columns-repeated="2"`) +
          column(`table:style-name="co1" table:visibility="filter"`) +
          column(`table:style-name="co1" table:default-cell-style-name="ce1" table:number-columns-repeated="2"`) +
          column(`table:style-name="co1" table:visibility="visible" table:number-columns-repeated="16378"`),
        row(float(1)),
      ),
    );
    expect(sheet.columns).toEqual([
      { min: 1, max: 2, hidden: true, sourceStyleName: "co1" },
      { min: 3, max: 3, hidden: true, sourceStyleName: "co1" },
      { min: 4, max: 5, styleId: id("ce1"), sourceStyleName: "co1" },
    ]);
    expect(styleAt(sheet, "E1")).toBe(id("ce1"));
    expect(styleAt(sheet, "A1")).toBe(0);
  });

  it("clamps a run reaching past XFD", async () => {
    const { sheet } = await read(
      table(column(`table:style-name="co1" table:number-columns-repeated="16380"`) + column(`table:style-name="co2" table:number-columns-repeated="10"`), row(float(1))),
    );
    expect(sheet.defaultColWidthPx).toBe(96);
    expect(sheet.columns).toEqual([{ min: 0, max: 16379, widthPx: 85.344, sourceStyleName: "co1" }]);
  });
});

describe("rows", () => {
  it("the last row run's height is the default; equal rows store nothing; others store their height", async () => {
    const { sheet } = await read(
      table(
        "",
        row(float(1), `table:style-name="ro1"`) +
          row(float(2), `table:style-name="ro2"`) +
          row(float(3), `table:style-name="ro3"`) +
          row(empty(16384), `table:style-name="ro2" table:number-rows-repeated="2"`) +
          row(empty(16384), `table:style-name="ro1" table:number-rows-repeated="1048570"`),
      ),
    );
    expect(sheet.defaultRowHeightPt).toBe(12.816);
    expect(sheet.rows).toEqual([
      { index: 0, cells: [{ col: 0, value: 1 }], sourceStyleName: "ro1" },
      { index: 1, cells: [{ col: 0, value: 2 }], sourceStyleName: "ro2", heightPt: 30, customHeight: true },
      { index: 2, cells: [{ col: 0, value: 3 }], sourceStyleName: "ro3", heightPt: 20 },
      { index: 3, cells: [], sourceStyleName: "ro2", heightPt: 30, customHeight: true },
      { index: 4, cells: [], sourceStyleName: "ro2", heightPt: 30, customHeight: true },
    ]);
  });

  it("with no styled last run the default is 12.8 pt", async () => {
    const { sheet } = await read(table("", row(float(1), `table:style-name="ro3"`) + row(empty(16384), `table:number-rows-repeated="1048575"`)));
    expect(sheet.defaultRowHeightPt).toBe(12.8);
    expect(sheet.rows[0]).toMatchObject({ heightPt: 20 });
    expect(sheet.rows).toHaveLength(1);
  });

  it("stores a hidden row (collapse or filter) and a row's default cell style, cells or not", async () => {
    const { sheet, id } = await read(
      table(
        "",
        row(float(1), `table:style-name="ro1" table:visibility="collapse"`) +
          row(empty(16384), `table:style-name="ro1" table:visibility="filter" table:number-rows-repeated="2"`) +
          row(empty(16384), `table:style-name="ro1" table:default-cell-style-name="ce2"`) +
          row(empty(16384), `table:style-name="ro1" table:default-cell-style-name="Default"`) +
          row(empty(16384), `table:style-name="ro1" table:number-rows-repeated="1048571"`),
      ),
    );
    expect(sheet.rows).toEqual([
      { index: 0, cells: [{ col: 0, value: 1 }], hidden: true, sourceStyleName: "ro1" },
      { index: 1, cells: [], hidden: true, sourceStyleName: "ro1" },
      { index: 2, cells: [], hidden: true, sourceStyleName: "ro1" },
      // "Default" as a row's default says nothing when no column sets a style.
      { index: 3, cells: [], styleId: id("ce2"), sourceStyleName: "ro1" },
    ]);
    expect(styleAt(sheet, "Z4")).toBe(id("ce2"));
  });

  it("a row's own default style beats the column's; a cell's beats both", async () => {
    const { sheet, id } = await read(
      table(
        column(`table:default-cell-style-name="ce1" table:number-columns-repeated="16384"`),
        row(float(1) + float(2, `table:style-name="ce3"`), `table:default-cell-style-name="Default"`) + row(float(3)),
      ),
    );
    expect(sheet.rows[0]!.styleId).toBe(0);
    expect(styleAt(sheet, "A1")).toBe(0);
    expect(styleAt(sheet, "B1")).toBe(id("ce3"));
    expect(styleAt(sheet, "A2")).toBe(id("ce1"));
  });

  it("costs nothing for LibreOffice's trailing 1048553-row run", async () => {
    const started = performance.now();
    const { sheet } = await read(
      table(
        column(`table:style-name="co1" table:number-columns-repeated="16384" table:default-cell-style-name="Default"`),
        row(float(1) + empty(16383), `table:style-name="ro1"`) +
          row(empty(16384), `table:style-name="ro2" table:number-rows-repeated="22"`) +
          row(empty(16384), `table:style-name="ro1" table:number-rows-repeated="1048553"`),
      ),
    );
    expect(performance.now() - started).toBeLessThan(500);
    expect(sheet.rows).toHaveLength(23);
    expect(sheet.rows.at(-1)!.index).toBe(22);
  });

  it("counts rows stored for their properties alone against the cell ceiling", async () => {
    // Two million hidden rows from a few hundred bytes: refused, not built.
    const body =
      table("", row(empty(16384), `table:visibility="collapse" table:number-rows-repeated="1048576"`)) +
      `<table:table table:name="U">${row(empty(16384), `table:visibility="collapse" table:number-rows-repeated="1048576"`)}</table:table>`;
    await expect(parseOds(new Uint8Array(0), workbook(body))).rejects.toThrow(WORKBOOK_TOO_LARGE_MESSAGE);
  });
});

describe("spans → merges", () => {
  it("reads number-columns-spanned / number-rows-spanned as merges, in document order", async () => {
    const { sheet } = await read(
      table(
        "",
        row(string("Title", `table:number-columns-spanned="3" table:number-rows-spanned="2"`) + `<table:covered-table-cell table:number-columns-repeated="2"/>`) +
          row(`<table:covered-table-cell table:number-columns-repeated="3"/>` + empty(1, undefined) + `<table:table-cell table:number-columns-spanned="2"/>`),
      ),
    );
    expect(sheet.merges).toEqual([parseRange("A1:C2"), parseRange("E2:F2")]);
  });

  it("keeps ranges of two or more cells overlapping no earlier merge; repeated rows repeat their spans", async () => {
    const { sheet } = await read(
      table(
        "",
        row(`<table:table-cell table:number-columns-spanned="1" table:number-rows-spanned="1"/><table:table-cell table:number-rows-spanned="3"/>`) +
          row(empty(1) + `<table:table-cell table:number-columns-spanned="2"/>`) +
          row(`<table:table-cell table:number-columns-spanned="2"/>`, `table:number-rows-repeated="2"`),
      ),
    );
    // B1:B3 is kept; B2:C2 overlaps it; A3:B3 overlaps it; A4:B4 is kept.
    expect(sheet.merges).toEqual([parseRange("B1:B3"), parseRange("A4:B4")]);
  });

  it("clamps a span at the grid's edge", async () => {
    const { sheet } = await read(
      table("", row(empty(16382) + `<table:table-cell table:number-columns-spanned="5"/>`) + row(empty(1), `table:number-rows-repeated="1048573"`) + row(`<table:table-cell table:number-rows-spanned="9"/>`)),
    );
    expect(sheet.merges).toEqual([parseRange("XFC1:XFD1"), parseRange("A1048575:A1048576")]);
  });

  it("decides merges band by band exactly as acceptMerges does over the whole list", () => {
    let seed = 7;
    const random = (n: number): number => {
      seed = (seed * 1103515245 + 12345) % 2147483648;
      return seed % n;
    };
    for (let trial = 0; trial < 50; trial += 1) {
      const candidates: GridRange[] = [];
      let top = 0;
      for (let i = 0; i < 40; i += 1) {
        top += random(3);
        const left = random(10);
        candidates.push({ top, left, bottom: top + random(4), right: left + random(3) });
      }
      expect(acceptMergesByBand(candidates)).toEqual(acceptMerges(candidates));
    }
  });
});

describe("covered cells", () => {
  it("are stored when they carry a value or a style, never when they carry neither", async () => {
    const { sheet, id } = await read(
      table(
        "",
        row(
          string("Merged", `table:number-columns-spanned="4"`) +
            `<table:covered-table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:covered-table-cell>` +
            `<table:covered-table-cell table:style-name="ce2" table:number-columns-repeated="2"/>` +
            `<table:covered-table-cell/>` +
            float(5),
        ),
      ),
    );
    expect(sheet.rows[0]!.cells).toEqual([
      { col: 0, value: "Merged" },
      { col: 1, value: 7 },
      { col: 2, value: null, styleId: id("ce2") },
      { col: 3, value: null, styleId: id("ce2") },
      { col: 5, value: 5 },
    ]);
    expect(sheet.merges).toEqual([parseRange("A1:D1")]);
  });
});

describe("text:span runs", () => {
  it("give a string cell runs, each span's font its text style's own", async () => {
    const { sheet } = await read(
      table("", row(string(`Mixed <text:span text:style-name="T1">bold</text:span> and <text:span text:style-name="T2">red</text:span>`))),
    );
    expect(getCell(sheet, 0, 0)).toEqual({
      col: 0,
      value: "Mixed bold and red",
      runs: [{ text: "Mixed " }, { text: "bold", font: { bold: true } }, { text: " and " }, { text: "red", font: { color: "#ff0000" } }],
    });
  });

  it("nest (the inner span overriding per property), cross paragraphs and keep text:s/tab/line-break", async () => {
    const { sheet } = await read(
      table(
        "",
        row(
          `<table:table-cell office:value-type="string"><text:p>a<text:span text:style-name="T1">b<text:span text:style-name="T2">c<text:s text:c="2"/></text:span></text:span></text:p>` +
            `<text:p><text:span text:style-name="T2">d<text:tab/>e<text:line-break/>f</text:span></text:p></table:table-cell>`,
        ),
      ),
    );
    expect(getCell(sheet, 0, 0)).toEqual({
      col: 0,
      value: "abc  \nd\te\nf",
      runs: [
        { text: "a" },
        { text: "b", font: { bold: true } },
        { text: "c  ", font: { bold: true, color: "#ff0000" } },
        { text: "\n" },
        { text: "d\te\nf", font: { color: "#ff0000" } },
      ],
    });
  });

  it("are left out when no span has a font, or the runs do not spell the value", async () => {
    const { sheet } = await read(
      table(
        "",
        row(
          string(`plain <text:span>span</text:span>`) +
            `<table:table-cell office:value-type="string" office:string-value="Declared"><text:p><text:span text:style-name="T1">Shown</text:span></text:p></table:table-cell>` +
            float(5).replace(`<text:p>5</text:p>`, `<text:p><text:span text:style-name="T1">5</text:span></text:p>`),
        ),
      ),
    );
    expect(sheet.rows[0]!.cells).toEqual([
      { col: 0, value: "plain span" },
      { col: 1, value: "Declared" },
      { col: 2, value: 5 },
    ]);
  });

  it("copies runs into every repeated cell", async () => {
    const { sheet } = await read(table("", row(string(`<text:span text:style-name="T1">x</text:span>`, `table:number-columns-repeated="2"`))));
    const [a, b] = sheet.rows[0]!.cells;
    expect(b).toEqual({ ...a, col: 1 });
    expect(b!.runs).not.toBe(a!.runs);
    expect(b!.runs![0]!.font).not.toBe(a!.runs![0]!.font);
  });
});

describe("styled empty cells", () => {
  it("a run whose style equals the effective column or row default stores nothing", async () => {
    const { sheet } = await read(
      table(
        column(`table:default-cell-style-name="ce1" table:number-columns-repeated="3"`) + column(`table:number-columns-repeated="16381"`),
        row(empty(3, "ce1") + float(1)) + row(empty(2, "ce2") + float(2), `table:default-cell-style-name="ce2"`),
      ),
    );
    expect(sheet.rows.map((r) => r.cells)).toEqual([[{ col: 3, value: 1 }], [{ col: 2, value: 2 }]]);
  });

  it("a run reaching XFD becomes the row's style and stores no cell — the 16380-column run", async () => {
    const { sheet, id } = await read(
      table(
        "",
        row(string("A", `table:style-name="ce1"`) + string("B", `table:style-name="ce1"`) + float(3, `table:style-name="ce1"`) + float(4, `table:style-name="ce1"`) + empty(16380, "ce1")) +
          row(empty(16384, "ce2"), `table:number-rows-repeated="2"`),
      ),
    );
    expect(sheet.rows[0]).toEqual({
      index: 0,
      styleId: id("ce1"),
      cells: [
        { col: 0, value: "A", styleId: id("ce1") },
        { col: 1, value: "B", styleId: id("ce1") },
        { col: 2, value: 3, styleId: id("ce1") },
        { col: 3, value: 4, styleId: id("ce1") },
      ],
    });
    expect(sheet.rows.slice(1)).toEqual([
      { index: 1, cells: [], styleId: id("ce2") },
      { index: 2, cells: [], styleId: id("ce2") },
    ]);
    expect(styleAt(sheet, "XFD1")).toBe(id("ce1"));
    expect(styleAt(sheet, "Q3")).toBe(id("ce2"));
  });

  it("the cells left of a row-style run keep the column default they showed", async () => {
    const { sheet, id } = await read(
      table(
        column(`table:default-cell-style-name="ce3"`) + column(`table:number-columns-repeated="16383"`),
        row(float(1) + float(2) + empty(1) + empty(16381, "ce2")),
      ),
    );
    expect(sheet.rows[0]).toEqual({
      index: 0,
      styleId: id("ce2"),
      cells: [
        { col: 0, value: 1, styleId: id("ce3") },
        { col: 1, value: 2, styleId: 0 },
        { col: 2, value: null, styleId: 0 },
      ],
    });
    expect(["A1", "B1", "C1", "D1", "XFD1"].map((a1) => styleAt(sheet, a1))).toEqual([id("ce3"), 0, 0, id("ce2"), id("ce2")]);
  });

  it("stores a short run to XFD as cells when that is cheaper than keeping what is left of it", async () => {
    const { sheet, id } = await read(table("", row(empty(16380) + empty(4, "ce2"))));
    expect(sheet.rows[0]!.styleId).toBeUndefined();
    expect(sheet.rows[0]!.cells).toEqual([16380, 16381, 16382, 16383].map((col) => ({ col, value: null, styleId: id("ce2") })));
  });

  it("stores every other styled empty position, where it differs from the default", async () => {
    const { sheet, id } = await read(
      table(
        column(`table:default-cell-style-name="ce1" table:number-columns-repeated="2"`) + column(`table:number-columns-repeated="16382"`),
        row(empty(4, "ce1") + float(9)) + row(empty(2, "ce2") + float(9), `table:default-cell-style-name="ce1"`),
      ),
    );
    expect(sheet.rows[0]!.cells).toEqual([
      { col: 2, value: null, styleId: id("ce1") },
      { col: 3, value: null, styleId: id("ce1") },
      { col: 4, value: 9 },
    ]);
    expect(sheet.rows[1]!.cells).toEqual([
      { col: 0, value: null, styleId: id("ce2") },
      { col: 1, value: null, styleId: id("ce2") },
      { col: 2, value: 9 },
    ]);
  });

  it("counts styled empty cells against the cell ceiling exactly as values", async () => {
    // 16,000 styled cells × 126 rows = 2,016,000 > 2,000,000: refused before it is built.
    const started = performance.now();
    await expect(parseOds(new Uint8Array(0), workbook(table("", row(empty(16000, "ce2") + empty(384), `table:number-rows-repeated="126"`))))).rejects.toThrow(
      WORKBOOK_TOO_LARGE_MESSAGE,
    );
    expect(performance.now() - started).toBeLessThan(2000);
    const { sheet } = await read(table("", row(empty(16000, "ce2") + empty(384), `table:number-rows-repeated="2"`)));
    expect(sheet.rows.map((r) => r.cells.length)).toEqual([16000, 16000]);
  });
});

/** Every path in `value` holding `undefined` (the model never stores one). */
function undefinedPaths(value: unknown, path: string): string[] {
  if (value === undefined) {
    return [path];
  }
  if (typeof value !== "object" || value === null) {
    return [];
  }
  return Object.entries(value).flatMap(([key, item]: [string, unknown]) => undefinedPaths(item, `${path}.${key}`));
}

/** The fixture's "Styled" table, and a style name's index. */
async function styled() {
  const parsed = await parseOds(await buildSpec117StylesOdsBytes());
  const sheet = parsed.sheets[0]!;
  const id = (name: string): number => parsed.styles.findIndex((style) => style.sourceStyleName === name);
  return { parsed, sheet, id };
}

describe("spec117-styles.ods (LibreOffice-shaped)", () => {
  const facts = SPEC117_STYLES_ODS.styled;

  it("reads B3:B12 and nothing in column B below", async () => {
    const { sheet } = await styled();
    const values = facts.numbers.values.map((_, i) => getCell(sheet, 2 + i, 1)?.value);
    expect(values).toEqual(facts.numbers.values);
    expect(sheet.rows.some((r) => r.index >= 12 && r.cells.some((c) => c.col === 1))).toBe(false);
    expect(getCell(sheet, 1, 1)?.value).toBe(facts.header.value);
  });

  it("reads the columns: A's default style, the wide B, the hidden G; co1 the default width", async () => {
    const { sheet, id } = await styled();
    expect(sheet.defaultColWidthPx).toBe(facts.defaultColWidthPx);
    expect(sheet.columns).toEqual([
      { min: 0, max: 0, styleId: id(facts.styledColumn.styleName), sourceStyleName: "co1" },
      { min: 1, max: 1, widthPx: facts.wideColumn.widthPx, sourceStyleName: facts.wideColumn.styleName },
      { min: facts.hiddenColumn.index, max: facts.hiddenColumn.index, hidden: true, sourceStyleName: "co1" },
    ]);
  });

  it("reads the rows: sized ones, the hidden one, the trailing run as nothing", async () => {
    const { sheet } = await styled();
    expect(sheet.defaultRowHeightPt).toBe(facts.defaultRowHeightPt);
    for (const sized of facts.sizedRows) {
      const stored = sheet.rows.find((r) => r.index === sized.index)!;
      expect({ heightPt: stored.heightPt, customHeight: stored.customHeight }, `row ${sized.index}`).toEqual({
        heightPt: sized.heightPt,
        customHeight: "customHeight" in sized ? true : undefined,
      });
    }
    expect(sheet.rows.filter((r) => r.hidden).map((r) => r.index)).toEqual([facts.hiddenRow.index]);
    expect(getCell(sheet, facts.hiddenRow.index, 3)?.value).toBe(facts.hiddenRow.value);
    expect(sheet.rows.find((r) => r.index === facts.emptySizedRow.index)?.cells).toEqual([]);
    expect(sheet.rows.at(-1)!.index).toBe(facts.emptySizedRow.index);
  });

  it("the banner row's 16380-column styled run is the row's style and stores no cell", async () => {
    const { sheet, id } = await styled();
    const banner = sheet.rows[0]!;
    expect(banner.styleId).toBe(id(facts.bannerRow.styleName));
    expect(banner.cells.map((c) => c.value)).toEqual(facts.bannerRow.values);
    expect(banner.cells.every((c) => c.styleId === id(facts.bannerRow.styleName))).toBe(true);
    expect(styleAt(sheet, "XFD1")).toBe(id(facts.bannerRow.styleName));
  });

  it("reads the merge, its styled covered cells and the rich text", async () => {
    const { sheet, id } = await styled();
    expect(sheet.merges).toEqual([parseRange(facts.merge.range)]);
    for (const a1 of facts.merge.covered) {
      const ref = parseA1(a1)!;
      expect(getCell(sheet, ref.row, ref.col), a1).toEqual({ col: ref.col, value: null, styleId: id(facts.merge.styleName) });
    }
    const rich = parseA1(facts.richText.a1)!;
    expect(getCell(sheet, rich.row, rich.col)).toEqual({ col: rich.col, value: facts.richText.value, runs: facts.richText.runs });
  });

  it("is plain JSON with no undefined-valued key, and keeps the sparse invariants", async () => {
    const { parsed } = await styled();
    expect(undefinedPaths({ styles: parsed.styles, sheets: parsed.sheets }, "parsed")).toEqual([]);
    for (const sheet of parsed.sheets) {
      sheet.rows.forEach((stored, i) => {
        expect(i === 0 || stored.index > sheet.rows[i - 1]!.index, `${sheet.name} row ${stored.index}`).toBe(true);
        expect(stored.cells.length > 0 || stored.heightPt !== undefined || stored.hidden || stored.styleId !== undefined).toBe(true);
        stored.cells.forEach((c, k) => {
          expect(k === 0 || c.col > stored.cells[k - 1]!.col).toBe(true);
          expect(c.value !== null || c.styleId !== undefined).toBe(true);
        });
      });
    }
  });

  it("reads the tab colour and keeps every table's splice exact", async () => {
    const { parsed } = await styled();
    expect(parsed.sheets.map((s) => s.tabColor)).toEqual([facts.tabColor, undefined, undefined]);
    expect(parsed.sheets.map((s) => s.name)).toEqual(SPEC117_STYLES_ODS.tableNames);
    const bytes = await buildSpec117StylesOdsBytes();
    const { unzip } = await import("../../core/zip");
    const content = new TextDecoder().decode((await unzip(bytes)).get("content.xml")!);
    const envelope = parsed.odsContentEnvelope!;
    expect(envelope.prefix + parsed.sheets.map((s) => s.sourceXml).join("") + envelope.suffix).toBe(content);
  });
});
