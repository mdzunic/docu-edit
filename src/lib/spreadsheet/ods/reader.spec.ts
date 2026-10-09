// SPEC-116 "Reading .ods": every rule, over the LibreOffice-shaped fixtures
// and synthetic content — tables in order, repeated runs (empty ones never
// materialised, non-empty ones expanded), covered cells, every value type,
// text:p/text:s/text:tab/text:line-break, calcext errors, hidden tables, the
// null date, the active table, properties, splice equality, out-of-grid
// addresses and the cell ceiling.
import { describe, expect, it } from "vitest";
import { getCell, parseA1, usedRange } from "../address";
import { SpreadsheetOpenError, WORKBOOK_TOO_LARGE_MESSAGE } from "../errors";
import { buildMiniBookOdsBytes, buildSpecWideOdsBytes, FIXTURE_WORKBOOK_AUTHOR, MINI_BOOK_ODS, SPEC116_WIDE } from "../fixtures/build";
import type { Sheet } from "../model";
import { dateToSerial } from "../serial";
import { unzip } from "../../core/zip";
import { parseOds, readActiveTableName } from "./reader";

const encoder = new TextEncoder();

const NAMESPACES =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" ` +
  `xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"`;

function at(sheet: Sheet, a1: string): unknown {
  const ref = parseA1(a1)!;
  return getCell(sheet, ref.row, ref.col)?.value;
}

/** A synthetic .ods as the parts map the reader accepts. */
function synthetic(options: {
  body: string;
  declarations?: string;
  styles?: string;
  activeTable?: string;
  mimetype?: string;
}): Map<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>();
  parts.set("mimetype", encoder.encode(options.mimetype ?? "application/vnd.oasis.opendocument.spreadsheet"));
  parts.set(
    "content.xml",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES}>` +
        `<office:automatic-styles>${options.styles ?? ""}</office:automatic-styles>` +
        `<office:body><office:spreadsheet>${options.declarations ?? ""}${options.body}</office:spreadsheet></office:body></office:document-content>`,
    ),
  );
  if (options.activeTable !== undefined) {
    parts.set(
      "settings.xml",
      encoder.encode(
        `<office:document-settings ${NAMESPACES}><office:settings><config:config-item-set config:name="ooo:view-settings">` +
          `<config:config-item config:name="ActiveTable" config:type="string">${options.activeTable}</config:config-item>` +
          `</config:config-item-set></office:settings></office:document-settings>`,
      ),
    );
  }
  return parts;
}

const table = (name: string, rows: string, style?: string): string =>
  `<table:table table:name="${name}"${style ? ` table:style-name="${style}"` : ""}>${rows}</table:table>`;
const row = (cells: string, repeat?: number): string =>
  `<table:table-row${repeat ? ` table:number-rows-repeated="${repeat}"` : ""}>${cells}</table:table-row>`;
const float = (value: number, repeat?: number): string =>
  `<table:table-cell office:value-type="float" office:value="${value}"${repeat ? ` table:number-columns-repeated="${repeat}"` : ""}><text:p>${value}</text:p></table:table-cell>`;
const empty = (repeat: number): string => `<table:table-cell table:number-columns-repeated="${repeat}"/>`;

async function oneTable(rows: string, declarations?: string): Promise<Sheet> {
  return (await parseOds(new Uint8Array(0), synthetic({ body: table("T", rows), declarations }))).sheets[0]!;
}

describe("mini-book.ods (LibreOffice-shaped)", () => {
  it("reads the tables in order: names, the hidden one, the active one", async () => {
    const parsed = await parseOds(await buildMiniBookOdsBytes());
    expect(parsed.sheets.map((sheet) => sheet.name)).toEqual(MINI_BOOK_ODS.tableNames);
    expect(parsed.sheets.map((sheet) => sheet.visibility)).toEqual(["hidden", "visible", "visible"]);
    expect(parsed.sheets.every((sheet) => sheet.kind === "grid")).toBe(true);
    expect(parsed.activeSheetIndex).toBe(1);
    expect(parsed.dateSystem).toBe("1900");
    expect(parsed.odsNullDate).toBeUndefined();
    expect(parsed.packageVariant).toBeUndefined();
  });

  it("reads every Inventory value — float, percentage, currency, date, time, boolean, string, error, formulas", async () => {
    const inventory = (await parseOds(await buildMiniBookOdsBytes())).sheets[1]!;
    for (const [a1, expected] of Object.entries(MINI_BOOK_ODS.inventory)) {
      expect(at(inventory, a1), a1).toEqual(expected);
    }
    expect(at(inventory, "B8")).toBe(dateToSerial(2026, 10, 8, "1900"));
  });

  it("never materialises the trailing empty runs; expands the repeated non-empty row", async () => {
    const inventory = (await parseOds(await buildMiniBookOdsBytes())).sheets[1]!;
    // 23 used rows, five of them an empty repeated run in the middle.
    expect(inventory.rows).toHaveLength(MINI_BOOK_ODS.inventoryRows - 5);
    expect(usedRange(inventory)).toEqual({ top: 0, left: 0, bottom: 22, right: 3 });
    expect(inventory.rows.slice(3, 6).map((r) => [r.index, r.cells.map((cell) => cell.value)])).toEqual([
      [3, ["Spare", 1, 0.5, 0.5]],
      [4, ["Spare", 1, 0.5, 0.5]],
      [5, ["Spare", 1, 0.5, 0.5]],
    ]);
    // The copies are distinct objects (copy-on-write needs that).
    expect(inventory.rows[3]!.cells[0]).not.toBe(inventory.rows[4]!.cells[0]);
  });

  it("lets a covered cell occupy its column and store nothing", async () => {
    const inventory = (await parseOds(await buildMiniBookOdsBytes())).sheets[1]!;
    expect(at(inventory, "A11")).toBe("Merged note");
    expect(getCell(inventory, 10, 1)).toBeUndefined();
    expect(at(inventory, "C11")).toBe(7);
  });

  it("reads meta.xml into properties", async () => {
    const parsed = await parseOds(await buildMiniBookOdsBytes());
    expect(parsed.properties.title).toBe("Mini book");
    expect(parsed.properties.author).toBe(FIXTURE_WORKBOOK_AUTHOR);
  });

  it("splices back to content.xml exactly", async () => {
    const bytes = await buildMiniBookOdsBytes();
    const parsed = await parseOds(bytes);
    const content = new TextDecoder().decode((await unzip(bytes)).get("content.xml")!);
    const envelope = parsed.odsContentEnvelope!;
    expect(envelope.prefix + parsed.sheets.map((sheet) => sheet.sourceXml).join("") + envelope.suffix).toBe(content);
    expect(parsed.sheets[0]!.sourceXml!.startsWith(`<table:table table:name="Archive"`)).toBe(true);
    expect(parsed.sheets[2]!.sourceXml!.endsWith("</table:table>")).toBe(true);
    expect(envelope.suffix.startsWith("<table:named-expressions/>")).toBe(true);
  });
});

describe("spec116-wide.ods", () => {
  it("reaches column 16,000 of row 900,000 through empty runs, storing two rows", async () => {
    const wide = (await parseOds(await buildSpecWideOdsBytes())).sheets[0]!;
    expect(wide.rows.map((r) => r.index)).toEqual([0, SPEC116_WIDE.row]);
    expect(getCell(wide, SPEC116_WIDE.row, SPEC116_WIDE.col)?.value).toBe(SPEC116_WIDE.value);
    expect(usedRange(wide)).toEqual({ top: 0, left: 0, bottom: SPEC116_WIDE.row, right: SPEC116_WIDE.col });
  });
});

describe("rows, cells and values", () => {
  it("counts rows inside header rows, row groups and table:table-rows in document order", async () => {
    const sheet = await oneTable(
      `<table:table-header-rows>${row(float(1))}</table:table-header-rows>` +
        `<table:table-row-group>${row(float(2))}<table:table-rows>${row(float(3))}</table:table-rows></table:table-row-group>${row(float(4))}`,
    );
    expect(sheet.rows.map((r) => [r.index, r.cells[0]!.value])).toEqual([
      [0, 1],
      [1, 2],
      [2, 3],
      [3, 4],
    ]);
  });

  it("expands a repeated non-empty cell and skips repeated empty ones", async () => {
    const sheet = await oneTable(row(empty(2) + float(5, 3) + empty(16379)));
    expect(sheet.rows[0]!.cells).toEqual([
      { col: 2, value: 5 },
      { col: 3, value: 5 },
      { col: 4, value: 5 },
    ]);
  });

  it("reads string values from office:string-value, else the text:p text", async () => {
    const sheet = await oneTable(
      row(
        `<table:table-cell office:value-type="string" office:string-value="declared"><text:p>shown</text:p></table:table-cell>` +
          `<table:table-cell office:value-type="string"><text:p>a<text:span>b</text:span><text:s/>c<text:s text:c="3"/>d</text:p><text:p/></table:table-cell>` +
          `<table:table-cell><text:p>no type</text:p></table:table-cell>`,
      ),
    );
    expect(sheet.rows[0]!.cells.map((cell) => cell.value)).toEqual(["declared", "ab c   d\n", "no type"]);
  });

  it("leaves an annotation's text out of the cell", async () => {
    const sheet = await oneTable(
      row(
        `<table:table-cell office:value-type="string"><office:annotation><text:p>a comment</text:p></office:annotation>` +
          `<text:p>value</text:p></table:table-cell>`,
      ),
    );
    expect(at(sheet, "A1")).toBe("value");
  });

  it("reads an error from calcext:value-type, unknown codes as #VALUE!", async () => {
    const sheet = await oneTable(
      row(
        `<table:table-cell table:formula="of:=NA()" office:value-type="string" office:string-value="" calcext:value-type="error"><text:p>#N/A</text:p></table:table-cell>` +
          `<table:table-cell office:value-type="float" office:value="0" calcext:value-type="error"><text:p>Err:502</text:p></table:table-cell>`,
      ),
    );
    expect(sheet.rows[0]!.cells.map((cell) => cell.value)).toEqual([{ error: "#N/A" }, { error: "#VALUE!" }]);
  });

  it("converts dates and durations, PT36H05M00S and fractional seconds included", async () => {
    const sheet = await oneTable(
      row(
        `<table:table-cell office:value-type="date" office:date-value="2026-10-08T13:45:00.5"/>` +
          `<table:table-cell office:value-type="time" office:time-value="PT36H05M00S"/>` +
          `<table:table-cell office:value-type="time" office:time-value="PT00H00M01.25S"/>`,
      ),
    );
    expect(sheet.rows[0]!.cells.map((cell) => cell.value)).toEqual([
      dateToSerial(2026, 10, 8, "1900") + (13 * 3600 + 45 * 60 + 0.5) / 86400,
      (36 * 60 + 5) / 1440,
      1.25 / 86400,
    ]);
  });
});

const declaring = (date: string): string =>
  `<table:calculation-settings><table:null-date table:date-value="${date}"/></table:calculation-settings>`;

describe("the null date", () => {
  const dateRow = row(`<table:table-cell office:value-type="date" office:date-value="1904-01-02"/>`);

  it("1899-12-30 (and absent) is the 1900 system; 1904-01-01 the 1904 system", async () => {
    for (const declarations of [undefined, declaring("1899-12-30")]) {
      const parsed = await parseOds(new Uint8Array(0), synthetic({ body: table("T", dateRow), declarations }));
      expect(parsed.dateSystem).toBe("1900");
      expect(parsed.odsNullDate).toBeUndefined();
      expect(at(parsed.sheets[0]!, "A1")).toBe(dateToSerial(1904, 1, 2, "1900"));
    }
    const parsed = await parseOds(new Uint8Array(0), synthetic({ body: table("T", dateRow), declarations: declaring("1904-01-01") }));
    expect(parsed.dateSystem).toBe("1904");
    expect(at(parsed.sheets[0]!, "A1")).toBe(1);
  });

  it("records any other null date and counts dates from it", async () => {
    const parsed = await parseOds(new Uint8Array(0), synthetic({ body: table("T", dateRow), declarations: declaring("1900-01-01") }));
    expect(parsed.dateSystem).toBe("1900");
    expect(parsed.odsNullDate).toBe("1900-01-01");
    // 1900–1903 are 4 × 365 days (1900 is no leap year), plus one.
    expect(at(parsed.sheets[0]!, "A1")).toBe(4 * 365 + 1);
  });
});

describe("tables, visibility, the active table and the envelope", () => {
  const hiddenStyle = `<style:style style:name="ta9" style:family="table"><style:table-properties table:display="false"/></style:style>`;

  it("hides a table whose automatic style says table:display=\"false\"", async () => {
    const parsed = await parseOds(
      new Uint8Array(0),
      synthetic({ styles: hiddenStyle, body: table("A", row(float(1))) + table("B", row(float(2)), "ta9") }),
    );
    expect(parsed.sheets.map((sheet) => sheet.visibility)).toEqual(["visible", "hidden"]);
  });

  it("reads ActiveTable by name; absent → 0; a hidden table → the first visible one", async () => {
    const body = table("A", row(float(1)), "ta9") + table("B", row(float(2))) + table("C", row(float(3)));
    expect((await parseOds(new Uint8Array(0), synthetic({ styles: hiddenStyle, body, activeTable: "C" }))).activeSheetIndex).toBe(2);
    expect((await parseOds(new Uint8Array(0), synthetic({ styles: hiddenStyle, body, activeTable: "A" }))).activeSheetIndex).toBe(1);
    expect((await parseOds(new Uint8Array(0), synthetic({ body: table("A", "") + table("B", "") }))).activeSheetIndex).toBe(0);
    expect(readActiveTableName(`<office:document-settings ${NAMESPACES}/>`)).toBeUndefined();
  });

  it("gives each sheet the whitespace before its table, so the splice is exact on indented content", async () => {
    const body = `\n   ${table("A", row(float(1)))}\n   ${table("B", "")}\n   <table:named-expressions/>\n  `;
    const parts = synthetic({ body, declarations: `\n   <table:calculation-settings/>` });
    const parsed = await parseOds(new Uint8Array(0), parts);
    const content = new TextDecoder().decode(parts.get("content.xml"));
    const envelope = parsed.odsContentEnvelope!;
    expect(parsed.sheets[0]!.sourceXml!.startsWith("\n   <table:table")).toBe(true);
    expect(parsed.sheets[1]!.sourceXml).toBe(`\n   <table:table table:name="B"></table:table>`);
    expect(envelope.prefix.endsWith("<table:calculation-settings/>")).toBe(true);
    expect(envelope.prefix + parsed.sheets.map((sheet) => sheet.sourceXml).join("") + envelope.suffix).toBe(content);
  });

  it("records an .ots as the template variant", async () => {
    const parsed = await parseOds(
      new Uint8Array(0),
      synthetic({ body: table("T", ""), mimetype: "application/vnd.oasis.opendocument.spreadsheet-template" }),
    );
    expect(parsed.packageVariant).toBe("template");
  });
});

/** The error parsing `parts` is refused with (failing if it is not). */
async function failure(parts: Map<string, Uint8Array>): Promise<Error> {
  try {
    await parseOds(new Uint8Array(0), parts);
  } catch (error) {
    return error as Error;
  }
  throw new Error("expected the workbook to be refused");
}

describe("refusals", () => {
  it("names a non-empty cell past column XFD", async () => {
    const error = await failure(synthetic({ body: table("Wide", row(empty(16384) + float(1))) }));
    expect(error).toBeInstanceOf(SpreadsheetOpenError);
    expect(error.message).toContain("XFE1");
  });

  it("names a non-empty row past row 1,048,576", async () => {
    const error = await failure(synthetic({ body: table("Tall", row(empty(1), 1048576) + row(float(1))) }));
    expect(error.message).toContain("A1048577");
  });

  it("tolerates empty runs that reach the grid's edge", async () => {
    const sheet = await oneTable(row(float(1) + empty(16383)) + row(empty(16384), 1048575));
    expect(sheet.rows).toHaveLength(1);
  });

  it("refuses a workbook over the ceiling before materialising a repeated run", async () => {
    // One non-empty cell × 16,384 columns × 123 rows = 2,015,232 cells from a
    // part of a few hundred bytes: refused arithmetically, never built.
    const started = performance.now();
    const error = await failure(synthetic({ body: table("Huge", row(float(1, 16384), 123)) }));
    expect(error.message).toBe(WORKBOOK_TOO_LARGE_MESSAGE);
    expect(performance.now() - started).toBeLessThan(2000);
  });

  it("refuses content.xml containing a DTD", async () => {
    const parts = synthetic({ body: table("T", "") });
    parts.set("content.xml", encoder.encode(`<!DOCTYPE office:document-content [<!ENTITY x "y">]><office:document-content ${NAMESPACES}/>`));
    expect((await failure(parts)).message).toMatch(/DTD/);
  });

  it("refuses a spreadsheet with no tables", async () => {
    expect(await failure(synthetic({ body: "" }))).toBeInstanceOf(SpreadsheetOpenError);
  });
});
