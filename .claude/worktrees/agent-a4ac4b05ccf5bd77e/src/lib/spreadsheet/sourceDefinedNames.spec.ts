// SPEC-130 "List sources": a list rule whose Source is a defined name lists
// what the source file defines the name as (sourceDefinedNames.ts,
// provisional until SPEC-133's names model). Both fixtures' B10 lists
// `Colours` (Lists!B1:B3) and enforces it; the name keeps meaning those cells
// after its sheet is renamed and after rows are inserted above them; a name
// scoped to a sheet answers there first and nowhere else; ODS named
// expressions and a table's own names resolve too; a name the source does not
// define is unavailable and accepts the entry.
import { describe, expect, it } from "vitest";
import { bytesToBase64, base64ToBytes } from "../core/packageParts";
import { parseA1 } from "./address";
import { buildSpec130ValidationOdsBytes, buildSpec130ValidationXlsxBytes, SPEC130_VALIDATION_ODS, SPEC130_VALIDATION_XLSX } from "./fixtures/build";
import { storedValueContext } from "./formula/evaluate";
import type { SpreadsheetModel } from "./model";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { workbookFormatContext } from "./numberFormat/render";
import { readSpreadsheetBytes } from "./open";
import { renameSheet } from "./sheets";
import { sourceDefinedName } from "./sourceDefinedNames";
import { applyStructuralEdit } from "./structure";
import { listItems, validateValue, validationAt, type ValidationDeps } from "./validation";

const COLOURS = [...SPEC130_VALIDATION_XLSX.lists.colours];
const REGIONS = [...SPEC130_VALIDATION_XLSX.lists.regions];

function deps(model: SpreadsheetModel): ValidationDeps {
  return { base: storedValueContext(model), ctx: workbookFormatContext(model, SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE]) };
}

/** The texts Form's rule at `a1` lists, or "unavailable". */
function items(model: SpreadsheetModel, a1: string): string[] | "unavailable" {
  const form = model.sheets[0]!;
  const ref = parseA1(a1)!;
  const list = listItems(model, form, validationAt(form, ref)!, ref, deps(model));
  return list.unavailable ? "unavailable" : list.items.map((item) => item.text);
}

function valid(model: SpreadsheetModel, a1: string, value: string): boolean {
  return validateValue(model, model.sheets[0]!, parseA1(a1)!, value, deps(model)).ok;
}

const open = {
  xlsx: async () => readSpreadsheetBytes(await buildSpec130ValidationXlsxBytes(), "xlsx", "book"),
  ods: async () => readSpreadsheetBytes(await buildSpec130ValidationOdsBytes(), "ods", "book"),
};

/** The model with its retained workbook part's text edited. */
function withWorkbookXml(model: SpreadsheetModel, edit: (xml: string) => string): SpreadsheetModel {
  const opaque = model.opaque.map((part) =>
    part.name === "xl/workbook.xml" ? { ...part, data: bytesToBase64(new TextEncoder().encode(edit(new TextDecoder().decode(base64ToBytes(part.data))))) } : part,
  );
  return { ...model, opaque };
}

describe.each(["xlsx", "ods"] as const)("SPEC-130: a defined-name list source (%s)", (format) => {
  it("B10 lists what Colours names, in source order, and enforces it", async () => {
    const model = await open[format]();
    expect(items(model, "B10")).toEqual(COLOURS);
    expect(valid(model, "B10", "green")).toBe(true);
    expect(valid(model, "B10", "Purple")).toBe(false);
    expect(sourceDefinedName(model, "colours", model.sheets[0]!.id)).toEqual({ formula: "Lists!$B$1:$B$3", origin: format === "xlsx" ? { row: 0, col: 0 } : { row: 0, col: 1 } });
  });

  it("follows a rename of its sheet (the source text keeps the source name)", async () => {
    const model = await open[format]();
    renameSheet(model, 1, "Palette");
    expect(model.sheets[1]!.name).toBe("Palette");
    expect(sourceDefinedName(model, "Colours", model.sheets[0]!.id)?.formula).toBe("Palette!$B$1:$B$3");
    expect(items(model, "B10")).toEqual(COLOURS);
  });

  it("follows rows inserted above its cells (SPEC-124's retained-reference pass)", async () => {
    const model = await open[format]();
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: model.sheets[1]!.id, at: 0, count: 2 })).toEqual({ ok: true });
    expect(sourceDefinedName(model, "Colours", model.sheets[0]!.id)?.formula).toBe("Lists!$B$3:$B$5");
    expect(items(model, "B10")).toEqual(COLOURS);
  });

  it("a name the source does not define is unavailable and accepts the entry", async () => {
    const model = await open[format]();
    const form = model.sheets[0]!;
    form.validations = form.validations.map((rule) => (rule.formula1 === "Colours" ? { ...rule, formula1: "Shades" } : rule));
    expect(sourceDefinedName(model, "Shades", form.id)).toBeUndefined();
    expect(items(model, "B10")).toBe("unavailable");
    expect(valid(model, "B10", "anything")).toBe(true);
  });
});

describe("SPEC-130: defined-name scopes", () => {
  it("xlsx: a name scoped to the host sheet wins over the workbook's; one scoped to another sheet answers only when qualified", async () => {
    const model = withWorkbookXml(await open.xlsx(), (xml) =>
      xml.replace(
        `</definedName></definedNames>`,
        `</definedName><definedName name="Colours" localSheetId="0">Lists!$A$1:$A$2</definedName>` +
          `<definedName name="Shades" localSheetId="1">Lists!$B$2:$B$3</definedName></definedNames>`,
      ),
    );
    const [form, lists] = model.sheets;
    expect(items(model, "B10")).toEqual(REGIONS.slice(0, 2));
    expect(sourceDefinedName(model, "Colours", lists!.id)?.formula).toBe("Lists!$B$1:$B$3");
    expect(sourceDefinedName(model, "Shades", form!.id)).toBeUndefined();
    expect(sourceDefinedName(model, "Shades", form!.id, "Lists")?.formula).toBe("Lists!$B$2:$B$3");
    form!.validations = form!.validations.map((rule) => (rule.formula1 === "Colours" ? { ...rule, formula1: "Lists!Shades" } : rule));
    expect(items(model, "B10")).toEqual(COLOURS.slice(1));
  });

  it("ods: a named expression resolves, and a table's own name wins over the spreadsheet's", async () => {
    const base = await open.ods();
    const envelope = base.odsContentEnvelope!;
    const expression = `<table:named-expression table:name="Fruits" table:base-cell-address="$Form.$A$1" table:expression="of:[$Lists.$A$1:.$A$2]"/>`;
    const model: SpreadsheetModel = {
      ...base,
      odsContentEnvelope: { ...envelope, suffix: envelope.suffix.replace(`</table:named-expressions>`, `${expression}</table:named-expressions>`) },
    };
    const form = model.sheets[0]!;
    expect(sourceDefinedName(model, "fruits", form.id)).toEqual({ formula: "Lists!$A$1:$A$2", origin: { row: 0, col: 0 } });
    const own = `<table:named-expressions><table:named-range table:name="Colours" table:base-cell-address="$Form.$E$2" table:cell-range-address="$Form.$E$2:.$E$3"/></table:named-expressions>`;
    model.sheets = [{ ...form, sourceXml: form.sourceXml!.replace(/<\/table:table>$/, `${own}</table:table>`) }, ...model.sheets.slice(1)];
    expect(items(model, "B10")).toEqual(SPEC130_VALIDATION_ODS.regions.slice(0, 2));
    expect(sourceDefinedName(model, "Colours", model.sheets[1]!.id)?.formula).toBe("Lists!$B$1:$B$3");
  });
});
