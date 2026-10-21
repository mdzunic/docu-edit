// SPEC-130's xlsx data-validation fixture (spec130-validation.xlsx, built by
// fixtures/build.ts): the committed e2e file is the builder's bytes, the
// package is whole, the producer shapes the reader and writer are proven on
// are really in the parts — `mc:Ignorable="x14ac xr xr2 xr3"`, an `xr:uid` on
// every base rule, `imeMode="noControl"`, `showDropDown="1"`, the cross-sheet
// list as an `x14:dataValidation` in the `{CCE6A557-…}` ext, the "Colours"
// definedName — and the facts the e2e flow leans on hold (B3 holds 5 and the
// rest of B2:B10 is empty, D2 has no rule, B2's list and input message, B8
// hides its arrow). The model the reader makes of the file is asserted once
// here too, every rule's fields, as the e2e opens it.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { getCell, parseA1, parseRange } from "./address";
import { buildSpec130ValidationXlsxBytes, SPEC130_VALIDATION_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { storedValueContext } from "./formula/evaluate";
import { DEFAULT_SPREADSHEET_LOCALE, SPREADSHEET_LOCALES } from "./numberFormat/locale";
import { workbookFormatContext } from "./numberFormat/render";
import { readSpreadsheetBytes } from "./open";
import { listItems, validationAt, type DataValidation } from "./validation";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();
const F = SPEC130_VALIDATION_XLSX;

async function parts(): Promise<Map<string, Uint8Array>> {
  return unzip(await buildSpec130ValidationXlsxBytes());
}

const text = (all: Map<string, Uint8Array>, name: string): string => decoder.decode(all.get(name));

/** A rule without its minted id. */
function withoutId(rule: DataValidation): Omit<DataValidation, "id"> {
  const { id: _id, ...rest } = rule;
  return rest;
}

const noMessage = { show: true, title: "", text: "" };
const stop = { show: true, style: "stop", title: "", text: "" } as const;

describe("spec130-validation.xlsx", () => {
  it("is the builder's bytes, whole, with Excel's parts in order", async () => {
    const committed = new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec130-validation.xlsx")));
    expect(equalBytes(committed, await buildSpec130ValidationXlsxBytes())).toBe(true);
    expect(await packageIntegrity(committed)).toEqual([]);
    expect([...(await parts()).keys()]).toEqual([...F.parts]);
  });

  it("holds Excel's shapes: the ignorable namespaces, xr:uid on every base rule, imeMode, showDropDown, the x14 cross-sheet list, the defined name", async () => {
    const all = await parts();
    const form = text(all, F.form.partName);
    expect(form).toContain(`mc:Ignorable="x14ac xr xr2 xr3"`);
    // The block after sheetData, before pageMargins (CT_Worksheet order); the ext last.
    expect(form).toContain(`</sheetData>${F.form.block}<pageMargins`);
    expect(form).toContain(`<extLst>${F.form.ext}</extLst></worksheet>`);
    const rules = [...form.matchAll(/<dataValidation [^>]*>/g)].map((match) => match[0]);
    expect(rules).toHaveLength(Object.keys(F.form.rules).length);
    expect(rules.every((rule) => /xr:uid="\{[0-9A-F-]{36}\}"/.test(rule))).toBe(true);
    expect(F.form.rules.amount).toContain(`imeMode="noControl"`);
    expect(F.form.rules.option).toContain(`showDropDown="1"`);
    expect(F.form.ext).toContain(`<x14:formula1><xm:f>Lists!$A$1:$A$4</xm:f></x14:formula1><xm:sqref>B9</xm:sqref>`);
    expect(F.form.ext).toMatch(/^<ext uri="\{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF\}" xmlns:x14="[^"]+"><x14:dataValidations count="1" xmlns:xm="[^"]+">/);
    expect(text(all, "xl/workbook.xml")).toContain(`</sheets><definedNames><definedName name="Colours">Lists!$B$1:$B$3</definedName></definedNames><calcPr`);
    expect(text(all, "docProps/app.xml")).toContain(`<vt:lpstr>Named Ranges</vt:lpstr>`);
  });

  it("keeps the facts the e2e flow leans on", async () => {
    const model = await readSpreadsheetBytes(await buildSpec130ValidationXlsxBytes(), "xlsx", "book");
    const form = model.sheets[0]!;
    expect(model.activeSheetIndex).toBe(F.activeSheetIndex);
    expect(model.sheets.map((sheet) => sheet.name)).toEqual([...F.sheetNames]);
    for (let row = 1; row <= 9; row += 1) {
      expect(getCell(form, row, 1)?.value ?? null, `B${row + 1}`).toBe(row === 2 ? F.form.score.value : null);
    }
    expect(validationAt(form, parseA1(F.form.free)!)).toBeUndefined();
    expect(getCell(form, 1, 3)).toBeUndefined();
    const deps = { base: storedValueContext(model), ctx: workbookFormatContext(model, SPREADSHEET_LOCALES[DEFAULT_SPREADSHEET_LOCALE]) };
    const items = (a1: string): string[] => {
      const at = parseA1(a1)!;
      return listItems(model, form, validationAt(form, at)!, at, deps).items.map((item) => item.text);
    };
    expect(items("B2")).toEqual(["Low", "Medium", "High"]);
    expect(items("B8")).toEqual([...F.form.options]);
    expect(items("B9")).toEqual([...F.lists.regions]);
    expect(validationAt(form, parseA1("B2")!)!.input).toEqual({ show: true, title: "Priority", text: "Pick one" });
    expect(validationAt(form, parseA1("B8")!)!.showDropdown).toBe(false);
  });

  it("reads every rule: the base block's in order with their source text, then the x14 list", async () => {
    const model = await readSpreadsheetBytes(await buildSpec130ValidationXlsxBytes(), "xlsx", "book");
    const [form, lists] = model.sheets;
    expect(lists!.validations).toEqual([]);
    expect(form!.validationPromptsDisabled).toBeUndefined();
    const at = (a1: string) => [parseRange(a1)!];
    expect(form!.validations.map(withoutId)).toEqual([
      {
        ranges: at("B2"),
        type: "list",
        operator: "between",
        formula1: `"Low,Medium,High"`,
        allowBlank: true,
        showDropdown: true,
        input: { show: true, title: "Priority", text: "Pick one" },
        error: { show: true, style: "stop", title: "Invalid priority", text: "Choose Low, Medium or High." },
        sourceXml: F.form.rules.priority,
      },
      {
        ranges: at("B3"),
        type: "whole",
        operator: "between",
        formula1: "1",
        formula2: "10",
        allowBlank: true,
        showDropdown: true,
        input: noMessage,
        error: { show: true, style: "warning", title: "Out of range", text: "Scores run from 1 to 10." },
        sourceXml: F.form.rules.score,
      },
      {
        ranges: at("B4"),
        type: "decimal",
        operator: "greaterThan",
        formula1: "0",
        allowBlank: true,
        showDropdown: true,
        input: noMessage,
        error: { show: true, style: "information", title: "", text: "" },
        sourceXml: F.form.rules.amount,
      },
      {
        ranges: at("B5"),
        type: "date",
        operator: "between",
        formula1: String(F.form.dates.from),
        formula2: String(F.form.dates.to),
        allowBlank: true,
        showDropdown: true,
        input: noMessage,
        error: stop,
        sourceXml: F.form.rules.due,
      },
      {
        ranges: at("B6"),
        type: "textLength",
        operator: "lessThanOrEqual",
        formula1: "5",
        allowBlank: true,
        showDropdown: true,
        input: noMessage,
        error: stop,
        sourceXml: F.form.rules.code,
      },
      {
        ranges: at("B7"),
        type: "custom",
        operator: "between",
        formula1: `ISNUMBER(SEARCH("@",B7))`,
        allowBlank: true,
        showDropdown: true,
        input: noMessage,
        error: stop,
        sourceXml: F.form.rules.email,
      },
      {
        ranges: at("B8"),
        type: "list",
        operator: "between",
        formula1: "$E$2:$E$6",
        allowBlank: true,
        showDropdown: false,
        input: noMessage,
        error: stop,
        sourceXml: F.form.rules.option,
      },
      {
        ranges: at("B10"),
        type: "list",
        operator: "between",
        formula1: "Colours",
        allowBlank: true,
        showDropdown: true,
        input: noMessage,
        error: stop,
        sourceXml: F.form.rules.colour,
      },
      {
        ranges: at("B9"),
        type: "list",
        operator: "between",
        formula1: "Lists!$A$1:$A$4",
        allowBlank: true,
        showDropdown: true,
        input: noMessage,
        error: stop,
        x14: true,
        sourceXml: F.form.x14Rule,
      },
    ]);
    expect(new Set(form!.validations.map((rule) => rule.id)).size).toBe(form!.validations.length);
  });
});
