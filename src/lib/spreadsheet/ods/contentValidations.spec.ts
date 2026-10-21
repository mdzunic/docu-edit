// SPEC-130 "ODS wire mapping" (ods/contentValidations.ts): every
// `table:condition` form read and written (`of:` and `oooc:`), each
// attribute and its default, messages as paragraphs, base-cell-address
// rebasing (one definition shared by two tables), ranges from the cells'
// names as row-run × column-run rectangles with nothing materialised, the
// error macro, orphans; the table writer's cell attribute and its repeat
// compression; the prefix block's regeneration rules — names kept, fresh
// `de-val<n>` names, a shared definition written once, orphans re-emitted,
// a renamed sheet's clean table and a duplicate's — and the round trips.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { zip } from "../../core/zipWriter";
import { parseRange, type GridRange } from "../address";
import { setCellEntry } from "../cells";
import { buildSpec130ValidationOdsBytes, SPEC130_VALIDATION_ODS } from "../fixtures/build";
import type { SpreadsheetModel } from "../model";
import { SPREADSHEET_LOCALES } from "../numberFormat/locale";
import { readSpreadsheetBytes } from "../open";
import { CellBudget } from "../parsed";
import { createBlankWorkbook, duplicateSheet, renameSheet } from "../sheets";
import { clearValidation, defaultValidationSettings, setValidation, type DataValidation, type ValidationSettings } from "../validation";
import { odsValidationXml, rangesFromCellNames, readOdsContentValidations, readTableValidations, tableValidationRanges } from "./contentValidations";
import { readOdsContent } from "./reader";
import { OdsStyleRegistry } from "./styles";
import { serializeTable } from "./tableXml";
import { writeOds } from "./writer";

const DECL = `<?xml version="1.0" encoding="UTF-8"?>`;
const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:oooc="http://openoffice.org/2004/calc"`;
const range = (a1: string): GridRange => parseRange(a1)!;
const decoder = new TextDecoder();

/** content.xml: `before` (the validations, label ranges…) in the prefix,
 * then the tables. */
function contentXml(before: string, tables: string): string {
  return `${DECL}<office:document-content ${NS} office:version="1.3"><office:body><office:spreadsheet><table:calculation-settings/>${before}${tables}</office:spreadsheet></office:body></office:document-content>`;
}
const prefixOf = (before: string): string => `${DECL}<office:document-content ${NS} office:version="1.3"><office:body><office:spreadsheet><table:calculation-settings/>${before}`;
const block = (definitions: string): string => `<table:content-validations>${definitions}</table:content-validations>`;
const tableXml = (name: string, rows: string): string => `<table:table table:name="${name}"><table:table-column table:number-columns-repeated="1024"/>${rows}</table:table>`;
const row = (cells: string, repeat = 1): string => `<table:table-row${repeat > 1 ? ` table:number-rows-repeated="${repeat}"` : ""}>${cells}</table:table-row>`;
const named = (name: string, repeat = 1): string => `<table:table-cell table:content-validation-name="${name}"${repeat > 1 ? ` table:number-columns-repeated="${repeat}"` : ""}/>`;
const empty = (repeat: number): string => (repeat > 1 ? `<table:table-cell table:number-columns-repeated="${repeat}"/>` : `<table:table-cell/>`);
const tail = (rows: number): string => row(empty(1024), 1_048_576 - rows);

/** A rule without its id, source and name. */
function shape(rule: DataValidation): Omit<DataValidation, "id" | "sourceXml" | "odsName"> {
  const { id: _id, sourceXml: _source, odsName: _name, ...rest } = rule;
  return rest;
}

/** The one rule a definition makes over B2. */
function readOne(definition: string, ranges: GridRange[] = [range("B2")]): DataValidation {
  const rules = readTableValidations(prefixOf(block(definition)), "", "S", new Map([["v", ranges]]));
  expect(rules).toHaveLength(1);
  return rules[0]!;
}

const condition = (text: string, extra = ""): string => `<table:content-validation table:name="v" table:condition="${text}"${extra} table:base-cell-address="S.B2"/>`;

/** A minimal ODS package around a content.xml. */
function odsPackage(content: string): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const manifest =
    `${DECL}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "META-INF/manifest.xml", data: encoder.encode(manifest) },
  ]);
}

async function contentOf(model: SpreadsheetModel): Promise<{ bytes: Uint8Array; content: string }> {
  const bytes = await writeOds(model);
  return { bytes, content: decoder.decode((await unzip(bytes)).get("content.xml")) };
}

const settings = (overrides: Partial<ValidationSettings>): ValidationSettings => ({ ...defaultValidationSettings(), ...overrides });

// --- reading ----------------------------------------------------------------------------------

describe("reading table:condition", () => {
  const forms: [string, Partial<DataValidation>][] = [
    ["of:cell-content-is-whole-number() and cell-content-is-between(1,10)", { type: "whole", operator: "between", formula1: "1", formula2: "10" }],
    ["of:cell-content-is-decimal-number() and cell-content-is-not-between(1;[.$Z$1])", { type: "decimal", operator: "notBetween", formula1: "1", formula2: "$Z$1" }],
    ["of:cell-content-is-date() and cell-content()&gt;=45658", { type: "date", operator: "greaterThanOrEqual", formula1: "45658" }],
    ["of:cell-content-is-time() and cell-content()&lt;0.5", { type: "time", operator: "lessThan", formula1: "0.5" }],
    ["of:cell-content-is-whole-number() and cell-content()=3", { type: "whole", operator: "equal", formula1: "3" }],
    ["of:cell-content-is-whole-number() and cell-content()!=3", { type: "whole", operator: "notEqual", formula1: "3" }],
    ["of:cell-content-is-whole-number() and cell-content()&lt;&gt;3", { type: "whole", operator: "notEqual", formula1: "3" }],
    ["of:cell-content-is-whole-number() and cell-content()&gt;[.C2]", { type: "whole", operator: "greaterThan", formula1: "C2" }],
    ["of:cell-content-is-whole-number() and cell-content()&lt;=3", { type: "whole", operator: "lessThanOrEqual", formula1: "3" }],
    ["of:cell-content-is-whole-number()", { type: "whole", operator: "between" }],
    ["of:cell-content-text-length()&lt;=5", { type: "textLength", operator: "lessThanOrEqual", formula1: "5" }],
    ["of:cell-content-text-length-is-between(2,8)", { type: "textLength", operator: "between", formula1: "2", formula2: "8" }],
    ["of:cell-content-text-length-is-not-between(2;8)", { type: "textLength", operator: "notBetween", formula1: "2", formula2: "8" }],
    ["of:cell-content-is-in-list(&quot;Low&quot;;&quot;Med&quot;;&quot;Say &quot;&quot;hi&quot;&quot;&quot;)", { type: "list", formula1: `"Low,Med,Say ""hi"""` }],
    ["of:cell-content-is-in-list(1;2;3)", { type: "list", formula1: `"1,2,3"` }],
    ["of:cell-content-is-in-list([$Lists.$A$1:.$A$4])", { type: "list", formula1: "Lists!$A$1:$A$4" }],
    ["of:cell-content-is-in-list([.$E$2:.$E$6])", { type: "list", formula1: "$E$2:$E$6" }],
    ["of:cell-content-is-in-list(Colours)", { type: "list", formula1: "Colours" }],
    ["of:cell-content-is-in-list(OFFSET([.$E$1];0;0;3;1))", { type: "list", formula1: "OFFSET($E$1,0,0,3,1)" }],
    ["of:is-true-formula(ISNUMBER(SEARCH(&quot;@&quot;;[.B2])))", { type: "custom", formula1: `ISNUMBER(SEARCH("@",B2))` }],
    // LibreOffice's older dialect.
    ["oooc:cell-content-is-whole-number() and cell-content-is-between(1,10)", { type: "whole", operator: "between", formula1: "1", formula2: "10" }],
    ["oooc:is-true-formula([.B2]&gt;0)", { type: "custom", formula1: "B2>0" }],
    ["oooc:cell-content-is-in-list(&quot;a&quot;;&quot;b&quot;)", { type: "list", formula1: `"a,b"` }],
  ];

  it.each(forms)("%s", (wire, expected) => {
    const rule = readOne(condition(wire));
    expect({ type: rule.type, operator: rule.operator, formula1: rule.formula1, formula2: rule.formula2 }).toEqual({
      operator: "between",
      formula1: undefined,
      formula2: undefined,
      ...expected,
    });
  });

  it("no condition is `any`; a condition the model cannot read is `any` with its element kept", () => {
    expect(readOne(`<table:content-validation table:name="v" table:base-cell-address="S.B2"/>`)).toMatchObject({ type: "any", operator: "between" });
    for (const wire of ["of:cell-content-is-mystery()", "msoxl:cell-content-is-whole-number()", "of:cell-content-is-whole-number() and cell-content-is-between(1)", "of:is-true-formula([.B2]"]) {
      const definition = condition(wire);
      const rule = readOne(definition);
      expect(rule, wire).toMatchObject({ type: "any", sourceXml: definition, odsName: "v" });
      expect(rule.formula1, wire).toBeUndefined();
    }
  });

  it("each written form reads back as the rule it was written from", () => {
    for (const [wire] of forms) {
      const rule = readOne(condition(wire));
      const written = odsValidationXml("v", rule, "S");
      expect(shape(readOne(written)), wire).toEqual(shape(rule));
    }
  });

  it("writes Calc's forms: `of:`, between operands joined with `,`, list items with `;`", () => {
    const at = { ranges: [range("B2")], id: "x" };
    const xml = (rule: ValidationSettings): string => /table:condition="([^"]*)"/.exec(odsValidationXml("v", { ...rule, ...at }, "S"))?.[1] ?? "";
    expect(xml(settings({ type: "whole", operator: "between", formula1: "1", formula2: "10" }))).toBe("of:cell-content-is-whole-number() and cell-content-is-between(1,10)");
    expect(xml(settings({ type: "date", operator: "notBetween", formula1: "45658", formula2: "$Z$1" }))).toBe(
      "of:cell-content-is-date() and cell-content-is-not-between(45658,[.$Z$1])",
    );
    expect(xml(settings({ type: "decimal", operator: "greaterThan", formula1: "0" }))).toBe("of:cell-content-is-decimal-number() and cell-content()&gt;0");
    expect(xml(settings({ type: "time", operator: "notEqual", formula1: "0.5" }))).toBe("of:cell-content-is-time() and cell-content()!=0.5");
    expect(xml(settings({ type: "textLength", operator: "lessThanOrEqual", formula1: "5" }))).toBe("of:cell-content-text-length()&lt;=5");
    expect(xml(settings({ type: "textLength", operator: "between", formula1: "2", formula2: "8" }))).toBe("of:cell-content-text-length-is-between(2,8)");
    expect(xml(settings({ type: "list", formula1: `"Low,Medium,High"` }))).toBe("of:cell-content-is-in-list(&quot;Low&quot;;&quot;Medium&quot;;&quot;High&quot;)");
    expect(xml(settings({ type: "list", formula1: "Lists!$A$1:$A$4" }))).toBe("of:cell-content-is-in-list([$Lists.$A$1:.$A$4])");
    expect(xml(settings({ type: "custom", formula1: `ISNUMBER(SEARCH("@",B2))` }))).toBe("of:is-true-formula(ISNUMBER(SEARCH(&quot;@&quot;;[.B2])))");
    expect(odsValidationXml("v", { ...settings({}), ...at }, "S")).toBe(
      `<table:content-validation table:name="v" table:base-cell-address="S.B2"><table:help-message table:display="true"></table:help-message>` +
        `<table:error-message table:message-type="stop" table:display="true"></table:error-message></table:content-validation>`,
    );
  });

  it("an ODS literal item holding a comma becomes two items once the rule is serialised", () => {
    const rule = readOne(condition("of:cell-content-is-in-list(&quot;a,b&quot;;&quot;c&quot;)"));
    expect(rule.formula1).toBe(`"a,b,c"`);
    expect(odsValidationXml("v", rule, "S")).toContain("of:cell-content-is-in-list(&quot;a&quot;;&quot;b&quot;;&quot;c&quot;)");
  });
});

describe("reading the attributes and messages", () => {
  it("defaults: allow-empty-cell true, display-list unsorted, no messages", () => {
    expect(readOne(condition("of:cell-content-is-in-list(&quot;a&quot;)"))).toMatchObject({
      allowBlank: true,
      showDropdown: true,
      input: { show: false, title: "", text: "" },
      error: { show: false, style: "stop", title: "", text: "" },
    });
    expect(readOne(condition("of:cell-content-is-in-list(&quot;a&quot;)")).odsListOrder).toBeUndefined();
  });

  it("allow-empty-cell false; display-list none, unsorted and sort-ascending", () => {
    const list = "of:cell-content-is-in-list(&quot;a&quot;)";
    expect(readOne(condition(list, ` table:allow-empty-cell="false"`)).allowBlank).toBe(false);
    expect(readOne(condition(list, ` table:display-list="none"`))).toMatchObject({ showDropdown: false });
    expect(readOne(condition(list, ` table:display-list="unsorted"`)).odsListOrder).toBeUndefined();
    expect(readOne(condition(list, ` table:display-list="sort-ascending"`))).toMatchObject({ showDropdown: true, odsListOrder: "sort-ascending" });
  });

  it("help and error messages: title, display (false), message-type (stop), paragraphs joined with \\n", () => {
    const definition =
      `<table:content-validation table:name="v" table:base-cell-address="S.B2">` +
      `<table:help-message table:title="Notes &amp; tips" table:display="true"><text:p>One<text:s text:c="2"/> two</text:p><text:p>a<text:tab/>b<text:line-break/>c</text:p><text:p/></table:help-message>` +
      `<table:error-message table:title="Careful"><text:p>Not that</text:p></table:error-message></table:content-validation>`;
    expect(readOne(definition)).toMatchObject({
      input: { show: true, title: "Notes & tips", text: "One   two\na\tb\nc\n" },
      error: { show: false, style: "stop", title: "Careful", text: "Not that" },
    });
    for (const style of ["warning", "information", "stop"] as const) {
      const error = `<table:content-validation table:name="v" table:base-cell-address="S.B2"><table:error-message table:message-type="${style}" table:display="true"/></table:content-validation>`;
      expect(readOne(error).error).toEqual({ show: true, style, title: "", text: "" });
    }
  });

  it("writes messages as one text:p per line (spaces kept) and reads them back", () => {
    const rule: DataValidation = {
      ...settings({ input: { show: true, title: "Say \"hi\"", text: "Line one\n  indented\n\nlast" }, error: { show: false, style: "warning", title: "", text: "a < b" } }),
      id: "x",
      ranges: [range("B2")],
    };
    const xml = odsValidationXml("v", rule, "S");
    expect(xml).toContain(
      `<table:help-message table:title="Say &quot;hi&quot;" table:display="true"><text:p>Line one</text:p><text:p><text:s text:c="2"/>indented</text:p><text:p/><text:p>last</text:p></table:help-message>`,
    );
    expect(xml).toContain(`<table:error-message table:message-type="warning" table:display="false"><text:p>a &lt; b</text:p></table:error-message>`);
    expect(shape(readOne(xml))).toEqual(shape(rule));
  });

  it("an error macro is kept in sourceXml and never run: error.show is false", () => {
    const definition =
      `<table:content-validation table:name="v" table:condition="of:cell-content-is-whole-number()" table:base-cell-address="S.B2">` +
      `<table:error-message table:message-type="stop" table:title="T" table:display="true"/><table:error-macro table:execute="true"/>` +
      `<office:event-listeners><script:event-listener script:event-name="dom:error" xlink:href="vnd.sun.star.script:Standard.M.Check?language=Basic&amp;location=application"/></office:event-listeners>` +
      `</table:content-validation>`;
    const rule = readOne(definition);
    expect(rule).toMatchObject({ type: "whole", error: { show: false, style: "stop", title: "T" }, sourceXml: definition });
  });
});

describe("ranges and rebasing", () => {
  it("rebases a definition's operands from its base cell onto each sheet's origin — one rule per table, same name and source", () => {
    const definition = condition("of:is-true-formula([.B2]&gt;[.$Z$1]+[.C1])");
    const { sheets, orphans } = readOdsContentValidations(prefixOf(block(definition)), [new Map([["v", [range("B2:B5")]]]), new Map([["v", [range("D5:E6")]]])]);
    expect(sheets.map((rules) => rules.map((rule) => [rule.ranges, rule.formula1, rule.odsName, rule.sourceXml]))).toEqual([
      [[[range("B2:B5")], "B2>$Z$1+C1", "v", definition]],
      [[[range("D5:E6")], "D5>$Z$1+E4", "v", definition]],
    ]);
    expect(orphans).toEqual([]);
    expect(sheets[0]![0]!.id).not.toBe(sheets[1]![0]!.id);
  });

  it("a missing base cell is the origin; a literal list is never rebased", () => {
    expect(readOne(`<table:content-validation table:name="v" table:condition="of:is-true-formula([.D4]&gt;0)"/>`, [range("D4")]).formula1).toBe("D4>0");
    expect(readOne(condition("of:cell-content-is-in-list(&quot;B2&quot;)"), [range("F9")]).formula1).toBe(`"B2"`);
  });

  it("a repeated run is ONE rectangle and stores no Cell; runs continuing on the next row extend it", () => {
    const rows =
      row(empty(1) + named("v") + empty(1022)) +
      row(empty(1) + named("v") + named("w", 3) + empty(1019), 999) +
      row(empty(1) + named("v") + empty(1) + named("w") + empty(1020)) +
      tail(1001);
    const content = readOdsContent(contentXml(block(condition("of:cell-content-is-whole-number()").replace(`"v"`, `"v"`) + condition("of:cell-content-is-decimal-number()").replace(`name="v"`, `name="w"`)), tableXml("S", rows)), new CellBudget())!;
    const sheet = content.sheets[0]!;
    expect(sheet.validations.map((rule) => [rule.odsName, rule.ranges])).toEqual([
      ["v", [range("B1:B1001")]],
      ["w", [range("C2:E1000"), range("D1001")]],
    ]);
    expect(sheet.rows).toEqual([]);
    expect(content.orphanValidations).toEqual([]);
  });

  it("LibreOffice's trailing run: a rule over 1,048,553 rows costs one rectangle", () => {
    const rows = row(named("v") + empty(1023), 23) + row(named("v") + empty(1023), 1_048_553);
    const sheet = readOdsContent(contentXml(block(condition("of:cell-content-is-whole-number()")), tableXml("S", rows)), new CellBudget())!.sheets[0]!;
    expect(sheet.validations[0]!.ranges).toEqual([{ top: 0, left: 0, bottom: 1_048_575, right: 0 }]);
    expect(sheet.rows).toEqual([]);
  });

  it("orphans: definitions no cell references, verbatim with the whitespace before them", () => {
    const orphan = `\n  ${condition("of:cell-content-is-decimal-number()").replace(`name="v"`, `name="lonely"`)}`;
    const content = readOdsContent(contentXml(block(condition("of:cell-content-is-whole-number()") + orphan), tableXml("S", row(named("v")) + tail(1))), new CellBudget())!;
    expect(content.orphanValidations).toEqual([orphan]);
    expect(content.sheets[0]!.validations.map((rule) => rule.odsName)).toEqual(["v"]);
  });

  it("readTableValidations: from the table's own cells, or from the ranges it is given", () => {
    const prefix = prefixOf(block(condition("of:is-true-formula([.B2]&gt;0)")));
    const table = tableXml("S", row(empty(1)) + `<table:table-header-rows>${row(empty(1) + named("v", 2))}</table:table-header-rows>` + tail(2));
    expect(tableValidationRanges(table)).toEqual(new Map([["v", [range("B2:C2")]]]));
    expect(readTableValidations(prefix, table, "S").map((rule) => [rule.ranges, rule.formula1])).toEqual([[[range("B2:C2")], "B2>0"]]);
    expect(readTableValidations(prefix, table, "S", new Map([["v", [range("D7:D9")]]])).map((rule) => [rule.ranges, rule.formula1])).toEqual([[[range("D7:D9")], "D7>0"]]);
    expect(readTableValidations(prefix, table, "S", new Map())).toEqual([]);
  });

  it("rangesFromCellNames accumulates single cells as the reader does", () => {
    const cells = [
      { row: 2, col: 1, name: "a" },
      { row: 1, col: 1, name: "a" },
      { row: 1, col: 2, name: "a" },
      { row: 2, col: 2, name: "a" },
      { row: 3, col: 1, name: "b" },
    ];
    expect(rangesFromCellNames(cells)).toEqual(new Map([["a", [range("B2:C3")]], ["b", [range("B4")]]]));
  });
});

// --- the table writer -------------------------------------------------------------------------

describe("the table writer's table:content-validation-name", () => {
  const EN_US = SPREADSHEET_LOCALES["en-US"];

  /** A new sheet's rows, style names stripped, with `rules` named A, B, …. */
  function rowsOf(rules: Omit<DataValidation, "id">[], cells: SpreadsheetModel["sheets"][number]["rows"] = []): string[] {
    const model = createBlankWorkbook();
    const sheet = { ...model.sheets[0]!, rows: cells, validations: rules.map((rule, index) => ({ ...rule, id: `r${index}` })) };
    const names = new Map(sheet.validations.map((rule, index) => [rule.id, String.fromCharCode(65 + index)]));
    const xml = serializeTable(sheet, { model, registry: new OdsStyleRegistry(), locale: EN_US, validationName: (_sheet, rule) => names.get(rule.id) });
    return [...xml.replace(/ table:style-name="[^"]*"/g, "").matchAll(/<table:table-row[^>]*>.*?<\/table:table-row>/g)].map((match) => match[0]);
  }
  const rule = (a1: string[], overrides: Partial<ValidationSettings> = {}): Omit<DataValidation, "id"> => ({ ...settings(overrides), ranges: a1.map(range) });

  it("names each validated cell, a run of empty validated cells and rows compressed, never materialised", () => {
    expect(rowsOf([rule(["B2:B4"]), rule(["C2:C1000"])], [{ index: 2, cells: [{ col: 1, value: 5 }] }])).toEqual([
      `<table:table-row><table:table-cell table:number-columns-repeated="3"/></table:table-row>`,
      `<table:table-row><table:table-cell/><table:table-cell table:content-validation-name="A"/><table:table-cell table:content-validation-name="B"/></table:table-row>`,
      `<table:table-row><table:table-cell/><table:table-cell table:content-validation-name="A" office:value-type="float" office:value="5" calcext:value-type="float"><text:p>5</text:p></table:table-cell>` +
        `<table:table-cell table:content-validation-name="B"/></table:table-row>`,
      `<table:table-row><table:table-cell/><table:table-cell table:content-validation-name="A"/><table:table-cell table:content-validation-name="B"/></table:table-row>`,
      `<table:table-row table:number-rows-repeated="996"><table:table-cell table:number-columns-repeated="2"/><table:table-cell table:content-validation-name="B"/></table:table-row>`,
    ]);
  });

  it("validated and unvalidated neighbours split runs; the first rule in model order wins an overlap", () => {
    expect(rowsOf([rule(["B1:D1"]), rule(["F1:G1", "D1"])])).toEqual([
      `<table:table-row><table:table-cell/><table:table-cell table:content-validation-name="A" table:number-columns-repeated="3"/><table:table-cell/>` +
        `<table:table-cell table:content-validation-name="B" table:number-columns-repeated="2"/></table:table-row>`,
    ]);
  });

  it("a whole-column rule over a 1,048,576-row extent is one repeated row", () => {
    const rows = rowsOf([rule(["A1:A1048576"])]);
    expect(rows).toEqual([`<table:table-row table:number-rows-repeated="1048576"><table:table-cell table:content-validation-name="A"/></table:table-row>`]);
  });

  it("a stale name in preservedAttrs is never written: the model's wins", () => {
    const rows = rowsOf([rule(["A1"])], [{ index: 0, cells: [{ col: 0, value: 1, preservedAttrs: ` table:content-validation-name="old" xml:id="c1"` }] }]);
    expect(rows[0]).toBe(
      `<table:table-row><table:table-cell table:content-validation-name="A" office:value-type="float" office:value="1" calcext:value-type="float" xml:id="c1"><text:p>1</text:p></table:table-cell></table:table-row>`,
    );
    const unvalidated = rowsOf([], [{ index: 0, cells: [{ col: 0, value: 1, preservedAttrs: ` table:content-validation-name="old"` }] }]);
    expect(unvalidated[0]).not.toContain("content-validation-name");
  });
});

// --- regenerating the block ---------------------------------------------------------------------

describe("regenerating table:content-validations", () => {
  const O = SPEC130_VALIDATION_ODS;

  it("an untouched workbook writes back byte-identical", async () => {
    const bytes = await odsPackage(contentXml(block(condition("of:cell-content-is-whole-number()")), tableXml("S", row(named("v")) + tail(1))));
    const out = await unzip(await writeOds(await readSpreadsheetBytes(bytes, "ods", "t")));
    for (const [name, data] of await unzip(bytes)) {
      expect(equalBytes(out.get(name)!, data), name).toBe(true);
    }
  });

  it("unchanged rules keep their name and source; a new one is serialised under a fresh de-val<n>", async () => {
    const model = await readSpreadsheetBytes(await buildSpec130ValidationOdsBytes(), "ods", "t");
    const form = model.sheets[0]!;
    setValidation(model, form.id, [range("D2:D3")], settings({ type: "list", formula1: `"Yes,No"` }));
    const { bytes, content } = await contentOf(model);
    expect(content).toContain(
      `<table:content-validations>${O.definitions.join("")}${O.orphan}<table:content-validation table:name="de-val1" table:condition="of:cell-content-is-in-list(&quot;Yes&quot;;&quot;No&quot;)" ` +
        `table:allow-empty-cell="true" table:display-list="unsorted" table:base-cell-address="Form.D2"><table:help-message table:display="true"></table:help-message>` +
        `<table:error-message table:message-type="stop" table:display="true"></table:error-message></table:content-validation></table:content-validations>`,
    );
    expect(content).toContain(`<table:table-cell table:content-validation-name="de-val1"/>`);
    // set → save → reopen: the same rules, ids, sources and minted names aside.
    const reread = await readSpreadsheetBytes(bytes, "ods", "t2");
    expect(reread.sheets[0]!.validations.map(shape)).toEqual(model.sheets[0]!.validations.map(shape));
    expect(reread.sheets[0]!.validations.at(-1)!.odsName).toBe("de-val1");
    expect(reread.odsOrphanValidations).toEqual([O.orphan]);
  });

  it("a fresh name collides with no name in use — a definition's, an orphan's or one a kept table carries", async () => {
    const definitions = condition("of:cell-content-is-whole-number()").replace(`name="v"`, `name="de-val1"`) + condition("of:cell-content-is-date()").replace(`name="v"`, `name="de-val2"`);
    const tables = tableXml("S", row(named("de-val1")) + tail(1)) + tableXml("T", row(named("de-val3")) + tail(1)) + tableXml("U", row(empty(1)) + tail(1));
    const model = await readSpreadsheetBytes(await odsPackage(contentXml(block(definitions), tables)), "ods", "t");
    expect(model.odsOrphanValidations).toEqual([condition("of:cell-content-is-date()").replace(`name="v"`, `name="de-val2"`)]);
    setValidation(model, model.sheets[2]!.id, [range("A1")], settings({ type: "decimal", operator: "greaterThan", formula1: "0" }));
    const { content } = await contentOf(model);
    expect(content).toContain(`table:name="de-val4"`);
    expect(content).toContain(`<table:table-cell table:content-validation-name="de-val4"/>`);
  });

  it("a rule shared by two clean tables is written once; a dirty table's unchanged copy shares it too", async () => {
    const shared = condition("of:is-true-formula([.B2]&gt;0)");
    const tables = tableXml("S", row(empty(1)) + row(empty(1) + named("v")) + tail(2)) + tableXml("T", row(named("v", 2)) + tail(1)) + tableXml("U", row(empty(1) + named("v")) + tail(1));
    const model = await readSpreadsheetBytes(await odsPackage(contentXml(block(shared), tables)), "ods", "t");
    expect(model.sheets.map((sheet) => sheet.validations.map((rule) => [rule.odsName, rule.formula1]))).toEqual([[["v", "B2>0"]], [["v", "A1>0"]], [["v", "B1>0"]]]);
    setCellEntry(model, 2, { row: 4, col: 0 }, { value: 1 });
    const { bytes, content } = await contentOf(model);
    expect(content.match(/table:name="v"/g)).toHaveLength(1);
    expect(content).toContain(block(shared));
    const reread = await readSpreadsheetBytes(bytes, "ods", "t2");
    expect(reread.sheets.map((sheet) => sheet.validations.map(shape))).toEqual(model.sheets.map((sheet) => sheet.validations.map(shape)));
  });

  it("a rule a command changed is serialised afresh; one cleared everywhere is dropped; an empty block is omitted", async () => {
    const tables = tableXml("S", row(named("v")) + tail(1));
    const model = await readSpreadsheetBytes(await odsPackage(contentXml(block(condition("of:cell-content-is-whole-number()")), tables)), "ods", "t");
    expect(clearValidation(model, model.sheets[0]!.id, [range("A1")])).toBe(true);
    const { content } = await contentOf(model);
    expect(content).not.toContain("content-validation");
    expect(content).toContain(`<table:calculation-settings/><table:table table:name="S">`);
  });

  it("a new block goes after table:calculation-settings, before table:label-ranges", async () => {
    const model = await readSpreadsheetBytes(await odsPackage(contentXml(`<table:label-ranges/>`, tableXml("S", row(empty(1)) + tail(1)))), "ods", "t");
    setValidation(model, model.sheets[0]!.id, [range("A1")], settings({ type: "whole", operator: "greaterThan", formula1: "0" }));
    const { bytes, content } = await contentOf(model);
    expect(content).toContain(`<table:calculation-settings/><table:content-validations><table:content-validation table:name="de-val1" `);
    expect(content).toContain(`</table:content-validations><table:label-ranges/><table:table table:name="S">`);
    expect((await readSpreadsheetBytes(bytes, "ods", "t2")).sheets[0]!.validations.map(shape)).toEqual(model.sheets[0]!.validations.map(shape));
  });

  it("rename → save → reopen: the clean sheet keeps every rule, the remapped one serialised under its own name", async () => {
    const model = await readSpreadsheetBytes(await buildSpec130ValidationOdsBytes(), "ods", "t");
    renameSheet(model, 1, "Data");
    const form = model.sheets[0]!;
    expect(form.dirty).toBeFalsy();
    const sorted = form.validations.find((rule) => rule.odsName === "val8")!;
    expect([sorted.formula1, sorted.sourceXml]).toEqual(["Data!$A$1:$A$4", undefined]);
    const { bytes, content } = await contentOf(model);
    expect(content).toContain(`<table:content-validation table:name="val8" table:condition="of:cell-content-is-in-list([$Data.$A$1:.$A$4])" table:allow-empty-cell="true" table:display-list="sort-ascending" table:base-cell-address="Form.B9">`);
    expect(content.match(/table:name="val8"/g)).toHaveLength(1);
    const reread = await readSpreadsheetBytes(bytes, "ods", "t2");
    expect(reread.sheets[0]!.validations.map(shape)).toEqual(form.validations.map(shape));
  });

  it("duplicate → save → reopen: the copy (clean, its rules unnamed) is rebuilt with the same rules under fresh names", async () => {
    const model = await readSpreadsheetBytes(await buildSpec130ValidationOdsBytes(), "ods", "t");
    duplicateSheet(model, 0);
    const copy = model.sheets[1]!;
    expect(copy.dirty).toBeFalsy();
    expect(copy.validations.every((rule) => rule.odsName === undefined)).toBe(true);
    const { bytes, content } = await contentOf(model);
    expect(content).toContain(`<table:content-validations>${O.definitions.join("")}${O.orphan}`);
    const reread = await readSpreadsheetBytes(bytes, "ods", "t2");
    expect(reread.sheets.map((sheet) => sheet.name)).toEqual(["Form", "Form (2)", "Lists"]);
    expect(reread.sheets[0]!.validations.map((rule) => rule.odsName)).toEqual(O.rules.map((rule) => rule.odsName));
    expect(reread.sheets[1]!.validations.map(shape)).toEqual(copy.validations.map(shape));
    expect(reread.sheets[1]!.validations.every((rule) => rule.odsName!.startsWith("de-val"))).toBe(true);
  });

  it("from scratch (a new workbook, or a copy from another format): every rule serialised, read back the same", async () => {
    const model = createBlankWorkbook();
    const id = model.sheets[0]!.id;
    setValidation(model, id, [range("B2:B5")], settings({ type: "list", formula1: `"Low,Medium,High"`, odsListOrder: "sort-ascending", input: { show: true, title: "Pick", text: "One\nof three" } }));
    setValidation(model, id, [range("C2:C9"), range("E1")], settings({ type: "custom", formula1: "ISNUMBER(C2)", allowBlank: false }));
    setValidation(model, id, [range("D1:D1048576")], settings({ type: "textLength", operator: "notBetween", formula1: "2", formula2: "$Z$1", showDropdown: false, error: { show: true, style: "warning", title: "Hm", text: "Sure?" } }));
    setValidation(model, id, [range("F3")], settings({ type: "date", operator: "lessThan", formula1: "TODAY()" }));
    const { bytes, content } = await contentOf(model);
    expect(content).toMatch(/<table:calculation-settings[^>]*>.*<\/table:calculation-settings><table:content-validations>|<office:spreadsheet><table:content-validations>/);
    const reread = await readSpreadsheetBytes(bytes, "ods", "t2");
    expect(reread.sheets[0]!.validations.map(shape)).toEqual(model.sheets[0]!.validations.map(shape));
    expect(reread.sheets[0]!.validations.map((rule) => rule.odsName)).toEqual(["de-val1", "de-val2", "de-val3", "de-val4"]);
  });
});
