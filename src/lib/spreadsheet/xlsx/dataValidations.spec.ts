// SPEC-130 "xlsx wire mapping" and "x14" (xlsx/dataValidations.ts and its
// hookup in xlsx/sheetXml.ts): every dataValidation attribute read and
// written per the table with Excel's defaults — `type="none"` ↔ `any`,
// `showDropDown="1"` ↔ `showDropdown: false` both ways — formulas through
// fromXlsxFormula/toXlsxFormula, `&#10;` line breaks, disablePrompts; the x14
// ext's rules read with `x14: true`; on write the block at its CT_Worksheet
// place (after conditionalFormatting, before hyperlinks) with `count` and
// the start tag's other attributes kept, created when absent, removed when
// the model has no base rule; unchanged rules emitted from their source
// text; x14 placement (a cross-sheet rule built now goes to the ext, read
// rules keep their location while unchanged); the ext after SPEC-129's, the
// other exts untouched; an untouched workbook byte-identical; and set → save
// → reopen reproducing the model.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { parseA1, parseRange, type GridRange } from "../address";
import { buildSpec130ValidationXlsxBytes, SPEC130_VALIDATION_XLSX } from "../fixtures/build";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import { MAX_ROWS, type Sheet, type SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { CellBudget } from "../parsed";
import { addSheet, createBlankWorkbook } from "../sheets";
import {
  defaultValidationSettings,
  setValidation,
  VALIDATION_OPERATORS,
  VALIDATION_TYPES,
  withoutValidationSource,
  type DataValidation,
  type ValidationSettings,
} from "../validation";
import { CF_EXT_URI, X14_NS, XM_NS } from "./conditionalFormats";
import {
  baseValidationXml,
  DV_EXT_URI,
  readValidationSource,
  readWorksheetValidations,
  writeDataValidations,
  x14ValidationXml,
} from "./dataValidations";
import { readWorksheet, serializeWorksheet, type XlsxSheetContext } from "./sheetXml";
import { writeXlsx } from "./writer";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const ROOT = `<worksheet xmlns="${MAIN}" xmlns:xr="${XR}">`;
const F = SPEC130_VALIDATION_XLSX;
const SPARKLINE_EXT =
  `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14_NS}"><x14:sparklineGroups xmlns:xm="${XM_NS}"><x14:sparklineGroup>` +
  `<x14:sparklines><x14:sparkline><xm:f>S!A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`;
const CF_EXT =
  `<ext uri="${CF_EXT_URI}" xmlns:x14="${X14_NS}"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="${XM_NS}">` +
  `<x14:cfRule type="expression" priority="2" id="{00000000-0000-4000-8000-000000000001}"><xm:f>Other!$A$1&gt;0</xm:f></x14:cfRule><xm:sqref>C1</xm:sqref>` +
  `</x14:conditionalFormatting></x14:conditionalFormattings></ext>`;
const CF_BLOCK = `<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>`;
const HYPERLINKS = `<hyperlinks><hyperlink ref="A1" location="S!B2" display="go"/></hyperlinks>`;
const MARGINS = `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`;

const ranges = (...texts: string[]): GridRange[] => texts.map((text) => parseRange(text)!);
const block = (rules: string, attrs = ""): string => `<dataValidations count="${(rules.match(/<dataValidation[\s>/]/g) ?? []).length}"${attrs}>${rules}</dataValidations>`;
const x14Ext = (rules: string, attrs = ""): string =>
  `<ext uri="${DV_EXT_URI}" xmlns:x14="${X14_NS}"><x14:dataValidations count="${(rules.match(/<x14:dataValidation[\s>]/g) ?? []).length}"${attrs} xmlns:xm="${XM_NS}">${rules}</x14:dataValidations></ext>`;
const x14Rule = (sqref: string, formula1: string, attrs = ` type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"`): string =>
  `<x14:dataValidation${attrs}><x14:formula1><xm:f>${formula1}</xm:f></x14:formula1><xm:sqref>${sqref}</xm:sqref></x14:dataValidation>`;

/** A worksheet part holding `body` after an empty sheetData — the elements
 * before it as an empty sheet's writer rebuilds them. */
const part = (body: string): string => `${ROOT}<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/>${body}</worksheet>`;

/** The rules the reader reads from `body` (a block and/or an extLst). */
function read(body: string): ReturnType<typeof readValidationSource> {
  return readWorksheetValidations(part(body));
}

/** A rule as the round trips compare it: no minted id, no source text. */
function shape(dv: DataValidation): Omit<DataValidation, "id" | "sourceXml"> {
  const { id: _id, sourceXml: _source, ...rest } = dv;
  return rest;
}

/** A complete rule with the defaults a bare `<dataValidation>` reads as. */
function rule(fields: Partial<DataValidation> & { ranges: GridRange[] }): DataValidation {
  return {
    id: "r",
    type: "any",
    operator: "between",
    allowBlank: false,
    showDropdown: true,
    input: { show: false, title: "", text: "" },
    error: { show: false, style: "stop", title: "", text: "" },
    ...fields,
  };
}

function sheetWith(validations: DataValidation[], extra: Partial<Sheet> = {}): Sheet {
  return { conditionalFormats: [], validations, id: "s", name: "S", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], dirty: true, ...extra };
}

function context(): XlsxSheetContext {
  return { active: false, theme: undefined, mdw: 7, strings: { indexFor: () => 0 }, styles: { xfFor: (id) => id }, dxfs: [] };
}

/** The part a dirty sheet holding `validations` writes over `source`. */
function written(validations: DataValidation[], source: string, extra: Partial<Sheet> = {}): string {
  return serializeWorksheet(sheetWith(validations, extra), source, context());
}

// --- reading ----------------------------------------------------------------------------------

describe("reading a dataValidation", () => {
  it("reads a bare rule with Excel's defaults: no type is any, between, nothing allowed or shown, the arrow shown, no formulas", () => {
    const element = `<dataValidation sqref="A1"/>`;
    const { validations, promptsDisabled } = read(block(element));
    expect(validations.map(shape)).toEqual([shape(rule({ ranges: ranges("A1") }))]);
    expect(validations[0]!.sourceXml).toBe(element);
    expect(validations[0]!.formula1).toBeUndefined();
    expect(validations[0]!.x14).toBeUndefined();
    expect(promptsDisabled).toBe(false);
  });

  it("reads every attribute of the table", () => {
    const element =
      `<dataValidation type="whole" errorStyle="warning" imeMode="hiragana" operator="notBetween" allowBlank="1" showDropDown="1" showInputMessage="1" ` +
      `showErrorMessage="1" errorTitle="Oops" error="Not that" promptTitle="Hint" prompt="Type a number" sqref="B2:B5 D1" xr:uid="{00000000-0000-4000-8000-0000000000AA}">` +
      `<formula1>1</formula1><formula2>$Z$1</formula2></dataValidation>`;
    const [read1] = read(block(element)).validations;
    expect(shape(read1!)).toEqual({
      ranges: ranges("B2:B5", "D1"),
      type: "whole",
      operator: "notBetween",
      formula1: "1",
      formula2: "$Z$1",
      allowBlank: true,
      showDropdown: false,
      input: { show: true, title: "Hint", text: "Type a number" },
      error: { show: true, style: "warning", title: "Oops", text: "Not that" },
    });
    // imeMode and xr:uid live only in the source text.
    expect(read1!.sourceXml).toBe(element);
  });

  it(`reads type="none" as any, every type and operator by name, an unknown one as the default, "true" as on`, () => {
    expect(read(block(`<dataValidation type="none" sqref="A1"/>`)).validations[0]!.type).toBe("any");
    for (const type of VALIDATION_TYPES.filter((each) => each !== "any")) {
      expect(read(block(`<dataValidation type="${type}" sqref="A1"/>`)).validations[0]!.type, type).toBe(type);
    }
    for (const operator of VALIDATION_OPERATORS) {
      expect(read(block(`<dataValidation type="whole" operator="${operator}" sqref="A1"/>`)).validations[0]!.operator, operator).toBe(operator);
    }
    const [odd] = read(block(`<dataValidation type="bogus" operator="near" errorStyle="loud" allowBlank="true" showErrorMessage="true" sqref="A1"/>`)).validations;
    expect([odd!.type, odd!.operator, odd!.error.style, odd!.allowBlank, odd!.error.show]).toEqual(["any", "between", "stop", true, true]);
    expect(read(block(`<dataValidation errorStyle="information" sqref="A1"/>`)).validations[0]!.error.style).toBe("information");
  });

  it(`showDropDown="1" HIDES the arrow: showDropdown false; absent or "0" shows it`, () => {
    const { validations } = read(
      block(
        [`showDropDown="1" sqref="A1"`, `showDropDown="0" sqref="A2"`, `sqref="A3"`].map((attrs) => `<dataValidation type="list" ${attrs}><formula1>"a,b"</formula1></dataValidation>`).join(""),
      ),
    );
    expect(validations.map((each) => each.showDropdown)).toEqual([false, true, true]);
  });

  it("decodes attribute text — &#10; is a line break — and reads formulas through fromXlsxFormula", () => {
    const [one] = read(
      block(
        `<dataValidation type="custom" showInputMessage="1" promptTitle="A &amp; B" prompt="Line one&#10;Line &quot;two&quot; &lt;3" sqref="C1">` +
          `<formula1>_xlfn.ISFORMULA(C1)</formula1></dataValidation>`,
      ),
    ).validations;
    expect(one!.input).toEqual({ show: true, title: "A & B", text: `Line one\nLine "two" <3` });
    expect(one!.formula1).toBe("ISFORMULA(C1)");
    const [list] = read(block(`<dataValidation type="list" sqref="A1"><formula1>"Low,Medium,High"</formula1></dataValidation>`)).validations;
    expect(list!.formula1).toBe(`"Low,Medium,High"`);
  });

  it("reads sqref lists, whole columns and rows; drops a rule whose sqref names nothing", () => {
    const { validations } = read(block(`<dataValidation sqref="A1 C3:D4"/><dataValidation sqref="B:B 4:5"/><dataValidation sqref="nowhere"/><dataValidation sqref=""/>`));
    expect(validations.map((each) => each.ranges)).toEqual([
      ranges("A1", "C3:D4"),
      [
        { top: 0, bottom: MAX_ROWS - 1, left: 1, right: 1 },
        { top: 3, bottom: 4, left: 0, right: 16383 },
      ],
    ]);
  });

  it("reads the x14 ext's rules with x14: true — formulas in xm:f, ranges in xm:sqref — after the base block's, each with its source text", () => {
    const x14 = x14Rule("B9", "Lists!$A$1:$A$4", ` type="whole" operator="greaterThan" showDropDown="1"`);
    const source = block(`<dataValidation sqref="A1"/>`) + `<extLst>${SPARKLINE_EXT}${x14Ext(x14)}</extLst>`;
    const { validations } = read(source);
    expect(validations.map(shape)).toEqual([
      shape(rule({ ranges: ranges("A1") })),
      { ...shape(rule({ ranges: ranges("B9"), type: "whole", operator: "greaterThan", formula1: "Lists!$A$1:$A$4", showDropdown: false })), x14: true },
    ]);
    expect(validations[1]!.sourceXml).toBe(x14);
    // Only the {CCE6A557…} ext is read.
    expect(read(`<extLst><ext uri="{00000000-0000-0000-0000-000000000000}" xmlns:x14="${X14_NS}"><x14:dataValidations xmlns:xm="${XM_NS}">${x14}</x14:dataValidations></ext></extLst>`).validations).toEqual([]);
  });

  it("disablePrompts on either container → promptsDisabled", () => {
    expect(read(block(`<dataValidation sqref="A1"/>`, ` disablePrompts="1"`)).promptsDisabled).toBe(true);
    expect(read(`<extLst>${x14Ext(x14Rule("A1", "S2!A1"), ` disablePrompts="1"`)}</extLst>`).promptsDisabled).toBe(true);
    expect(read(block(`<dataValidation sqref="A1"/>`, ` xWindow="10" yWindow="20"`)).promptsDisabled).toBe(false);
  });

  it("readWorksheet fills validations and validationPromptsDisabled; a whole part reads the same through readWorksheetValidations; a prefixed part too", () => {
    const xml = part(block(`<dataValidation type="list" sqref="A1"><formula1>$E$1:$E$3</formula1></dataValidation>`, ` disablePrompts="1"`) + `<extLst>${x14Ext(x14Rule("B1", "S2!A1"))}</extLst>`);
    const content = readWorksheet(xml, "S", { sharedStrings: [], dateSystem: "1900", budget: new CellBudget(), mdw: 7 })!;
    expect(content.validations!.map(shape)).toEqual(readWorksheetValidations(xml).validations.map(shape));
    expect(content.validations!.map((each) => each.x14 === true)).toEqual([false, true]);
    expect(content.validationPromptsDisabled).toBe(true);
    const prefixed = `<x:worksheet xmlns:x="${MAIN}"><x:sheetData/><x:dataValidations count="1"><x:dataValidation type="whole" sqref="A1"><x:formula1>3</x:formula1></x:dataValidation></x:dataValidations></x:worksheet>`;
    expect(readWorksheetValidations(prefixed).validations.map(shape)).toEqual([shape(rule({ ranges: ranges("A1"), type: "whole", formula1: "3" }))]);
  });
});

// --- writing ----------------------------------------------------------------------------------

describe("writing a dataValidation", () => {
  it("writes a default rule bare, as Excel does", () => {
    expect(baseValidationXml(rule({ ranges: ranges("A1") }), "")).toBe(`<dataValidation sqref="A1"/>`);
    expect(x14ValidationXml(rule({ ranges: ranges("A1") }))).toBe(`<x14:dataValidation><xm:sqref>A1</xm:sqref></x14:dataValidation>`);
  });

  it("writes every attribute in Excel's order, text XML-escaped with &#10; line breaks, formulas through toXlsxFormula", () => {
    const full = rule({
      ranges: ranges("B2:B5", "D1"),
      type: "whole",
      operator: "notBetween",
      formula1: "ISFORMULA(A1)",
      formula2: "$Z$1",
      allowBlank: true,
      showDropdown: false,
      input: { show: true, title: "A & B", text: `Line one\nLine "two" <3` },
      error: { show: true, style: "warning", title: "Oops", text: "Not that" },
    });
    expect(baseValidationXml(full, "x:")).toBe(
      `<x:dataValidation type="whole" errorStyle="warning" operator="notBetween" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" ` +
        `errorTitle="Oops" error="Not that" promptTitle="A &amp; B" prompt="Line one&#10;Line &quot;two&quot; &lt;3" sqref="B2:B5 D1">` +
        `<x:formula1>_xlfn.ISFORMULA(A1)</x:formula1><x:formula2>$Z$1</x:formula2></x:dataValidation>`,
    );
    expect(x14ValidationXml(full)).toBe(
      `<x14:dataValidation type="whole" errorStyle="warning" operator="notBetween" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" ` +
        `errorTitle="Oops" error="Not that" promptTitle="A &amp; B" prompt="Line one&#10;Line &quot;two&quot; &lt;3">` +
        `<x14:formula1><xm:f>_xlfn.ISFORMULA(A1)</xm:f></x14:formula1><x14:formula2><xm:f>$Z$1</xm:f></x14:formula2><xm:sqref>B2:B5 D1</xm:sqref></x14:dataValidation>`,
    );
  });

  it(`any ↔ no type; showDropdown: false ↔ showDropDown="1", true ↔ no attribute`, () => {
    expect(baseValidationXml(rule({ ranges: ranges("A1"), type: "any", showDropdown: true }), "")).not.toMatch(/type=|showDropDown=/);
    expect(baseValidationXml(rule({ ranges: ranges("A1"), type: "list", formula1: `"a,b"`, showDropdown: false }), "")).toContain(` showDropDown="1"`);
    expect(baseValidationXml(rule({ ranges: ranges("A1"), type: "list", formula1: `"a,b"`, showDropdown: true }), "")).not.toContain("showDropDown");
  });

  it("every type × operator, flag and style reads back as written", () => {
    const rules: DataValidation[] = [];
    for (const type of VALIDATION_TYPES) {
      for (const operator of VALIDATION_OPERATORS) {
        rules.push(rule({ ranges: ranges(`A${rules.length + 1}`), type, operator, formula1: "1", formula2: "2" }));
      }
    }
    for (const style of ["stop", "warning", "information"] as const) {
      for (const flag of [false, true]) {
        rules.push(
          rule({
            ranges: ranges(`B${rules.length + 1}`),
            type: "list",
            formula1: `"x, y"`,
            allowBlank: flag,
            showDropdown: !flag,
            input: { show: flag, title: flag ? "T" : "", text: "a\nb" },
            error: { show: !flag, style, title: "E", text: flag ? "" : "e" },
          }),
        );
      }
    }
    const xml = block(rules.map((each) => baseValidationXml(each, "")).join(""));
    expect(read(xml).validations.map(shape)).toEqual(rules.map(shape));
    const x14 = `<extLst>${x14Ext(rules.map(x14ValidationXml).join(""))}</extLst>`;
    expect(read(x14).validations.map(shape)).toEqual(rules.map((each) => ({ ...shape(each), x14: true })));
  });
});

// --- the sheet writer -------------------------------------------------------------------------

describe("the sheet writer: the dataValidations block", () => {
  const listRule = rule({ ranges: ranges("D2:D9"), type: "list", formula1: `"Yes,No"`, allowBlank: true, input: { show: true, title: "", text: "" }, error: { show: true, style: "stop", title: "", text: "" } });
  const LIST_XML = `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D2:D9"><formula1>"Yes,No"</formula1></dataValidation>`;

  it("creates the block after conditionalFormatting, before hyperlinks, with its count", () => {
    const out = written([listRule], part(CF_BLOCK + HYPERLINKS + MARGINS), { conditionalFormats: undefined });
    expect(out).toContain(`${CF_BLOCK}<dataValidations count="1">${LIST_XML}</dataValidations>${HYPERLINKS}`);
    // No neighbours: still in CT_Worksheet order, before pageMargins.
    expect(written([listRule], part(MARGINS))).toContain(`<sheetData/><dataValidations count="1">${LIST_XML}</dataValidations>${MARGINS}`);
  });

  it("rewrites a source block in place with count updated and its start tag's other attributes kept; disablePrompts follows the model", () => {
    const source = part(`${block(`<dataValidation sqref="A1"/>`, ` xWindow="120" yWindow="40"`)}${MARGINS}`);
    const out = written([listRule, rule({ ranges: ranges("E1"), type: "whole", operator: "greaterThan", formula1: "0" })], source);
    expect(out).toContain(
      `<sheetData/><dataValidations count="2" xWindow="120" yWindow="40">${LIST_XML}<dataValidation type="whole" operator="greaterThan" sqref="E1"><formula1>0</formula1></dataValidation></dataValidations>${MARGINS}`,
    );
    expect(written([listRule], source, { validationPromptsDisabled: true })).toContain(`<dataValidations count="1" xWindow="120" yWindow="40" disablePrompts="1">`);
    const disabled = part(`${block(`<dataValidation sqref="A1"/>`, ` disablePrompts="1" xWindow="120"`)}${MARGINS}`);
    expect(written([listRule], disabled)).toContain(`<dataValidations count="1" xWindow="120">`);
    expect(written([listRule], part(MARGINS), { validationPromptsDisabled: true })).toContain(`<dataValidations count="1" disablePrompts="1">`);
  });

  it("removes the block when the model has no base rule; leaves an empty source block alone", () => {
    expect(written([], part(`${CF_BLOCK}${block(LIST_XML)}${HYPERLINKS}`), { conditionalFormats: undefined })).toBe(part(`${CF_BLOCK}${HYPERLINKS}`));
    const empty = part(`<dataValidations count="0"/>${MARGINS}`);
    expect(written([], empty)).toBe(empty);
  });

  it("a dirty sheet whose rules are unchanged keeps the block byte for byte; an edit re-emits the unchanged rules from their source text", () => {
    const first = `<dataValidation type="whole" allowBlank="1" sqref="A1:A5" xr:uid="{00000000-0000-4000-8000-000000000001}"><formula1>1</formula1></dataValidation>`;
    const second = `<dataValidation type="list" imeMode="noControl" sqref="B1" xr:uid="{00000000-0000-4000-8000-000000000002}"><formula1>$E$1:$E$3</formula1></dataValidation>`;
    const source = part(`\n  <dataValidations count="2" xWindow="5">\n    ${first}\n    ${second}\n  </dataValidations>${MARGINS}`);
    const model = readWorksheetValidations(source).validations;
    expect(written(model, source)).toBe(source);

    const edited = [model[0]!, withoutValidationSource({ ...model[1]!, formula1: "$E$1:$E$4" })];
    expect(written(edited, source)).toContain(
      `<dataValidations count="2" xWindow="5">${first}<dataValidation type="list" sqref="B1"><formula1>$E$1:$E$4</formula1></dataValidation></dataValidations>`,
    );
    // A rule gone: the block keeps the others verbatim.
    expect(written([model[1]!], source)).toContain(`<dataValidations count="1" xWindow="5">${second}</dataValidations>`);
  });

  it("writes in the part's prefix, its source rules verbatim", () => {
    const own = `<x:dataValidation type="whole" sqref="A1"><x:formula1>3</x:formula1></x:dataValidation>`;
    const source = `<x:worksheet xmlns:x="${MAIN}"><x:sheetData/><x:dataValidations count="1">${own}</x:dataValidations></x:worksheet>`;
    const [kept] = readWorksheetValidations(source).validations;
    expect(written([kept!, listRule], source)).toContain(
      `<x:dataValidations count="2">${own}<x:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D2:D9"><x:formula1>"Yes,No"</x:formula1></x:dataValidation></x:dataValidations>`,
    );
  });

  it("a source text that does not fit the part — another format's, a rule moved from x14 — is serialised instead", () => {
    const foreign = { ...listRule, sourceXml: `<table:content-validation table:name="val1"/>` };
    expect(written([foreign], part(MARGINS))).toContain(`<dataValidations count="1">${LIST_XML}</dataValidations>`);
    const x14Text = x14Rule("D2:D9", `"Yes,No"`);
    const misplaced = { ...listRule, sourceXml: x14Text };
    expect(written([misplaced], part(MARGINS))).toContain(`<dataValidations count="1">${LIST_XML}</dataValidations>`);
  });

  it("a sheet stored before SPEC-130 (no validations) leaves the block and the ext as they are", () => {
    const source = part(`${block(LIST_XML)}<extLst>${x14Ext(x14Rule("B1", "S2!A1"))}</extLst>`);
    expect(written(undefined as unknown as DataValidation[], source)).toBe(source);
  });
});

describe("the sheet writer: the x14 ext", () => {
  const cross = rule({ ranges: ranges("B9"), type: "list", formula1: "Lists!$A$1:$A$4", allowBlank: true, input: { show: true, title: "", text: "" }, error: { show: true, style: "stop", title: "", text: "" } });
  const CROSS_XML = x14Rule("B9", "Lists!$A$1:$A$4");

  it("puts a rule built now whose formulas reference another sheet into the x14 ext, every other rule into the base block", () => {
    const local = rule({ ranges: ranges("C1"), type: "custom", formula1: "S!A1>0" });
    const out = written([cross, local], part(MARGINS));
    expect(out).toContain(`<dataValidations count="1"><dataValidation type="custom" sqref="C1"><formula1>S!A1&gt;0</formula1></dataValidation></dataValidations>${MARGINS}`);
    expect(out).toContain(`${MARGINS}<extLst><ext uri="${DV_EXT_URI}" xmlns:x14="${X14_NS}"><x14:dataValidations count="1" xmlns:xm="${XM_NS}">${CROSS_XML}</x14:dataValidations></ext></extLst></worksheet>`);
    expect(writeDataValidations(sheetWith([cross]), undefined, { prefix: "" })).toEqual({ block: "", ext: `<ext uri="${DV_EXT_URI}" xmlns:x14="${X14_NS}"><x14:dataValidations count="1" xmlns:xm="${XM_NS}">${CROSS_XML}</x14:dataValidations></ext>` });
  });

  it("read rules keep their location while unchanged: an x14 rule naming no other sheet stays in the ext, a base rule naming one stays in the block", () => {
    const x14Local = x14Rule("A1", "$E$1:$E$3");
    const baseCross = `<dataValidation type="list" sqref="A2"><formula1>Lists!$A$1:$A$4</formula1></dataValidation>`;
    const source = part(`${block(baseCross)}${MARGINS}<extLst>${x14Ext(x14Local)}</extLst>`);
    const model = readWorksheetValidations(source).validations;
    expect(model.map((each) => each.x14 === true)).toEqual([false, true]);
    const added = rule({ ranges: ranges("C1"), type: "whole", formula1: "1" });
    const out = written([...model, added], source);
    expect(out).toContain(`<dataValidations count="2">${baseCross}<dataValidation type="whole" sqref="C1"><formula1>1</formula1></dataValidation></dataValidations>`);
    expect(out).toContain(`<extLst>${x14Ext(x14Local)}</extLst>`);
    // Once edited, they go where a rule built now goes.
    const edited = model.map((each) => withoutValidationSource({ ...each, allowBlank: true }));
    const moved = written(edited, source);
    expect(moved).toContain(
      `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1"><formula1>$E$1:$E$3</formula1></dataValidation></dataValidations>`,
    );
    expect(moved).toContain(`<x14:dataValidations count="1" xmlns:xm="${XM_NS}"><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Lists!$A$1:$A$4</xm:f></x14:formula1><xm:sqref>A2</xm:sqref></x14:dataValidation></x14:dataValidations>`);
  });

  it("creates the ext after SPEC-129's {78C0D931…}, else first; every other ext stays as found", () => {
    expect(written([cross], part(`${MARGINS}<extLst>${CF_EXT}${SPARKLINE_EXT}</extLst>`), { conditionalFormats: undefined })).toContain(
      `<extLst>${CF_EXT}<ext uri="${DV_EXT_URI}" xmlns:x14="${X14_NS}"><x14:dataValidations count="1" xmlns:xm="${XM_NS}">${CROSS_XML}</x14:dataValidations></ext>${SPARKLINE_EXT}</extLst>`,
    );
    expect(written([cross], part(`${MARGINS}<extLst>${SPARKLINE_EXT}</extLst>`))).toContain(
      `<extLst><ext uri="${DV_EXT_URI}" xmlns:x14="${X14_NS}"><x14:dataValidations count="1" xmlns:xm="${XM_NS}">${CROSS_XML}</x14:dataValidations></ext>${SPARKLINE_EXT}</extLst>`,
    );
    expect(written([cross], part(`${MARGINS}<extLst/>`))).toContain(`<extLst><ext uri="${DV_EXT_URI}"`);
  });

  it("rewrites only its own ext — in place, its start tags kept, count updated — and drops it when empty, with the extLst when nothing else is left", () => {
    const one = x14Rule("B9", "Lists!$A$1:$A$4");
    const source = part(`${MARGINS}<extLst>${SPARKLINE_EXT}<ext uri="${DV_EXT_URI}" xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}"><x14:dataValidations count="1" xWindow="3">${one}</x14:dataValidations></ext></extLst>`);
    const [kept] = readWorksheetValidations(source).validations;
    const second = rule({ ranges: ranges("B10"), type: "list", formula1: "Other!$A$1:$A$2" });
    expect(written([kept!, second], source)).toContain(
      `<extLst>${SPARKLINE_EXT}<ext uri="${DV_EXT_URI}" xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}"><x14:dataValidations count="2" xWindow="3">${one}` +
        `<x14:dataValidation type="list"><x14:formula1><xm:f>Other!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>B10</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`,
    );
    expect(written([], source)).toBe(part(`${MARGINS}<extLst>${SPARKLINE_EXT}</extLst>`));
    expect(written([], part(`${MARGINS}<extLst>${x14Ext(one)}</extLst>`))).toBe(part(MARGINS));
  });

  it("works beside SPEC-129's ext in one set of splices: the CF ext dropped and the DV ext created, or both rewritten", () => {
    // The model has no conditional formats: the CF ext goes; the DV ext comes.
    const out = written([cross], part(`${CF_BLOCK}${MARGINS}<extLst>${CF_EXT}</extLst>`), { conditionalFormats: [] });
    expect(out).toBe(part(`${MARGINS}<extLst><ext uri="${DV_EXT_URI}" xmlns:x14="${X14_NS}"><x14:dataValidations count="1" xmlns:xm="${XM_NS}">${CROSS_XML}</x14:dataValidations></ext></extLst>`));
    // Neither has anything: the extLst goes whole.
    expect(written([], part(`${MARGINS}<extLst>${CF_EXT}${x14Ext(CROSS_XML)}</extLst>`), { conditionalFormats: [] })).toBe(part(MARGINS));
  });

  it("disablePrompts follows the model on a rewritten x14 container", () => {
    const out = written([cross], part(MARGINS), { validationPromptsDisabled: true });
    expect(out).toContain(`<x14:dataValidations count="1" disablePrompts="1" xmlns:xm="${XM_NS}">`);
    expect(readWorksheetValidations(out).promptsDisabled).toBe(true);
  });
});

// --- whole workbooks -------------------------------------------------------------------------

/** The dialog's defaults with `fields` over them. */
const settings = (fields: Partial<ValidationSettings>): ValidationSettings => ({ ...defaultValidationSettings(), ...fields });

/** Every rule of a model's sheets as the round trips compare them. */
function rulesOf(model: SpreadsheetModel): unknown[] {
  return model.sheets.map((sheet) => ({ name: sheet.name, disabled: sheet.validationPromptsDisabled, rules: sheet.validations.map(shape) }));
}

describe("whole workbooks", () => {
  it("an untouched spec130-validation.xlsx downloads byte-identical, part for part", async () => {
    const bytes = await buildSpec130ValidationXlsxBytes();
    const source = await unzip(bytes);
    const out = await unzip(await writeXlsx(await readSpreadsheetBytes(bytes, "xlsx", "t")));
    expect([...out.keys()]).toEqual([...source.keys()]);
    for (const [name, data] of source) {
      expect(equalBytes(out.get(name)!, data), name).toBe(true);
    }
  });

  it("a dirty sheet whose validations were not touched writes its block and ext byte for byte", async () => {
    const model = await readSpreadsheetBytes(await buildSpec130ValidationXlsxBytes(), "xlsx", "t");
    model.sheets[0] = { ...model.sheets[0]!, dirty: true };
    const out = await unzip(await writeXlsx(model));
    const form = new TextDecoder().decode(out.get(F.form.partName));
    expect(form).toContain(`${F.form.block}<pageMargins`);
    expect(form).toContain(`<extLst>${F.form.ext}</extLst>`);
  });

  it("set → save → reopen reproduces the validations (ids and source text aside), cross-sheet rules in x14, an edited fixture rule serialised", async () => {
    const model = await readSpreadsheetBytes(await buildSpec130ValidationXlsxBytes(), "xlsx", "t");
    const form = model.sheets[0]!;
    setValidation(model, form.id, ranges("D2"), settings({ type: "list", formula1: `"Yes,No"` }));
    setValidation(model, form.id, ranges("D3:D4", "F1"), settings({ type: "time", operator: "lessThan", formula1: "0.5", error: { show: true, style: "information", title: "Late", text: "Before\nnoon" } }));
    setValidation(model, form.id, ranges("D5"), settings({ type: "list", formula1: "Lists!$B$1:$B$3", showDropdown: false, input: { show: true, title: "Colour", text: "One of\nthree" } }));
    // B3's rule re-set: the fixture's source text goes, the rule is serialised.
    setValidation(model, form.id, ranges("B3"), settings({ type: "whole", operator: "between", formula1: "1", formula2: "20", error: { show: true, style: "warning", title: "", text: "" } }));
    model.sheets = model.sheets.map((sheet) => (sheet.id === form.id ? { ...sheet, validationPromptsDisabled: true } : sheet));
    const bytes = await writeXlsx(model);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const reopened = await readSpreadsheetBytes(bytes, "xlsx", "t2");
    // The file keeps the base rules, then the x14 ones.
    const byLocation = (each: SpreadsheetModel): unknown[] =>
      rulesOf({ ...each, sheets: each.sheets.map((sheet) => ({ ...sheet, validations: sheet.validations.toSorted((a, b) => Number(a.x14 === true) - Number(b.x14 === true)) })) });
    expect(rulesOf(reopened)).toEqual(byLocation(model));
    const d5 = parseA1("D5")!;
    expect(reopened.sheets[0]!.validations.find((each) => each.ranges[0]!.top === d5.row && each.ranges[0]!.left === d5.col)?.x14).toBe(true);
    const formXml = new TextDecoder().decode((await unzip(bytes)).get(F.form.partName));
    // The untouched fixture rules are still their source text.
    for (const key of ["priority", "amount", "due", "code", "email", "option", "colour"] as const) {
      expect(formXml).toContain(F.form.rules[key]);
    }
    expect(formXml).not.toContain(F.form.rules.score);
    expect(formXml).toContain(F.form.x14Rule);
  });

  it("a workbook built from scratch writes valid validations that read back", async () => {
    const model = createBlankWorkbook();
    addSheet(model, 0);
    const [first, second] = model.sheets;
    setValidation(model, first!.id, ranges("A1:A3"), { ...defaultValidationSettings(), type: "decimal", operator: "greaterThanOrEqual", formula1: "2.5" });
    setValidation(model, first!.id, ranges("B1"), { ...defaultValidationSettings(), type: "list", formula1: `${second!.name}!$A$1:$A$3` });
    setValidation(model, second!.id, ranges("C1"), { ...defaultValidationSettings(), type: "textLength", operator: "equal", formula1: "4", allowBlank: false });
    const bytes = await writeXlsx(model);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const reopened = await readSpreadsheetBytes(bytes, "xlsx", "t2");
    expect(rulesOf(reopened)).toEqual(rulesOf(model));
    expect(reopened.sheets[0]!.validations.map((each) => each.x14 === true)).toEqual([false, true]);
  });

  it("a clean sheet's retained part keeps taking the sheet-name pass, validation formulas included; a dirty sheet's unchanged rules take it too", async () => {
    for (const dirty of [false, true]) {
      const model = await readSpreadsheetBytes(await buildSpec130ValidationXlsxBytes(), "xlsx", "t");
      model.sheets = model.sheets.map((sheet, index) => (index === 1 ? { ...sheet, name: "Sources" } : index === 0 && dirty ? { ...sheet, dirty: true } : sheet));
      model.structureDirty = true;
      const out = await unzip(await writeXlsx(model));
      const form = new TextDecoder().decode(out.get(F.form.partName));
      expect(form, `dirty ${dirty}`).toContain(`<xm:f>Sources!$A$1:$A$4</xm:f>`);
      expect(form, `dirty ${dirty}`).not.toContain(`Lists!`);
    }
  });
});
