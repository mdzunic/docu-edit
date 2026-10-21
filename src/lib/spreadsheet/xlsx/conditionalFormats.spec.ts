// SPEC-129 "xlsx wire mapping" and "x14" (xlsx/conditionalFormats.ts and its
// hookup in xlsx/sheetXml.ts): every cfRule type read per the table — the
// text, blank, error and time-period kinds from their attributes, their
// formula ignored — with each block kept verbatim; what the model cannot hold
// preserved verbatim with its sqref still modelled; the x14 ext — a base data
// bar paired with its twin by GUID and overlaid, the x14-only icon sets and
// cross-sheet rules (inline dxf interned) modelled, custom icon sets
// preserved — in one global priority order. Writing a dirty sheet: blocks
// verbatim while unedited; after an edit, serialised rules carry the model's
// priorities and a verbatim block that would no longer rank is patched in its
// priority attributes only; a data bar always a base rule plus an x14 twin
// sharing an upper-case GUID; the regenerated formulas per kind; the
// `{78C0D931…}` ext first in the extLst with its namespaces on the ext, every
// other ext untouched, dropped when empty; the part's prefix; and whole
// workbooks — brand new and a stripped cross-format copy — that write valid
// conditional formats and dxfs.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { formatRange, parseRange, type GridRange } from "../address";
import {
  addConditionalFormat,
  deleteRule,
  modelledDxf,
  moveRule,
  rulesByPriority,
  type CfDataBar,
  type CfRule,
  type CfRuleDraft,
  type ConditionalFormat,
  type Dxf,
} from "../conditionalFormats";
import { spreadsheetDownload } from "../download";
import { buildSpec129CfXlsxBytes } from "../fixtures/build";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import type { Sheet, SheetColorRef, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { CellBudget } from "../parsed";
import { createBlankWorkbook, stripRetention } from "../sheets";
import { internDxf } from "../styles";
import { CF_EXT_URI, readConditionalFormatSource, regeneratedFormula, RULE_ID_EXT_URI, writeConditionalFormats, X14_NS, XM_NS } from "./conditionalFormats";
import { readWorksheet, serializeWorksheet, type XlsxSheetContext } from "./sheetXml";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const ROOT = `<worksheet xmlns="${MAIN}" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">`;
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;
const SPARKLINE_EXT =
  `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14_NS}"><x14:sparklineGroups xmlns:xm="${XM_NS}"><x14:sparklineGroup>` +
  `<x14:sparklines><x14:sparkline><xm:f>S!A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`;
const DV_EXT = `<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14_NS}"><x14:dataValidations count="0" xmlns:xm="${XM_NS}"/></ext>`;
const RED: Dxf = { font: { color: "#9c0006" }, fill: { kind: "solid", color: "#ffc7ce" } };

const cf = (sqref: string, rules: string, attrs = ""): string => `<conditionalFormatting sqref="${sqref}"${attrs}>${rules}</conditionalFormatting>`;
const x14Block = (sqref: string, rules: string): string =>
  `<x14:conditionalFormatting xmlns:xm="${XM_NS}">${rules}<xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`;
const cfExt = (blocks: string): string => `<ext uri="${CF_EXT_URI}" xmlns:x14="${X14_NS}"><x14:conditionalFormattings>${blocks}</x14:conditionalFormattings></ext>`;
const range = (text: string): GridRange => parseRange(text)!;

/** The formats the reader makes of `blocks` and an `<extLst>`. */
function read(blocks: string[], extLst?: string, holder: { dxfs: Dxf[] } = { dxfs: [] }, sheetName = "S"): ConditionalFormat[] {
  return readConditionalFormatSource({ rootTag: ROOT, blocks, ...(extLst === undefined ? {} : { extLst }) }, sheetName, holder);
}

/** A format as compared: its sqref and its rules without minted ids. */
function shape(format: ConditionalFormat): { sqref: string; rules: Omit<CfRule, "id">[] } {
  return { sqref: format.ranges.map(formatRange).join(" "), rules: format.rules.map(({ id: _id, ...rule }) => rule) };
}

/** shape() without the x14 GUID a save mints. */
function withoutGuid(format: ConditionalFormat): { sqref: string; rules: Omit<CfRule, "id" | "x14Id">[] } {
  const { sqref, rules } = shape(format);
  return { sqref, rules: rules.map(({ x14Id: _guid, ...rule }) => rule) };
}

/** A worksheet part holding `body` after an empty sheetData. */
const part = (body: string): string => `${ROOT}<sheetData/>${body}</worksheet>`;

/** The formats readWorksheet reads from a whole part. */
function readPart(xml: string, holder: { dxfs: Dxf[] } = { dxfs: [] }): ConditionalFormat[] {
  return readWorksheet(xml, "S", { sharedStrings: [], dateSystem: "1900", budget: new CellBudget(), mdw: 7, dxfs: holder })!.conditionalFormats ?? [];
}

function sheetWith(conditionalFormats: ConditionalFormat[]): Sheet {
  return { conditionalFormats, validations: [], id: "s", name: "S", visibility: "visible", kind: "grid", rows: [], columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [], dirty: true };
}

function context(dxfs: Dxf[] = []): XlsxSheetContext {
  return { active: false, theme: undefined, mdw: 7, strings: { indexFor: () => 0 }, styles: { xfFor: (id) => id }, dxfs };
}

/** A blank xlsx workbook whose one sheet is `sheet`. */
function modelWith(sheet: Sheet, dxfs: Dxf[] = []): SpreadsheetModel {
  return { ...createBlankWorkbook(), sheets: [sheet], dxfs };
}

/** The rules of a whole model as the round trips compare them: no minted
 * ids, no x14 GUID, no source text, each dxfId replaced by the dxf it names;
 * formats by their first rule's priority (the file keeps the base and the
 * x14 formats in two lists, so their interleaving in the array is not
 * carried — precedence is the priority). */
function rulesOf(model: SpreadsheetModel): unknown[] {
  return model.sheets.map((sheet) =>
    sheet.conditionalFormats.toSorted((a, b) => (a.rules[0]?.priority ?? 0) - (b.rules[0]?.priority ?? 0)).map((format) => ({
      ranges: format.ranges,
      pivot: format.pivot,
      rules: format.rules.map(({ id: _id, x14Id: _guid, dxfId, ...rule }) => ({ ...rule, dxf: dxfId === undefined ? undefined : modelledDxf(model.dxfs[dxfId]!) })),
    })),
  );
}

/** A model as a cross-format copy carries it: stripRetention drops every
 * conditional-format source text, x14 GUID and dxf source, and the
 * preserved rules (they cannot cross formats). */
function crossFormat(model: SpreadsheetModel): SpreadsheetModel {
  const stripped = stripRetention(model);
  for (const sheet of stripped.sheets) {
    for (const format of sheet.conditionalFormats) {
      expect(format.sourceXml).toBeUndefined();
      expect(format.rules.filter((rule) => rule.x14Id !== undefined || rule.kind === "preserved")).toEqual([]);
    }
  }
  expect(stripped.dxfs.filter((dxf) => dxf.sourceXml !== undefined)).toEqual([]);
  return stripped;
}

/** How many of a model's rules `keep` accepts. */
function ruleCount(model: SpreadsheetModel, keep: (rule: CfRule) => boolean): number {
  return model.sheets.flatMap((sheet) => sheet.conditionalFormats.flatMap((format) => format.rules.filter(keep))).length;
}

/** Accent 6, darker 25%, as the xlsx reader keeps a tinted theme colour. */
const ACCENT6_DARKER: SheetColorRef = { theme: "accent6", lumMod: 0.75, xlsxTint: "-0.25" };

const BAR: CfDataBar = {
  min: { type: "autoMin" },
  max: { type: "autoMax" },
  color: "#638ec6",
  gradient: true,
  axis: "automatic",
  direction: "context",
  minLengthPct: 0,
  maxLengthPct: 100,
  showValue: true,
};

const STARS: CfRuleDraft = { kind: "iconSet", icons: { set: "3Stars", thresholds: [0, 33, 67].map((at) => ({ at: { type: "percent", value: String(at) }, gte: true })), reverse: false, showValue: true } };

describe("reading every cfRule type per the table", () => {
  const blocks = [
    `\n  ${cf("A1:A5 C1:C5", `<cfRule type="cellIs" dxfId="0" priority="1" operator="between"><formula>1</formula><formula>$Z$1</formula></cfRule><cfRule type="cellIs" dxfId="1" priority="2" stopIfTrue="1" operator="lessThan"><formula>"x"</formula></cfRule>`, ` xr:uid="{00000000-0000-0000-0000-000000000001}"`)}`,
    cf("B2:B9", `<cfRule type="expression" dxfId="2" priority="3"><formula>_xlfn.ISOWEEKNUM(B2)=1</formula></cfRule>`),
    cf(
      "C1:C9",
      `<cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="formula" val="$Z$2*2"/>` +
        `<color rgb="FFF8696B"/><color theme="0" tint="-0.249977111117893"/><color indexed="10"/></colorScale></cfRule>`,
    ),
    cf("D1:D9", `<cfRule type="iconSet" priority="5"><iconSet showValue="0" reverse="1"><cfvo type="percent" val="0"/><cfvo type="num" val="5" gte="0"/><cfvo type="percentile" val="80"/></iconSet></cfRule>`),
    cf("E1:E9", `<cfRule type="top10" dxfId="0" priority="6" bottom="1" rank="3"/><cfRule type="top10" dxfId="0" priority="7" percent="1" rank="25"/>`),
    cf("F1:F9", `<cfRule type="aboveAverage" dxfId="0" priority="8" aboveAverage="0" equalAverage="1"/><cfRule type="aboveAverage" dxfId="0" priority="9" stdDev="2"/>`),
    cf("G1:G9", `<cfRule type="duplicateValues" dxfId="0" priority="10"/><cfRule type="uniqueValues" dxfId="0" priority="11"/>`),
    cf(
      "H1:H9",
      `<cfRule type="containsText" dxfId="0" priority="12" operator="containsText" text="an"><formula>stale</formula></cfRule>` +
        `<cfRule type="notContainsText" dxfId="0" priority="13" operator="notContains" text="x"><formula>ISERROR(SEARCH("x",H1))</formula></cfRule>` +
        `<cfRule type="beginsWith" dxfId="0" priority="14" operator="beginsWith" text="A&quot;b"><formula>LEFT(H1,LEN("A""b"))="A""b"</formula></cfRule>` +
        `<cfRule type="endsWith" dxfId="0" priority="15" operator="endsWith" text="z"/>`,
    ),
    cf(
      "I1:I9",
      `<cfRule type="containsBlanks" dxfId="0" priority="16"><formula>LEN(TRIM(I1))=0</formula></cfRule><cfRule type="notContainsBlanks" dxfId="0" priority="17"><formula>LEN(TRIM(I1))&gt;0</formula></cfRule>` +
        `<cfRule type="containsErrors" dxfId="0" priority="18"><formula>ISERROR(I1)</formula></cfRule><cfRule type="notContainsErrors" dxfId="0" priority="19"><formula>NOT(ISERROR(I1))</formula></cfRule>`,
    ),
    cf("J1:J9", `<cfRule type="timePeriod" dxfId="0" priority="20" timePeriod="yesterday"><formula>FLOOR(J1,1)=TODAY()-1</formula></cfRule><cfRule type="timePeriod" dxfId="0" priority="21" timePeriod="nextMonth"/>`),
  ];

  it("models each kind from its attributes and children (formulas through fromXlsxFormula), keeping each block with the whitespace before it", () => {
    const formats = read(blocks);
    expect(formats.map(shape)).toEqual([
      {
        sqref: "A1:A5 C1:C5",
        rules: [
          { priority: 1, dxfId: 0, kind: "cellIs", operator: "between", formulas: ["1", "$Z$1"] },
          { priority: 2, dxfId: 1, stopIfTrue: true, kind: "cellIs", operator: "lessThan", formulas: [`"x"`] },
        ],
      },
      { sqref: "B2:B9", rules: [{ priority: 3, dxfId: 2, kind: "expression", formula: "ISOWEEKNUM(B2)=1" }] },
      {
        sqref: "C1:C9",
        rules: [
          {
            priority: 4,
            kind: "colorScale",
            stops: [
              { at: { type: "min" }, color: "#f8696b" },
              { at: { type: "percentile", value: "50" }, color: { theme: "lt1", lumMod: 1 - 0.249977111117893, xlsxTint: "-0.249977111117893" } },
              { at: { type: "formula", value: "$Z$2*2" }, color: "#ff0000" },
            ],
          },
        ],
      },
      {
        sqref: "D1:D9",
        rules: [
          {
            priority: 5,
            kind: "iconSet",
            icons: {
              set: "3TrafficLights1",
              thresholds: [
                { at: { type: "percent", value: "0" }, gte: true },
                { at: { type: "num", value: "5" }, gte: false },
                { at: { type: "percentile", value: "80" }, gte: true },
              ],
              reverse: true,
              showValue: false,
            },
          },
        ],
      },
      {
        sqref: "E1:E9",
        rules: [
          { priority: 6, dxfId: 0, kind: "top10", bottom: true, percent: false, rank: 3 },
          { priority: 7, dxfId: 0, kind: "top10", bottom: false, percent: true, rank: 25 },
        ],
      },
      {
        sqref: "F1:F9",
        rules: [
          { priority: 8, dxfId: 0, kind: "aboveAverage", below: true, equalAverage: true },
          { priority: 9, dxfId: 0, kind: "aboveAverage", below: false, equalAverage: false, stdDev: 2 },
        ],
      },
      {
        sqref: "G1:G9",
        rules: [
          { priority: 10, dxfId: 0, kind: "duplicates", unique: false },
          { priority: 11, dxfId: 0, kind: "duplicates", unique: true },
        ],
      },
      {
        sqref: "H1:H9",
        rules: [
          { priority: 12, dxfId: 0, kind: "text", operator: "contains", text: "an" },
          { priority: 13, dxfId: 0, kind: "text", operator: "notContains", text: "x" },
          { priority: 14, dxfId: 0, kind: "text", operator: "beginsWith", text: `A"b` },
          { priority: 15, dxfId: 0, kind: "text", operator: "endsWith", text: "z" },
        ],
      },
      {
        sqref: "I1:I9",
        rules: [
          { priority: 16, dxfId: 0, kind: "blanks", not: false },
          { priority: 17, dxfId: 0, kind: "blanks", not: true },
          { priority: 18, dxfId: 0, kind: "errors", not: false },
          { priority: 19, dxfId: 0, kind: "errors", not: true },
        ],
      },
      {
        sqref: "J1:J9",
        rules: [
          { priority: 20, dxfId: 0, kind: "timePeriod", period: "yesterday" },
          { priority: 21, dxfId: 0, kind: "timePeriod", period: "nextMonth" },
        ],
      },
    ]);
    expect(formats.map((format) => format.sourceXml)).toEqual(blocks);
  });

  it("reads a base-only data bar as Excel 2007 draws it: gradient, no axis, its own lengths", () => {
    const formats = read([cf("K1:K9", `<cfRule type="dataBar" priority="1"><dataBar minLength="5" maxLength="80" showValue="0"><cfvo type="num" val="0"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>`)]);
    expect(formats.map(shape)).toEqual([
      {
        sqref: "K1:K9",
        rules: [
          {
            priority: 1,
            kind: "dataBar",
            bar: { min: { type: "num", value: "0" }, max: { type: "max" }, color: "#638ec6", gradient: true, axis: "none", direction: "context", minLengthPct: 5, maxLengthPct: 80, showValue: false },
          },
        ],
      },
    ]);
  });

  it("keeps what the model cannot hold as a preserved rule (where base) verbatim, its sqref still modelled; a block naming no range is not modelled", () => {
    const rules = [
      `<cfRule type="futureKind" priority="1" dxfId="0"/>`,
      `<cfRule type="cellIs" priority="2" operator="sideways"><formula>1</formula></cfRule>`,
      `<cfRule type="iconSet" priority="3"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="50"/></iconSet></cfRule>`,
      `<cfRule type="expression" priority="4"><formula>A1</formula><extLst><ext uri="{X}"/></extLst></cfRule>`,
      `<cfRule type="iconSet" priority="5"><iconSet iconSet="9Planets"><cfvo type="percent" val="0"/></iconSet></cfRule>`,
    ];
    const formats = read([cf("A1:B2 D4", rules.join("")), cf("", `<cfRule type="expression" priority="6"><formula>1</formula></cfRule>`)]);
    expect(formats.map(shape)).toEqual([{ sqref: "A1:B2 D4", rules: rules.map((xml, index) => ({ priority: index + 1, kind: "preserved", xml, where: "base" })) }]);
  });
});

describe("x14", () => {
  const LOWER = "{8a1e2f3c-0129-4d5e-9f60-7a8b9c0d1e2f}";
  const UPPER = LOWER.toUpperCase();
  const baseBar = (guid: string, priority = 2): string =>
    cf(
      "C2:C13",
      `<cfRule type="dataBar" priority="${priority}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>` +
        `<extLst><ext uri="${RULE_ID_EXT_URI}" xmlns:x14="${X14_NS}"><x14:id>${guid}</x14:id></ext></extLst></cfRule>`,
    );

  it("pairs a base data bar with the x14 rule of its GUID: every x14 setting overlays the base's, the base priority counts and the GUID is kept", () => {
    const twin =
      `<x14:cfRule type="dataBar" id="${LOWER}"><x14:dataBar minLength="5" maxLength="95" gradient="0" border="1" direction="rightToLeft" negativeBarColorSameAsPositive="1" ` +
      `negativeBarBorderColorSameAsPositive="0" axisPosition="middle"><x14:cfvo type="num"><xm:f>Other!$A$1</xm:f></x14:cfvo><x14:cfvo type="autoMax"/>` +
      `<x14:borderColor rgb="FF00FF00"/><x14:negativeBorderColor rgb="FF0000FF"/><x14:axisColor theme="1"/></x14:dataBar></x14:cfRule>`;
    const formats = read([baseBar(UPPER, 7)], `<extLst>${cfExt(x14Block("C2:C13", twin))}${SPARKLINE_EXT}</extLst>`);
    expect(formats.map(shape)).toEqual([
      {
        sqref: "C2:C13",
        rules: [
          {
            priority: 1,
            kind: "dataBar",
            x14Id: UPPER,
            bar: {
              min: { type: "num", value: "Other!$A$1" },
              max: { type: "autoMax" },
              color: "#638ec6",
              gradient: false,
              border: "#00ff00",
              negativeColor: "#638ec6",
              negativeBorderColor: "#0000ff",
              axis: "middle",
              axisColor: { theme: "dk1" },
              direction: "rightToLeft",
              minLengthPct: 5,
              maxLengthPct: 95,
              showValue: true,
            },
          },
        ],
      },
    ]);
    expect(formats[0]!.sourceXml).toBe(baseBar(UPPER, 7));
  });

  it("applies the x14 defaults a twin leaves out", () => {
    const twin = `<x14:cfRule type="dataBar" id="${UPPER}"><x14:dataBar><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>`;
    const [format] = read([baseBar(UPPER)], `<extLst>${cfExt(x14Block("C2:C13", twin))}</extLst>`);
    expect(format!.rules[0]).toMatchObject({
      kind: "dataBar",
      bar: { min: { type: "autoMin" }, max: { type: "autoMax" }, gradient: true, axis: "automatic", direction: "context", minLengthPct: 10, maxLengthPct: 90, negativeColor: "#ff0000", axisColor: "#000000" },
    });
  });

  it("models the x14-only icon sets and cross-sheet cell/expression rules (inline dxf interned) with x14: true; custom icon sets and anything else stay preserved (where x14); one global priority order", () => {
    const fiveBoxes =
      `<x14:cfRule type="iconSet" priority="30" id="{11111111-1111-1111-1111-111111111111}"><x14:iconSet iconSet="5Boxes" showValue="0">` +
      [0, 20, 40, 60, 80].map((at) => `<x14:cfvo type="percent"><xm:f>${at}</xm:f></x14:cfvo>`).join("") +
      `</x14:iconSet></x14:cfRule>`;
    const crossCell = `<x14:cfRule type="cellIs" priority="10" operator="greaterThan" id="{22222222-2222-2222-2222-222222222222}"><xm:f>Lists!$A$1</xm:f><x14:dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></x14:dxf></x14:cfRule>`;
    const crossExpression = `<x14:cfRule type="expression" priority="40" id="{33333333-3333-3333-3333-333333333333}"><xm:f>B1&gt;Lists!$B$1</xm:f><x14:dxf><font><b/></font></x14:dxf></x14:cfRule>`;
    const custom =
      `<x14:cfRule type="iconSet" priority="20" id="{44444444-4444-4444-4444-444444444444}"><x14:iconSet iconSet="3Arrows" custom="1">` +
      [0, 33, 67].map((at) => `<x14:cfvo type="percent"><xm:f>${at}</xm:f></x14:cfvo>`).join("") +
      `<x14:cfIcon iconSet="3Flags" iconId="0"/><x14:cfIcon iconSet="3Symbols" iconId="1"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet></x14:cfRule>`;
    const sameSheet = `<x14:cfRule type="expression" priority="50" id="{55555555-5555-5555-5555-555555555555}"><xm:f>H1&gt;1</xm:f></x14:cfRule>`;
    const holder = { dxfs: [{ ...RED, sourceXml: "<dxf/>" }] as Dxf[] };
    const extLst = `<extLst>${cfExt(x14Block("F2:F9", fiveBoxes) + x14Block("A1:A9", crossCell + crossExpression) + x14Block("G2:G9", custom) + x14Block("H1", sameSheet))}${SPARKLINE_EXT}</extLst>`;
    const formats = read([cf("Z1", `<cfRule type="expression" dxfId="0" priority="25"><formula>TRUE</formula></cfRule>`)], extLst, holder);
    expect(formats.map(shape)).toEqual([
      { sqref: "Z1", rules: [{ priority: 3, dxfId: 0, kind: "expression", formula: "TRUE" }] },
      {
        sqref: "F2:F9",
        rules: [
          {
            priority: 4,
            x14: true,
            x14Id: "{11111111-1111-1111-1111-111111111111}",
            kind: "iconSet",
            icons: { set: "5Boxes", thresholds: [0, 20, 40, 60, 80].map((at) => ({ at: { type: "percent", value: String(at) }, gte: true })), reverse: false, showValue: false },
          },
        ],
      },
      {
        sqref: "A1:A9",
        rules: [
          { priority: 1, dxfId: 0, x14: true, x14Id: "{22222222-2222-2222-2222-222222222222}", kind: "cellIs", operator: "greaterThan", formulas: ["Lists!$A$1"] },
          { priority: 5, dxfId: 1, x14: true, x14Id: "{33333333-3333-3333-3333-333333333333}", kind: "expression", formula: "B1>Lists!$B$1" },
        ],
      },
      { sqref: "G2:G9", rules: [{ priority: 2, kind: "preserved", xml: custom, where: "x14" }] },
      { sqref: "H1", rules: [{ priority: 6, kind: "preserved", xml: sameSheet, where: "x14" }] },
    ]);
    // The equal inline dxf reused the table's entry; the new one was appended.
    expect(holder.dxfs).toEqual([{ ...RED, sourceXml: "<dxf/>" }, { font: { bold: true } }]);
    // An x14 block keeps its element (with the whitespace before it) as source.
    expect(formats[1]!.sourceXml).toBe(x14Block("F2:F9", fiveBoxes));
  });
});

describe("writing a dirty sheet", () => {
  const blockA = `\n  ${cf("A1:A9", `<cfRule type="cellIs" dxfId="0" priority="3" operator="equal"><formula>1</formula></cfRule>`, ` xr:uid="{00000000-0000-0000-0000-00000000000A}"`)}`;
  const blockB = `\n  ${cf("B1:B9", `<cfRule type="expression" dxfId="0" priority="7"><formula>$B1&gt;2</formula></cfRule>`)}`;
  const blockC = `\n  ${cf("C1:C9", `<cfRule type="top10" dxfId="0" priority="9" rank="2"/>`)}`;
  const twin = `<x14:cfRule type="dataBar" id="{8A1E2F3C-0129-4D5E-9F60-7A8B9C0D1E2F}"><x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>`;
  const blockD = `\n  ${cf(
    "D1:D9",
    `<cfRule type="dataBar" priority="11"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="${RULE_ID_EXT_URI}" xmlns:x14="${X14_NS}"><x14:id>{8A1E2F3C-0129-4D5E-9F60-7A8B9C0D1E2F}</x14:id></ext></extLst></cfRule>`,
  )}`;
  const extLst = `<extLst>${cfExt(x14Block("D1:D9", twin))}${SPARKLINE_EXT}</extLst>`;
  const base = part(`${blockA}${blockB}${blockC}${blockD}\n  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${extLst}`);

  it("re-emits every block and the extLst verbatim while no rule was edited", () => {
    const formats = readPart(base);
    expect(formats).toHaveLength(4);
    const out = serializeWorksheet(sheetWith(formats), base, context([RED]));
    expect(out.slice(out.indexOf(blockA))).toBe(base.slice(base.indexOf(blockA)));
  });

  it("after an edit: serialised rules take the model's priorities; a verbatim block whose written ones would no longer rank is patched in its priority attributes only; the part re-reads in the model's order", () => {
    const model = modelWith(sheetWith(readPart(base)), [RED]);
    const ruleC = model.sheets[0]!.conditionalFormats[2]!.rules[0]!;
    moveRule(model, "s", ruleC.id, "up");
    const formats = model.sheets[0]!.conditionalFormats;
    expect(formats.map((format) => format.sourceXml !== undefined)).toEqual([true, false, false, true]);
    const out = serializeWorksheet(model.sheets[0]!, base, context([RED]));
    // A (model 1, written 3) would rank below C (serialised at 2): patched alone.
    expect(out).toContain(blockA.replace(`priority="3"`, `priority="1"`));
    expect(out).toContain(`<conditionalFormatting sqref="C1:C9"><cfRule type="top10" dxfId="0" priority="2" rank="2"/></conditionalFormatting>`);
    expect(out).toContain(`<conditionalFormatting sqref="B1:B9"><cfRule type="expression" dxfId="0" priority="3"><formula>$B1&gt;2</formula></cfRule></conditionalFormatting>`);
    // D (model 4, written 11) still ranks last: untouched.
    expect(out).toContain(blockD);
    // The data bar's twin is rewritten from the model with its GUID; the sparkline ext stays.
    expect(out).toContain(`<extLst>${cfExt(`<x14:conditionalFormatting>${twin}<xm:sqref>D1:D9</xm:sqref></x14:conditionalFormatting>`).replace(`xmlns:x14="${X14_NS}"`, `xmlns:x14="${X14_NS}" xmlns:xm="${XM_NS}"`)}${SPARKLINE_EXT}</extLst>`);
    const back = readPart(out);
    expect(rulesByPriority(back).map(({ rule }) => rule.kind)).toEqual(rulesByPriority(formats).map(({ rule }) => rule.kind));
    expect(back.map(shape)).toEqual(formats.map(shape));
  });

  it("deleting the last-ranked rule leaves the other blocks byte-verbatim and drops its twin's emptied ext, the sparkline ext kept", () => {
    const model = modelWith(sheetWith(readPart(base)), [RED]);
    deleteRule(model, "s", model.sheets[0]!.conditionalFormats[3]!.rules[0]!.id);
    const out = serializeWorksheet(model.sheets[0]!, base, context([RED]));
    expect(out).toContain(`<sheetData/>${blockA}${blockB}${blockC}\n  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst>${SPARKLINE_EXT}</extLst></worksheet>`);
  });

  it("writes a new data bar as a base rule plus an x14 twin sharing a fresh upper-case GUID, with the new bar's defaults", () => {
    const model = modelWith(sheetWith([]));
    addConditionalFormat(model, "s", [range("C2:C9")], { kind: "dataBar", bar: BAR });
    const write = () => writeConditionalFormats(model.sheets[0]!.conditionalFormats, "S", { prefix: "", theme: undefined, dxfs: [], numFmtId: () => 164 });
    const out = write();
    const guid = /<x14:id>([^<]*)<\/x14:id>/.exec(out.run)?.[1] ?? "";
    expect(guid).toMatch(GUID);
    expect(out.run).toBe(
      `<conditionalFormatting sqref="C2:C9"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>` +
        `<extLst><ext uri="${RULE_ID_EXT_URI}" xmlns:x14="${X14_NS}"><x14:id>${guid}</x14:id></ext></extLst></cfRule></conditionalFormatting>`,
    );
    expect(out.x14).toEqual([
      `<x14:conditionalFormatting><x14:cfRule type="dataBar" id="${guid}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>` +
        `<x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>C2:C9</xm:sqref></x14:conditionalFormatting>`,
    ]);
    // Each save of a rule without a GUID mints a fresh one.
    expect(/<x14:id>([^<]*)<\/x14:id>/.exec(write().run)?.[1]).not.toBe(guid);
  });

  it("regenerates the text, blank, error and time-period formulas per kind with A1 the origin; the LibreOffice-only year periods are expression rules", () => {
    const cases: [CfRuleDraft, string][] = [
      [{ kind: "text", operator: "contains", text: `a"b` }, `NOT(ISERROR(SEARCH("a""b",C3)))`],
      [{ kind: "text", operator: "notContains", text: "x" }, `ISERROR(SEARCH("x",C3))`],
      [{ kind: "text", operator: "beginsWith", text: "x" }, `LEFT(C3,LEN("x"))="x"`],
      [{ kind: "text", operator: "endsWith", text: "x" }, `RIGHT(C3,LEN("x"))="x"`],
      [{ kind: "blanks", not: false }, `LEN(TRIM(C3))=0`],
      [{ kind: "blanks", not: true }, `LEN(TRIM(C3))>0`],
      [{ kind: "errors", not: false }, `ISERROR(C3)`],
      [{ kind: "errors", not: true }, `NOT(ISERROR(C3))`],
      [{ kind: "timePeriod", period: "today" }, `FLOOR(C3,1)=TODAY()`],
      [{ kind: "timePeriod", period: "yesterday" }, `FLOOR(C3,1)=TODAY()-1`],
      [{ kind: "timePeriod", period: "tomorrow" }, `FLOOR(C3,1)=TODAY()+1`],
      [{ kind: "timePeriod", period: "last7Days" }, `AND(TODAY()-FLOOR(C3,1)<=6,FLOOR(C3,1)<=TODAY())`],
      [{ kind: "timePeriod", period: "thisWeek" }, `AND(TODAY()-ROUNDDOWN(C3,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(C3,0)-TODAY()<=7-WEEKDAY(TODAY()))`],
      [{ kind: "timePeriod", period: "lastWeek" }, `AND(TODAY()-ROUNDDOWN(C3,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(C3,0)<(WEEKDAY(TODAY())+7))`],
      [{ kind: "timePeriod", period: "nextWeek" }, `AND(ROUNDDOWN(C3,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(C3,0)-TODAY()<(15-WEEKDAY(TODAY())))`],
      [{ kind: "timePeriod", period: "thisMonth" }, `AND(MONTH(C3)=MONTH(TODAY()),YEAR(C3)=YEAR(TODAY()))`],
      [{ kind: "timePeriod", period: "lastMonth" }, `AND(MONTH(C3)=MONTH(EDATE(TODAY(),0-1)),YEAR(C3)=YEAR(EDATE(TODAY(),0-1)))`],
      [{ kind: "timePeriod", period: "nextMonth" }, `AND(MONTH(C3)=MONTH(EDATE(TODAY(),0+1)),YEAR(C3)=YEAR(EDATE(TODAY(),0+1)))`],
      [{ kind: "timePeriod", period: "thisYear" }, `YEAR(C3)=YEAR(TODAY())`],
      [{ kind: "timePeriod", period: "lastYear" }, `YEAR(C3)=YEAR(TODAY())-1`],
      [{ kind: "timePeriod", period: "nextYear" }, `YEAR(C3)=YEAR(TODAY())+1`],
    ];
    for (const [draft, formula] of cases) {
      expect(regeneratedFormula({ ...draft, id: "r", priority: 1 } as CfRule, "C3"), JSON.stringify(draft)).toBe(formula);
    }
    const model = modelWith(sheetWith([]), [RED]);
    addConditionalFormat(model, "s", [range("C3:C9"), range("A1")], { kind: "text", operator: "contains", text: `a"b`, dxfId: 0 });
    addConditionalFormat(model, "s", [range("C3:C9")], { kind: "timePeriod", period: "thisYear", dxfId: 0 });
    const { run } = writeConditionalFormats(model.sheets[0]!.conditionalFormats, "S", { prefix: "", theme: undefined, dxfs: [RED], numFmtId: () => 164 });
    expect(run).toBe(
      `<conditionalFormatting sqref="C3:C9 A1"><cfRule type="containsText" dxfId="0" priority="2" operator="containsText" text="a&quot;b"><formula>NOT(ISERROR(SEARCH("a""b",C3)))</formula></cfRule></conditionalFormatting>` +
        `<conditionalFormatting sqref="C3:C9"><cfRule type="expression" dxfId="0" priority="1"><formula>YEAR(C3)=YEAR(TODAY())</formula></cfRule></conditionalFormatting>`,
    );
    expect(readPart(part(run)).map(shape)).toEqual([
      { sqref: "C3:C9 A1", rules: [{ priority: 2, dxfId: 0, kind: "text", operator: "contains", text: `a"b` }] },
      { sqref: "C3:C9", rules: [{ priority: 1, dxfId: 0, kind: "expression", formula: "YEAR(C3)=YEAR(TODAY())" }] },
    ]);
  });

  it("puts the run after phoneticPr, before dataValidations, and the x14 ext first in the extLst — namespaces on the ext, every other ext verbatim", () => {
    const source = part(`<phoneticPr fontId="1"/><dataValidations count="0"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst>${DV_EXT}${SPARKLINE_EXT}</extLst>`);
    const model = modelWith(sheetWith(readPart(source)), [RED]);
    addConditionalFormat(model, "s", [range("A1:A9")], { kind: "cellIs", operator: "greaterThan", formulas: ["5"], dxfId: 0 });
    addConditionalFormat(model, "s", [range("F2:F9")], STARS);
    const out = serializeWorksheet(model.sheets[0]!, source, context([RED]));
    expect(out).toContain(`<phoneticPr fontId="1"/><conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting><dataValidations count="0"/>`);
    const ext = /<extLst><ext uri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}" xmlns:x14="[^"]+" xmlns:xm="[^"]+"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="iconSet" priority="1" id="([^"]+)">/.exec(out);
    expect(ext?.[1]).toMatch(GUID);
    expect(out).toContain(`</x14:conditionalFormattings></ext>${DV_EXT}${SPARKLINE_EXT}</extLst>`);
    const back = readPart(out);
    expect(back.map(withoutGuid)).toEqual(model.sheets[0]!.conditionalFormats.map(withoutGuid));
    expect(back[1]!.rules[0]!.x14Id).toBe(ext![1]);
  });

  it("drops the ext when nothing x14 is left — and the extLst when that was all it held", () => {
    const stars = (others: string) => part(`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst>${cfExt(x14Block("F2:F9", `<x14:cfRule type="iconSet" priority="1" id="{11111111-1111-1111-1111-111111111111}"><x14:iconSet iconSet="3Stars">${[0, 33, 67].map((at) => `<x14:cfvo type="percent"><xm:f>${at}</xm:f></x14:cfvo>`).join("")}</x14:iconSet></x14:cfRule>`))}${others}</extLst>`);
    for (const others of [SPARKLINE_EXT, ""]) {
      const source = stars(others);
      const model = modelWith(sheetWith(readPart(source)));
      expect(model.sheets[0]!.conditionalFormats).toHaveLength(1);
      deleteRule(model, "s", model.sheets[0]!.conditionalFormats[0]!.rules[0]!.id);
      const out = serializeWorksheet(model.sheets[0]!, source, context());
      expect(out).not.toContain("x14:conditionalFormattings");
      expect(out.endsWith(others === "" ? `footer="0.3"/></worksheet>` : `footer="0.3"/><extLst>${SPARKLINE_EXT}</extLst></worksheet>`)).toBe(true);
    }
  });

  it("writes with the part's own prefix when its main namespace is prefixed, an x14 cross-sheet rule with its dxf inline", () => {
    const source = `<x:worksheet xmlns:x="${MAIN}"><x:sheetData/><x:pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></x:worksheet>`;
    const model = modelWith(sheetWith([]), [RED]);
    addConditionalFormat(model, "s", [range("A1:A9")], { kind: "cellIs", operator: "greaterThan", formulas: ["5"], dxfId: 0 });
    addConditionalFormat(model, "s", [range("B1:B9")], { kind: "expression", formula: "B1>Lists!$A$1", dxfId: 0 });
    const out = serializeWorksheet(model.sheets[0]!, source, context([RED]));
    expect(out).toContain(`<x:conditionalFormatting sqref="A1:A9"><x:cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><x:formula>5</x:formula></x:cfRule></x:conditionalFormatting><x:pageMargins`);
    expect(out).toMatch(
      /<x:extLst><x:ext uri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}" xmlns:x14="[^"]+" xmlns:xm="[^"]+"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression" priority="1" id="\{[0-9A-F-]+\}"><xm:f>B1&gt;Lists!\$A\$1<\/xm:f><x14:dxf><x:font><x:color rgb="FF9C0006"\/><\/x:font><x:fill><x:patternFill><x:bgColor rgb="FFFFC7CE"\/><\/x:patternFill><\/x:fill><\/x14:dxf><\/x14:cfRule><xm:sqref>B1:B9<\/xm:sqref><\/x14:conditionalFormatting><\/x14:conditionalFormattings><\/x:ext><\/x:extLst><\/x:worksheet>$/,
    );
    const holder = { dxfs: [RED] };
    const back = readWorksheet(out, "S", { sharedStrings: [], dateSystem: "1900", budget: new CellBudget(), mdw: 7, dxfs: holder })!.conditionalFormats!;
    expect(back.map(withoutGuid)).toEqual([
      { sqref: "A1:A9", rules: [{ priority: 2, dxfId: 0, kind: "cellIs", operator: "greaterThan", formulas: ["5"] }] },
      { sqref: "B1:B9", rules: [{ priority: 1, dxfId: 0, x14: true, kind: "expression", formula: "B1>Lists!$A$1" }] },
    ]);
    expect(holder.dxfs).toEqual([RED]);
  });
});

describe("whole workbooks write valid conditional formats and dxfs", () => {
  it("a brand-new workbook: every rule and its dxf read back", async () => {
    const model = createBlankWorkbook();
    const sheetId = model.sheets[0]!.id;
    const green = internDxf(model, { font: { color: "#006100" }, fill: { kind: "solid", color: "#c6efce" } });
    const bold = internDxf(model, { font: { bold: true, underline: "none" }, numberFormat: "0.0%" });
    addConditionalFormat(model, sheetId, [range("B2:B13")], { kind: "cellIs", operator: "between", formulas: ["10", "$Z$1"], dxfId: green });
    addConditionalFormat(model, sheetId, [range("C2:C13")], { kind: "dataBar", bar: { ...BAR, gradient: false, negativeColor: "#ff00ff", axisColor: "#00ff00" } });
    addConditionalFormat(model, sheetId, [range("D2:D13")], { kind: "iconSet", icons: { set: "3Arrows", thresholds: [0, 33, 67].map((at) => ({ at: { type: "percent", value: String(at) }, gte: at === 0 })), reverse: true, showValue: true } });
    addConditionalFormat(model, sheetId, [range("E2:E13")], { kind: "colorScale", stops: [{ at: { type: "min" }, color: "#f8696b" }, { at: { type: "max" }, color: ACCENT6_DARKER }] });
    addConditionalFormat(model, sheetId, [range("F2:F13")], STARS);
    addConditionalFormat(model, sheetId, [range("G2:G13")], { kind: "expression", formula: "G2>Sheet9!$A$1", dxfId: bold });
    addConditionalFormat(model, sheetId, [range("H2:H13")], { kind: "top10", bottom: true, percent: true, rank: 15, dxfId: bold, stopIfTrue: true });
    // LibreOffice's smilies have no Excel form: written by name, as its own xlsx export does.
    for (const [set, column] of [["3Smilies", "I"], ["3ColorSmilies", "J"]] as const) {
      addConditionalFormat(model, sheetId, [range(`${column}2:${column}13`)], { kind: "iconSet", icons: { set, thresholds: [0, 33, 67].map((at) => ({ at: { type: "percent", value: String(at) }, gte: true })), reverse: false, showValue: true } });
    }
    const bytes = (await spreadsheetDownload(model)).bytes;
    expect(await packageIntegrity(bytes)).toEqual([]);
    const parts = await unzip(bytes);
    const styles = new TextDecoder().decode(parts.get("xl/styles.xml"));
    expect(styles).toContain(`<dxfs count="2">`);
    const sheet = new TextDecoder().decode(parts.get("xl/worksheets/sheet1.xml"));
    expect(sheet).toContain(`<conditionalFormatting sqref="I2:I13"><cfRule type="iconSet" priority="2"><iconSet iconSet="3Smilies"><cfvo type="percent" val="0"/>`);
    expect(sheet).toContain(`<iconSet iconSet="3ColorSmilies">`);
    const back = await readSpreadsheetBytes(bytes, "xlsx", "book");
    expect(rulesOf(back)).toEqual(rulesOf(model));
  });

  it("a stripped cross-format copy of spec129-cf.xlsx (no sources, no GUIDs, dxfs without their XML)", async () => {
    const model = await readSpreadsheetBytes(await buildSpec129CfXlsxBytes(), "xlsx", "book");
    const stripped = crossFormat(model);
    // Every rule but the preserved custom-icon one crosses.
    expect(ruleCount(stripped, () => true)).toBe(ruleCount(model, (rule) => rule.kind !== "preserved"));
    const bytes = (await spreadsheetDownload(stripped)).bytes;
    expect(await packageIntegrity(bytes)).toEqual([]);
    const back = await readSpreadsheetBytes(bytes, "xlsx", "book");
    expect(rulesOf(back)).toEqual(rulesOf(stripped));
  });
});
