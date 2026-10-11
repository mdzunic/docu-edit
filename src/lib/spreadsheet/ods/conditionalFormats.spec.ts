// SPEC-129 "ODS wire mapping": reading `calcext:conditional-formats` (every
// `calcext:value` form, colour scales, data bars, icon sets, date-is), the
// applied named styles as dxfs (ancestors short of Default, the source
// name kept, one per style), operands rebased from
// `calcext:base-cell-address`, legacy `style:map` conditions over cells,
// repeated runs and column defaults; writing in priority order, the
// generated `de-cf` styles spliced into an otherwise verbatim styles.xml,
// `xmlns:calcext` declared when absent, mapped cell styles re-emitted as
// `de-ce` styles without their maps, unedited formats verbatim, and a model
// without retention (another format's, a new workbook's) written valid.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { zip } from "../../core/zipWriter";
import { parseRange, type GridRange } from "../address";
import {
  addConditionalFormat,
  deleteRule,
  rulesByPriority,
  type CfRule,
  type CfRuleBody,
  type CfRuleDraft,
  type ConditionalFormat,
  type Dxf,
} from "../conditionalFormats";
import { buildSpec129CfOdsBytes, buildSpec129LegacyOdsBytes, SPEC129_CF_ODS, SPEC129_LEGACY_ODS } from "../fixtures/build";
import type { SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { createBlankWorkbook, stripRetention } from "../sheets";
import { internDxf } from "../styles";
import { odsCfReadContext, readOdsTableConditionalFormats } from "./conditionalFormats";
import { parseOds } from "./reader";
import { writeOds } from "./writer";

const CALCEXT = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ` +
  `xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"`;
const DECL = `<?xml version="1.0" encoding="UTF-8"?>`;

/** Common styles: Default sets a colour and italics no dxf may inherit. */
const STYLES_XML =
  `${DECL}<office:document-styles ${NS} office:version="1.3"><office:styles>` +
  `<style:default-style style:family="table-cell"><style:text-properties fo:font-size="10pt"/></style:default-style>` +
  `<number:number-style style:name="N2"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/></number:number-style>` +
  `<style:style style:name="Default" style:family="table-cell"><style:text-properties fo:color="#123456" fo:font-style="italic"/></style:style>` +
  `<style:style style:name="Accent" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold"/></style:style>` +
  `<style:style style:name="Accent_20_3" style:display-name="Accent 3" style:family="table-cell" style:parent-style-name="Accent">` +
  `<style:table-cell-properties fo:background-color="#dddddd" fo:border-bottom="0.74pt solid #ff0000"/></style:style>` +
  `<style:style style:name="Bad" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ffc7ce"/>` +
  `<style:text-properties fo:color="#9c0006"/></style:style>` +
  `<style:style style:name="Money" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N2">` +
  `<style:text-properties fo:font-weight="normal" style:text-underline-style="solid" style:text-underline-type="single" style:text-line-through-style="solid"/></style:style>` +
  `</office:styles></office:document-styles>`;

/** content.xml up to the first table: automatic cell styles with legacy maps. */
const PREFIX =
  `${DECL}<office:document-content ${NS} xmlns:calcext="${CALCEXT}" office:version="1.3"><office:automatic-styles>` +
  `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default">` +
  `<style:map style:condition="cell-content()&gt;=10" style:apply-style-name="Bad" style:base-cell-address="S.B2"/>` +
  `<style:map style:condition="cell-content-is-not-between(1,[.$Z$1])" style:apply-style-name="Accent_20_3" style:base-cell-address="S.B2"/></style:style>` +
  `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default">` +
  `<style:map style:condition="is-true-formula([.A3]&lt;&gt;&quot;&quot;)" style:apply-style-name="Bad" style:base-cell-address="S.A3"/></style:style>` +
  `<style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default"/>` +
  `</office:automatic-styles><office:body><office:spreadsheet>`;

const range = (a1: string): GridRange => parseRange(a1)!;

/** A table "S" ending with `trailing` (its conditional formats). */
function table(trailing: string, rows = `<table:table-row table:number-rows-repeated="1048576"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>`): string {
  return `\n<table:table table:name="S"><table:table-column table:number-columns-repeated="1024"/>${rows}${trailing}</table:table>`;
}

/** One format over S.B2:S.B10 holding `entries`. */
function oneFormat(entries: string, target = "S.B2:S.B10"): string {
  return `<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="${target}">${entries}</calcext:conditional-format></calcext:conditional-formats>`;
}

function read(tableXml: string): { formats: ConditionalFormat[]; dxfs: Dxf[] } {
  const holder = { dxfs: [] as Dxf[] };
  const formats = readOdsTableConditionalFormats(tableXml, odsCfReadContext(STYLES_XML, PREFIX), holder);
  return { formats, dxfs: holder.dxfs };
}

/** A rule's body: without its id, priority and dxf. */
function body(rule: CfRule): CfRuleBody {
  const { id: _id, priority: _priority, dxfId: _dxf, ...rest } = rule;
  return rest as CfRuleBody;
}

const condition = (value: string, base = "S.B2", style = "Bad"): string =>
  `<calcext:condition calcext:apply-style-name="${style}" calcext:value="${value}" calcext:base-cell-address="${base}"/>`;

describe("reading calcext:conditional-formats", () => {
  const forms: [string, CfRuleBody][] = [
    ["=5", { kind: "cellIs", operator: "equal", formulas: ["5"] }],
    ["!=5", { kind: "cellIs", operator: "notEqual", formulas: ["5"] }],
    ["&gt;5", { kind: "cellIs", operator: "greaterThan", formulas: ["5"] }],
    ["&gt;=5", { kind: "cellIs", operator: "greaterThanOrEqual", formulas: ["5"] }],
    ["&lt;5", { kind: "cellIs", operator: "lessThan", formulas: ["5"] }],
    ["&lt;=[.$Z$1]", { kind: "cellIs", operator: "lessThanOrEqual", formulas: ["$Z$1"] }],
    ["between(1,[.C2])", { kind: "cellIs", operator: "between", formulas: ["1", "C2"] }],
    ["not-between(&quot;a&quot;,&quot;b,c&quot;)", { kind: "cellIs", operator: "notBetween", formulas: [`"a"`, `"b,c"`] }],
    ["formula-is(ISEVEN([.B2]))", { kind: "expression", formula: "ISEVEN(B2)" }],
    ["duplicate", { kind: "duplicates", unique: false }],
    ["unique", { kind: "duplicates", unique: true }],
    ["top-elements(3)", { kind: "top10", bottom: false, percent: false, rank: 3 }],
    ["bottom-elements(4)", { kind: "top10", bottom: true, percent: false, rank: 4 }],
    ["top-percent(10)", { kind: "top10", bottom: false, percent: true, rank: 10 }],
    ["bottom-percent(25)", { kind: "top10", bottom: true, percent: true, rank: 25 }],
    ["above-average", { kind: "aboveAverage", below: false, equalAverage: false }],
    ["below-average", { kind: "aboveAverage", below: true, equalAverage: false }],
    ["above-equal-average", { kind: "aboveAverage", below: false, equalAverage: true }],
    ["below-equal-average", { kind: "aboveAverage", below: true, equalAverage: true }],
    ["begins-with(&quot;ab&quot;)", { kind: "text", operator: "beginsWith", text: "ab" }],
    ["ends-with(&quot;yz&quot;)", { kind: "text", operator: "endsWith", text: "yz" }],
    ["contains-text(&quot;say &quot;&quot;hi&quot;&quot;&quot;)", { kind: "text", operator: "contains", text: `say "hi"` }],
    ["not-contains-text(&quot;x&quot;)", { kind: "text", operator: "notContains", text: "x" }],
    ["is-error", { kind: "errors", not: false }],
    ["is-no-error", { kind: "errors", not: true }],
  ];

  it("every calcext:value form, one dxf for the style they all apply, priorities in document order", () => {
    const { formats, dxfs } = read(table(oneFormat(forms.map(([value]) => condition(value)).join(""))));
    expect(formats).toHaveLength(1);
    expect(formats[0]!.ranges).toEqual([range("B2:B10")]);
    expect(formats[0]!.rules.map(body)).toEqual(forms.map(([, expected]) => expected));
    expect(formats[0]!.rules.map((rule) => rule.priority)).toEqual(forms.map((_, index) => index + 1));
    expect(formats[0]!.rules.every((rule) => rule.dxfId === 0)).toBe(true);
    expect(dxfs).toHaveLength(1);
  });

  it("an unknown value form or entry is a preserved rule holding the element verbatim", () => {
    const odd = condition("frobnicate(1)");
    const element = `<calcext:sparkle calcext:x="1"><calcext:y/></calcext:sparkle>`;
    const { formats } = read(table(oneFormat(odd + element)));
    expect(formats[0]!.rules.map(body)).toEqual([
      { kind: "preserved", xml: odd, where: "ods" },
      { kind: "preserved", xml: element, where: "ods" },
    ]);
  });

  it("operands are rebased from calcext:base-cell-address onto the first range's top-left; value objects are absolute", () => {
    const entries =
      condition("formula-is([.B5]&gt;[.$Z$1])", "S.B5") +
      condition("&gt;[.C5]", "S.B5") +
      `<calcext:color-scale><calcext:color-scale-entry calcext:value="[.$Z$1]" calcext:type="formula" calcext:color="#000000"/>` +
      `<calcext:color-scale-entry calcext:value="0" calcext:type="maximum" calcext:color="#FFFFFF"/></calcext:color-scale>`;
    const { formats } = read(table(oneFormat(entries, "S.B2:S.B10 S.D2:S.D10")));
    expect(formats[0]!.ranges).toEqual([range("B2:B10"), range("D2:D10")]);
    expect(formats[0]!.rules.map(body)).toEqual([
      { kind: "expression", formula: "B2>$Z$1" },
      { kind: "cellIs", operator: "greaterThan", formulas: ["C2"] },
      {
        kind: "colorScale",
        stops: [
          { at: { type: "formula", value: "$Z$1" }, color: "#000000" },
          { at: { type: "max" }, color: "#ffffff" },
        ],
      },
    ]);
  });

  it("colour scales, data bars, icon sets and date-is", () => {
    const entries =
      `<calcext:color-scale><calcext:color-scale-entry calcext:value="0" calcext:type="minimum" calcext:color="#f8696b"/>` +
      `<calcext:color-scale-entry calcext:value="50" calcext:type="percentile" calcext:color="#ffeb84"/>` +
      `<calcext:color-scale-entry calcext:value="7" calcext:type="number" calcext:color="#63be7b"/></calcext:color-scale>` +
      `<calcext:data-bar calcext:positive-color="#638ec6" calcext:negative-color="#ff0000" calcext:axis-color="#000000" calcext:gradient="false" ` +
      `calcext:axis-position="middle" calcext:show-value="false" calcext:min-length="10" calcext:max-length="90">` +
      `<calcext:formatting-entry calcext:value="0" calcext:type="auto-minimum"/><calcext:formatting-entry calcext:value="0" calcext:type="auto-maximum"/></calcext:data-bar>` +
      `<calcext:data-bar calcext:positive-color="#638ec6"><calcext:formatting-entry calcext:value="0" calcext:type="minimum"/>` +
      `<calcext:formatting-entry calcext:value="90" calcext:type="percent"/></calcext:data-bar>` +
      `<calcext:icon-set calcext:icon-set-type="3Arrows" calcext:show-value="false"><calcext:formatting-entry calcext:value="0" calcext:type="percent"/>` +
      `<calcext:formatting-entry calcext:value="33" calcext:type="percent" calcext:gte="false"/><calcext:formatting-entry calcext:value="67" calcext:type="percent"/></calcext:icon-set>` +
      `<calcext:icon-set calcext:icon-set-type="3Smilies"><calcext:formatting-entry calcext:value="0" calcext:type="percent"/>` +
      `<calcext:formatting-entry calcext:value="33" calcext:type="percent"/></calcext:icon-set>` +
      `<calcext:date-is calcext:style="Bad" calcext:date="last-7-days"/><calcext:date-is calcext:style="Bad" calcext:date="next-month"/>`;
    const { formats } = read(table(oneFormat(entries)));
    const rules = formats[0]!.rules;
    expect(rules.map(body)).toEqual([
      {
        kind: "colorScale",
        stops: [
          { at: { type: "min" }, color: "#f8696b" },
          { at: { type: "percentile", value: "50" }, color: "#ffeb84" },
          { at: { type: "num", value: "7" }, color: "#63be7b" },
        ],
      },
      {
        kind: "dataBar",
        bar: {
          min: { type: "autoMin" },
          max: { type: "autoMax" },
          color: "#638ec6",
          negativeColor: "#ff0000",
          axisColor: "#000000",
          gradient: false,
          axis: "middle",
          direction: "context",
          minLengthPct: 10,
          maxLengthPct: 90,
          showValue: false,
        },
      },
      {
        kind: "dataBar",
        bar: {
          min: { type: "min" },
          max: { type: "percent", value: "90" },
          color: "#638ec6",
          gradient: true,
          axis: "automatic",
          direction: "context",
          minLengthPct: 0,
          maxLengthPct: 100,
          showValue: true,
        },
      },
      {
        kind: "iconSet",
        icons: {
          set: "3Arrows",
          thresholds: [
            { at: { type: "percent", value: "0" }, gte: true },
            { at: { type: "percent", value: "33" }, gte: false },
            { at: { type: "percent", value: "67" }, gte: true },
          ],
          reverse: false,
          showValue: false,
        },
      },
      // Two thresholds for a three-icon set: not modelled, carried.
      expect.objectContaining({ kind: "preserved", where: "ods" }),
      { kind: "timePeriod", period: "last7Days" },
      { kind: "timePeriod", period: "nextMonth" },
    ]);
    expect(rules[5]!.dxfId).toBe(0);
    expect(rules[0]!.dxfId).toBeUndefined();
  });

  it("an applied style resolves to a dxf from its own and its ancestors' properties short of Default, its source name kept", () => {
    const entries = condition("=1", "S.B2", "Bad") + condition("=2", "S.B2", "Accent 3") + condition("=3", "S.B2", "Money") + condition("=4", "S.B2", "Accent_20_3");
    const { formats, dxfs } = read(table(oneFormat(entries)));
    expect(formats[0]!.rules.map((rule) => rule.dxfId)).toEqual([0, 1, 2, 1]);
    expect(dxfs[0]).toEqual({ font: { color: "#9c0006" }, fill: { kind: "solid", color: "#ffc7ce" }, sourceStyleName: "Bad" });
    expect(dxfs[1]).toMatchObject({ font: { bold: true }, fill: { kind: "solid", color: "#dddddd" }, sourceStyleName: "Accent_20_3" });
    expect(dxfs[1]!.borders?.bottom?.color).toBe("#ff0000");
    expect(dxfs[1]!.font).toEqual({ bold: true });
    expect(dxfs[2]).toEqual({ font: { bold: false, underline: "single", strike: true }, numberFormat: "0.00", sourceStyleName: "Money" });
  });

  it("each format keeps its source (preceding whitespace included) and priorities run across the table's formats", () => {
    const first = `\n  <calcext:conditional-format calcext:target-range-address="S.A1:S.A3">${condition("=1")}${condition("=2")}</calcext:conditional-format>`;
    const second = `\n  <calcext:conditional-format calcext:target-range-address="S.C1">${condition("=3")}</calcext:conditional-format>`;
    const { formats } = read(table(`<calcext:conditional-formats>${first}${second}\n</calcext:conditional-formats>`));
    expect(formats.map((format) => format.sourceXml)).toEqual([first, second]);
    expect(formats.map((format) => format.rules.map((rule) => rule.priority))).toEqual([[1, 2], [3]]);
    expect(formats[1]!.ranges).toEqual([range("C1")]);
  });

  it("a container the table does not own (inside a cell's annotation text) is not read", () => {
    const rows = `<table:table-row><table:table-cell><office:annotation><text:p>calcext:conditional-formats</text:p></office:annotation></table:table-cell></table:table-row>`;
    expect(read(table("", rows)).formats).toEqual([]);
  });
});

describe("reading legacy style:map conditions", () => {
  // Column C's default is ce2; B2 and the 1000-row repeated run below it use ce1.
  const columns = `<table:table-column table:number-columns-repeated="2"/><table:table-column table:default-cell-style-name="ce2"/><table:table-column table:number-columns-repeated="1021"/>`;
  const styledRow = `<table:table-cell/><table:table-cell table:style-name="ce1"/><table:table-cell table:number-columns-repeated="1022"/>`;
  const rows =
    `<table:table-row><table:table-cell table:number-columns-repeated="1024"/></table:table-row>` +
    `<table:table-row>${styledRow}</table:table-row>` +
    `<table:table-row table:number-rows-repeated="1000">${styledRow}</table:table-row>` +
    `<table:table-row table:number-rows-repeated="1048574"><table:table-cell table:style-name="ce3"/><table:table-cell table:number-columns-repeated="1023"/></table:table-row>`;
  const legacyTable = `\n<table:table table:name="S">${columns}${rows}</table:table>`;

  it("a table with no calcext rules reads its cells' maps as cellIs and expression rules over row-run × column-run rectangles", () => {
    const { formats, dxfs } = read(legacyTable);
    expect(formats.map((format) => format.ranges)).toEqual([[range("C1:C1048576")], [range("B2:B1002")]]);
    expect(formats.map((format) => format.rules.map(body))).toEqual([
      [{ kind: "expression", formula: `C1<>""` }],
      [
        { kind: "cellIs", operator: "greaterThanOrEqual", formulas: ["10"] },
        { kind: "cellIs", operator: "notBetween", formulas: ["1", "$Z$1"] },
      ],
    ]);
    expect(formats.flatMap((format) => format.rules.map((rule) => rule.priority))).toEqual([1, 2, 3]);
    expect(formats.flatMap((format) => format.rules.map((rule) => dxfs[rule.dxfId!]!.sourceStyleName))).toEqual(["Bad", "Bad", "Accent_20_3"]);
    expect(formats.every((format) => format.sourceXml === undefined)).toBe(true);
  });

  it("a row's default cell style applies to its unstyled cells", () => {
    const rowDefault = `\n<table:table table:name="S">${columns}<table:table-row table:default-cell-style-name="ce1" table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="2"/><table:table-cell table:style-name="ce3" table:number-columns-repeated="1022"/></table:table-row></table:table>`;
    expect(read(rowDefault).formats.map((format) => format.ranges)).toEqual([[range("A1:B3")]]);
  });

  it("cells a row omits take the column defaults", () => {
    const short = `\n<table:table table:name="S">${columns}<table:table-row table:number-rows-repeated="5"/><table:table-row><table:table-cell/></table:table-row></table:table>`;
    expect(read(short).formats.map((format) => format.ranges)).toEqual([[range("C1:C6")]]);
  });

  it("calcext rules win: the maps are not read beside them", () => {
    const both = legacyTable.replace("</table:table>", `${oneFormat(condition("=1"))}</table:table>`);
    expect(read(both).formats.map((format) => format.ranges)).toEqual([[range("B2:B10")]]);
  });
});

describe("reading one retained table (SPEC-129's reader upgrade over a dirty sheet)", () => {
  it("the envelope around a sheet's sourceXml, with the retained styles.xml, yields its formats and the dxfs they apply", async () => {
    const source = await buildSpec129CfOdsBytes();
    const model = await readSpreadsheetBytes(source, "ods", "book");
    const parts = await unzip(source);
    const envelope = model.odsContentEnvelope!;
    parts.set("content.xml", new TextEncoder().encode(envelope.prefix + model.sheets[1]!.sourceXml! + envelope.suffix));
    const parsed = await parseOds(new Uint8Array(0), parts);
    expect(parsed.sheets.map((sheet) => sheet.name)).toEqual(["More"]);
    const formats = parsed.sheets[0]!.conditionalFormats;
    expect(formats.map((format) => format.sourceXml)).toEqual(SPEC129_CF_ODS.moreFormats);
    expect(formats.flatMap((format) => format.rules.map((rule) => parsed.dxfs[rule.dxfId!]!.sourceStyleName))).toEqual(["Good", "Accent_20_3", "Bad", "Bad", "Accent_20_3", "Good"]);
    expect(parsed.dxfs).toEqual([SPEC129_CF_ODS.dxfs[2], SPEC129_CF_ODS.dxfs[1], SPEC129_CF_ODS.dxfs[0]]);
  });
});

// --- writing ---------------------------------------------------------------------------------

const decoder = new TextDecoder();
const MODIFIED = new Date(Date.UTC(2026, 9, 11, 12, 0, 0));
type Parts = Map<string, Uint8Array>;

async function open(bytes: Uint8Array): Promise<{ source: Parts; model: SpreadsheetModel }> {
  return { source: await unzip(bytes), model: await readSpreadsheetBytes(bytes, "ods", "book") };
}

async function save(model: SpreadsheetModel): Promise<{ bytes: Uint8Array; parts: Parts; content: string; styles: string }> {
  const bytes = await writeOds(model, { modified: MODIFIED });
  const parts = await unzip(bytes);
  return { bytes, parts, content: decoder.decode(parts.get("content.xml")), styles: decoder.decode(parts.get("styles.xml")) };
}

function markDirty(model: SpreadsheetModel, index: number): void {
  model.sheets = model.sheets.map((sheet, at) => (at === index ? { ...sheet, dirty: true } : sheet));
}

/** A table's text in a written content.xml. */
function tableText(content: string, name: string): string {
  const start = content.indexOf(`<table:table table:name="${name}"`);
  return content.slice(start, content.indexOf("</table:table>", start) + "</table:table>".length);
}

/** What formats read back as: no ids, no source, in priority order (ODS
 * numbers priorities by document order, so a save writes them that way —
 * a format a command appended to the array with priority 1 comes first). */
function comparable(formats: readonly ConditionalFormat[]) {
  return formats.toSorted((a, b) => firstPriority(a) - firstPriority(b)).map(({ ranges, rules }) => ({ ranges, rules: rules.map(({ id: _id, ...rule }) => rule) }));
}

function firstPriority(format: ConditionalFormat): number {
  return Math.min(...format.rules.map((rule) => rule.priority));
}

/** A dxf as compared across a save: modelled fields only. */
function modelled(dxf: Dxf | undefined): Dxf | undefined {
  if (dxf === undefined) {
    return undefined;
  }
  const { sourceXml: _xml, sourceStyleName: _name, ...rest } = dxf;
  return rest;
}

/** Each rule's dxf (modelled), in priority order. */
function appliedDxfs(model: SpreadsheetModel, formats: readonly ConditionalFormat[]): (Dxf | undefined)[] {
  return rulesByPriority(formats).map(({ rule }) => (rule.dxfId === undefined ? undefined : modelled(model.dxfs[rule.dxfId])));
}

/** The formats with their rules' dxfIds dropped (the tables differ). */
function withoutDxfIds(formats: ReturnType<typeof comparable>) {
  return formats.map((format) => ({ ...format, rules: format.rules.map(({ dxfId: _dxf, ...rule }) => rule) }));
}

function expectSameRules(back: SpreadsheetModel, model: SpreadsheetModel, sheetIndex: number): void {
  const written = model.sheets[sheetIndex]!.conditionalFormats;
  const reread = back.sheets[sheetIndex]!.conditionalFormats;
  expect(withoutDxfIds(comparable(reread))).toEqual(withoutDxfIds(comparable(written)));
  expect(appliedDxfs(back, reread)).toEqual(appliedDxfs(model, written));
}

const scoresDraft = (dxfId: number): CfRuleDraft => ({ kind: "cellIs", operator: "greaterThan", formulas: ["120"], dxfId });

describe("writing", () => {
  it("an untouched workbook is every part byte-identical", async () => {
    for (const build of [buildSpec129CfOdsBytes, buildSpec129LegacyOdsBytes]) {
      const { source, model } = await open(await build());
      const { parts } = await save(model);
      expect([...parts.keys()]).toEqual([...source.keys()]);
      for (const [name, bytes] of source) {
        expect(decoder.decode(parts.get(name)), name).toBe(decoder.decode(bytes));
      }
    }
  });

  it("a dirty table whose rules were not edited re-emits each format's source verbatim, styles.xml untouched", async () => {
    const { source, model } = await open(await buildSpec129CfOdsBytes());
    markDirty(model, 0);
    markDirty(model, 1);
    const { content, styles } = await save(model);
    expect(tableText(content, "Scores")).toContain(`<calcext:conditional-formats>${SPEC129_CF_ODS.scoresFormats.join("")}</calcext:conditional-formats></table:table>`);
    expect(tableText(content, "More")).toContain(`<calcext:conditional-formats>${SPEC129_CF_ODS.moreFormats.join("")}</calcext:conditional-formats></table:table>`);
    expect(styles).toBe(decoder.decode(source.get("styles.xml")));
  });

  it("formats are written in priority order; a new dxf becomes a de-cf named style spliced before </office:styles>", async () => {
    const { source, model } = await open(await buildSpec129CfOdsBytes());
    const scores = model.sheets[0]!;
    const dxf: Dxf = { font: { bold: true, color: "#006100" }, fill: { kind: "solid", color: "#c6efce" }, numberFormat: "0.0%" };
    const dxfId = internDxf(model, dxf);
    addConditionalFormat(model, scores.id, [range("B2:B13")], scoresDraft(dxfId));
    const { content, styles } = await save(model);

    const scoresXml = tableText(content, "Scores");
    const first = scoresXml.indexOf("<calcext:conditional-format ");
    expect(scoresXml.slice(first)).toMatch(
      /^<calcext:conditional-format calcext:target-range-address="Scores\.B2:Scores\.B13"><calcext:condition calcext:apply-style-name="de-cf1" calcext:value="&gt;120" calcext:base-cell-address="Scores\.B2"\/><\/calcext:conditional-format>/,
    );
    // The source's own styles are kept by name.
    expect(scoresXml).toContain(`calcext:apply-style-name="Bad"`);
    expect(scoresXml).not.toContain("style:map");

    const sourceStyles = decoder.decode(source.get("styles.xml"));
    const at = sourceStyles.lastIndexOf("</office:styles>");
    expect(styles.slice(0, at)).toBe(sourceStyles.slice(0, at));
    expect(styles.endsWith(sourceStyles.slice(at))).toBe(true);
    const added = styles.slice(at, styles.length - (sourceStyles.length - at));
    expect(added).toContain(
      `<style:style style:name="de-cf1" style:display-name="Conditional 1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="de-N1">`,
    );
    expect(added).toMatch(/^<number:percentage-style style:name="de-N1"/);

    const back = await readSpreadsheetBytes((await save(model)).bytes, "ods", "book");
    expect(comparable(back.sheets[0]!.conditionalFormats).map((format) => format.rules.map((rule) => rule.priority))).toEqual(
      comparable(model.sheets[0]!.conditionalFormats).map((format) => format.rules.map((rule) => rule.priority)),
    );
    expectSameRules(back, model, 0);
    expectSameRules(back, model, 1);
    const reread = back.dxfs[back.sheets[0]!.conditionalFormats[0]!.rules[0]!.dxfId!]!;
    expect(modelled(reread)).toEqual(dxf);
    expect(reread.sourceStyleName).toBe("de-cf1");
  });

  it("deleting every rule of a table removes its container", async () => {
    const { model } = await open(await buildSpec129CfOdsBytes());
    for (const format of model.sheets[1]!.conditionalFormats) {
      for (const rule of format.rules) {
        deleteRule(model, model.sheets[1]!.id, rule.id);
      }
    }
    const { content } = await save(model);
    expect(tableText(content, "More")).not.toContain("conditional-format");
    expect(tableText(content, "Scores")).toContain("<calcext:conditional-formats>");
  });

  it("a rewritten legacy table writes calcext rules and re-emits its mapped cell styles as de-ce styles without style:map", async () => {
    const { model } = await open(await buildSpec129LegacyOdsBytes());
    expect(model.sheets[0]!.conditionalFormats.map((format) => format.ranges)).toEqual(SPEC129_LEGACY_ODS.formats.map((format) => format.ranges));
    markDirty(model, 0);
    const { content, bytes } = await save(model);
    const legacy = tableText(content, "Legacy");
    expect(legacy).not.toMatch(/style-name="ce[67]"/);
    expect(legacy).toContain("<calcext:conditional-formats>");
    // Only the source's own (now unused) automatic styles still carry maps.
    const generated = [...content.matchAll(/<style:style style:name="de-ce\d+"[\s\S]*?(?:\/>|<\/style:style>)/g)].map((match) => match[0]);
    expect(generated.length).toBeGreaterThan(0);
    expect(generated.some((style) => style.includes("style:map"))).toBe(false);

    const back = await readSpreadsheetBytes(bytes, "ods", "book");
    expectSameRules(back, model, 0);
  });

  it("xmlns:calcext is declared on the content root when the source lacks it", async () => {
    const content =
      `${DECL}<office:document-content ${NS} office:version="1.2"><office:automatic-styles/><office:body><office:spreadsheet>` +
      `<table:table table:name="S"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>` +
      `<table:table-cell/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const { model } = await open(await odsPackage(content));
    expect(content).not.toContain("xmlns:calcext");
    const dxfId = internDxf(model, { fill: { kind: "solid", color: "#ffc7ce" } });
    addConditionalFormat(model, model.sheets[0]!.id, [range("A1:A2")], scoresDraft(dxfId));
    const written = await save(model);
    expect(/<office:document-content\b[^>]*>/.exec(written.content)![0]).toContain(`xmlns:calcext="${CALCEXT}"`);
    const back = await readSpreadsheetBytes(written.bytes, "ods", "book");
    expectSameRules(back, model, 0);
  });

  it("the calcext namespace bound under another prefix reads, and its table's rules write back", async () => {
    const content =
      `${DECL}<office:document-content ${NS} xmlns:ce="${CALCEXT}" office:version="1.2"><office:automatic-styles/><office:body><office:spreadsheet>` +
      `<table:table table:name="S"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>` +
      `<table:table-cell/></table:table-row><ce:conditional-formats><ce:conditional-format ce:target-range-address="S.A1:S.A2">` +
      `<ce:condition ce:apply-style-name="Bad" ce:value="&lt;3" ce:base-cell-address="S.A1"/></ce:conditional-format></ce:conditional-formats></table:table>` +
      `</office:spreadsheet></office:body></office:document-content>`;
    const { model } = await open(await odsPackage(content));
    expect(comparable(model.sheets[0]!.conditionalFormats)).toEqual([
      { ranges: [range("A1:A2")], rules: [{ priority: 1, dxfId: 0, kind: "cellIs", operator: "lessThan", formulas: ["3"] }] },
    ]);
    addConditionalFormat(model, model.sheets[0]!.id, [range("B1:B2")], scoresDraft(0));
    const back = await readSpreadsheetBytes((await save(model)).bytes, "ods", "book");
    expectSameRules(back, model, 0);
  });

  it("a model without retention (another format's copy) writes valid conditional formats from scratch", async () => {
    const { model: source } = await open(await buildSpec129CfOdsBytes());
    const stripped = stripRetention(source);
    stripped.sheets = stripped.sheets.map((sheet) => ({ ...sheet, conditionalFormats: sheet.conditionalFormats.map(({ sourceXml: _source, ...format }) => format) }));
    stripped.dxfs = stripped.dxfs.map((dxf) => modelled(dxf)!);
    const { content, styles, bytes } = await save(stripped);
    expect(content).toContain(`xmlns:calcext="${CALCEXT}"`);
    expect(styles).toContain(`<style:style style:name="de-cf1" style:display-name="Conditional 1" style:family="table-cell" style:parent-style-name="Default">`);
    const back = await readSpreadsheetBytes(bytes, "ods", "book");
    expectSameRules(back, stripped, 0);
    expectSameRules(back, stripped, 1);
  });

  it("a new workbook writes every modelled kind; blanks and a stdDev average go as formula-is, reversed icons and stopIfTrue have no wire form", async () => {
    const model = createBlankWorkbook();
    const sheet = model.sheets[0]!;
    const dxfId = internDxf(model, { font: { italic: true, underline: "double" }, borders: { top: { style: "thin", color: "#ff0000" } } });
    const drafts: CfRuleDraft[] = [
      { kind: "blanks", not: false, dxfId },
      { kind: "aboveAverage", below: true, equalAverage: false, stdDev: 2, dxfId },
      { kind: "iconSet", icons: { set: "4Rating", thresholds: [0, 25, 50, 75].map((value) => ({ at: { type: "percent", value: String(value) }, gte: true })), reverse: true, showValue: true } },
      { kind: "cellIs", operator: "between", formulas: ["B$1", "10"], dxfId, stopIfTrue: true },
    ];
    for (const draft of drafts) {
      addConditionalFormat(model, sheet.id, [range("B2:C5")], draft);
    }
    const { content, bytes } = await save(model);
    expect(content).toContain(`calcext:value="formula-is(LEN(TRIM([.B2]))=0)"`);
    expect(content).toContain(`calcext:value="formula-is(AND(ISNUMBER([.B2]);[.B2]&lt;AVERAGE([.$B$2:.$C$5])-2*STDEVP([.$B$2:.$C$5])))"`);
    const back = await readSpreadsheetBytes(bytes, "ods", "book");
    const rules = back.sheets[0]!.conditionalFormats.flatMap((format) => format.rules);
    expect(rules.map((rule) => rule.kind)).toEqual(["cellIs", "iconSet", "expression", "expression"]);
    expect(rules.map((rule) => rule.priority)).toEqual([1, 2, 3, 4]);
    expect(body(rules[0]!)).toEqual({ kind: "cellIs", operator: "between", formulas: ["B$1", "10"] });
    expect(rules[1]!.kind === "iconSet" && rules[1]!.icons.reverse).toBe(false);
    expect(body(rules[3]!)).toEqual({ kind: "expression", formula: "LEN(TRIM(B2))=0" });
    expect(modelled(back.dxfs[rules[0]!.dxfId!])).toEqual(model.dxfs[dxfId]);
  });
});

/** A minimal ODS package: content.xml and STYLES_XML (Bad and friends). */
function odsPackage(content: string): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const manifest =
    `${DECL}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
    `</manifest:manifest>`;
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(STYLES_XML) },
    { name: "META-INF/manifest.xml", data: encoder.encode(manifest) },
  ]);
}
