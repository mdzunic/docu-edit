// What a cross-format copy leaves out (SPEC-121 "Cross-format copy"):
// UNCONVERTIBLE_CONTENT is the normative table — each row a label and, per
// source format, a marker that finds the structure in the source workbook.
// sheets.ts's unconvertibleContent returns the labels whose marker is
// present, in table order; the download names them in one toast ("Not
// included in the .ods copy: charts, comments").
//
// Every row is a structure the model does not carry yet — or carries only
// verbatim, in its source format's terms — so stripRetention's copy cannot
// either. Each later spec that models a structure deletes its row when the
// structure crosses through the model (the "Row removed by" column of the
// spec's table): SPEC-128 deleted "filters" — an autofilter's range,
// criteria and sort cross — and added "colour and icon filters", the
// criteria it keeps verbatim (an xlsx `preserved` criterion, an ods
// `odsUnmappedFilter`). SPEC-129 replaced "conditional formatting" — rules,
// their ranges and differential styles cross — with "some conditional
// formatting", what its stripRetention drops: a `preserved` rule (another
// format's XML, in either direction) and, going to .ods, an icon set's
// reversed order (no wire form there). SPEC-130 deleted "data validation":
// every rule type crosses through the model (stripRetention's
// crossValidations; an ODS sort-ascending list becomes an ordinary xlsx
// list, the one property that does not cross, and is not named here).
// SPEC-131 replaced "images and shapes" — image and shape drawings cross
// through the model's canonical writers — with its item-level row: each
// drawing stripRetention drops (a `foreign` or `group` one) is named with
// its kind and sheet ("Form control on Sheet1"), once per distinct label; a
// placeholder `chart` drawing is dropped too, and named by the "charts" row
// (SPEC-121's toast reads "charts, comments"), not twice. SPEC-132 deleted
// the "charts" and "chart sheets" rows: a catalogue chart crosses through
// the model (stripRetention's crossDrawings; a chart sheet becomes, in an
// .ods, a grid sheet holding its chart), and each chart a copy drops — out
// of the catalogue, read-only, or an Excel chartex chart (a `foreign` "Chart"
// drawing) — is named on its own, `Chart "Chart 5" on Data`, as is a chart
// sheet left a foreign placeholder (`Chart sheet "Chart1"`). A row yields
// item-level labels by returning them from its marker. SPEC-135 deleted the
// "comments" and "hyperlinks" rows: notes, threads and links cross through
// the model (stripRetention's crossAnnotations — a thread becomes a note on
// the way to .ods, its replies as replyParagraphText paragraphs). SPEC-136
// deleted "page setup": a sheet's setup crosses field for field
// (stripRetention keeps `pageSetup`, drops `pageSetupDirty`), its print area
// and titles as the reserved names both formats map.
import { isModelledChart } from "./charts";
import { sheetFormats } from "./conditionalFormats";
import { isChartSheet } from "./foreign";
import type { SheetDrawing, SpreadsheetModel } from "./model";
import { readSourceWorkbook, type PartLookup } from "./xlsx/package";

/** What a marker looks at: the source model and its retained parts. */
export interface ConversionSource {
  model: SpreadsheetModel;
  parts: PartLookup;
}

/** True when the structure is present, or the item-level labels it yields. */
export type Marker = (source: ConversionSource) => boolean | string[];

export interface UnconvertibleRow {
  label: string;
  xlsx?: Marker;
  ods?: Marker;
}

// --- xlsx --------------------------------------------------------------------------------

function anyPart(pattern: RegExp): Marker {
  return ({ parts }) => [...parts.keys()].some((name) => pattern.test(name));
}

const xlsxNamedRanges: Marker = ({ parts }) => {
  const workbook = readSourceWorkbook(parts);
  return workbook !== undefined && [...workbook.xml.matchAll(/<(?:\w+:)?definedName\b[^>]*\bname="([^"]*)"/g)].some((match) => !match[1]!.startsWith("_xlnm."));
};

// --- ods --------------------------------------------------------------------------------------

function tableTexts({ model }: ConversionSource): string[] {
  return model.sheets.map((sheet) => sheet.sourceXml ?? "");
}

function anyTableHas(...fragments: string[]): Marker {
  return (source) => tableTexts(source).some((text) => fragments.some((fragment) => text.includes(fragment)));
}

function envelopeHas(fragment: string): Marker {
  return ({ model }) => (model.odsContentEnvelope?.prefix ?? "").includes(fragment) || (model.odsContentEnvelope?.suffix ?? "").includes(fragment);
}

/** A database range no sheet's autofilter claims (SPEC-128: the one a sheet's
 * autofilter is read from crosses through the model). */
const odsTables: Marker = ({ model }) => {
  const claimed = new Set(model.sheets.flatMap((sheet) => (sheet.autoFilter?.odsName === undefined ? [] : [sheet.autoFilter.odsName])));
  const text = (model.odsContentEnvelope?.prefix ?? "") + (model.odsContentEnvelope?.suffix ?? "");
  return [...text.matchAll(/<table:database-range\b[^>]*?\btable:name="([^"]*)"/g)].some((match) => !claimed.has(match[1]!));
};

// --- both ---------------------------------------------------------------------------------------

/** SPEC-131: what a drawing the copy drops is called ("Form control",
 * "Group"), or undefined for one that crosses. SPEC-132: a chart the copy
 * drops — out of the catalogue, read-only or chartex — by its own name. */
function droppedDrawingKind(drawing: SheetDrawing): string | undefined {
  switch (drawing.content.kind) {
    case "foreign":
      return drawing.content.label === "Chart" ? `Chart "${drawing.name}"` : drawing.content.label;
    case "group":
      return "Group";
    case "chart":
      return isModelledChart(drawing.content) && drawing.content.readOnlyReason === undefined ? undefined : `Chart "${drawing.name}"`;
    default:
      return undefined;
  }
}

/** SPEC-131: every drawing the copy drops, as "<kind> on <sheet>", once per
 * distinct label, in sheet and z-order; SPEC-132: then each chart sheet left
 * a foreign placeholder. */
const droppedDrawings: Marker = ({ model }) => {
  const labels = new Set<string>();
  for (const sheet of model.sheets) {
    if (sheet.kind === "foreign") {
      if (isChartSheet(model, sheet)) {
        labels.add(`Chart sheet "${sheet.name}"`);
      }
      continue;
    }
    for (const drawing of sheet.drawings ?? []) {
      const kind = droppedDrawingKind(drawing);
      if (kind !== undefined) {
        labels.add(`${kind} on ${sheet.name}`);
      }
    }
  }
  return [...labels];
};

/** SPEC-128: an autofilter criterion kept verbatim — an xlsx colour or icon
 * filter (`preserved`), an ods filter that does not decompose per column. */
const verbatimFilters: Marker = ({ model }) =>
  model.sheets.some((sheet) => sheet.autoFilter !== undefined && (sheet.autoFilter.odsUnmappedFilter !== undefined || sheet.autoFilter.columns.some((column) => column.criteria?.kind === "preserved")));

/** SPEC-129: a conditional-format rule the copy drops — a `preserved` rule,
 * or (`toOds`) an icon set in reversed order. */
function someConditionalFormatting(toOds: boolean): Marker {
  return ({ model }) =>
    model.sheets.some((sheet) =>
      sheetFormats(sheet).some((format) => format.rules.some((rule) => rule.kind === "preserved" || (toOds && rule.kind === "iconSet" && rule.icons.reverse))),
    );
}

/** A formula that cannot cross: unparsed text or a data table. */
const untranslatableFormulas: Marker = ({ model }) =>
  model.sheets.some((sheet) => sheet.rows.some((row) => row.cells.some((cell) => cell.formulaUnparsed === true || cell.dataTableXml !== undefined)));

/** SPEC-121's table, in its order. */
export const UNCONVERTIBLE_CONTENT: readonly UnconvertibleRow[] = [
  {
    label: "macros",
    xlsx: ({ model }) => model.packageVariant === "macroEnabled" || model.packageVariant === "macroEnabledTemplate",
    ods: anyPart(/^(?:Basic|Scripts)\//),
  },
  { label: "drawings", xlsx: droppedDrawings, ods: droppedDrawings },
  { label: "some conditional formatting", xlsx: someConditionalFormatting(true), ods: someConditionalFormatting(false) },
  { label: "named ranges", xlsx: xlsxNamedRanges, ods: (source) => envelopeHas("<table:named-expressions")(source) || anyTableHas("<table:named-expressions")(source) },
  { label: "tables", xlsx: anyPart(/^xl\/tables\/[^/]+\.xml$/), ods: odsTables },
  { label: "colour and icon filters", xlsx: verbatimFilters, ods: verbatimFilters },
  { label: "pivot tables", xlsx: anyPart(/^xl\/pivotTables\//), ods: envelopeHas("<table:data-pilot-tables") },
  { label: "untranslatable formulas", xlsx: untranslatableFormulas, ods: untranslatableFormulas },
  { label: "custom date origin", ods: ({ model }) => model.odsNullDate !== undefined },
];
