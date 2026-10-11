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
// reversed order (no wire form there). A row may also yield item-level
// labels (SPEC-131's "Form control on Sheet1") by returning them from its
// marker.
import { sheetFormats } from "./conditionalFormats";
import { isChartSheet } from "./foreign";
import type { SpreadsheetModel } from "./model";
import { partTextOf, readSourceWorkbook, relationshipKind, relationshipsOf, targetPart, type PartLookup } from "./xlsx/package";

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

const WORKSHEET_TEXTS = new WeakMap<ConversionSource, string[]>();

/** The retained text of every grid sheet the model holds. */
function worksheetTexts(source: ConversionSource): string[] {
  let texts = WORKSHEET_TEXTS.get(source);
  if (!texts) {
    texts = source.model.sheets.flatMap((sheet) => {
      const text = sheet.kind === "grid" && sheet.partName !== undefined ? partTextOf(source.parts, sheet.partName) : undefined;
      return text === undefined ? [] : [text];
    });
    WORKSHEET_TEXTS.set(source, texts);
  }
  return texts;
}

function anyWorksheetHas(fragment: string): Marker {
  return (source) => worksheetTexts(source).some((text) => text.includes(fragment));
}

/** Whether a worksheet of the model has a relationship of `kind`. */
function worksheetRelationship(kind: string): Marker {
  return ({ model, parts }) =>
    model.sheets.some(
      (sheet) => sheet.partName !== undefined && relationshipsOf(parts, sheet.partName).some((rel) => relationshipKind(rel.type) === kind),
    );
}

function anyPart(pattern: RegExp): Marker {
  return ({ parts }) => [...parts.keys()].some((name) => pattern.test(name));
}

/** A worksheet drawing holding a picture or shape — not only charts, which
 * are their own row. */
const xlsxImagesAndShapes: Marker = ({ model, parts }) =>
  model.sheets.some((sheet) => {
    if (sheet.partName === undefined) {
      return false;
    }
    return relationshipsOf(parts, sheet.partName).some((rel) => {
      if (relationshipKind(rel.type) !== "drawing" || rel.external) {
        return false;
      }
      const drawing = partTextOf(parts, targetPart(sheet.partName!, rel) ?? "") ?? "";
      return /<(?:\w+:)?(?:pic|sp|grpSp|cxnSp)[\s>]/.test(drawing);
    });
  });

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

/** Shapes, or a frame that is not an embedded object's. */
const odsImagesAndShapes: Marker = (source) =>
  tableTexts(source).some((text) => {
    if (/<draw:(?:custom-shape|rect|line|ellipse|circle|path|polygon|polyline|connector|g)[\s>]/.test(text)) {
      return true;
    }
    const frames = text.split("<draw:frame").length - 1;
    const objects = text.split("<draw:object").length - 1;
    return frames > objects;
  });

const odsPageSetup: Marker = ({ model, parts }) => {
  const styles = partTextOf(parts, "styles.xml") ?? "";
  if (/style:print-orientation="landscape"|style:scale-to(?:-pages)?="/.test(styles)) {
    return true;
  }
  return model.sheets.some((sheet) => /<table:table\b[^>]*\btable:print-ranges="/.test(sheet.sourceXml ?? ""));
};

// --- both ---------------------------------------------------------------------------------------

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
  { label: "chart sheets", xlsx: ({ model }) => model.sheets.some((sheet) => isChartSheet(model, sheet)) },
  {
    label: "macros",
    xlsx: ({ model }) => model.packageVariant === "macroEnabled" || model.packageVariant === "macroEnabledTemplate",
    ods: anyPart(/^(?:Basic|Scripts)\//),
  },
  { label: "images and shapes", xlsx: xlsxImagesAndShapes, ods: odsImagesAndShapes },
  { label: "charts", xlsx: anyPart(/^xl\/charts\/chart[^/]*\.xml$/), ods: anyTableHas("<draw:object") },
  { label: "comments", xlsx: worksheetRelationship("comments"), ods: anyTableHas("<office:annotation") },
  { label: "some conditional formatting", xlsx: someConditionalFormatting(true), ods: someConditionalFormatting(false) },
  { label: "data validation", xlsx: anyWorksheetHas("<dataValidations"), ods: envelopeHas("<table:content-validations") },
  { label: "named ranges", xlsx: xlsxNamedRanges, ods: (source) => envelopeHas("<table:named-expressions")(source) || anyTableHas("<table:named-expressions")(source) },
  { label: "tables", xlsx: anyPart(/^xl\/tables\/[^/]+\.xml$/), ods: odsTables },
  { label: "colour and icon filters", xlsx: verbatimFilters, ods: verbatimFilters },
  { label: "hyperlinks", xlsx: anyWorksheetHas("<hyperlinks"), ods: anyTableHas("<text:a") },
  { label: "pivot tables", xlsx: anyPart(/^xl\/pivotTables\//), ods: envelopeHas("<table:data-pilot-tables") },
  { label: "page setup", xlsx: (source) => anyWorksheetHas("<pageSetup")(source) || anyWorksheetHas("<headerFooter")(source), ods: odsPageSetup },
  { label: "untranslatable formulas", xlsx: untranslatableFormulas, ods: untranslatableFormulas },
  { label: "custom date origin", ods: ({ model }) => model.odsNullDate !== undefined },
];
