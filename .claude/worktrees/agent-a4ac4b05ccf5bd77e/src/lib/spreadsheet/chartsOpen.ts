// Sheet charts at open and at a reader upgrade (SPEC-132 "Open trusts
// caches"; open.ts's version 13) — loaded on demand by open.ts, which must
// not pay for either format's chart codec until a workbook holds a chart.
//
// - refreshUncachedCharts: open trusts a chart part's caches exactly as
//   SPEC-120 trusts cached cell values, so nothing is refreshed — except a
//   modelled chart whose part carries a reference with no cache beside it,
//   refreshed for display only (no `cachesDirty`: its part is still written
//   verbatim, and an untouched workbook still downloads byte-identical).
// - upgradedChartDrawings: a SPEC-131 record's chart placeholders
//   (`{ kind: "chart", sourcePart? }`) parsed from the retained package —
//   see open.ts's version history.
import { chartContentOf, isModelledChart, refreshSheetChart } from "./charts";
import { isEmptyRemap, remapSheetReferences } from "./formula/sheetRefs";
import type { Sheet, SheetChartContent, SheetChartRefs, SheetDrawing, SpreadsheetModel } from "./model";
import { readOdsChart } from "./ods/chartXml";
import { objectsReferenced } from "./ods/package";
import { copyRemap, sourcePartsOf, sourceSheetRemap } from "./sourceNames";
import { readChartPartRefs, readXlsxChart } from "./xlsx/chartXml";
import { LazyParts, partTextOf, relationshipsOf, targetPart, type PartLookup } from "./xlsx/package";

/** Refreshes, for display, every modelled chart whose part names a
 * reference without a cache (see the module header). */
export function refreshUncachedCharts(model: SpreadsheetModel): void {
  const parts = sourcePartsOf(model);
  for (const sheet of model.sheets) {
    for (const drawing of sheet.drawings ?? []) {
      const content = chartContentOf(drawing);
      if (!content?.refs || content.sourcePart === undefined || content.readOnlyReason !== undefined || !isModelledChart(content)) {
        continue;
      }
      const text = partTextOf(parts, content.sourcePart);
      if (text === undefined) {
        continue;
      }
      const uncached = model.format === "xlsx" ? readChartPartRefs(text)?.uncached === true : readOdsChart(text).uncached === true;
      if (uncached) {
        refreshSheetChart(model, sheet.id, drawing.id, { forDisplay: true });
      }
    }
  }
}

/** The part a SPEC-131 chart drawing shows, found through its own source
 * XML: xlsx — its `c:chart r:id` among the sheet's drawing part
 * relationships; ods — its frame's object (`Object N/content.xml`). */
function chartPartOf(model: SpreadsheetModel, sheet: Sheet, drawing: SheetDrawing, parts: PartLookup): string | undefined {
  const xml = drawing.sourceXml;
  if (xml === undefined) {
    return undefined;
  }
  if (model.format === "ods") {
    const object = objectsReferenced(xml)[0];
    return object === undefined ? undefined : `${object}/content.xml`;
  }
  const rId = /<(?:[\w.-]+:)?chart\b[^>]*?\s(?:[\w.-]+:)?id\s*=\s*["']([^"']+)["']/.exec(xml)?.[1];
  if (rId === undefined || sheet.drawingPart === undefined) {
    return undefined;
  }
  const rel = relationshipsOf(parts, sheet.drawingPart).find((each) => each.id === rId && !each.external);
  return rel ? targetPart(sheet.drawingPart, rel) : undefined;
}

/** Chart references through `rename`. */
function renamedRefs(refs: SheetChartRefs, rename: (text: string) => string): SheetChartRefs {
  const each = <T extends string | undefined>(text: T): T => (text === undefined ? text : (rename(text) as T));
  const series = refs.series.map((ref) => Object.fromEntries(Object.entries(ref).map(([key, text]) => [key, each(text)])));
  return refs.plotArea === undefined ? { series, plotVisibleOnly: refs.plotVisibleOnly } : { series, plotArea: each(refs.plotArea), plotVisibleOnly: refs.plotVisibleOnly };
}

/** A sheet's drawings with every SPEC-131 chart placeholder upgraded (see
 * open.ts's version history). `record` is the stored model: its sheet list
 * and retained package give the source-time names to remap. */
export function upgradedChartDrawings(model: SpreadsheetModel, sheet: Sheet, sourceParts: Map<string, Uint8Array>, record: SpreadsheetModel): SheetDrawing[] {
  const remap = copyRemap(sourceSheetRemap(record, new LazyParts(record.opaque)), sheet);
  const rename = isEmptyRemap(remap) ? (text: string) => text : (text: string) => remapSheetReferences(text, remap);
  return (sheet.drawings ?? []).map((drawing): SheetDrawing => {
    if (drawing.content.kind !== "chart" || "visual" in drawing.content) {
      return drawing;
    }
    const placeholder: { sourcePart?: string } = drawing.content;
    const part = chartPartOf(model, sheet, drawing, sourceParts) ?? placeholder.sourcePart;
    const bytes = part === undefined ? undefined : sourceParts.get(part);
    if (part === undefined || bytes === undefined) {
      return { ...drawing, content: { kind: "foreign", label: "Chart" } };
    }
    const text = new TextDecoder().decode(bytes);
    const content: SheetChartContent = { kind: "chart", visual: { chartType: "columnClustered", categories: [], series: [], options: { legend: "none", dataLabels: false, gridlines: false } }, sourcePart: part };
    if (model.format === "xlsx") {
      const read = readXlsxChart(text);
      content.visual = read.visual;
      if (read.refs) {
        content.refs = renamedRefs(read.refs, rename);
      }
      if (read.readOnlyReason !== undefined) {
        content.readOnlyReason = read.readOnlyReason;
      }
    } else {
      const read = readOdsChart(text);
      content.visual = read.visual;
      if (read.refs) {
        content.refs = renamedRefs(read.refs, rename);
      }
    }
    return { ...drawing, content };
  });
}
