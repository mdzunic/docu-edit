// The ods chart write paths (SPEC-132 "Writing", "New chart parts") for the
// embedded chart objects a model's chart drawings show — the planner
// ods/writer.ts runs once per save, on the substitution path and from
// scratch. For each modelled chart drawing, in this order of precedence:
// 1. no `sourcePart` (created here, a duplicate's new chart, a cross-format
//    copy): a new `Object {n}` (n = 1 + the highest across the package and
//    the objects minted in this save) — its parts newOdsChartObjectParts over
//    odsChartContentXml, its manifest entries odsChartManifestEntries — the
//    dir shared with the table writer (`chartObjectDir`), which writes the
//    chart's canonical frame around chartObjectElementXml;
// 2. `edited` → its object's content.xml regenerated (odsChartContentXml at
//    the anchor's size); the object's other parts stay;
// 3. `refsDirty` or `cachesDirty` → patched in place (patchOdfChartContent:
//    the five addresses, and the `local-table` when the caches changed) over
//    the text the sheet-name pass gives, so its other references keep being
//    renamed;
// 4. otherwise its bytes pass through, with the pass, as before.
// A regenerated or patched object is model text: the writer never passes
// it again. An out-of-catalogue chart is never patched or regenerated.
import { chartContentOf, isModelledChart } from "../charts";
import { drawingRectPx } from "../drawings";
import { buildGridMetrics } from "../geometry";
import type { Sheet, SheetChartContent, SheetDrawing, SpreadsheetModel } from "../model";
import { newOdsChartObjectParts, nextOdsObjectDir, odsChartContentXml, odsChartManifestEntries, patchOdfChartContent } from "./chartXml";

/** cm per CSS px at 96 dpi. */
const CM_PER_PX = 2.54 / 96;

export interface OdsChartPlan {
  /** Each created chart drawing (by identity) → its minted `Object {n}`. */
  minted: Map<SheetDrawing, string>;
  /** The minted objects' parts. */
  newParts: { name: string; text: string }[];
  /** The minted objects' manifest entries. */
  manifestEntries: string[];
  /** An existing object's content.xml regenerated or patched from the text
   * the sheet-name pass gives (`passed`), or undefined to pass it through. */
  rewrite(partName: string, passed: string): string | undefined;
}

/** A drawing's modelled chart content, or undefined (a record a reader
 * upgrade has not reached yet holds a chart without a visual). */
function modelledChart(drawing: SheetDrawing): SheetChartContent | undefined {
  const content = chartContentOf(drawing);
  return content && "visual" in content && isModelledChart(content) ? content : undefined;
}

/** The anchor's size in cm under the sheet's geometry. */
function sizeCm(sheet: Sheet, drawing: SheetDrawing): { width: number; height: number } {
  const rect = drawingRectPx(drawing.anchor, buildGridMetrics(sheet));
  return { width: rect.width * CM_PER_PX, height: rect.height * CM_PER_PX };
}

/** Whether a save has chart work beyond verbatim parts. */
export function hasOdsChartWork(model: SpreadsheetModel): boolean {
  return model.sheets.some((sheet) =>
    (sheet.drawings ?? []).some((drawing) => {
      const content = modelledChart(drawing);
      return content !== undefined && (content.sourcePart === undefined || content.edited === true || content.refsDirty === true || content.cachesDirty === true);
    }),
  );
}

/** The save's chart plan (see the module header); `taken` names every part
 * the package already holds. */
export function planOdsCharts(model: SpreadsheetModel, taken: Iterable<string>): OdsChartPlan {
  const names = new Set(taken);
  const minted = new Map<SheetDrawing, string>();
  const newParts: { name: string; text: string }[] = [];
  const manifestEntries: string[] = [];
  const edits = new Map<string, (passed: string) => string>();
  for (const sheet of model.sheets) {
    for (const drawing of sheet.drawings ?? []) {
      const content = modelledChart(drawing);
      if (!content) {
        continue;
      }
      if (content.sourcePart === undefined) {
        const dir = nextOdsObjectDir(names);
        names.add(`${dir}/content.xml`);
        minted.set(drawing, dir);
        const { width, height } = sizeCm(sheet, drawing);
        newParts.push(...newOdsChartObjectParts(dir, odsChartContentXml(content.visual, width, height, model.theme, content.refs)));
        manifestEntries.push(odsChartManifestEntries(dir));
      } else if (content.edited && content.readOnlyReason === undefined) {
        const { width, height } = sizeCm(sheet, drawing);
        const text = odsChartContentXml(content.visual, width, height, model.theme, content.refs);
        edits.set(content.sourcePart, () => text);
      } else if (content.refs && (content.refsDirty || content.cachesDirty)) {
        const refs = content.refs;
        const visual = content.cachesDirty && content.readOnlyReason === undefined ? content.visual : undefined;
        edits.set(content.sourcePart, (passed) => patchOdfChartContent(passed, refs, visual));
      }
    }
  }
  return { minted, newParts, manifestEntries, rewrite: (partName, passed) => edits.get(partName)?.(passed) };
}
