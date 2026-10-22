// The xlsx sheet-chart write planner (SPEC-132 "Writing", "New chart
// parts"): THIS FILE is the normative source for the bytes a save writes
// each sheet chart's part with. For every chart drawing on every sheet (grid
// sheets and chart sheets), in this order of precedence:
//
//  1. No part to write at — a chart created here, a duplicated chart, a
//     cross-format copy (no `sourcePart`, or one the package lacks): a NEW
//     part `xl/charts/chart{n}.xml` (chartXml.ts nextChartPartName over every
//     part name the package holds — dropped ones included — and every one
//     this save mints), its text xlsxChartPartXml(visual, refs), its
//     CHART_CONTENT_TYPE Override. Its anchor is written in the canonical
//     form (drawingXml.ts canonicalAnchorXml) with a CHART_RELATIONSHIP in
//     the drawing part's `.rels`, so a part is minted only when the drawing
//     layer relating it is written (partFor). An unmodelled chart (out of
//     the catalogue) has no wire form of its own: it is left out.
//  2. `edited` on a modelled chart that is not read-only → the part
//     regenerated: xlsxChartPartXml(visual, refs); its `.rels` and siblings
//     (chartStyle/colors, embeddings) stay as they are (SPEC-092's answer).
//  3. `refsDirty` or `cachesDirty` on a modelled chart with refs → the part
//     patched in place: patchXlsxChartPart over its retained text — after
//     SPEC-121's sheet-name pass (sheetNamePass.ts), which renames the
//     part's own retained references (a linked title's, a data-label
//     range's) exactly as a dirty sheet's retained text is renamed before
//     the model's elements are spliced into it — with `cellsOf` =
//     charts.ts chartReferenceCells on the chart's sheet; a read-only chart
//     gets no `cellsOf`, so only its `c:f` texts change.
//  4. Otherwise the part's bytes pass through as the package holds them —
//     the part, its `.rels`, chartStyle/colors and embeddings verbatim (the
//     writer gives a passing-through part the sheet-name pass, as every
//     retained part gets it).
//
// A regenerated or patched part is MODEL TEXT: the writer never puts it
// through the sheet-name pass (its series references already name sheets as
// the model does — a swap of two names would undo itself). An unmodelled
// chart is never patched or regenerated: its part passes through, already
// adjusted by SPEC-124's and SPEC-125's retained rewrites. The part a chart
// drawing is written at is the one its anchor relates (drawingXml.ts
// chartPartResolver), so a duplicated sheet's chart writes its own closure
// copy. A deleted chart's part leaves the package by SPEC-131's reachability
// rule (writer.ts): nothing relates it once its anchor is gone.
import { CHART_CONTENT_TYPE } from "../../core/chartPartOoxml";
import { chartContentOf, chartReferenceCells, isModelledChart } from "../charts";
import { isEmptyRemap, type SheetRemap } from "../formula/sheetRefs";
import type { Sheet, SheetChartRefs, SheetDrawing, SpreadsheetModel } from "../model";
import { nextChartPartName, patchXlsxChartPart, xlsxChartPartXml } from "./chartXml";
import { chartPartResolver } from "./drawingXml";
import { decodeUtf8, encodeUtf8, type PartLookup } from "./package";
import { mentionsAnyName, passChart, remappedNames } from "./sheetNamePass";

/** What a save writes for the workbook's charts. */
export interface XlsxChartWrites {
  /** Retained chart parts written from the model (regenerated or patched):
   * model text, never given the sheet-name pass. */
  replaced: Map<string, Uint8Array>;
  /** New chart parts, in minting order. */
  added: { name: string; data: Uint8Array }[];
  /** The new parts' Overrides. */
  overrides: { partName: string; contentType: string }[];
}

/** Whether a sheet's charts need a save to write anything (the writer's
 * "nothing changed" fast path is not taken): a chart drawing that is
 * `edited`, `cachesDirty` or `refsDirty`, or has no `sourcePart`. A part
 * change alone needs no dirty sheet — a refresh sets `cachesDirty` only. */
export function sheetChartsChanged(sheet: Sheet): boolean {
  return (sheet.drawings ?? []).some((drawing) => {
    const content = chartContentOf(drawing);
    return content !== undefined && (content.edited === true || content.cachesDirty === true || content.refsDirty === true || content.sourcePart === undefined);
  });
}

/** Whether a sheet holds a chart drawing with no retained anchor XML — one
 * its drawing layer must be written to relate (the canonical form). */
export function sheetHasCreatedChart(sheet: Sheet): boolean {
  return (sheet.drawings ?? []).some((drawing) => drawing.content.kind === "chart" && drawing.sourceXml === undefined);
}

/** One save's chart writes (see the module header): `partFor` while the
 * drawing layers are written, then `writes()`. `parts` is the retained
 * package (empty from scratch); `taken` every part name in it and every one
 * this save has minted so far (joined by the chart parts minted here);
 * `remapOf` the sheet-name remap a retained part's pass takes (SPEC-121's
 * sourceSheetRemap, a copy's own for its closure). */
export class XlsxChartWriter {
  private readonly resolve: ReturnType<typeof chartPartResolver>;
  private readonly minted = new Map<string, { name: string; drawing: SheetDrawing }>();

  constructor(
    private readonly model: SpreadsheetModel,
    private readonly parts: PartLookup,
    private readonly taken: Set<string>,
    private readonly remapOf: (partName: string) => SheetRemap = () => ({ renamed: {}, deleted: [], orderBefore: [] }),
  ) {
    this.resolve = chartPartResolver(parts);
  }

  /** The chart part a chart drawing written in the canonical form relates:
   * the part it already has (written by paths 2–4), else a newly minted one
   * (once per drawing); undefined for an unmodelled chart without a part,
   * or any other content. */
  partFor(sheet: Sheet, drawing: SheetDrawing): string | undefined {
    const content = chartContentOf(drawing);
    if (!content) {
      return undefined;
    }
    const existing = this.resolve(sheet, drawing);
    if (existing !== undefined) {
      return existing;
    }
    if (!isModelledChart(content)) {
      return undefined;
    }
    const key = `${sheet.id}\u0000${drawing.id}`;
    const known = this.minted.get(key);
    if (known) {
      return known.name;
    }
    const name = nextChartPartName(this.taken);
    this.taken.add(name);
    this.minted.set(key, { name, drawing });
    return name;
  }

  /** Every retained chart part the model rewrites, and the minted parts. */
  writes(): XlsxChartWrites {
    const out: XlsxChartWrites = { replaced: new Map(), added: [], overrides: [] };
    for (const sheet of this.model.sheets) {
      for (const drawing of sheet.drawings ?? []) {
        const content = chartContentOf(drawing);
        const part = content && isModelledChart(content) ? this.resolve(sheet, drawing) : undefined;
        if (!content || part === undefined || out.replaced.has(part)) {
          continue;
        }
        if (content.edited && content.readOnlyReason === undefined) {
          out.replaced.set(part, encodeUtf8(xlsxChartPartXml(content.visual, content.refs)));
        } else if ((content.refsDirty || content.cachesDirty) && content.refs) {
          out.replaced.set(part, this.patched(sheet, part, content.refs, content.readOnlyReason === undefined));
        }
      }
    }
    for (const { name, drawing } of this.minted.values()) {
      const content = chartContentOf(drawing)!;
      out.added.push({ name, data: encodeUtf8(xlsxChartPartXml(content.visual, content.refs)) });
      out.overrides.push({ partName: name, contentType: CHART_CONTENT_TYPE });
    }
    return out;
  }

  /** Path 3: the retained part through the pass, then patched. */
  private patched(sheet: Sheet, part: string, refs: SheetChartRefs, refresh: boolean): Uint8Array {
    const bytes = this.parts.get(part)!;
    const source = decodeUtf8(bytes);
    const remap = this.remapOf(part);
    const renamed = isEmptyRemap(remap) || !mentionsAnyName(bytes, remappedNames(remap)) ? source : passChart(source, remap);
    const next = refresh ? patchXlsxChartPart(renamed, refs, (canonical) => chartReferenceCells(this.model, canonical, sheet.id)) : patchXlsxChartPart(renamed, refs);
    return next === source ? bytes : encodeUtf8(next);
  }
}
