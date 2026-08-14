// Plans the package additions/changes a .docx save needs for the model's
// charts (SPEC-052): a relationship per chart (word/_rels/document.xml.rels),
// a `word/charts/chartN.xml` part per chart — either the preserved original
// bytes (untouched chart, handled by the caller re-emitting `sourcePart`
// verbatim from `DocumentModel.opaque`) or freshly regenerated content (a
// created or edited chart, `zipEntries` here) — and the Content-Type
// Override a genuinely new part needs. Mirrors media.ts's MediaWritePlan
// shape, but keyed by a chart's *location* (section/block index) rather than
// a stable id: Chart carries no field named `id` (IGNORED_KEYS, see
// documentModel.ts), and a location is stable for the lifetime of one write.
import { isChart, type DocumentModel } from "../documentModel";
import { chartPartXml, CHART_CONTENT_TYPE, CHART_RELATIONSHIP } from "./chart";
import { nextFreeRelId } from "./media";
import type { ZipEntryInput } from "./zipWriter";

export interface ChartPlanEntry {
  relId: string;
  partName: string;
  /** Regenerated part text to write, or null to leave the part alone (an
   * untouched chart re-emitted verbatim from `DocumentModel.opaque` by the
   * caller's existing preserved-parts pass-through). */
  xml: string | null;
}

export interface ChartWritePlan {
  planFor(section: number, block: number): ChartPlanEntry | undefined;
  /** `<Relationship>` elements to append to word/_rels/document.xml.rels. */
  relationshipsXml: string;
  /** `<Override>` elements for `[Content_Types].xml` — only charts with no
   * `sourcePart` (freshly created here) need one; a reused part's Override
   * already exists in the preserved `[Content_Types].xml`. */
  contentTypesXml: string;
  /** Brand-new chart parts (no `sourcePart`) to add to the package. */
  zipEntries: ZipEntryInput[];
  /** Regenerated content for an *edited* chart whose part name is reused
   * in place (it has a `sourcePart`) — keyed by that part name, for the
   * caller's preserved-parts pass-through loop to substitute in place of
   * the verbatim original bytes (mirrors how patchedStyles/patchedSettings
   * substitute a named preserved part elsewhere in writeDocx). */
  regeneratedPartXml: Map<string, string>;
  /** Preserved `word/charts/chartN.xml` part names no live Chart block
   * references any more — a chart the user deleted. The caller excludes
   * these from the same-format save's verbatim preserved-parts pass-through
   * so SPEC-008's part-preservation rule does not resurrect a deleted chart
   * (SPEC-052 AC). */
  orphanedPartNames: Set<string>;
  /** This plan's own claimed relationship ids, so later allocators never
   * collide (mirrors MediaWritePlan.relIds/HyperlinkWritePlan.relIds). */
  relIds: Set<string>;
}

const CHART_PART_NAME_RE = /^word\/charts\/chart\d+\.xml$/;

export function planChartWrite(
  model: DocumentModel,
  takenRelIds: Set<string>,
  existingPartNames: Set<string>,
  /** True for a brand-new document or a cross-format save (the top,
   * `preservedParts.size === 0` branch of writeDocx): there is no preserved
   * `word/charts/*.xml` part to reuse — whatever `sourcePart` says (if
   * anything; a chart read from `.odt` names an ODF-format part there,
   * meaningless as a `.docx` target) is ignored, every catalogue-type chart
   * is regenerated fresh, and an unsupported chart (`unsupportedChartXml`)
   * is dropped entirely rather than written — already reported via
   * `opaque.ts`'s `collectPackageDropRecords` "chart" DropRecord. */
  forceRegenerate = false,
  /** The preserved package's own chart relationships, part-name -> relId
   * (built by the caller from the preserved word/_rels/document.xml.rels).
   * A same-format chart reusing its `sourcePart` reuses that exact relId
   * too and adds no new relationship — a chart the user never touched (not
   * edited, not moved/resized/re-anchored, since none of those set
   * `edited`) then leaves `word/_rels/document.xml.rels` byte-identical as
   * well as its own part, not just the part. Only a chart with no matching
   * preserved relationship (defensive — should not normally happen) or no
   * `sourcePart` at all mints a fresh one. */
  existingRelIdForPart: Map<string, string> = new Map(),
): ChartWritePlan {
  const taken = new Set(takenRelIds);
  const relIds = new Set<string>();
  const usedNames = new Set(existingPartNames);
  const relParts: string[] = [];
  const overrideParts: string[] = [];
  const zipEntries: ZipEntryInput[] = [];
  const regeneratedPartXml = new Map<string, string>();
  const plans = new Map<string, ChartPlanEntry>();
  const liveSourceParts = new Set<string>();
  const encoder = new TextEncoder();
  let nextChartNumber = 1;

  const freshPartName = (): string => {
    let name = `word/charts/chart${nextChartNumber}.xml`;
    while (usedNames.has(name)) {
      nextChartNumber++;
      name = `word/charts/chart${nextChartNumber}.xml`;
    }
    usedNames.add(name);
    return name;
  };

  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (!isChart(block)) {
        return;
      }
      if (!forceRegenerate && block.sourcePart) {
        liveSourceParts.add(block.sourcePart);
      }
      if (forceRegenerate && block.unsupportedChartXml !== undefined) {
        // Cross-format save of an unsupported chart: nothing to regenerate
        // from and no bytes to carry across formats — dropped entirely.
        return;
      }
      // A same-format save reuses the source part's own name so the
      // caller's preserved-parts pass-through naturally keeps re-emitting
      // it (verbatim when untouched, substituted when edited). A
      // cross-format save ignores `sourcePart` — it names a part in the
      // *other* format, meaningless as a target here — and always mints a
      // fresh name.
      const reuseName = !forceRegenerate && !!block.sourcePart;
      const partName = reuseName ? block.sourcePart! : freshPartName();
      const existingRelId = reuseName ? existingRelIdForPart.get(partName) : undefined;
      let relId: string;
      if (existingRelId) {
        relId = existingRelId;
        taken.add(relId);
        relIds.add(relId);
      } else {
        relId = nextFreeRelId(taken);
        taken.add(relId);
        relIds.add(relId);
        relParts.push(
          `<Relationship Id="${relId}" Type="${CHART_RELATIONSHIP}" Target="${partName.replace(/^word\//, "")}"/>`,
        );
      }
      // Regenerate content for: a cross-format save (always), a chart with
      // no source (freshly created/inserted here), or one the user
      // actually edited (SPEC-052 "Rewriting a chart"). An unsupported
      // chart is never edited (AC-30 disables it) and already returned
      // above for the cross-format case; for a same-format save its
      // placeholder model fields are stand-ins, never regenerated.
      const needsRegeneration =
        forceRegenerate || !block.sourcePart || (block.edited === true && block.unsupportedChartXml === undefined);
      const xml = needsRegeneration ? chartPartXml(block) : null;
      if (xml !== null) {
        if (reuseName) {
          regeneratedPartXml.set(partName, xml);
        } else {
          zipEntries.push({ name: partName, data: encoder.encode(xml) });
        }
      }
      if (!reuseName) {
        overrideParts.push(`<Override PartName="/${partName}" ContentType="${CHART_CONTENT_TYPE}"/>`);
      }
      plans.set(`${sectionIndex}.${blockIndex}`, { relId, partName, xml });
    });
  });

  const orphanedPartNames = new Set<string>();
  for (const name of existingPartNames) {
    if (CHART_PART_NAME_RE.test(name) && !liveSourceParts.has(name)) {
      orphanedPartNames.add(name);
    }
  }

  return {
    planFor: (section, block) => plans.get(`${section}.${block}`),
    relationshipsXml: relParts.join(""),
    contentTypesXml: overrideParts.join(""),
    zipEntries,
    regeneratedPartXml,
    orphanedPartNames,
    relIds,
  };
}

export const EMPTY_CHART_PLAN: ChartWritePlan = {
  planFor: () => undefined,
  relationshipsXml: "",
  contentTypesXml: "",
  zipEntries: [],
  regeneratedPartXml: new Map(),
  orphanedPartNames: new Set(),
  relIds: new Set(),
};
