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
import { isChart, isDrawingGroup, type Chart, type DocumentModel, type DrawingMember } from "../documentModel";
import { chartPartXml, CHART_CONTENT_TYPE, CHART_RELATIONSHIP } from "./chart";
import { nextFreeRelId } from "./media";
import type { ZipEntryInput } from "../core/zipWriter";

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
  /** The plan entry for a Chart member of a container at (section, block),
   * addressed by its member path (SPEC-057) — the group writer looks up the
   * relationship id it must reference in the member's `graphicFrame`. */
  planForMember(section: number, block: number, path: number[]): ChartPlanEntry | undefined;
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

  /** Plans one chart's package additions and records it under `key`. Shared by
   * a top-level Chart block and a Chart nested inside a DrawingGroup's members
   * (SPEC-057 AC-55/AC-59): both mint a part + relationship the same way. */
  const planChartAt = (chart: Chart, key: string): void => {
    if (!forceRegenerate && chart.sourcePart) {
      liveSourceParts.add(chart.sourcePart);
    }
    if (forceRegenerate && chart.unsupportedChartXml !== undefined) {
      return;
    }
    const reuseName = !forceRegenerate && !!chart.sourcePart;
    const partName = reuseName ? chart.sourcePart! : freshPartName();
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
    const needsRegeneration =
      forceRegenerate || !chart.sourcePart || (chart.edited === true && chart.unsupportedChartXml === undefined);
    const xml = needsRegeneration ? chartPartXml(chart) : null;
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
    plans.set(key, { relId, partName, xml });
  };

  /** Recurses into a DrawingGroup's members, planning any Chart member and
   * keying it by its member path so the group writer can look it up. A sealed
   * container re-emits its preserved subtree verbatim, so its members are not
   * re-planned. */
  const planGroupMembers = (members: DrawingMember[], keyPrefix: string): void => {
    members.forEach((member, index) => {
      if (isChart(member)) {
        planChartAt(member, `${keyPrefix}.m${index}`);
      } else if (isDrawingGroup(member) && member.preservedSubtreeXml === undefined) {
        planGroupMembers(member.members, `${keyPrefix}.m${index}`);
      }
    });
  };

  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (isChart(block)) {
        planChartAt(block, `${sectionIndex}.${blockIndex}`);
      } else if (isDrawingGroup(block) && block.preservedSubtreeXml === undefined) {
        planGroupMembers(block.members, `${sectionIndex}.${blockIndex}`);
      }
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
    planForMember: (section, block, path) => plans.get(`${section}.${block}.${path.map((p) => `m${p}`).join(".")}`),
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
  planForMember: () => undefined,
  relationshipsXml: "",
  contentTypesXml: "",
  zipEntries: [],
  regeneratedPartXml: new Map(),
  orphanedPartNames: new Set(),
  relIds: new Set(),
};
