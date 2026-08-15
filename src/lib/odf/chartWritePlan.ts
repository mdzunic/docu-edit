// Plans the package additions/changes a .odt save needs for the model's
// charts (SPEC-052): an embedded object subpackage per chart
// ("Object N/content.xml", the preserved original bytes for an untouched
// chart or freshly regenerated content for a created/edited one — mirrors
// docx/chartWritePlan.ts's own split, adapted to ODF's embedded-object
// package shape) plus its META-INF/manifest.xml file-entries. Keyed by a
// chart's *location* (section/block index), not a stable id — Chart carries
// no field named `id` (IGNORED_KEYS, see documentModel.ts).
import { isChart, type DocumentModel } from "../documentModel";
import { CHART_OBJECT_CONTENT_TYPE, odfChartContentXml } from "./chart";
import type { OdtManifestEntry } from "./media";
import type { ZipEntryInput } from "../core/zipWriter";

export interface OdfChartPlanEntry {
  /** The embedded object's directory, e.g. "Object 1" (no trailing slash);
   * `draw:object`'s `xlink:href` is `"./" + objectDir`. */
  objectDir: string;
  /** Regenerated content.xml text to write, or null to leave the part alone
   * (an untouched chart re-emitted verbatim from `DocumentModel.opaque` by
   * the caller's existing preserved-parts pass-through). */
  xml: string | null;
}

export interface OdfChartWritePlan {
  planFor(section: number, block: number): OdfChartPlanEntry | undefined;
  /** Manifest file-entries for brand-new embedded objects (both the
   * directory entry and its content.xml entry) — a reused object's entries
   * already exist in the preserved manifest. */
  manifestEntries: OdtManifestEntry[];
  /** Brand-new embedded object content.xml parts to add to the package. */
  zipEntries: ZipEntryInput[];
  /** Regenerated content for an *edited* chart whose object directory is
   * reused in place (it has a `sourcePart`) — keyed by that part name
   * ("Object N/content.xml"), for the caller's preserved-parts pass-through
   * loop to substitute in place of the verbatim original bytes. */
  regeneratedPartXml: Map<string, string>;
  /** Preserved "Object N/content.xml" (and their directory manifest entry)
   * names no live Chart block references any more — a chart the user
   * deleted. Derived from the *preserved manifest*'s own chart-typed
   * embedded-object entries, not a blanket "Object N" name guess, so a
   * non-chart embedded object (a real OLE object, another app's embed)
   * sharing the same naming convention is never mistaken for an orphan. */
  orphanedPartNames: Set<string>;
  orphanedObjectDirs: Set<string>;
}

/** Every embedded-object directory the *preserved* manifest already declares
 * as a chart (`manifest:media-type="application/vnd.oasis.opendocument.chart"`)
 * — a plain regex scan over already-decoded manifest text, not full XML
 * parsing, since this is the one thing this module needs from it. */
export function chartObjectDirsInManifest(manifestXml: string): Set<string> {
  const dirs = new Set<string>();
  const re = /<manifest:file-entry\b[^>]*manifest:full-path="([^"]+)\/"[^>]*manifest:media-type="application\/vnd\.oasis\.opendocument\.chart"[^>]*\/>/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(manifestXml))) {
    dirs.add(match[1]);
  }
  // Attribute order in a real file is not guaranteed — try the reverse order too.
  const reReversed = /<manifest:file-entry\b[^>]*manifest:media-type="application\/vnd\.oasis\.opendocument\.chart"[^>]*manifest:full-path="([^"]+)\/"[^>]*\/>/g;
  while ((match = reReversed.exec(manifestXml))) {
    dirs.add(match[1]);
  }
  return dirs;
}

export function planOdfChartWrite(
  model: DocumentModel,
  /** Chart object directories already known to be charts in the preserved
   * package (see chartObjectDirsInManifest) — empty for a fresh document or
   * a cross-format save, where there is nothing preserved to reuse or orphan. */
  preservedChartObjectDirs: Set<string>,
  forceRegenerate = false,
): OdfChartWritePlan {
  const manifestEntries: OdtManifestEntry[] = [];
  const zipEntries: ZipEntryInput[] = [];
  const regeneratedPartXml = new Map<string, string>();
  const plans = new Map<string, OdfChartPlanEntry>();
  const liveObjectDirs = new Set<string>();
  let nextObjectNumber = 1;
  const usedDirs = new Set<string>();

  const freshObjectDir = (): string => {
    let dir = `Object ${nextObjectNumber}`;
    while (usedDirs.has(dir)) {
      nextObjectNumber++;
      dir = `Object ${nextObjectNumber}`;
    }
    usedDirs.add(dir);
    return dir;
  };

  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (!isChart(block)) {
        return;
      }
      if (forceRegenerate && block.unsupportedChartXml !== undefined) {
        // Cross-format save of an unsupported chart: nothing to regenerate
        // from and no bytes to carry across formats — dropped entirely
        // (already reported via opaque.ts's collectPackageDropRecords).
        return;
      }
      const sourceDir = !forceRegenerate && block.sourcePart ? block.sourcePart.replace(/\/content\.xml$/, "") : undefined;
      if (sourceDir) {
        liveObjectDirs.add(sourceDir);
      }
      const objectDir = sourceDir ?? freshObjectDir();
      const needsRegeneration = forceRegenerate || !sourceDir || (block.edited === true && block.unsupportedChartXml === undefined);
      const xml = needsRegeneration ? odfChartContentXml(block) : null;
      const partName = `${objectDir}/content.xml`;
      if (xml !== null) {
        if (sourceDir) {
          regeneratedPartXml.set(partName, xml);
        } else {
          zipEntries.push({ name: partName, data: new TextEncoder().encode(xml) });
          manifestEntries.push({ path: `${objectDir}/`, mediaType: CHART_OBJECT_CONTENT_TYPE });
          manifestEntries.push({ path: partName, mediaType: "text/xml" });
        }
      }
      plans.set(`${sectionIndex}.${blockIndex}`, { objectDir, xml });
    });
  });

  const orphanedObjectDirs = new Set<string>();
  for (const dir of preservedChartObjectDirs) {
    if (!liveObjectDirs.has(dir)) {
      orphanedObjectDirs.add(dir);
    }
  }
  const orphanedPartNames = new Set<string>();
  for (const dir of orphanedObjectDirs) {
    orphanedPartNames.add(`${dir}/content.xml`);
  }

  return {
    planFor: (section, block) => plans.get(`${section}.${block}`),
    manifestEntries,
    zipEntries,
    regeneratedPartXml,
    orphanedPartNames,
    orphanedObjectDirs,
  };
}

export const EMPTY_ODF_CHART_PLAN: OdfChartWritePlan = {
  planFor: () => undefined,
  manifestEntries: [],
  zipEntries: [],
  regeneratedPartXml: new Map(),
  orphanedPartNames: new Set(),
  orphanedObjectDirs: new Set(),
};
