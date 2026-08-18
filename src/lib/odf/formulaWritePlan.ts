// Plans the package additions a .odt save needs for the model's equations
// (SPEC-056): one embedded formula subpackage per equation run
// ("Formula N/content.xml" holding MathML) plus its META-INF/manifest.xml
// file-entries. Mirrors chartWritePlan.ts, adapted to the fact that an
// equation rides on a Run (inline or the sole run of a display paragraph)
// rather than a Block: the plan is keyed by the Run object's identity, since
// buildContentXml walks the same model instance the plan was built from.
//
// An equation is fully modelled (its MathNode body is the source of truth), so
// — unlike a chart, which preserves unmodelled features by re-emitting the
// original bytes — every save regenerates the formula content.xml fresh and
// orphans any preserved formula object dir. That is lossless on the model
// (the equation round-trips structurally) and keeps this module simple.
import { equationAltText } from "../math/linearForm";
import type { DocumentModel, Run } from "../documentModel";
import { odfFormulaContentXml } from "./formula";
import { FORMULA_OBJECT_CONTENT_TYPE } from "./formula";
import type { OdtManifestEntry } from "./media";
import type { ZipEntryInput } from "../core/zipWriter";

export interface OdfFormulaWritePlan {
  /** The object href (e.g. "./Formula 1") for an equation run, or undefined
   * when the run carries no equation. */
  hrefFor(run: Run): string | undefined;
  manifestEntries: OdtManifestEntry[];
  zipEntries: ZipEntryInput[];
  /** Preserved "Formula N/content.xml" names no live equation reuses (every
   * one, since equations always regenerate) — dropped from the rebuilt
   * package like a deleted chart's object. */
  orphanedPartNames: Set<string>;
  orphanedObjectDirs: Set<string>;
}

function* equationRuns(model: DocumentModel): Generator<Run> {
  for (const section of model.sections) {
    for (const block of section.blocks) {
      const paragraphs =
        block.type === "table"
          ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
          : block.type === "paragraph"
            ? [block]
            : [];
      for (const paragraph of paragraphs) {
        for (const run of paragraph.runs) {
          if (run.equation) yield run;
        }
      }
    }
  }
}

export function planOdfFormulaWrite(
  model: DocumentModel,
  /** Formula object directories the preserved package already declares — all
   * are orphaned, since equations always regenerate to fresh dirs. */
  preservedFormulaObjectDirs: Set<string> = new Set(),
): OdfFormulaWritePlan {
  const manifestEntries: OdtManifestEntry[] = [];
  const zipEntries: ZipEntryInput[] = [];
  const hrefs = new Map<Run, string>();
  const encoder = new TextEncoder();

  // Fresh "Formula N" directories that never collide with a preserved dir.
  let next = 1;
  const nextDir = (): string => {
    let dir = `Formula ${next}`;
    while (preservedFormulaObjectDirs.has(dir)) {
      next += 1;
      dir = `Formula ${next}`;
    }
    next += 1;
    return dir;
  };

  for (const run of equationRuns(model)) {
    if (!run.equation) continue;
    const objectDir = nextDir();
    const partName = `${objectDir}/content.xml`;
    zipEntries.push({ name: partName, data: encoder.encode(odfFormulaContentXml(run.equation)) });
    manifestEntries.push({ path: `${objectDir}/`, mediaType: FORMULA_OBJECT_CONTENT_TYPE });
    manifestEntries.push({ path: partName, mediaType: "text/xml" });
    hrefs.set(run, `./${objectDir}`);
  }

  const orphanedObjectDirs = new Set(preservedFormulaObjectDirs);
  const orphanedPartNames = new Set<string>();
  for (const dir of orphanedObjectDirs) orphanedPartNames.add(`${dir}/content.xml`);

  return {
    hrefFor: (run) => hrefs.get(run),
    manifestEntries,
    zipEntries,
    orphanedPartNames,
    orphanedObjectDirs,
  };
}

/** The alt text a formula frame's `svg:desc` should carry — omitted when it is
 * only the default linear form (that is regenerated on read), emitted when the
 * author set an explicit override. */
export function formulaFrameAltText(run: Run): string {
  if (!run.equation) return "";
  return run.equation.altText !== undefined ? equationAltText(run.equation) : "";
}
