// SPEC-048 — planning "Embed fonts in this file" (AC-15/16/17/18). Walks the
// document to learn which family+style each run uses and the code points it
// needs, then — via a caller-supplied face resolver — decides what to embed
// (subset to the used glyphs), what it costs, and which faces are skipped
// because their licence bits forbid embedding. Framework-free: loading the
// actual face bytes (bundled fetch / embedded extraction) is the composable's
// job, injected as `resolveFace`.
import {
  isBreak,
  isChart,
  isImage,
  isShape,
  isToc,
  isTof,
  type DocumentModel,
} from "../documentModel";
import { subsetFont } from "./subset";

export interface StyleUsage {
  family: string;
  bold: boolean;
  italic: boolean;
  codepoints: Set<number>;
}

/** A resolved source for one family+style: the raw face bytes plus whether
 * its licence permits embedding. Null when we have no face for it at all. */
export interface ResolvedFace {
  bytes: Uint8Array;
  embeddable: boolean;
}

export type FaceResolver = (family: string, bold: boolean, italic: boolean) => ResolvedFace | null;

export interface EmbedFace {
  family: string;
  bold: boolean;
  italic: boolean;
  /** Already subset to the used glyphs, ready to write into the package. */
  bytes: Uint8Array;
}

export interface FontEmbedPlan {
  faces: EmbedFace[];
  /** Families skipped because their licence bits forbid embedding (AC-17/18),
   * named for the user. De-duplicated, sorted. */
  skipped: string[];
  /** Total embedded byte size, for the "size cost" disclosure (AC-16). */
  totalBytes: number;
}

function styleKey(bold: boolean, italic: boolean): string {
  return `${bold ? "b" : ""}${italic ? "i" : ""}` || "r";
}

/** Every family+style a run actually uses, with the code points needed. The
 * document default backs a run with no explicit family. */
export function collectFontUsage(model: DocumentModel): StyleUsage[] {
  const byKey = new Map<string, StyleUsage>();
  const add = (family: string, bold: boolean, italic: boolean, text: string) => {
    if (!family || text.length === 0) {
      return;
    }
    const key = `${family.toLowerCase()}|${styleKey(bold, italic)}`;
    let usage = byKey.get(key);
    if (!usage) {
      usage = { family, bold, italic, codepoints: new Set() };
      byKey.set(key, usage);
    }
    for (const ch of text) {
      usage.codepoints.add(ch.codePointAt(0) as number);
    }
  };
  const defaultFamily = model.defaultStyles.fontFamily;
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (isImage(block) || isShape(block) || isChart(block) || isBreak(block) || isToc(block) || isTof(block)) {
        continue;
      }
      const paragraphs =
        block.type === "table"
          ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
          : [block];
      for (const paragraph of paragraphs) {
        for (const run of paragraph.runs) {
          add(
            run.formatting.fontFamily ?? defaultFamily,
            run.formatting.bold === true,
            run.formatting.italic === true,
            run.text,
          );
        }
      }
    }
  }
  return Array.from(byKey.values());
}

/** Builds the embed plan for a document: resolves each used family+style,
 * subsets embeddable faces to the glyphs they need, and collects the
 * restricted families to disclose. */
export function planFontEmbedding(model: DocumentModel, resolveFace: FaceResolver): FontEmbedPlan {
  const faces: EmbedFace[] = [];
  const skipped = new Set<string>();
  let totalBytes = 0;
  for (const usage of collectFontUsage(model)) {
    const resolved = resolveFace(usage.family, usage.bold, usage.italic);
    if (!resolved) {
      continue;
    }
    if (!resolved.embeddable) {
      skipped.add(usage.family);
      continue;
    }
    const bytes = subsetFont(resolved.bytes, usage.codepoints);
    totalBytes += bytes.length;
    faces.push({ family: usage.family, bold: usage.bold, italic: usage.italic, bytes });
  }
  return { faces, skipped: Array.from(skipped).sort((a, b) => a.localeCompare(b)), totalBytes };
}

/** A human "size cost" string for the embed disclosure (AC-16). */
export function formatEmbedSize(totalBytes: number): string {
  if (totalBytes < 1024) {
    return `${totalBytes} B`;
  }
  const kb = totalBytes / 1024;
  if (kb < 1024) {
    return `${kb.toFixed(kb < 10 ? 1 : 0)} KB`;
  }
  return `${(kb / 1024).toFixed(1)} MB`;
}
