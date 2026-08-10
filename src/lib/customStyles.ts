// "Import styles from another document" (SPEC-042 AC-19..24): reads a
// chosen .docx/.odt file's style catalogue, plans a merge against the
// current document (which source styles conflict by user-visible name, AC-
// 21), and applies the user's per-conflict keep/replace/skip choice plus
// every non-conflicting style automatically (AC-22, AC-23). Framework-free
// like every lib module — the dialog drives this through a composable.
import type { DocumentModel, NamedStyle } from "./documentModel";
import { parseDocx } from "./docx/reader";
import { detectOdf, parseOdt } from "./odf/reader";
import { allStyles, builtInStyles } from "./styles";

/** Reads every style a `.docx`/`.odt` file declares (AC-19, AC-20) — the
 * same detect-and-parse path `openDocument.ts` uses for opening a document
 * outright, since both readers already resolve their whole style catalogue
 * as one step; only `.styles` is kept, the rest of the parse is discarded. */
export async function readStylesFromFile(file: File): Promise<NamedStyle[]> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const odf = await detectOdf(bytes);
  if (odf === "flat") {
    throw new Error("Flat ODF (.fodt) files are not supported for style import");
  }
  if (odf === "package") {
    const parsed = await parseOdt(bytes);
    return parsed.styles;
  }
  const parsed = await parseDocx(bytes);
  return parsed.styles;
}

/** A source style whose user-visible name matches an existing style in the
 * current document (AC-21) — `existing` is the current document's style the
 * user is asked to keep, replace with `source`, or skip (AC-22). */
export interface StyleImportConflict {
  source: NamedStyle;
  existing: NamedStyle;
}

export interface StyleImportPlan {
  /** Source styles merged automatically, no name collision (AC-23). */
  nonConflicting: NamedStyle[];
  conflicts: StyleImportConflict[];
}

function normalizedName(style: NamedStyle): string | undefined {
  return style.name?.trim().toLowerCase();
}

/** Plans a merge of `sourceStyles` into `model` (AC-20, AC-21): every named
 * source style either collides with an existing style's name (a conflict
 * the dialog must show, AC-22) or merges automatically (AC-23). A source
 * style with no name (an id-only stub some documents declare) is always
 * treated as non-conflicting — there is no user-visible name to collide on. */
export function planStyleImport(model: DocumentModel, sourceStyles: NamedStyle[]): StyleImportPlan {
  const existingByName = new Map<string, NamedStyle>();
  for (const style of allStyles(model)) {
    const name = normalizedName(style);
    if (name !== undefined) {
      existingByName.set(name, style);
    }
  }
  const nonConflicting: NamedStyle[] = [];
  const conflicts: StyleImportConflict[] = [];
  for (const style of sourceStyles) {
    const name = normalizedName(style);
    const existing = name !== undefined ? existingByName.get(name) : undefined;
    if (existing) {
      conflicts.push({ source: style, existing });
    } else {
      nonConflicting.push(style);
    }
  }
  return { nonConflicting, conflicts };
}

export type StyleConflictResolution = "keep" | "replace" | "skip";

/** The user's choice for one conflicting source style, keyed by the source
 * style's own (source-document) id — conflicts are matched back to their
 * plan entry by that id, since two source styles could share a display
 * name only by construction the reader never produces (ids are unique
 * within one parsed document). */
export interface StyleImportResolution {
  sourceId: string;
  resolution: StyleConflictResolution;
}

export interface StyleImportSummary {
  added: string[];
  replaced: string[];
  skipped: string[];
}

function slugStyleId(name: string): string {
  return name.trim().replace(/[^A-Za-z0-9]+/g, "") || "Style";
}

function uniqueStyleId(takenIds: Set<string>, name: string): string {
  const base = slugStyleId(name);
  if (!takenIds.has(base)) {
    return base;
  }
  let suffix = 2;
  while (takenIds.has(`${base}${suffix}`)) {
    suffix += 1;
  }
  return `${base}${suffix}`;
}

/** Applies an import plan (AC-22, AC-23, AC-24): every non-conflicting
 * source style is added under a fresh id (never colliding with the current
 * document's own ids); a conflicting style is skipped, kept (no change), or
 * replaces the existing style's definition in place — reusing the existing
 * style's own id, so anything in the current document already referencing
 * it keeps resolving correctly, just against the source's formatting now.
 * A source style's `basedOn`/`next` is remapped to the fresh id when its
 * own parent/next style was imported in the same pass; otherwise it is left
 * as-is (chains to a same-named style already in the document, or degrades
 * gracefully to "no parent" if nothing resolves it — the same fallback
 * `resolveNamedStyle` already gives any dangling reference). */
export function importStyles(
  model: DocumentModel,
  sourceStyles: NamedStyle[],
  resolutions: StyleImportResolution[],
): { model: DocumentModel; summary: StyleImportSummary } {
  const plan = planStyleImport(model, sourceStyles);
  const resolutionBySourceId = new Map(resolutions.map((r) => [r.sourceId, r.resolution]));

  const catalogue = model.styles ? [...model.styles] : builtInStyles();
  const takenIds = new Set(catalogue.map((style) => style.id));
  const sourceIdToTargetId = new Map<string, string>();
  const summary: StyleImportSummary = { added: [], replaced: [], skipped: [] };

  const toApply: Array<{ source: NamedStyle; targetId: string; replaceIndex?: number }> = [];
  for (const source of plan.nonConflicting) {
    const targetId = uniqueStyleId(takenIds, source.name ?? source.id);
    takenIds.add(targetId);
    sourceIdToTargetId.set(source.id, targetId);
    toApply.push({ source, targetId });
  }
  for (const { source, existing } of plan.conflicts) {
    const resolution = resolutionBySourceId.get(source.id) ?? "keep";
    if (resolution === "skip") {
      summary.skipped.push(source.name ?? source.id);
      continue;
    }
    if (resolution === "keep") {
      summary.skipped.push(source.name ?? source.id);
      continue;
    }
    // "replace": keeps the existing style's own id.
    sourceIdToTargetId.set(source.id, existing.id);
    const replaceIndex = catalogue.findIndex((style) => style.id === existing.id);
    toApply.push({ source, targetId: existing.id, replaceIndex: replaceIndex >= 0 ? replaceIndex : undefined });
  }

  for (const { source, targetId, replaceIndex } of toApply) {
    const remapped: NamedStyle = { ...source, id: targetId };
    if (source.basedOn !== undefined) {
      remapped.basedOn = sourceIdToTargetId.get(source.basedOn) ?? source.basedOn;
    }
    if (source.next !== undefined) {
      remapped.next = sourceIdToTargetId.get(source.next) ?? source.next;
    }
    if (source.listDefinitionId !== undefined) {
      // The source's own ListDefinition is not imported (SPEC-042 scopes
      // list-style import to the catalogue entry's text formatting) — a
      // list style comes across as a plain paragraph style, matching how
      // it round-trips through .docx/.odt in the first place (neither
      // format has a first-class list-style family of its own).
      delete remapped.listDefinitionId;
      remapped.type = "paragraph";
    }
    if (replaceIndex !== undefined) {
      catalogue[replaceIndex] = remapped;
      summary.replaced.push(remapped.name ?? remapped.id);
    } else {
      catalogue.push(remapped);
      summary.added.push(remapped.name ?? remapped.id);
    }
  }

  return { model: { ...model, styles: catalogue }, summary };
}
