// --- paragraph styles (SPEC-019) --------------------------------------------
import {
  orderedRange,
  selectAll,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import { definitionIn, styleLinkFor } from "./commands.multilevel";
import { touchedBlocks, type BlockAddress } from "./commands.paragraph";
import { definedLevelType } from "./core/listDefinitions";
import {
  cloneContainers,
  mergeParagraphFormatting,
  mergeRunFormatting,
  type DocumentModel,
  type NamedStyle,
  type Paragraph,
  type ParagraphFormatting,
  type RunFormatting,
} from "./documentModel";
import {
  builtInStyles,
  diffParagraphFormatting,
  namedStyleIdFor,
  paragraphStyleBase,
  resolveNamedStyle,
  splitRunFormatting,
  styleById,
  styleChainIncludes,
  styleDisplayName,
  type ResolvedNamedStyle,
} from "./styles";
import { paragraphAt, paragraphAtCaret, replaceParagraphAt } from "./textEditing";

/** A run's two layers relative to its old style base (see
 * styles.splitRunFormatting): the default-derived underlay and the genuine
 * direct-formatting diff. Recorded per run so a catalogue change can
 * re-bake the run as `underlay ⊕ new style ⊕ diff`. */
// RunLayers/runLayersFor/restyledParagraph are shared with
// commands.customStyles and commands.clearFormatting (their re-bakes are
// deliberately built on the same layering) — NOT re-exported by the
// commands.ts barrel; see the barrel's header.
export interface RunLayers {
  underlay: RunFormatting;
  diff: RunFormatting;
}

export function runLayersFor(model: DocumentModel, block: Paragraph, base: ResolvedNamedStyle): RunLayers[] {
  return block.runs.map((run) => splitRunFormatting(model, run.formatting, base.runFormatting));
}

/** A paragraph rebuilt around a style resolution, with each run's layers
 * put back in order — underlay beneath the style, direct diff on top — and
 * the paragraph formatting as `resolved ⊕ diff`; the heading level follows
 * the resolution. Shared by apply (layers split against the old style) and
 * update (layers recorded before the catalogue changed). */
export function restyledParagraph(
  block: Paragraph,
  resolved: ResolvedNamedStyle,
  paragraphDiff: ParagraphFormatting,
  runLayers: RunLayers[],
): Paragraph {
  const replacement: Paragraph = {
    ...block,
    runs: block.runs.map((run, index) => {
      const layers = runLayers[index] ?? { underlay: {}, diff: {} };
      return {
        ...run,
        formatting: mergeRunFormatting(
          mergeRunFormatting(layers.underlay, resolved.runFormatting),
          layers.diff,
        ),
      };
    }),
  };
  const formatting = mergeParagraphFormatting(resolved.paragraphFormatting, paragraphDiff);
  if (Object.keys(formatting).length > 0) {
    replacement.formatting = formatting;
  } else {
    delete replacement.formatting;
  }
  if (resolved.headingLevel !== undefined) {
    replacement.headingLevel = resolved.headingLevel;
  } else {
    delete replacement.headingLevel;
  }
  return replacement;
}

/** Applies a paragraph style to every paragraph the selection touches
 * (AC-3): each paragraph's baked formatting becomes the new style's
 * resolution with its *direct* formatting — the diff between its old baked
 * formatting and its old style's resolution — layered back on top, so
 * direct formatting survives the style change (AC-9). A built-in style
 * applied to a document that has a catalogue of its own is added to it, so
 * the picker lists it and a save declares it. */
export function applyParagraphStyle(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  // Materialize the applied style into the model's catalogue: a document
  // without one gets the full built-in set (the picker's implicit list
  // becomes explicit), a document with its own gets the built-in definition
  // appended. This is what lets the writers declare the style — a paragraph
  // must never reference a style the saved file doesn't define (AC-15).
  if (!next.styles?.some((style) => style.id === styleId)) {
    const catalogue = next.styles ? [...next.styles] : builtInStyles();
    if (!catalogue.some((style) => style.id === styleId)) {
      const builtIn = builtInStyles().find((style) => style.id === styleId);
      if (builtIn) {
        catalogue.push(builtIn);
      }
    }
    next.styles = catalogue;
  }
  const resolved = resolveNamedStyle(next, styleId);
  // Heading numbering (SPEC-032 AC-13): applying a style linked to a list
  // definition makes the paragraph a defined item at the linked level (so a
  // new Heading 2 numbers itself, and re-styling Heading 1 -> Heading 2
  // demotes its number, AC-31); applying an unlinked style removes a
  // membership that only existed through such a link.
  const link = styleLinkFor(next, styleId);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const base = paragraphStyleBase(next, block);
    const replacement = restyledParagraph(
      block,
      resolved,
      diffParagraphFormatting(block.formatting ?? {}, base.paragraphFormatting),
      runLayersFor(next, block, base),
    );
    replacement.styleId = styleId;
    if (link) {
      replacement.list = {
        type: definedLevelType(link.def, link.level),
        level: link.level,
        defId: link.def.defId,
        listId: link.def.defId,
      };
    } else if (replacement.list?.defId !== undefined) {
      const memberOf = definitionIn(next, replacement.list.defId);
      if (memberOf?.styleLinks) {
        delete replacement.list;
      }
    }
    replaceParagraphAt(next, address, replacement);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** "Reset style" (AC-14): removes direct formatting from every touched
 * paragraph, leaving exactly what its style's resolution produces. A `.odt`
 * paragraph referencing an automatic style is re-pointed at the named style
 * underneath (the automatic style *was* the direct formatting). A heading
 * level the style chain doesn't set (a direct outline level) is kept — it
 * is structure, not appearance. */
export function resetParagraphStyle(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const namedId = namedStyleIdFor(next, block.styleId);
    const base = paragraphStyleBase(next, block);
    // Empty diffs: the underlay (the baked document defaults) survives the
    // reset, the direct formatting does not.
    const replacement = restyledParagraph(
      block,
      base,
      {},
      runLayersFor(next, block, base).map((layers) => ({ underlay: layers.underlay, diff: {} })),
    );
    if (base.headingLevel === undefined && block.headingLevel !== undefined) {
      replacement.headingLevel = block.headingLevel;
    }
    if (namedId !== undefined) {
      replacement.styleId = namedId;
    } else {
      delete replacement.styleId;
    }
    replaceParagraphAt(next, address, replacement);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** "Update <style> to match" (AC-12, AC-13): redefines `styleId` so its
 * resolution reproduces the current (first touched) paragraph's formatting,
 * then re-bakes every paragraph whose style chain passes through it —
 * keeping each one's direct formatting layered on the new resolution. A
 * document without a catalogue of its own gets one seeded from the
 * built-ins, so the picker keeps its full list. */
export function updateStyleFromSelection(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const preserved = { model: next, selection: { anchor: start, focus: end } };
  const target = styleById(next, styleId);
  const source = paragraphAtCaret(next, start);
  if (!target || !source) {
    return preserved;
  }

  // The style's new own formatting: the source paragraph's effective
  // formatting minus what the style's parents already provide. Run
  // formatting comes from the paragraph's first run; an empty paragraph
  // keeps the style's current run formatting.
  const parentResolved = resolveNamedStyle(next, target.basedOn);
  const ownParagraph = diffParagraphFormatting(source.formatting ?? {}, parentResolved.paragraphFormatting);
  const firstRun = source.runs[0];
  const ownRun = firstRun
    ? splitRunFormatting(next, firstRun.formatting, parentResolved.runFormatting).diff
    : (target.runFormatting ?? {});

  // Per-paragraph direct-formatting layers, recorded against the OLD
  // catalogue before it changes below.
  const affected: {
    address: BlockAddress;
    paragraphDiff: ParagraphFormatting;
    runLayers: RunLayers[];
  }[] = [];
  for (const address of touchedBlocks(next, selectAll(next))) {
    const block = paragraphAt(next, address);
    if (!block || !styleChainIncludes(next, block.styleId, styleId)) {
      continue;
    }
    const base = paragraphStyleBase(next, block);
    affected.push({
      address,
      paragraphDiff: diffParagraphFormatting(block.formatting ?? {}, base.paragraphFormatting),
      runLayers: runLayersFor(next, block, base),
    });
  }

  const updated: NamedStyle = { id: target.id };
  if (target.name !== undefined) {
    updated.name = target.name;
  }
  if (target.basedOn !== undefined) {
    updated.basedOn = target.basedOn;
  }
  if (Object.keys(ownParagraph).length > 0) {
    updated.paragraphFormatting = ownParagraph;
  }
  if (Object.keys(ownRun).length > 0) {
    updated.runFormatting = ownRun;
  }
  if (target.headingLevel !== undefined) {
    updated.headingLevel = target.headingLevel;
  }
  const catalogue = next.styles ? [...next.styles] : builtInStyles();
  const index = catalogue.findIndex((style) => style.id === styleId);
  if (index >= 0) {
    catalogue[index] = updated;
  } else {
    catalogue.push(updated);
  }
  next.styles = catalogue;

  // Re-bake with the new catalogue: resolution changed, the layers did not.
  for (const { address, paragraphDiff, runLayers } of affected) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    replaceParagraphAt(
      next,
      address,
      restyledParagraph(block, paragraphStyleBase(next, block), paragraphDiff, runLayers),
    );
  }
  return preserved;
}

/** What the style picker shows for the selection (AC-1): the first touched
 * paragraph's style — its resolved *named* style id and display name. */
export interface ToolbarStyleState {
  /** The named style id the current paragraph resolves to; "" when it
   * resolves to none (displays as Normal). */
  id: string;
  name: string;
}

/** The initial/no-document state — a typed seed for `.vue` refs (espree
 * bars the annotation, same constraint as noSelection). */
export function emptyToolbarStyleState(): ToolbarStyleState {
  return { id: "", name: "Normal" };
}

export function toolbarStyleState(
  model: DocumentModel,
  selection: ModelSelection,
): ToolbarStyleState {
  const { start } = orderedRange(model, selection);
  const paragraph = paragraphAtCaret(model, start);
  return {
    id: namedStyleIdFor(model, paragraph?.styleId) ?? "",
    name: styleDisplayName(model, paragraph?.styleId),
  };
}
