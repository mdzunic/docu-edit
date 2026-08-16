// --- custom styles manager (SPEC-042) ---------------------------------------
//
// Paragraph styles were SPEC-019's; this section adds the other three
// families (character, list, table), plus catalogue CRUD (new/modify/
// rename/delete) driven by the Styles pane. The "modify re-renders every
// paragraph using it, respecting inheritance" requirement (AC-4, AC-5,
// AC-11) is deliberately built the same way updateStyleFromSelection above
// already does it: re-bake every paragraph/run whose style CHAIN passes
// through the changed style, using restyledParagraph, and never touch that
// paragraph's/run's own `styleId` — only its baked formatting changes. A
// child style's descendants keep inheriting through their own basedOn link;
// they are never re-pointed at the style that was modified.
import {
  isCollapsed,
  orderedRange,
  selectAll,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import { cloneDefinitions, newListObjectId } from "./commands.multilevel";
import { touchedBlocks, type BlockAddress } from "./commands.paragraph";
import {
  applyParagraphStyle,
  restyledParagraph,
  runLayersFor,
  type RunLayers,
  type ToolbarStyleState,
} from "./commands.styles";
import { presetLevels } from "./core/listDefinitions";
import { type CellBorders } from "./core/tables";
import {
  cloneContainers,
  isTable,
  mergeRunFormatting,
  type DocumentModel,
  type NamedStyle,
  type Paragraph,
  type ParagraphFormatting,
  type RunFormatting,
  type StyleType,
  type Table,
  type TableCell,
} from "./documentModel";
import {
  allStyles,
  builtInStyles,
  diffParagraphFormatting,
  paragraphStyleBase,
  resolveCharacterStyle,
  resolveNamedStyle,
  splitRunFormatting,
  styleById,
  styleChainIncludes,
  styleTypeOf,
} from "./styles";
import {
  mapRunsInRange,
  normalizeRuns,
  paragraphAt,
  paragraphAtCaret,
  replaceParagraphAt,
  runStyleIdAtOffset,
  type CaretPosition,
} from "./textEditing";

/** A character style's resolved run formatting layered onto a paragraph's
 * own style resolution — OOXML's precedence (character style overrides
 * paragraph style; direct run formatting overrides both). */
function characterStyleBase(model: DocumentModel, paragraph: Paragraph, styleId: string | undefined): RunFormatting {
  return mergeRunFormatting(paragraphStyleBase(model, paragraph).runFormatting, resolveCharacterStyle(model, styleId));
}

/** Applies a character style to a real selection (AC-9): independent of the
 * touched paragraphs' own style, split at the range's boundaries like every
 * other run-formatting command. `styleId === ""` clears the character style
 * (back to plain paragraph-style formatting) while keeping direct
 * formatting the user typed, the same "diff survives a style change"
 * behaviour applyParagraphStyle gives paragraphs. A collapsed selection is a
 * no-op — there is no text to carry the reference. */
export function applyCharacterStyle(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const targetId = styleId === "" ? undefined : styleId;
  if (!isCollapsed(selection)) {
    mapRunsInRange(next, start, end, (run, paragraph) => {
      const oldBase = characterStyleBase(next, paragraph, run.styleId);
      const { underlay, diff } = splitRunFormatting(next, run.formatting, oldBase);
      const newBase = characterStyleBase(next, paragraph, targetId);
      const formatting = mergeRunFormatting(mergeRunFormatting(underlay, newBase), diff);
      const { styleId: _old, ...rest } = run;
      return targetId !== undefined ? { ...rest, styleId: targetId, formatting } : { ...rest, formatting };
    });
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** The character style active at the caret (AC-10's second readout,
 * alongside `toolbarStyleState`'s paragraph style): the run just before the
 * caret, else just after — same convention as `caretFormatting`. `id: ""`
 * means no character style is active. */
export function toolbarCharacterStyleState(
  model: DocumentModel,
  selection: ModelSelection,
): ToolbarStyleState {
  const { start } = orderedRange(model, selection);
  const paragraph = paragraphAtCaret(model, start);
  const styleId = paragraph ? runStyleIdAtOffset(paragraph, start.offset) : undefined;
  if (!styleId) {
    return { id: "", name: "" };
  }
  return { id: styleId, name: styleById(model, styleId)?.name ?? styleId };
}

// --- table styles -------------------------------------------------------

interface ResolvedTableStyle {
  cellBorders?: CellBorders;
  cellShadingFill?: string;
  runFormatting: RunFormatting;
}

/** A table style's basedOn chain resolved into the properties it bakes onto
 * cells: borders/shading replace an ancestor's (no partial merge — a table
 * style sets a border set outright, matching how CellBorders itself is
 * already a per-cell all-or-nothing value elsewhere in the model), run
 * formatting merges like any other style chain. */
function resolveTableStyle(model: DocumentModel, styleId: string | undefined): ResolvedTableStyle {
  const chain: NamedStyle[] = [];
  const seen = new Set<string>();
  let currentId = styleId;
  while (currentId && !seen.has(currentId)) {
    seen.add(currentId);
    const style = styleById(model, currentId);
    if (!style) {
      break;
    }
    chain.unshift(style);
    currentId = style.basedOn;
  }
  let cellBorders: CellBorders | undefined;
  let cellShadingFill: string | undefined;
  let runFormatting: RunFormatting = {};
  for (const style of chain) {
    if (style.cellBorders) {
      cellBorders = style.cellBorders;
    }
    if (style.cellShadingFill !== undefined) {
      cellShadingFill = style.cellShadingFill;
    }
    runFormatting = mergeRunFormatting(runFormatting, style.runFormatting ?? {});
  }
  return { cellBorders, cellShadingFill, runFormatting };
}

/** Bakes a resolved table style onto every non-covered cell (AC-4/AC-5's
 * "table styles ... full formatting set ... borders, shading" and
 * "re-renders every ... using it"): no per-cell direct-formatting layer is
 * tracked (table banding is out of scope — see NamedStyle.cellBorders), so
 * this simply sets the resolved border/shading on every cell and layers the
 * resolved run formatting under whatever a cell's runs already carry. */
function bakeTableStyle(table: Table, resolved: ResolvedTableStyle, styleId: string | undefined): Table {
  const hasRunFormatting = Object.keys(resolved.runFormatting).length > 0;
  const rows = table.rows.map((row) => ({
    ...row,
    cells: row.cells.map((cell) => {
      if (cell.covered) {
        return cell;
      }
      const updated: TableCell = { ...cell };
      if (resolved.cellBorders) {
        updated.borders = resolved.cellBorders;
      }
      if (resolved.cellShadingFill !== undefined) {
        updated.shadingFill = resolved.cellShadingFill;
      }
      if (hasRunFormatting) {
        updated.blocks = cell.blocks.map((paragraph) => ({
          ...paragraph,
          runs: paragraph.runs.map((run) => ({
            ...run,
            formatting: mergeRunFormatting(resolved.runFormatting, run.formatting),
          })),
        }));
      }
      return updated;
    }),
  }));
  const next: Table = { ...table, rows };
  if (styleId !== undefined) {
    next.styleId = styleId;
  } else {
    delete next.styleId;
  }
  return next;
}

function allTableAddresses(model: DocumentModel): Array<{ section: number; block: number }> {
  const addresses: Array<{ section: number; block: number }> = [];
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (isTable(block)) {
        addresses.push({ section: sectionIndex, block: blockIndex });
      }
    });
  });
  return addresses;
}

function tableAddressAtCaret(
  model: DocumentModel,
  caret: CaretPosition,
): { section: number; block: number } | undefined {
  const block = model.sections[caret.section]?.blocks[caret.block];
  return block && isTable(block) ? { section: caret.section, block: caret.block } : undefined;
}

function applyTableStyleCore(
  next: DocumentModel,
  section: number,
  block: number,
  styleId: string | undefined,
): void {
  const table = next.sections[section]?.blocks[block];
  if (!table || !isTable(table)) {
    return;
  }
  const resolved = resolveTableStyle(next, styleId);
  const updated = bakeTableStyle(table, resolved, styleId);
  next.sections[section] = {
    ...next.sections[section],
    blocks: next.sections[section].blocks.map((existing, index) => (index === block ? updated : existing)),
  };
}

/** Applies a table style to the table the caret is in/on (AC "apply"),
 * baking its borders/shading/run formatting onto every cell. `styleId ===
 * ""` clears the reference and (since there is no direct-formatting layer
 * to fall back on for tables) leaves the cells' current borders/shading as
 * plain direct formatting. */
export function applyTableStyle(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const address = tableAddressAtCaret(next, start);
  if (address) {
    applyTableStyleCore(next, address.section, address.block, styleId === "" ? undefined : styleId);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

// --- catalogue CRUD -------------------------------------------------------

function slugStyleId(name: string): string {
  return name.trim().replace(/[^A-Za-z0-9]+/g, "") || "Style";
}

function uniqueStyleId(model: DocumentModel, name: string): string {
  const base = slugStyleId(name);
  const existingIds = new Set(allStyles(model).map((style) => style.id));
  if (!existingIds.has(base)) {
    return base;
  }
  let suffix = 2;
  while (existingIds.has(`${base}${suffix}`)) {
    suffix += 1;
  }
  return `${base}${suffix}`;
}

/** The Styles pane's New/Modify dialog's field set — a superset of what any
 * one style type uses; `createStyle`/`modifyStyle` read only the fields
 * relevant to `type`. */
export interface StyleFormInput {
  name: string;
  type: StyleType;
  basedOn?: string;
  next?: string;
  paragraphFormatting?: ParagraphFormatting;
  runFormatting?: RunFormatting;
  /** List styles only: a lib/listDefinitions.ts gallery preset key
   * ("decimal", "outline", "legal", "bulletMixed") the style's own
   * ListDefinition is (re)built from. */
  listPresetKey?: string;
  /** Table styles only. */
  cellBorders?: CellBorders;
  cellShadingFill?: string;
}

/** Typed empty seed for a `.vue` view's default parameter value (espree
 * bars a generic default-parameter/`as` cast — see
 * documentModel.asCharacterFormatKey) — DocumentView.vue's saveStyleEditor
 * handler needs `payload`'s inferred type to already be `StyleFormInput`,
 * not a narrower object-literal type TypeScript would otherwise infer. */
export function noStyleFormInput(): StyleFormInput {
  return { name: "", type: "paragraph" };
}

/** "New style" (AC-3): registers a catalogue entry from the dialog's
 * (selection-seeded — the caller pre-fills `input` from the current
 * selection/caret) values, then applies it: a paragraph or list style to
 * the touched paragraphs (reusing `applyParagraphStyle`, which already
 * links a list style's numbering via `styleLinkFor` once the ListDefinition
 * below exists), a character style to a real selection, a table style to
 * the table the caret is in. */
export function createStyle(
  model: DocumentModel,
  selection: ModelSelection,
  input: StyleFormInput,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const ranged = { anchor: start, focus: end };
  const id = uniqueStyleId(next, input.name);
  const entry: NamedStyle = { id, name: input.name.trim() };
  if (input.type !== "paragraph") {
    entry.type = input.type;
  }
  if (input.basedOn) {
    entry.basedOn = input.basedOn;
  }
  if (input.next) {
    entry.next = input.next;
  }
  if (input.type === "paragraph") {
    if (input.paragraphFormatting && Object.keys(input.paragraphFormatting).length > 0) {
      entry.paragraphFormatting = input.paragraphFormatting;
    }
    if (input.runFormatting && Object.keys(input.runFormatting).length > 0) {
      entry.runFormatting = input.runFormatting;
    }
  } else if (input.type === "character") {
    if (input.runFormatting && Object.keys(input.runFormatting).length > 0) {
      entry.runFormatting = input.runFormatting;
    }
  } else if (input.type === "table") {
    if (input.cellBorders) {
      entry.cellBorders = input.cellBorders;
    }
    if (input.cellShadingFill) {
      entry.cellShadingFill = input.cellShadingFill;
    }
    if (input.runFormatting && Object.keys(input.runFormatting).length > 0) {
      entry.runFormatting = input.runFormatting;
    }
  } else if (input.type === "list") {
    const defId = newListObjectId();
    const definitions = cloneDefinitions(next);
    definitions.push({ defId, levels: presetLevels(input.listPresetKey ?? "decimal"), styleLinks: { [id]: 1 } });
    next.listDefinitions = definitions;
    entry.listDefinitionId = defId;
  }
  const catalogue = next.styles ? [...next.styles] : builtInStyles();
  catalogue.push(entry);
  next.styles = catalogue;

  if (input.type === "paragraph" || input.type === "list") {
    return applyParagraphStyle(next, ranged, id);
  }
  if (input.type === "character" && !isCollapsed(ranged)) {
    return applyCharacterStyle(next, ranged, id);
  }
  if (input.type === "table") {
    const address = tableAddressAtCaret(next, start);
    if (address) {
      applyTableStyleCore(next, address.section, address.block, id);
    }
  }
  return { model: next, selection: ranged };
}

/** "Modify style" (AC-4): rewrites the catalogue entry from the dialog's
 * values, then re-bakes every paragraph/run whose style chain passes
 * through it (AC-5), the same `styleChainIncludes`-scoped re-bake
 * `updateStyleFromSelection` already performs — which is what keeps AC-11's
 * inheritance intact: a descendant style's own paragraphs get their
 * *formatting* recomputed through the (unchanged) chain, never their
 * `styleId` rewritten to the modified style. `Modify` is a no-op for a
 * style the editor could not fully model on read (`unmodeled`) — the editor
 * has no properties of its own to show for it. */
export function modifyStyle(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
  input: StyleFormInput,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const ranged = { anchor: start, focus: end };
  const target = styleById(next, styleId);
  if (!target || target.unmodeled) {
    return { model: next, selection: ranged };
  }
  const type = styleTypeOf(target);

  const affectedParagraphs: Array<{
    address: BlockAddress;
    paragraphDiff: ParagraphFormatting;
    runLayers: RunLayers[];
  }> = [];
  if (type === "paragraph" || type === "list") {
    for (const address of touchedBlocks(next, selectAll(next))) {
      const block = paragraphAt(next, address);
      if (!block || !styleChainIncludes(next, block.styleId, styleId)) {
        continue;
      }
      const base = paragraphStyleBase(next, block);
      affectedParagraphs.push({
        address,
        paragraphDiff: diffParagraphFormatting(block.formatting ?? {}, base.paragraphFormatting),
        runLayers: runLayersFor(next, block, base),
      });
    }
  }

  const affectedRuns: Array<{
    address: BlockAddress;
    runIndex: number;
    underlay: RunFormatting;
    diff: RunFormatting;
  }> = [];
  if (type === "character") {
    for (const address of touchedBlocks(next, selectAll(next))) {
      const block = paragraphAt(next, address);
      if (!block) {
        continue;
      }
      block.runs.forEach((run, runIndex) => {
        if (!run.styleId || !styleChainIncludes(next, run.styleId, styleId)) {
          return;
        }
        const base = characterStyleBase(next, block, run.styleId);
        const { underlay, diff } = splitRunFormatting(next, run.formatting, base);
        affectedRuns.push({ address, runIndex, underlay, diff });
      });
    }
  }

  const updated: NamedStyle = { id: target.id, name: (input.name?.trim() || target.name) ?? target.id };
  if (type !== "paragraph") {
    updated.type = type;
  }
  const basedOn = input.basedOn !== undefined ? input.basedOn : target.basedOn;
  if (basedOn) {
    updated.basedOn = basedOn;
  }
  const nextId = input.next !== undefined ? input.next : target.next;
  if (nextId) {
    updated.next = nextId;
  }
  if (target.headingLevel !== undefined) {
    updated.headingLevel = target.headingLevel;
  }
  if (type === "paragraph") {
    if (input.paragraphFormatting && Object.keys(input.paragraphFormatting).length > 0) {
      updated.paragraphFormatting = input.paragraphFormatting;
    }
    if (input.runFormatting && Object.keys(input.runFormatting).length > 0) {
      updated.runFormatting = input.runFormatting;
    }
  } else if (type === "character") {
    if (input.runFormatting && Object.keys(input.runFormatting).length > 0) {
      updated.runFormatting = input.runFormatting;
    }
  } else if (type === "table") {
    if (input.cellBorders) {
      updated.cellBorders = input.cellBorders;
    }
    if (input.cellShadingFill) {
      updated.cellShadingFill = input.cellShadingFill;
    }
    if (input.runFormatting && Object.keys(input.runFormatting).length > 0) {
      updated.runFormatting = input.runFormatting;
    }
  } else if (type === "list") {
    if (input.runFormatting && Object.keys(input.runFormatting).length > 0) {
      updated.runFormatting = input.runFormatting;
    }
    updated.listDefinitionId = target.listDefinitionId;
  }

  const catalogue = next.styles ? [...next.styles] : builtInStyles();
  const index = catalogue.findIndex((style) => style.id === styleId);
  if (index >= 0) {
    catalogue[index] = updated;
  } else {
    catalogue.push(updated);
  }
  next.styles = catalogue;

  // List numbering (AC-5): markers are derived live from the ListDefinition
  // at render time (lib/listDefinitions.ts), so regenerating its levels is
  // the entire re-bake — no paragraph needs touching for the numbering
  // itself (affectedParagraphs above still re-bakes the style's run/
  // paragraph formatting, e.g. its font).
  if (type === "list" && input.listPresetKey && target.listDefinitionId) {
    next.listDefinitions = (next.listDefinitions ?? []).map((definition) =>
      definition.defId === target.listDefinitionId
        ? { ...definition, levels: presetLevels(input.listPresetKey!) }
        : definition,
    );
  }

  for (const { address, paragraphDiff, runLayers } of affectedParagraphs) {
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
  for (const { address, runIndex, underlay, diff } of affectedRuns) {
    const block = paragraphAt(next, address);
    const run = block?.runs[runIndex];
    if (!block || !run) {
      continue;
    }
    const base = characterStyleBase(next, block, run.styleId);
    const formatting = mergeRunFormatting(mergeRunFormatting(underlay, base), diff);
    replaceParagraphAt(next, address, {
      ...block,
      runs: block.runs.map((existing, index) => (index === runIndex ? { ...existing, formatting } : existing)),
    });
  }
  if (type === "table") {
    for (const address of allTableAddresses(next)) {
      const table = next.sections[address.section].blocks[address.block] as Table;
      if (!styleChainIncludes(next, table.styleId, styleId)) {
        continue;
      }
      applyTableStyleCore(next, address.section, address.block, table.styleId);
    }
  }

  return { model: next, selection: ranged };
}

/** Renaming a style (AC-6): only the catalogue entry's display name
 * changes — every reference is by `id`, so no paragraph/run/table needs
 * touching, and no re-render is needed since resolved formatting is
 * unaffected. */
export function renameStyle(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
  name: string,
): CommandResult {
  const next = cloneContainers(model);
  const catalogue = next.styles ? [...next.styles] : builtInStyles();
  const index = catalogue.findIndex((style) => style.id === styleId);
  if (index >= 0) {
    catalogue[index] = { ...catalogue[index], name: name.trim() };
    next.styles = catalogue;
  }
  return { model: next, selection };
}

/** Deleting a style (AC-8): every paragraph/run/table that referenced it
 * reverts to its `basedOn` parent (re-baked so the parent's formatting
 * actually takes effect), or to no style (displays as Normal/no character
 * style) when it has none, then the catalogue entry — and a list style's
 * ListDefinition — is removed. Confirmation is a Styles-pane concern, not
 * this command's (AC-7 asks the user before this ever runs). */
export function deleteStyle(
  model: DocumentModel,
  selection: ModelSelection,
  styleId: string,
): CommandResult {
  const next = cloneContainers(model);
  const target = styleById(next, styleId);
  if (!target) {
    return { model: next, selection };
  }
  const type = styleTypeOf(target);
  // Only a parent that still resolves to a real catalogue entry is a valid
  // fallback (AC-8) — a basedOn left dangling by an earlier delete degrades
  // to "no parent", the same as never having had one.
  const fallbackId = target.basedOn && styleById(next, target.basedOn) ? target.basedOn : undefined;

  if (type === "paragraph" || type === "list") {
    for (const address of touchedBlocks(next, selectAll(next))) {
      const block = paragraphAt(next, address);
      if (!block || block.styleId !== styleId) {
        continue;
      }
      const base = fallbackId
        ? resolveNamedStyle(next, fallbackId)
        : { paragraphFormatting: {}, runFormatting: {} };
      const oldBase = paragraphStyleBase(next, block);
      const replacement = restyledParagraph(
        block,
        base,
        diffParagraphFormatting(block.formatting ?? {}, oldBase.paragraphFormatting),
        runLayersFor(next, block, oldBase),
      );
      if (fallbackId) {
        replacement.styleId = fallbackId;
      } else {
        delete replacement.styleId;
      }
      if (target.listDefinitionId && block.list?.defId === target.listDefinitionId) {
        delete replacement.list;
      }
      replaceParagraphAt(next, address, replacement);
    }
  } else if (type === "character") {
    for (const address of touchedBlocks(next, selectAll(next))) {
      const block = paragraphAt(next, address);
      if (!block) {
        continue;
      }
      let touched = false;
      const runs = block.runs.map((run) => {
        if (run.styleId !== styleId) {
          return run;
        }
        touched = true;
        const oldBase = characterStyleBase(next, block, styleId);
        const { underlay, diff } = splitRunFormatting(next, run.formatting, oldBase);
        const newBase = characterStyleBase(next, block, fallbackId);
        const formatting = mergeRunFormatting(mergeRunFormatting(underlay, newBase), diff);
        const { styleId: _old, ...rest } = run;
        return fallbackId ? { ...rest, styleId: fallbackId, formatting } : { ...rest, formatting };
      });
      if (touched) {
        replaceParagraphAt(next, address, { ...block, runs: normalizeRuns(runs) });
      }
    }
  } else if (type === "table") {
    for (const address of allTableAddresses(next)) {
      const table = next.sections[address.section].blocks[address.block] as Table;
      if (table.styleId !== styleId) {
        continue;
      }
      const updatedTable: Table = { ...table };
      if (fallbackId) {
        updatedTable.styleId = fallbackId;
      } else {
        delete updatedTable.styleId;
      }
      next.sections[address.section] = {
        ...next.sections[address.section],
        blocks: next.sections[address.section].blocks.map((existing, index) =>
          index === address.block ? updatedTable : existing,
        ),
      };
    }
  }

  next.styles = (next.styles ?? []).filter((style) => style.id !== styleId);
  if (target.listDefinitionId) {
    next.listDefinitions = (next.listDefinitions ?? []).filter(
      (definition) => definition.defId !== target.listDefinitionId,
    );
  }
  return { model: next, selection };
}
