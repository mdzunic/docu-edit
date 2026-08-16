// --- multilevel lists and numbering control (SPEC-032) ----------------------
import {
  orderedRange,
  selectAll,
  type CommandResult,
  type ModelSelection,
} from "./commands.editing";
import {
  firstTouchedFormatting,
  INDENT_STEP_CM,
  snapCm,
  touchedBlocks,
  updateTouchedParagraphs,
  type BlockAddress,
} from "./commands.paragraph";
import {
  defaultDefinitionLevels,
  definedLevelType,
  headingLinkedLevels,
  headingStyleLinks,
  presetLevels,
} from "./core/listDefinitions";
import {
  clampListLevel,
  isBulletStyle,
  planListLevels,
  type ListDefinition,
  type ListFormatting,
  type ListLevelDefinition,
  type PlannedListLevel,
} from "./core/lists";
import { cmToTwips } from "./core/units";
import {
  cloneContainers,
  DEFAULT_TAB_STOP_CM,
  type DocumentModel,
  type ParagraphFormatting,
  type TabStop,
} from "./documentModel";
import { namedStyleIdFor } from "./styles";
import {
  paragraphAt,
  paragraphAtCaret,
  replaceParagraphAt,
  type CaretPosition,
} from "./textEditing";

// newListObjectId/cloneDefinitions (and definitionIn/styleLinkFor below) are
// shared with commands.styles/commands.customStyles — deliberately NOT
// re-exported by the commands.ts barrel; see the barrel's header.
export function newListObjectId(): string {
  return crypto.randomUUID();
}

/** The model's definitions array cloned for mutation (the listDefinitions
 * sibling of cloneContainers — the input model's array is never touched). */
export function cloneDefinitions(model: DocumentModel): ListDefinition[] {
  return [...(model.listDefinitions ?? [])];
}

export function definitionIn(model: DocumentModel, defId: string | undefined): ListDefinition | undefined {
  if (defId === undefined) {
    return undefined;
  }
  return model.listDefinitions?.find((def) => def.defId === defId);
}

/** The definition (and level) a paragraph style is linked to, scanning
 * newest-first so re-linking headings supersedes an older link (AC-13). */
export function styleLinkFor(
  model: DocumentModel,
  styleId: string,
): { def: ListDefinition; level: number } | undefined {
  const defs = model.listDefinitions ?? [];
  for (let i = defs.length - 1; i >= 0; i--) {
    const level = defs[i].styleLinks?.[styleId];
    if (level !== undefined) {
      return { def: defs[i], level: clampListLevel(level) };
    }
  }
  return undefined;
}

/** Normalises the Define dialog's nine levels into canonical stored form:
 * defaults are elided (empty prefix/suffix, includeLevels 1, startAt 1,
 * left alignment, editor-default positions), lengths snap to the
 * cross-format grid — so an untouched save/re-open reproduces the
 * definition bit-for-bit. */
export function normalizeDefinitionLevels(levels: ListLevelDefinition[]): ListLevelDefinition[] {
  const out: ListLevelDefinition[] = [];
  for (let level = 1; level <= 9; level++) {
    const input = levels[level - 1] ?? { style: "decimal" };
    const cleaned: ListLevelDefinition = { style: input.style };
    if (input.prefix !== undefined && input.prefix !== "") {
      cleaned.prefix = input.prefix;
    }
    if (input.suffix !== undefined && input.suffix !== "") {
      cleaned.suffix = input.suffix;
    }
    const include = Math.round(input.includeLevels ?? 1);
    if (Number.isFinite(include) && include > 1) {
      cleaned.includeLevels = Math.min(include, level);
    }
    const startAt = Math.round(input.startAt ?? 1);
    if (Number.isFinite(startAt) && startAt >= 0 && startAt !== 1) {
      cleaned.startAt = startAt;
    }
    if (input.alignment !== undefined && input.alignment !== "left") {
      cleaned.alignment = input.alignment;
    }
    const defaultText = level * INDENT_STEP_CM;
    let textIndent = defaultText;
    if (input.textIndentCm !== undefined && Number.isFinite(input.textIndentCm)) {
      textIndent = snapCm(Math.max(0, input.textIndentCm));
      if (textIndent !== snapCm(defaultText)) {
        cleaned.textIndentCm = textIndent;
      } else {
        textIndent = defaultText;
      }
    }
    if (input.numberIndentCm !== undefined && Number.isFinite(input.numberIndentCm)) {
      const numberIndent = snapCm(Math.max(0, input.numberIndentCm));
      if (numberIndent !== snapCm(Math.max(0, textIndent - 0.635))) {
        cleaned.numberIndentCm = numberIndent;
      }
    }
    out.push(cleaned);
  }
  return out;
}

/** Every paragraph address in the document, in document order, descending
 * into table cells — the walk the whole-document list operations share. */
function allParagraphAddresses(model: DocumentModel): BlockAddress[] {
  return touchedBlocks(model, selectAll(model));
}

function sameAddress(a: BlockAddress, b: BlockAddress): boolean {
  if (a.section !== b.section || a.block !== b.block) {
    return false;
  }
  if (!a.cell || !b.cell) {
    return !a.cell && !b.cell;
  }
  return a.cell.row === b.cell.row && a.cell.col === b.cell.col && a.cell.block === b.cell.block;
}

/** Whether two addresses share a numbering context (SPEC-022's rule): both
 * top-level, or both inside the same table cell. */
function sameListContext(a: BlockAddress, b: BlockAddress): boolean {
  if (!a.cell && !b.cell) {
    return true;
  }
  if (!a.cell || !b.cell) {
    return false;
  }
  return (
    a.section === b.section &&
    a.block === b.block &&
    a.cell.row === b.cell.row &&
    a.cell.col === b.cell.col
  );
}

/** The addresses of the target's numbering context at or after (`forward`)
 * or strictly before, nearest first (`backward`) the target. */
function contextAddresses(
  model: DocumentModel,
  target: BlockAddress,
  direction: "forward" | "backward",
): BlockAddress[] {
  const all = allParagraphAddresses(model).filter((address) => sameListContext(address, target));
  const index = all.findIndex((address) => sameAddress(address, target));
  if (index === -1) {
    return [];
  }
  if (direction === "forward") {
    return all.slice(index);
  }
  return all.slice(0, index).reverse();
}

/** The BlockAddress a caret's paragraph lives at. */
function addressOfCaret(caret: CaretPosition): BlockAddress {
  if (caret.cell) {
    return { section: caret.section, block: caret.block, cell: { ...caret.cell } };
  }
  return { section: caret.section, block: caret.block };
}

function levelsFromPlanned(planned: PlannedListLevel[]): ListLevelDefinition[] {
  return planned.map((level, index) => {
    const def: ListLevelDefinition = { style: level.style };
    if (!isBulletStyle(level.style)) {
      def.suffix = ".";
    }
    if (level.startAt !== 1) {
      def.startAt = level.startAt;
    }
    if (cmToTwips(level.indentCm) !== (index + 1) * 720) {
      def.textIndentCm = level.indentCm;
    }
    return def;
  });
}

/** Upgrades the contiguous ad-hoc (SPEC-018) run around `target` into a
 * defined list: a definition planned from the run's own level styles plus a
 * fresh instance, so the numbering-control operations below always work on
 * instances. Returns the minted ids, or null when the target is not an
 * ad-hoc list paragraph. */
function upgradeAdHocRun(
  model: DocumentModel,
  target: BlockAddress,
): { defId: string; listId: string } | null {
  const block = paragraphAt(model, target);
  if (!block?.list || block.list.defId !== undefined) {
    return null;
  }
  const forward = contextAddresses(model, target, "forward");
  const backward = contextAddresses(model, target, "backward");
  const isAdHocItem = (address: BlockAddress) => {
    const paragraph = paragraphAt(model, address);
    return !!paragraph?.list && paragraph.list.defId === undefined;
  };
  // Contiguity within the context: the run extends while addresses are
  // consecutive blocks of the same block list and stay ad-hoc items.
  const run: BlockAddress[] = [target];
  let previous = target;
  for (const address of forward.slice(1)) {
    const at = address.cell ? address.cell.block : address.block;
    const prev = previous.cell ? previous.cell.block : previous.block;
    if (address.section !== previous.section || at !== prev + 1 || !isAdHocItem(address)) {
      break;
    }
    run.push(address);
    previous = address;
  }
  previous = target;
  for (const address of backward) {
    const at = address.cell ? address.cell.block : address.block;
    const prev = previous.cell ? previous.cell.block : previous.block;
    if (address.section !== previous.section || at !== prev - 1 || !isAdHocItem(address)) {
      break;
    }
    run.unshift(address);
    previous = address;
  }
  const items = run
    .map((address) => paragraphAt(model, address)?.list)
    .filter((list): list is ListFormatting => list !== undefined);
  const def: ListDefinition = {
    defId: newListObjectId(),
    levels: levelsFromPlanned(planListLevels(items)),
  };
  const listId = newListObjectId();
  model.listDefinitions = [...cloneDefinitions(model), def];
  for (const address of run) {
    const paragraph = paragraphAt(model, address);
    if (!paragraph?.list) {
      continue;
    }
    const list: ListFormatting = {
      type: paragraph.list.type,
      level: clampListLevel(paragraph.list.level),
      defId: def.defId,
      listId,
    };
    replaceParagraphAt(model, address, { ...paragraph, list });
  }
  return { defId: def.defId, listId };
}

/** Re-points `target` and every following item of its instance at new
 * instance ids — the shared core of restart/continue/set-value. `adoptDef`
 * (continue) also switches the definition, recomputing each item's type
 * from the adopted levels; `startAtFirst` (set value) restarts the target
 * item at an explicit ordinal. */
function reinstanceForward(
  model: DocumentModel,
  target: BlockAddress,
  oldKey: { defId: string; listId?: string },
  next: { defId: string; listId: string },
  startAtFirst?: number,
): void {
  const def = definitionIn(model, next.defId);
  let first = true;
  for (const address of contextAddresses(model, target, "forward")) {
    const paragraph = paragraphAt(model, address);
    const list = paragraph?.list;
    if (!paragraph || !list || list.defId !== oldKey.defId || list.listId !== oldKey.listId) {
      continue;
    }
    const level = clampListLevel(list.level);
    const updated: ListFormatting = {
      type: def ? definedLevelType(def, level) : list.type,
      level,
      defId: next.defId,
      listId: next.listId,
    };
    if (first && startAtFirst !== undefined) {
      updated.startAt = startAtFirst;
    }
    first = false;
    replaceParagraphAt(model, address, { ...paragraph, list: updated });
  }
}

/** Applies a gallery preset (AC-1..4): every touched paragraph becomes an
 * item of a NEW definition + instance minted from the preset — new, so
 * editing this list's levels later never touches another list started from
 * the same preset (AC-18). Paragraphs already in a list keep their level;
 * plain paragraphs join at level 1. */
export function applyMultilevelPreset(
  model: DocumentModel,
  selection: ModelSelection,
  presetKey: string,
): CommandResult {
  return applyDefinedList(model, selection, presetLevels(presetKey));
}

/** "Define new multilevel list…" OK (AC-5..12): with the selection in a
 * defined list, the list's own definition is updated in place — every item
 * of that list at every level re-renders from the new levels (AC-17), and
 * no other list changes (AC-18). Otherwise the touched paragraphs become a
 * new defined list with these levels (upgrading an ad-hoc SPEC-018 list or
 * starting one from plain paragraphs). */
export function applyListDefinition(
  model: DocumentModel,
  selection: ModelSelection,
  levels: ListLevelDefinition[],
): CommandResult {
  const cleaned = normalizeDefinitionLevels(levels);
  const { start } = orderedRange(model, selection);
  const target = paragraphAtCaret(model, start);
  const existing = definitionIn(model, target?.list?.defId);
  if (!existing) {
    return applyDefinedList(model, selection, cleaned);
  }
  const next = cloneContainers(model);
  const range = orderedRange(next, selection);
  const updated: ListDefinition = { ...existing, levels: cleaned };
  next.listDefinitions = cloneDefinitions(next).map((def) =>
    def.defId === existing.defId ? updated : def,
  );
  // A level's style may have crossed the bullet/number line: re-derive
  // every member item's type from the new levels.
  for (const address of allParagraphAddresses(next)) {
    const paragraph = paragraphAt(next, address);
    const list = paragraph?.list;
    if (!paragraph || !list || list.defId !== existing.defId) {
      continue;
    }
    const type = definedLevelType(updated, list.level);
    if (type === list.type) {
      continue;
    }
    replaceParagraphAt(next, address, { ...paragraph, list: { ...list, type } });
  }
  return { model: next, selection: { anchor: range.start, focus: range.end } };
}

function applyDefinedList(
  model: DocumentModel,
  selection: ModelSelection,
  levels: ListLevelDefinition[],
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const def: ListDefinition = { defId: newListObjectId(), levels };
  next.listDefinitions = [...cloneDefinitions(next), def];
  const listId = newListObjectId();
  for (const address of touchedBlocks(next, { anchor: start, focus: end })) {
    const block = paragraphAt(next, address);
    if (!block) {
      continue;
    }
    const level = clampListLevel(block.list?.level ?? 1);
    const list: ListFormatting = {
      type: definedLevelType(def, level),
      level,
      defId: def.defId,
      listId,
    };
    replaceParagraphAt(next, address, { ...block, list });
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** "Number headings 1–3" (AC-13): creates a definition whose levels 1–3 are
 * linked to the built-in Heading 1–3 styles, then numbers every existing
 * heading paragraph. All linked headings share one instance (keyed by the
 * definition itself), so their numbering is continuous document-wide and
 * survives delete/move/demote (AC-29..31) — markers are derived, never
 * stored. Future paragraphs pick the link up through applyParagraphStyle. */
export function linkHeadingNumbering(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const def: ListDefinition = {
    defId: newListObjectId(),
    levels: headingLinkedLevels(),
    styleLinks: headingStyleLinks(),
  };
  next.listDefinitions = [...cloneDefinitions(next), def];
  for (const address of allParagraphAddresses(next)) {
    const paragraph = paragraphAt(next, address);
    if (!paragraph) {
      continue;
    }
    const namedId = namedStyleIdFor(next, paragraph.styleId);
    const level = namedId !== undefined ? def.styleLinks?.[namedId] : undefined;
    if (level === undefined) {
      continue;
    }
    const list: ListFormatting = {
      type: definedLevelType(def, level),
      level: clampListLevel(level),
      defId: def.defId,
      listId: def.defId,
    };
    replaceParagraphAt(next, address, { ...paragraph, list });
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** "Restart numbering at 1" (AC-14): the clicked item and every following
 * item of the same instance move to a fresh instance, whose counters start
 * from the definition's own start values. An ad-hoc SPEC-018 list is
 * upgraded to a defined one first, since only instances can restart
 * mid-run. */
export function restartListNumbering(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  return reinstanceCommand(model, selection, undefined);
}

/** "Set numbering value…" (AC-16): like a restart, but the clicked item
 * carries an explicit start ordinal. */
export function setListNumberingValue(
  model: DocumentModel,
  selection: ModelSelection,
  value: number,
): CommandResult {
  const startAt = Math.max(0, Math.round(value));
  if (!Number.isFinite(startAt)) {
    return { model: cloneContainers(model), selection };
  }
  return reinstanceCommand(model, selection, startAt);
}

function reinstanceCommand(
  model: DocumentModel,
  selection: ModelSelection,
  startAtFirst: number | undefined,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const preserved = { model: next, selection: { anchor: start, focus: end } };
  const target = addressOfCaret(start);
  const paragraph = paragraphAt(next, target);
  if (!paragraph?.list) {
    return preserved;
  }
  if (paragraph.list.defId === undefined) {
    upgradeAdHocRun(next, target);
  }
  const list = paragraphAt(next, target)?.list;
  if (!list?.defId) {
    return preserved;
  }
  reinstanceForward(
    next,
    target,
    { defId: list.defId, listId: list.listId },
    { defId: list.defId, listId: newListObjectId() },
    startAtFirst,
  );
  return preserved;
}

/** "Continue previous numbering" (AC-15): the clicked item (and the rest of
 * its instance) rejoins the nearest preceding list's instance, continuing
 * its numbering across whatever separated them. Works from either side of
 * the SPEC-018 divide: an ad-hoc predecessor is upgraded to a defined list
 * first, and an ad-hoc target run adopts the predecessor's definition
 * outright. No-op when nothing numbered precedes the item. */
export function continueListNumbering(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const preserved = { model: next, selection: { anchor: start, focus: end } };
  const target = addressOfCaret(start);
  const targetParagraph = paragraphAt(next, target);
  if (!targetParagraph?.list) {
    return preserved;
  }
  // The nearest preceding list paragraph outside the target's own instance
  // (for an ad-hoc target: outside its contiguous run).
  const targetKey = targetParagraph.list.defId
    ? { defId: targetParagraph.list.defId, listId: targetParagraph.list.listId }
    : null;
  let predecessor: BlockAddress | undefined;
  let previous = target;
  let insideOwnRun = targetKey === null;
  for (const address of contextAddresses(next, target, "backward")) {
    const list = paragraphAt(next, address)?.list;
    if (!list) {
      insideOwnRun = false;
      previous = address;
      continue;
    }
    if (targetKey) {
      if (list.defId !== targetKey.defId || list.listId !== targetKey.listId) {
        predecessor = address;
        break;
      }
    } else {
      // Ad-hoc target: skip the target's own contiguous run, then take the
      // first list paragraph found.
      const at = address.cell ? address.cell.block : address.block;
      const prev = previous.cell ? previous.cell.block : previous.block;
      const contiguous = insideOwnRun && address.section === previous.section && at === prev - 1;
      if (contiguous && list.defId === undefined) {
        previous = address;
        continue;
      }
      predecessor = address;
      break;
    }
    previous = address;
  }
  if (!predecessor) {
    return preserved;
  }
  let adoptedList = paragraphAt(next, predecessor)?.list;
  if (adoptedList && adoptedList.defId === undefined) {
    upgradeAdHocRun(next, predecessor);
    adoptedList = paragraphAt(next, predecessor)?.list;
  }
  if (!adoptedList?.defId || !adoptedList.listId) {
    return preserved;
  }
  const adopted = { defId: adoptedList.defId, listId: adoptedList.listId };
  if (targetKey) {
    reinstanceForward(next, target, targetKey, adopted);
    return preserved;
  }
  // Ad-hoc target: the target and the rest of its contiguous run adopt the
  // predecessor's definition and instance directly.
  const def = definitionIn(next, adopted.defId);
  previous = target;
  for (const address of contextAddresses(next, target, "forward")) {
    const paragraph = paragraphAt(next, address);
    const list = paragraph?.list;
    const at = address.cell ? address.cell.block : address.block;
    const prev = previous.cell ? previous.cell.block : previous.block;
    const contiguous =
      sameAddress(address, target) || (address.section === previous.section && at === prev + 1);
    if (!paragraph || !list || list.defId !== undefined || !contiguous) {
      break;
    }
    const level = clampListLevel(list.level);
    replaceParagraphAt(next, address, {
      ...paragraph,
      list: {
        type: def ? definedLevelType(def, level) : list.type,
        level,
        defId: adopted.defId,
        listId: adopted.listId,
      },
    });
    previous = address;
  }
  return preserved;
}

/** The Define dialog's initial levels (a query): the selection's own
 * definition when it sits in a defined list, the levels planned from its
 * ad-hoc run when in a SPEC-018 list, else the decimal preset. */
export function selectionListDefinitionLevels(
  model: DocumentModel,
  selection: ModelSelection,
): ListLevelDefinition[] {
  const { start } = orderedRange(model, selection);
  const list = paragraphAtCaret(model, start)?.list;
  const def = definitionIn(model, list?.defId);
  if (def) {
    return structuredClone(def.levels);
  }
  if (list) {
    const target = addressOfCaret(start);
    const run = contextAddresses(model, target, "forward")
      .concat(contextAddresses(model, target, "backward"))
      .map((address) => paragraphAt(model, address)?.list)
      .filter((item): item is ListFormatting => !!item && item.defId === undefined);
    if (run.length > 0) {
      return levelsFromPlanned(planListLevels(run));
    }
  }
  return defaultDefinitionLevels();
}

// --- custom tab stops (SPEC-032 AC-19..25) ----------------------------------

/** Replaces the custom tab stops of every touched paragraph (the tab
 * dialog's OK, and the ruler's add/remove built on top of it). Stops are
 * snapped to the cross-format grid, sorted by position and deduplicated
 * (last one at a position wins, like Word); an empty list removes the
 * field. */
export function setTabStops(
  model: DocumentModel,
  selection: ModelSelection,
  stops: TabStop[],
): CommandResult {
  const byPosition = new Map<number, TabStop>();
  for (const stop of stops) {
    if (!Number.isFinite(stop.positionCm) || stop.positionCm < 0) {
      continue;
    }
    const positionCm = snapCm(stop.positionCm);
    const cleaned: TabStop = { positionCm, alignment: stop.alignment };
    if (stop.leader !== undefined) {
      cleaned.leader = stop.leader;
    }
    byPosition.set(positionCm, cleaned);
  }
  const normalized = [...byPosition.values()].sort((a, b) => a.positionCm - b.positionCm);
  return updateTouchedParagraphs(model, selection, (formatting) => {
    const nextFormatting: ParagraphFormatting = { ...formatting };
    if (normalized.length > 0) {
      nextFormatting.tabStops = normalized;
    } else {
      delete nextFormatting.tabStops;
    }
    return nextFormatting;
  });
}

/** The first touched paragraph's custom tab stops (a query) — what the tab
 * dialog opens with and the ruler paints. */
export function selectionTabStops(model: DocumentModel, selection: ModelSelection): TabStop[] {
  return (firstTouchedFormatting(model, selection).tabStops ?? []).map((stop) => ({ ...stop }));
}

/** The Tabs dialog's OK (SPEC-054): sets the touched paragraphs' stops and,
 * in the same undo step, the document's default tab interval. A value equal
 * to the 1.27 cm default is stored as absent (deleting the field) so a
 * document that never changed it round-trips without a defaultTabStop. */
export function setTabStopsWithDefault(
  model: DocumentModel,
  selection: ModelSelection,
  stops: TabStop[],
  defaultTabStopCm: number,
): CommandResult {
  const result = setTabStops(model, selection, stops);
  const nextModel: DocumentModel = {
    ...result.model,
    defaultStyles: { ...result.model.defaultStyles },
  };
  if (Number.isFinite(defaultTabStopCm) && Math.abs(defaultTabStopCm - DEFAULT_TAB_STOP_CM) > 1e-9) {
    // Stored as the user's raw centimetre value (unlike per-stop positions,
    // which snap to the twip grid): it round-trips through twips (.docx) /
    // centimetre strings (.odt) within tolerance anyway, and the dialog
    // should re-open showing exactly what was typed (AC-14).
    nextModel.defaultStyles.defaultTabStopCm = defaultTabStopCm;
  } else {
    delete nextModel.defaultStyles.defaultTabStopCm;
  }
  return { model: nextModel, selection: result.selection };
}

/** The document's effective default tab interval in cm (SPEC-054): the stored
 * value, or the 1.27 cm default when absent — what the Tabs dialog prefills. */
export function documentDefaultTabStopCm(model: DocumentModel): number {
  return model.defaultStyles.defaultTabStopCm ?? DEFAULT_TAB_STOP_CM;
}
