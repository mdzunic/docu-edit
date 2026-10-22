// The spreadsheet undo stack and edit pipeline (SPEC-121 "The undo stack",
// "The edit pipeline") — THIS FILE is their normative home in code. Every
// workbook command is one step: the caller `record`s BEFORE mutating the live
// model and `commit`s after; undo/redo swap snapshots. Per open view (a fresh
// stack per mounted view), never persisted, capped at MAX_UNDO_STEPS.
//
// Snapshots are by reference. A snapshot holds, for EVERY sheet, a shallow
// copy whose object/array fields are structuredCloned — except
// SHEET_BY_REFERENCE_FIELDS (`rows`, `sourceXml`, SPEC-124's `columns`,
// `merges` and `frozen`, SPEC-128's `autoFilter` and `sortState`,
// SPEC-129's `conditionalFormats`, SPEC-130's `validations`, SPEC-131's
// `drawings`, SPEC-133's `tables`, SPEC-135's `notes`, `threads` and
// `hyperlinks` and SPEC-136's `pageSetup`, which every
// command replaces and never mutates), kept by reference because
// copy-on-write makes that safe — so an undo
// restores them by identity — plus every WORKBOOK_SNAPSHOT_FIELDS value by
// reference (each a primitive or copy-on-write: `styles` and SPEC-129's
// `dxfs` grow by replacement, `opaque` and `odsContentEnvelope` are replaced
// when they change, SPEC-131's `media` is replaced when an image goes in,
// SPEC-133's `names` is copy-on-write, SPEC-135's `persons` is replaced
// when a new author joins). Taking one costs O(sheets + small metadata), never O(cells) or
// O(package). A later spec that adds a sheet field able to grow with the
// cell count makes it copy-on-write and appends it to
// SHEET_BY_REFERENCE_FIELDS; one that adds workbook-level state appends to
// WORKBOOK_SNAPSHOT_FIELDS. `forceFullCalcOnLoad` is deliberately absent
// (SPEC-120 never clears it).
//
// The pipeline (`commit`, and every undo/redo):
// 1. Every sheet whose `rows` differ (by reference) between before and after
//    contributes SPEC-120's changedCellsBetween keys; a sheet on one side
//    only contributes all its stored cells. When the sheet list, a name or
//    the order changed, invalidateDependencyGraph runs first; otherwise
//    updateGraphCells follows the keys.
//    SPEC-133: every formula cell mentioning a name or table the step
//    changed (nameDependents.ts) joins the keys, its graph node rebuilt.
//    When a step hid or showed a row (recalc.ts's rowVisibilityChanged —
//    the user's `hidden` or SPEC-128's `filterHidden`, so every filter
//    command counts), every formula cell calling SUBTOTAL or AGGREGATE joins
//    the keys (visibilitySensitiveCells, SPEC-127): their results follow
//    which rows show, though no cell they read changed.
// 2. deps.recalc(keys ∪ alsoRecalc, done) — an "edit" pass, possibly sliced.
// 3. Every sheet whose `rows` reference differs from the reference snapshot
//    (the step's pre-command state for a command; the installed state for an
//    undo/redo, whose own dirty flags it restores) is replaced with
//    `dirty: true` — once now (the command's changes) and again when the
//    pass is done (its recalculated values).
// 4. deps.onApplied() — re-render and write through; after the command and
//    again when a sliced pass completes.
// Undo/redo push the live state as the opposite snapshot, install fresh
// copies of the stored one (snapshots stay immutable), run the pipeline with
// the step's own alsoRecalc, and re-activate the sheet the step was recorded
// on (its selection rides in the snapshot). `commit()` on a command that
// changed nothing pops its step. Cmd/Ctrl+Z, Shift+Cmd/Ctrl+Z and
// Cmd/Ctrl+Y (the registry's edit.undo / edit.redo) are ignored while focus
// is in an input or a dialog.
import { onScopeDispose, ref, type Ref, type ShallowRef } from "vue";
import { findShortcut, isOutsideDocumentInput } from "../../lib/shortcuts";
import { cellKey, invalidateDependencyGraph, refreshGraphCells, updateGraphCells, type CellKey } from "../../lib/spreadsheet/formula/dependencies";
import { changedCellsBetween, rowVisibilityChanged, visibilitySensitiveCells } from "../../lib/spreadsheet/formula/recalc";
import type { DefinedName, Row, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { nameTableDependents } from "../../lib/spreadsheet/nameDependents";

export const MAX_UNDO_STEPS = 100;

export type UndoScope = "workbook" | { sheetId: string };

export const SHEET_BY_REFERENCE_FIELDS: readonly (keyof Sheet)[] = [
  "rows",
  "sourceXml",
  "columns",
  "merges",
  "frozen",
  "autoFilter",
  "sortState",
  "conditionalFormats",
  "validations",
  "drawings",
  "tables",
  "notes",
  "threads",
  "hyperlinks",
  "pageSetup",
];

export const WORKBOOK_SNAPSHOT_FIELDS: readonly (keyof SpreadsheetModel)[] = [
  "activeSheetIndex",
  "structureDirty",
  "styles",
  "opaque",
  "odsContentEnvelope",
  "dxfs",
  "media",
  "names",
  "namesDirty",
  "persons",
];

export interface SpreadsheetUndoDeps {
  model: ShallowRef<SpreadsheetModel | null>;
  /** SPEC-120's useRecalc with reason "edit": runs (possibly sliced) and
   *  calls back with the pass's changedSheetIds when it completes. */
  recalc: (changed: CellKey[], done: (changedSheetIds: string[]) => void) => void;
  /** Completes a running pass now, so steps never interleave. */
  settle?: () => void;
  /** triggerRef + write-through (touchDocument). */
  onApplied: () => void;
  activateSheet: (sheetId: string) => void;
}

export interface SpreadsheetUndo {
  /** BEFORE mutating. */
  record(label: string, scope: UndoScope): void;
  /** AFTER mutating: the edit pipeline. */
  commit(alsoRecalc?: CellKey[]): void;
  undo(): void;
  redo(): void;
  reset(): void;
  canUndo: Ref<boolean>;
  canRedo: Ref<boolean>;
  undoLabel: Ref<string>;
  redoLabel: Ref<string>;
  /** SPEC-125: bumped by every committed step, undo, redo and reset — "any
   * edit" (the marching ants and a pending cut end on it). */
  revision: Ref<number>;
  /** SPEC-129: calls `listener` on every committed step, undo and redo,
   * with its label, after the model changed and before the pipeline
   * re-renders. */
  onStep(listener: (kind: "commit" | "undo" | "redo", label: string) => void): void;
}

/** A stored state: every sheet (small fields cloned, the by-reference
 * fields shared) and the workbook fields. */
interface Snapshot {
  sheets: Sheet[];
  workbook: Partial<Record<keyof SpreadsheetModel, unknown>>;
}

interface Step {
  label: string;
  scope: UndoScope;
  snapshot: Snapshot;
  alsoRecalc: CellKey[];
}

/** A sheet copy: object/array fields cloned, SHEET_BY_REFERENCE_FIELDS
 * shared. */
function copySheet(sheet: Sheet): Sheet {
  const copy: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(sheet)) {
    copy[key] =
      value !== null && typeof value === "object" && !SHEET_BY_REFERENCE_FIELDS.includes(key as keyof Sheet) ? structuredClone(value) : value;
  }
  return copy as unknown as Sheet;
}

/** The model's state as a snapshot (see the module header). */
export function takeSnapshot(model: SpreadsheetModel): Snapshot {
  const workbook: Partial<Record<keyof SpreadsheetModel, unknown>> = {};
  for (const field of WORKBOOK_SNAPSHOT_FIELDS) {
    if (model[field] !== undefined) {
      workbook[field] = model[field];
    }
  }
  return { sheets: model.sheets.map(copySheet), workbook };
}

/** Installs fresh copies of a snapshot into the live model. */
function install(model: SpreadsheetModel, snapshot: Snapshot): void {
  model.sheets = snapshot.sheets.map(copySheet);
  const fields = model as unknown as Record<string, unknown>;
  for (const field of WORKBOOK_SNAPSHOT_FIELDS) {
    if (field in snapshot.workbook) {
      fields[field] = snapshot.workbook[field];
    } else {
      delete fields[field];
    }
  }
}

/** Marks dirty every sheet whose rows differ from the reference's. */
function markDirty(model: SpreadsheetModel, reference: Snapshot): void {
  const rowsById = new Map(reference.sheets.map((sheet) => [sheet.id, sheet.rows] as const));
  let changed = false;
  const sheets = model.sheets.map((sheet) => {
    const before = rowsById.get(sheet.id);
    if (!sheet.dirty && before !== undefined && before !== sheet.rows) {
      changed = true;
      return { ...sheet, dirty: true as const };
    }
    return sheet;
  });
  if (changed) {
    model.sheets = sheets;
  }
}

/** Every stored cell's key of a sheet's rows. */
function allCellKeys(sheetId: string, rows: readonly Row[]): CellKey[] {
  return rows.flatMap((row) => row.cells.map((cell) => cellKey(sheetId, row.index, cell.col)));
}

export function useSpreadsheetUndo(deps: SpreadsheetUndoDeps): SpreadsheetUndo {
  const undoStack: Step[] = [];
  const redoStack: Step[] = [];
  const canUndo = ref(false);
  const canRedo = ref(false);
  const undoLabel = ref("");
  const redoLabel = ref("");
  const revision = ref(0);
  const stepListeners: ((kind: "commit" | "undo" | "redo", label: string) => void)[] = [];
  /** The step being recorded, and what the model held when it was. */
  let pending: { step: Step; sheets: Sheet[]; fields: unknown[] } | null = null;

  function refresh(): void {
    canUndo.value = undoStack.length > 0;
    canRedo.value = redoStack.length > 0;
    undoLabel.value = undoStack.at(-1)?.label ?? "";
    redoLabel.value = redoStack.at(-1)?.label ?? "";
    revision.value += 1;
  }

  /** The edit pipeline (see the module header). */
  function pipeline(
    model: SpreadsheetModel,
    before: readonly Sheet[],
    reference: Snapshot,
    alsoRecalc: readonly CellKey[],
    namesBefore: readonly DefinedName[],
  ): void {
    const beforeById = new Map(before.map((sheet) => [sheet.id, sheet] as const));
    const afterById = new Map(model.sheets.map((sheet) => [sheet.id, sheet] as const));
    const keys: CellKey[] = [];
    let visibility = false;
    for (const [id, was] of beforeById) {
      const now = afterById.get(id);
      if (!now) {
        keys.push(...allCellKeys(id, was.rows));
      } else if (now.rows !== was.rows) {
        keys.push(...changedCellsBetween(id, was.rows, now.rows));
        visibility ||= rowVisibilityChanged(was.rows, now.rows);
      }
    }
    for (const [id, now] of afterById) {
      if (!beforeById.has(id)) {
        keys.push(...allCellKeys(id, now.rows));
      }
    }
    const listChanged =
      before.length !== model.sheets.length || before.some((sheet, index) => sheet.id !== model.sheets[index]!.id || sheet.name !== model.sheets[index]!.name);
    // SPEC-133: the cells whose formulas mention a changed name or table.
    const named = nameTableDependents(model, namesBefore, before);
    if (listChanged) {
      invalidateDependencyGraph(model);
    } else {
      updateGraphCells(model, keys);
      refreshGraphCells(model, named);
    }
    keys.push(...named);
    markDirty(model, reference);
    const sensitive = visibility ? visibilitySensitiveCells(model) : [];
    deps.recalc([...new Set([...keys, ...alsoRecalc, ...sensitive])], () => {
      if (deps.model.value === model) {
        markDirty(model, reference);
        deps.onApplied();
      }
    });
    deps.onApplied();
  }

  function record(label: string, scope: UndoScope): void {
    const model = deps.model.value;
    if (!model) {
      return;
    }
    deps.settle?.();
    pending = {
      step: { label, scope, snapshot: takeSnapshot(model), alsoRecalc: [] },
      sheets: model.sheets,
      fields: WORKBOOK_SNAPSHOT_FIELDS.map((field) => model[field]),
    };
  }

  function commit(alsoRecalc: CellKey[] = []): void {
    const model = deps.model.value;
    const recorded = pending;
    pending = null;
    if (!model || !recorded) {
      return;
    }
    const unchanged =
      model.sheets === recorded.sheets && WORKBOOK_SNAPSHOT_FIELDS.every((field, index) => model[field] === recorded.fields[index]);
    if (unchanged) {
      return;
    }
    recorded.step.alsoRecalc = alsoRecalc;
    undoStack.push(recorded.step);
    if (undoStack.length > MAX_UNDO_STEPS) {
      undoStack.shift();
    }
    redoStack.length = 0;
    refresh();
    for (const listener of stepListeners) {
      listener("commit", recorded.step.label);
    }
    pipeline(model, recorded.step.snapshot.sheets, recorded.step.snapshot, alsoRecalc, recorded.step.snapshot.workbook.names as DefinedName[]);
  }

  /** Swaps the live state for a stored step's, pushing the live one onto
   * the opposite stack. */
  function swap(from: Step[], to: Step[], kind: "undo" | "redo"): void {
    const model = deps.model.value;
    if (!model || from.length === 0) {
      return;
    }
    deps.settle?.();
    const step = from.pop()!;
    const live = takeSnapshot(model);
    to.push({ ...step, snapshot: live });
    const before = model.sheets;
    install(model, step.snapshot);
    refresh();
    for (const listener of stepListeners) {
      listener(kind, step.label);
    }
    pipeline(model, before, step.snapshot, step.alsoRecalc, live.workbook.names as DefinedName[]);
    const scopeId = step.scope === "workbook" ? undefined : step.scope.sheetId;
    const target = model.sheets.find((sheet) => sheet.id === scopeId) ?? model.sheets[model.activeSheetIndex];
    if (target) {
      deps.activateSheet(target.id);
    }
  }

  function undo(): void {
    swap(undoStack, redoStack, "undo");
  }

  function redo(): void {
    swap(redoStack, undoStack, "redo");
  }

  function reset(): void {
    undoStack.length = 0;
    redoStack.length = 0;
    pending = null;
    refresh();
  }

  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !deps.model.value || isOutsideDocumentInput(event.target)) {
      return;
    }
    const id = findShortcut(event)?.id;
    if (id === "edit.undo") {
      event.preventDefault();
      undo();
    } else if (id === "edit.redo") {
      event.preventDefault();
      redo();
    }
  }

  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown);
    }
  });

  function onStep(listener: (kind: "commit" | "undo" | "redo", label: string) => void): void {
    stepListeners.push(listener);
  }

  return { record, commit, undo, redo, reset, canUndo, canRedo, undoLabel, redoLabel, revision, onStep };
}
