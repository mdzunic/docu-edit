// Incremental recalculation (SPEC-120 "Recalculation", "Cycles", "Stale
// cells", "Time slicing") — DOM-free: a RecalcPass is a resumable state
// machine (graph build, dirty set, ordering and evaluation are all steps)
// that useRecalc drives in RECALC_SLICE_MS slices; recalculate runs one to
// the end. Nothing here assumes the main thread.
//
// 1. Update the graph for changed cells whose formula changed.
// 2. Dirty set = changed formula cells ∪ the transitive dependents of every
//    changed cell ∪ every volatile cell and its transitive dependents — or
//    every formula cell for a `full` pass.
// 3. Order the dirty subgraph by Kahn's algorithm, always taking the first
//    ready cell in sheet order then row-major — deterministic, so the order
//    of RAND draws and the cycle reported are reproducible. Leftover cells
//    are cycles and what reads them: Tarjan names the components, which are
//    appended in dependency order.
// 4. Evaluate. A precedent still dirty when read is evaluated first (a
//    dynamic dependency); past MAX_EVAL_DEPTH nested evaluations the outer
//    cell is re-queued behind its precedent instead of recursing. A cell
//    found in progress is a (dynamic) cycle.
// 5. Write changed values copy-on-write, batched per row per step: one new
//    Row per touched row, one new `rows` array and Sheet per touched sheet,
//    one new `sheets` array; an unchanged value writes nothing and every
//    untouched row keeps its identity. An array anchor's result is
//    distributed over its `arrayRange` with the broadcasting rule — every
//    member's value written, the anchor's included. A changed value drops
//    what was bound to the old one (SPEC-122's drop rules, model.ts
//    dropValueBound): rich-text runs, the shared-string index, an ods cell's
//    preserved text:p (a HYPERLINK result's link), an xlsx ` vm`.
//
// Cycles: after an "edit" or "manual" pass every cell of a cycle is 0 and its
// dependents compute from that 0; an "open" pass leaves their cached values
// alone; an iterating workbook's cycle cells are uncomputable ("iterative")
// and never reported. `circular` lists each cycle's first cell, in sheet
// order then row-major.
//
// Stale cells: an uncomputable cell is stale when it is a transitive
// dependent of a cell whose value or formula changed in the pass; after an
// open-time full pass every uncomputable cell is; after an "edit" or
// "manual" pass every cell of a data-table range is. A pass leaving any stale
// cell sets `model.forceFullCalcOnLoad` (never cleared here). SPEC-127: an
// uncomputable cell with NO cached value (a newly typed formula) gets the
// value #NAME? and is stale — never a guessed value.
//
// SPEC-127: visibilitySensitiveCells lists the formula cells calling SUBTOTAL
// or AGGREGATE (visibility.ts's memo), which hiding or showing rows hands to
// a pass, since their results depend on which rows show.
import { now as clockNow } from "../../core/clock";
import { windowOf } from "../../core/arrayWindow";
import { getCell, parseRange } from "../address";
import { dropValueBound, type Cell, type CellValue, type Row, type Sheet, type SpreadsheetModel } from "../model";
import { formulaAnchorAt, startTagAttribute } from "./anchors";
import { elementAt, sheetById } from "./areas";
import {
  cellKey,
  GraphBuilder,
  hasDependencyGraph,
  internalGraph,
  parseCellKey,
  updateGraphCells,
  type CellKey,
  type DependencyGraph,
} from "./dependencies";
import {
  evalSettingsFor,
  evaluateCell,
  isUncomputable,
  MAX_EVAL_DEPTH,
  rowVisibilityIn,
  type EvalContext,
  type EvalSettings,
  type UncomputableReason,
} from "./evaluate";
import { fromCellValue, isError, toCellValue, type Area, type ScalarValue } from "./values";
import { callsVisibilitySensitive } from "./visibility";

export type RecalcReason = "open" | "edit" | "manual";

export interface RecalcOptions {
  /** Default "edit". */
  reason?: RecalcReason;
  /** Every formula cell is dirty. */
  full?: true;
  /** Locale defaults to en-US. */
  settings?: Partial<EvalSettings>;
  now?: () => Date;
  random?: () => number;
}

export interface RecalcResult {
  valueChanged: CellKey[];
  changedSheetIds: string[];
  /** The first cell of each cycle, report order. */
  circular: CellKey[];
  stale: CellKey[];
  evaluated: number;
}

export interface RecalcPass {
  /** True when done. */
  step(budgetMs: number): boolean;
  /** The cells it had not evaluated. */
  cancel(): CellKey[];
  readonly result: RecalcResult;
}

export const RECALC_SLICE_MS = 12;
export const CALCULATING_STATUS_DELAY_MS = 250;

/** The identity diff of a sheet's rows (core/arrayWindow.ts's windowOf over
 * the rows, then over each changed row's cells): the keys of every cell an
 * edit touched, row-major. */
export function changedCellsBetween(sheetId: string, before: Row[], after: Row[]): CellKey[] {
  const rows = windowOf(before, after);
  const beforeRows = new Map(rows.beforeMid.map((row) => [row.index, row] as const));
  const afterRows = new Map(rows.afterMid.map((row) => [row.index, row] as const));
  const indices = [...new Set([...beforeRows.keys(), ...afterRows.keys()])].toSorted((a, b) => a - b);
  const keys: CellKey[] = [];
  for (const index of indices) {
    const was = beforeRows.get(index);
    const now = afterRows.get(index);
    if (was === now) {
      continue;
    }
    const cells = windowOf(was?.cells ?? [], now?.cells ?? []);
    const cols = new Set([...cells.beforeMid, ...cells.afterMid].map((cell) => cell.col));
    for (const col of [...cols].toSorted((a, b) => a - b)) {
      keys.push(cellKey(sheetId, index, col));
    }
  }
  return keys;
}

/** Whether an edit hid or showed a row: some row's `hidden` differs between
 * the two row lists (the identity diff changedCellsBetween takes) — when it
 * did, visibilitySensitiveCells must recalculate (SPEC-127). */
export function rowVisibilityChanged(before: Row[], after: Row[]): boolean {
  const rows = windowOf(before, after);
  const hiddenBefore = new Set(rows.beforeMid.filter((row) => row.hidden).map((row) => row.index));
  const hiddenAfter = rows.afterMid.filter((row) => row.hidden).map((row) => row.index);
  return hiddenAfter.length !== hiddenBefore.size || hiddenAfter.some((index) => !hiddenBefore.has(index));
}

/** Whether opening a workbook must recalculate it whole: its file asks
 * (`fullCalcOnLoad`) or a formula cell has no cached value. */
export function needsFullCalcOnOpen(model: SpreadsheetModel): boolean {
  if (model.calc?.fullCalcOnLoad) {
    return true;
  }
  return model.sheets.some((sheet) => sheet.rows.some((row) => row.cells.some((cell) => cell.formula !== undefined && cell.value === null)));
}

/** Every formula cell's key, sheet order then row-major. */
function formulaCellKeys(model: SpreadsheetModel): CellKey[] {
  return model.sheets.flatMap((sheet) =>
    sheet.rows.flatMap((row) => row.cells.flatMap((cell) => (cell.formula === undefined ? [] : [cellKey(sheet.id, row.index, cell.col)]))),
  );
}

/** Every data-table range of the workbook, with its anchor's sheet. */
function dataTableAreas(model: SpreadsheetModel): Area[] {
  const areas: Area[] = [];
  for (const sheet of model.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.dataTableXml !== undefined) {
          const range = parseRange(startTagAttribute(cell.dataTableXml, "ref")?.trim() ?? "") ?? {
            top: row.index,
            left: cell.col,
            bottom: row.index,
            right: cell.col,
          };
          areas.push({ sheetId: sheet.id, ...range });
        }
      }
    }
  }
  return areas;
}

/** Why a cell cannot be recalculated here (the stale `title`), if it
 * cannot: its formula's reason, an iterating cycle, or a data table. */
export function uncomputableReasonOf(model: SpreadsheetModel, key: CellKey): UncomputableReason | undefined {
  const reason = hasDependencyGraph(model) ? internalGraph(model).uncomputable.get(key) : undefined;
  if (reason) {
    return reason;
  }
  const { sheetId, row, col } = parseCellKey(key);
  const sheet = sheetById(model, sheetId);
  return sheet && formulaAnchorAt(sheet, row, col)?.kind === "dataTable" ? { kind: "dataTable" } : undefined;
}

function sameValue(a: CellValue, b: CellValue): boolean {
  return a === b || (isError(a) && isError(b) && a.error === b.error);
}

/** A min-heap of keys by rank. */
class RankHeap {
  private readonly items: { rank: number; key: CellKey }[] = [];

  get size(): number {
    return this.items.length;
  }

  push(rank: number, key: CellKey): void {
    const items = this.items;
    items.push({ rank, key });
    let i = items.length - 1;
    while (i > 0) {
      const parent = (i - 1) >> 1;
      if (items[parent]!.rank <= rank) {
        break;
      }
      [items[parent], items[i]] = [items[i]!, items[parent]!];
      i = parent;
    }
  }

  pop(): CellKey | undefined {
    const items = this.items;
    const top = items[0];
    const last = items.pop();
    if (items.length > 0 && last) {
      items[0] = last;
      let i = 0;
      for (;;) {
        const left = i * 2 + 1;
        const right = left + 1;
        let smallest = i;
        if (left < items.length && items[left]!.rank < items[smallest]!.rank) {
          smallest = left;
        }
        if (right < items.length && items[right]!.rank < items[smallest]!.rank) {
          smallest = right;
        }
        if (smallest === i) {
          break;
        }
        [items[smallest], items[i]] = [items[i]!, items[smallest]!];
        i = smallest;
      }
    }
    return top?.key;
  }
}

/** Thrown past MAX_EVAL_DEPTH: re-queue the outer cell behind `target`. */
class Requeue {
  constructor(readonly target: CellKey) {}
}

/** Work units between deadline checks while building and ordering. */
const CHECK_EVERY = 256;

class Pass implements RecalcPass {
  readonly result: RecalcResult = { valueChanged: [], changedSheetIds: [], circular: [], stale: [], evaluated: 0 };
  private readonly reason: RecalcReason;
  private readonly full: boolean;
  private readonly settings: EvalSettings;
  private readonly base: Omit<EvalContext, "host" | "origin" | "arrayMode">;
  private phase: "graph" | "dirty" | "order" | "evaluate" | "done" = "graph";
  private builder: GraphBuilder | null = null;
  private graph!: ReturnType<typeof internalGraph>;
  private readonly dirty = new Set<CellKey>();
  private queue: CellKey[] = [];
  private queued = 0;
  private successors = new Map<CellKey, CellKey[]>();
  private order: CellKey[] = [];
  private position = 0;
  private readonly front: CellKey[] = [];
  /** A dirty cell's static cycle (its members, sheet order then row-major). */
  private readonly cycleOf = new Map<CellKey, CellKey[]>();
  /** Members of cycles met in progress during evaluation. */
  private readonly dynamicCycle = new Set<CellKey>();
  private readonly done = new Set<CellKey>();
  private readonly inProgress = new Set<CellKey>();
  private readonly stack: CellKey[] = [];
  private readonly recorded: Area[][] = [];
  private readonly pending = new Map<CellKey, CellValue>();
  /** Set when a write created a cell the model does not store yet: the step
   * ends there so iteration (which walks stored cells) sees it. */
  private flushSoon = false;
  /** Cells whose dependents may now be stale: changed inputs and cells
   * whose value changed or that went stale. */
  private readonly affected = new Set<CellKey>();
  private readonly staleKeys = new Set<CellKey>();
  private readonly circular = new Set<CellKey>();
  private readonly valueChanged = new Set<CellKey>();
  private readonly touchedSheets = new Set<string>();

  constructor(
    private readonly model: SpreadsheetModel,
    private readonly changed: readonly CellKey[],
    options: RecalcOptions,
  ) {
    this.reason = options.reason ?? "edit";
    this.full = options.full === true;
    this.settings = evalSettingsFor(model, options.settings);
    this.base = {
      model,
      settings: this.settings,
      now: options.now ?? clockNow,
      random: options.random ?? Math.random,
      valueAt: (sheetId, row, col) => this.valueAt(sheetId, row, col),
      recordDependency: (area) => {
        this.recorded.at(-1)?.push(area);
      },
      rowVisibility: (sheetId, row) => rowVisibilityIn(this.model, sheetId, row),
    };
  }

  step(budgetMs: number): boolean {
    const deadline = performance.now() + budgetMs;
    const overBudget = (): boolean => performance.now() >= deadline;
    let first = true;
    while (this.phase !== "done") {
      if (!first && overBudget()) {
        break;
      }
      first = false;
      switch (this.phase) {
        case "graph":
          this.buildGraph(overBudget);
          break;
        case "dirty":
          this.collectDirty(overBudget);
          break;
        case "order":
          this.orderDirty(overBudget);
          break;
        case "evaluate":
          this.evaluateSome(overBudget);
          break;
      }
      if (this.flushSoon) {
        break;
      }
    }
    this.flush();
    if (this.phase === "done") {
      this.finish();
    }
    return this.phase === "done";
  }

  cancel(): CellKey[] {
    this.flush();
    let left: CellKey[];
    if (this.phase === "evaluate") {
      left = [...this.front, ...this.order.slice(this.position)].filter((key) => !this.done.has(key));
    } else if (this.phase === "done") {
      left = [];
    } else if (this.full) {
      // Nothing evaluated yet, and every formula cell was to be.
      left = formulaCellKeys(this.model);
    } else {
      // Nothing evaluated yet: the whole job carries over.
      left = [...this.changed, ...this.dirty];
    }
    this.phase = "done";
    return [...new Set(left)];
  }

  // --- 1: the graph ------------------------------------------------------------------

  private buildGraph(overBudget: () => boolean): void {
    if (!hasDependencyGraph(this.model)) {
      this.builder ??= new GraphBuilder(this.model);
      if (!this.builder.step(overBudget)) {
        return;
      }
      this.builder = null;
    }
    updateGraphCells(this.model, this.changed);
    if (!hasDependencyGraph(this.model)) {
      return; // an array anchor changed: rebuild
    }
    this.graph = internalGraph(this.model);
    this.phase = "dirty";
    if (this.full) {
      for (const key of this.graph.formulaCells()) {
        this.dirty.add(key);
      }
      this.queue = [];
    } else {
      for (const key of this.changed) {
        if (this.graph.has(key)) {
          this.dirty.add(key);
        }
      }
      for (const key of this.graph.formulaCells()) {
        if (this.graph.volatile.has(key)) {
          this.dirty.add(key);
        }
      }
      this.queue = [...this.changed, ...this.dirty];
    }
    for (const key of this.changed) {
      this.affected.add(key);
      this.markAffected(key);
    }
  }

  // --- 2: the dirty set ----------------------------------------------------------------

  private collectDirty(overBudget: () => boolean): void {
    while (this.queued < this.queue.length) {
      const { sheetId, row, col } = parseCellKey(this.queue[this.queued]!);
      this.queued += 1;
      for (const dependent of this.graph.dependentsOf(sheetId, row, col)) {
        if (!this.dirty.has(dependent)) {
          this.dirty.add(dependent);
          this.queue.push(dependent);
        }
      }
      if (this.queued % CHECK_EVERY === 0 && overBudget()) {
        return;
      }
    }
    this.queue = [...this.graph.formulaCells()].filter((key) => this.dirty.has(key));
    this.queued = 0;
    this.phase = "order";
  }

  // --- 3: the order ------------------------------------------------------------------

  private orderDirty(overBudget: () => boolean): void {
    // Each dirty cell's dirty dependents, sorted — the edges.
    while (this.queued < this.queue.length) {
      const key = this.queue[this.queued]!;
      this.queued += 1;
      const { sheetId, row, col } = parseCellKey(key);
      const next = this.graph.dependentsOf(sheetId, row, col).filter((dependent) => this.dirty.has(dependent));
      this.successors.set(key, this.byRank(next));
      if (this.queued % CHECK_EVERY === 0 && overBudget()) {
        return;
      }
    }
    const keys = this.queue;
    const indegree = new Map<CellKey, number>(keys.map((key) => [key, 0]));
    for (const key of keys) {
      for (const dependent of this.successors.get(key)!) {
        indegree.set(dependent, indegree.get(dependent)! + 1);
      }
    }
    const heap = new RankHeap();
    for (const key of keys) {
      if (indegree.get(key) === 0) {
        heap.push(this.graph.rank(key), key);
      }
    }
    const order: CellKey[] = [];
    for (let key = heap.pop(); key !== undefined; key = heap.pop()) {
      order.push(key);
      for (const dependent of this.successors.get(key)!) {
        const left = indegree.get(dependent)! - 1;
        indegree.set(dependent, left);
        if (left === 0) {
          heap.push(this.graph.rank(dependent), dependent);
        }
      }
    }
    if (order.length < keys.length) {
      order.push(...this.orderLeftovers(keys.filter((key) => indegree.get(key)! > 0)));
    }
    this.order = order;
    this.successors = new Map();
    this.queue = [];
    this.phase = "evaluate";
  }

  /** The cells Kahn's algorithm could not order — cycles and their readers —
   * in dependency order of their strongly connected components (Tarjan). */
  private orderLeftovers(left: CellKey[]): CellKey[] {
    const position = new Map(left.map((key, i) => [key, i] as const));
    const next = left.map((key) => this.successors.get(key)!.map((d) => position.get(d)).filter((i): i is number => i !== undefined));
    const index = new Int32Array(left.length).fill(-1);
    const low = new Int32Array(left.length);
    const onStack = new Uint8Array(left.length);
    const stack: number[] = [];
    const components: number[][] = [];
    let counter = 0;
    for (let root = 0; root < left.length; root += 1) {
      if (index[root] !== -1) {
        continue;
      }
      const frames: { node: number; at: number }[] = [];
      const open = (node: number): void => {
        index[node] = low[node] = counter;
        counter += 1;
        stack.push(node);
        onStack[node] = 1;
        frames.push({ node, at: 0 });
      };
      open(root);
      while (frames.length > 0) {
        const frame = frames.at(-1)!;
        const edges = next[frame.node]!;
        if (frame.at < edges.length) {
          const to = edges[frame.at]!;
          frame.at += 1;
          if (index[to] === -1) {
            open(to);
          } else if (onStack[to]) {
            low[frame.node] = Math.min(low[frame.node]!, index[to]!);
          }
          continue;
        }
        frames.pop();
        const parent = frames.at(-1);
        if (parent) {
          low[parent.node] = Math.min(low[parent.node]!, low[frame.node]!);
        }
        if (low[frame.node] === index[frame.node]) {
          const component: number[] = [];
          let member: number;
          do {
            member = stack.pop()!;
            onStack[member] = 0;
            component.push(member);
          } while (member !== frame.node);
          components.push(component.toSorted((a, b) => a - b));
        }
      }
    }
    // Tarjan completes a component after everything it reaches: reversed,
    // precedents come first.
    const ordered: CellKey[] = [];
    for (const component of components.toReversed()) {
      const members = component.map((i) => left[i]!);
      const cyclic = component.length > 1 || next[component[0]!]!.includes(component[0]!);
      if (cyclic) {
        for (const member of members) {
          this.cycleOf.set(member, members);
        }
      }
      ordered.push(...members);
    }
    return ordered;
  }

  // --- 4: evaluation -------------------------------------------------------------------

  private evaluateSome(overBudget: () => boolean): void {
    for (;;) {
      const key = this.front.pop() ?? this.order[this.position++];
      if (key === undefined) {
        this.phase = "done";
        return;
      }
      try {
        this.process(key);
      } catch (signal) {
        if (!(signal instanceof Requeue)) {
          throw signal;
        }
        this.front.push(key, signal.target);
      }
      if (this.flushSoon || overBudget()) {
        return;
      }
    }
  }

  private process(key: CellKey): void {
    if (this.done.has(key)) {
      return;
    }
    const cycle = this.cycleOf.get(key);
    if (cycle) {
      this.settleCycle(cycle);
      return;
    }
    this.evaluateNode(key);
  }

  /** A static cycle: 0 after an edit or F9, the cached values at open,
   * uncomputable in an iterating workbook. */
  private settleCycle(members: CellKey[]): void {
    const iterating = this.model.calc?.iterate === true;
    if (iterating) {
      this.graph.markIterative(members);
    } else {
      this.circular.add(members[0]!);
    }
    for (const member of members) {
      this.done.add(member);
      if (iterating || this.graph.uncomputable.has(member)) {
        this.settleUncomputable(member);
      } else if (this.reason !== "open") {
        this.writeResult(member, 0);
      }
    }
  }

  private evaluateNode(key: CellKey): void {
    if (this.graph.hasStaticReason(key)) {
      this.done.add(key);
      this.settleUncomputable(key);
      return;
    }
    const { sheetId, row, col } = parseCellKey(key);
    this.inProgress.add(key);
    this.stack.push(key);
    this.recorded.push([]);
    let value: ReturnType<typeof evaluateCell>;
    let areas: Area[];
    try {
      value = evaluateCell(this.model, sheetId, row, col, this.base);
    } finally {
      this.inProgress.delete(key);
      this.stack.pop();
      areas = this.recorded.pop()!;
    }
    this.done.add(key);
    this.result.evaluated += 1;
    if (this.graph.dynamic.has(key)) {
      this.graph.setDynamic(key, areas);
    }
    if (isUncomputable(value)) {
      this.graph.setRuntimeReason(key, value.uncomputable);
      this.settleUncomputable(key);
      return;
    }
    this.graph.setRuntimeReason(key, undefined);
    if (this.dynamicCycle.has(key)) {
      if (this.model.calc?.iterate) {
        this.graph.markIterative([key]);
        this.settleUncomputable(key);
      } else if (this.reason !== "open") {
        this.writeResult(key, 0);
      }
      return;
    }
    if (typeof value === "object" && value !== null && "kind" in value) {
      // An array anchor: its result over its range, broadcast.
      const range = getCell(sheetById(this.model, sheetId)!, row, col)?.arrayRange;
      if (!range) {
        return;
      }
      let changed = false;
      for (let r = range.top; r <= range.bottom; r += 1) {
        for (let c = range.left; c <= range.right; c += 1) {
          changed = this.writeResult(cellKey(sheetId, r, c), toCellValue(elementAt(value, r - range.top, c - range.left))) || changed;
        }
      }
      if (changed) {
        this.markAffected(key);
      }
      return;
    }
    this.writeResult(key, value);
  }

  /** A dirty precedent is evaluated before it is read; a cell found in
   * progress closes a dynamic cycle. */
  private valueAt(sheetId: string, row: number, col: number): ScalarValue {
    const key = cellKey(sheetId, row, col);
    const pending = this.pending.get(key);
    if (pending !== undefined) {
      return fromCellValue(pending);
    }
    const target = this.graph.has(key) ? key : this.graph.arrayAnchorAt(sheetId, row, col);
    if (target !== undefined && this.dirty.has(target) && !this.done.has(target)) {
      if (this.inProgress.has(target)) {
        const members = this.stack.slice(this.stack.indexOf(target));
        for (const member of members) {
          this.dynamicCycle.add(member);
        }
        if (!this.model.calc?.iterate) {
          this.circular.add(members.reduce((first, member) => (this.graph.rank(member) < this.graph.rank(first) ? member : first)));
        }
        if (this.reason !== "open" && !this.model.calc?.iterate) {
          return 0;
        }
      } else {
        const cycle = this.cycleOf.get(target);
        if (cycle) {
          this.settleCycle(cycle);
        } else if (this.stack.length >= MAX_EVAL_DEPTH) {
          throw new Requeue(target);
        } else {
          this.evaluateNode(target);
        }
        const settled = this.pending.get(key);
        if (settled !== undefined) {
          return fromCellValue(settled);
        }
      }
    }
    const sheet = sheetById(this.model, sheetId);
    return fromCellValue(sheet ? (getCell(sheet, row, col)?.value ?? null) : null);
  }

  // --- 5: writes and staleness --------------------------------------------------------------

  /** Records a cell's new value; an unchanged one writes nothing. True when
   * the value changed. */
  private writeResult(key: CellKey, value: CellValue): boolean {
    const { sheetId, row, col } = parseCellKey(key);
    const pending = this.pending.get(key);
    const cell = pending === undefined ? getCell(sheetById(this.model, sheetId)!, row, col) : undefined;
    const current = pending !== undefined ? pending : (cell?.value ?? null);
    if (sameValue(current, value)) {
      return false;
    }
    if (pending === undefined && !cell) {
      this.flushSoon = true;
    }
    this.pending.set(key, value);
    this.valueChanged.add(key);
    this.touchedSheets.add(sheetId);
    this.markAffected(key);
    return true;
  }

  /** The dependents of a changed (or stale) cell may now be stale. */
  private markAffected(key: CellKey): void {
    const { sheetId, row, col } = parseCellKey(key);
    for (const dependent of this.graph.dependentsOf(sheetId, row, col)) {
      this.affected.add(dependent);
    }
  }

  /** An uncomputable cell keeps its cached value; one with none — a newly
   * typed formula — gets #NAME? and is stale (SPEC-127), never a guess. */
  private settleUncomputable(key: CellKey): void {
    const { sheetId, row, col } = parseCellKey(key);
    const pending = this.pending.get(key);
    const sheet = sheetById(this.model, sheetId);
    const current = pending !== undefined ? pending : sheet ? (getCell(sheet, row, col)?.value ?? null) : null;
    if (current === null) {
      this.writeResult(key, { error: "#NAME?" });
      this.affected.add(key);
    }
    this.checkStale(key);
  }

  private checkStale(key: CellKey): void {
    if (this.staleKeys.has(key)) {
      return;
    }
    if ((this.reason === "open" && this.full) || this.affected.has(key)) {
      this.staleKeys.add(key);
      this.markAffected(key);
    }
  }

  /** Writes the step's values copy-on-write (see the module header). */
  private flush(): void {
    this.flushSoon = false;
    if (this.pending.size === 0) {
      return;
    }
    const bySheet = new Map<string, Map<number, Map<number, CellValue>>>();
    for (const [key, value] of this.pending) {
      const { sheetId, row, col } = parseCellKey(key);
      const rows = bySheet.get(sheetId) ?? new Map<number, Map<number, CellValue>>();
      const cells = rows.get(row) ?? new Map<number, CellValue>();
      cells.set(col, value);
      rows.set(row, cells);
      bySheet.set(sheetId, rows);
    }
    this.pending.clear();
    this.model.sheets = this.model.sheets.map((sheet): Sheet => {
      const writes = bySheet.get(sheet.id);
      if (!writes) {
        return sheet;
      }
      const rows = sheet.rows.slice();
      for (const [index, cells] of [...writes].toSorted((a, b) => a[0] - b[0])) {
        const at = rowPosition(rows, index);
        if (rows[at]?.index === index) {
          rows[at] = withValues(rows[at]!, cells);
        } else {
          rows.splice(at, 0, withValues({ index, cells: [] }, cells));
        }
      }
      return { ...sheet, rows };
    });
  }

  /** Keys in sheet order, then row-major. */
  private byRank(keys: Iterable<CellKey>): CellKey[] {
    return [...keys]
      .map((key) => [this.graph.rank(key), key] as const)
      .toSorted((a, b) => a[0] - b[0])
      .map(([, key]) => key);
  }

  private finish(): void {
    if (this.reason !== "open") {
      for (const area of dataTableAreas(this.model)) {
        for (let row = area.top; row <= area.bottom; row += 1) {
          for (let col = area.left; col <= area.right; col += 1) {
            this.staleKeys.add(cellKey(area.sheetId, row, col));
          }
        }
      }
    }
    if (this.staleKeys.size > 0) {
      this.model.forceFullCalcOnLoad = true;
    }
    this.result.valueChanged = this.byRank(this.valueChanged);
    this.result.circular = this.byRank(this.circular);
    this.result.stale = this.byRank(this.staleKeys);
    this.result.changedSheetIds = this.model.sheets.filter((sheet) => this.touchedSheets.has(sheet.id)).map((sheet) => sheet.id);
  }
}

/** Where row `index` is (or would go) in sorted rows. */
function rowPosition(rows: readonly Row[], index: number): number {
  let lo = 0;
  let hi = rows.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (rows[mid]!.index < index) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  return lo;
}

/** A row with new values at the given columns — new Cell objects for them
 * (a changed value drops what was bound to the old one: rich-text runs, and
 * SPEC-122's drop rules — the shared-string index, an ods cell's preserved
 * text:p, an xlsx ` vm`), new cells where the row stored none. */
function withValues(row: Row, values: ReadonlyMap<number, CellValue>): Row {
  const cells: Cell[] = row.cells.map((cell) => {
    if (!values.has(cell.col)) {
      return cell;
    }
    const next: Cell = { ...cell, value: values.get(cell.col)! };
    dropValueBound(next);
    return next;
  });
  let added = false;
  for (const [col, value] of values) {
    if (!row.cells.some((cell) => cell.col === col)) {
      cells.push({ col, value });
      added = true;
    }
  }
  return { ...row, cells: added ? cells.toSorted((a, b) => a.col - b.col) : cells };
}

/** Every formula cell calling a visibility-sensitive function (SUBTOTAL,
 * AGGREGATE), sheet order then row-major: what hiding or showing rows hands
 * to a pass (SPEC-124's hide/unhide; SPEC-128's filters). */
export function visibilitySensitiveCells(model: SpreadsheetModel): CellKey[] {
  const keys: CellKey[] = [];
  for (const sheet of model.sheets) {
    for (const row of sheet.rows) {
      for (const cell of row.cells) {
        if (cell.formula !== undefined && callsVisibilitySensitive(cell.formula)) {
          keys.push(cellKey(sheet.id, row.index, cell.col));
        }
      }
    }
  }
  return keys;
}

/** Starts a pass over the model (see the module header). */
export function startRecalc(model: SpreadsheetModel, changed: CellKey[], options: RecalcOptions = {}): RecalcPass {
  return new Pass(model, changed, options);
}

/** Runs a pass to the end, synchronously. */
export function recalculate(model: SpreadsheetModel, changed: CellKey[], options: RecalcOptions = {}): RecalcResult {
  const pass = startRecalc(model, changed, options);
  while (!pass.step(Number.POSITIVE_INFINITY)) {
    // every step ends at a flush; keep going
  }
  return pass.result;
}

/** The graph the passes keep for a model (re-exported for the view). */
export function recalcGraph(model: SpreadsheetModel): DependencyGraph {
  return internalGraph(model);
}
