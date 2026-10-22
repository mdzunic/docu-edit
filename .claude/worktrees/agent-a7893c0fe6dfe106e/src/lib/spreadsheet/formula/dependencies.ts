// The dependency graph (SPEC-120 "The dependency graph") — which formula
// cells read which cells, so a recalculation evaluates only what a change
// reaches. Lib state, never model data: the graph lives in a WeakMap keyed
// by the model, built lazily (graphFor) or step by step (GraphBuilder, which
// a RecalcPass drives in time slices).
//
// - Nodes are formula cells; an array anchor stands for its whole range;
//   data tables are not nodes. A node's precedents are the areas of its
//   references, resolved at its host (a 3-D reference as one area per
//   sheet), plus — for OFFSET and INDIRECT — the areas its last evaluation
//   recorded, replaced on each evaluation.
// - A single-cell precedent goes in a map; a range goes in its sheet's
//   BAND_WIDTH-column band lists (sorted by `top`), or — when it spans more
//   than WIDE_RANGE_BANDS bands — the sheet's wide list. dependentsOf unions
//   the map hit with the band and wide-list entries containing the cell. A
//   precedent area overlapping an `arrayRange` also makes the formula depend
//   on that anchor, so an anchor is evaluated before every reader of its
//   members.
// - Sheet add/delete/rename/move (SPEC-121) and row/column edits (SPEC-124)
//   call invalidateDependencyGraph; cell edits call updateGraphCells.
// - findCycles is Tarjan's algorithm over the whole graph, iterative (a
//   100,000-cell chain does not overflow the stack), in a deterministic
//   order: every strongly connected component of more than one node, and
//   every self-reference, its cells in sheet order then row-major, the
//   cycles ordered by their first cell.
import { getCell } from "../address";
import { MAX_COLS, MAX_ROWS, type Cell, type SpreadsheetModel } from "../model";
import { resolveReference, sheetById, sheetIndexOf } from "./areas";
import { formulaFacts, type UncomputableReason } from "./evaluate";
import { FormulaSyntaxError, parseFormulaAt } from "./parse";
import { referenceAreas } from "./referenceHooks";
import { isError, type Area } from "./values";

/** A cell's key, `${sheetId}!${row},${col}` — built only by cellKey. */
export type CellKey = string;

export function cellKey(sheetId: string, row: number, col: number): CellKey {
  return `${sheetId}!${row},${col}`;
}

/** A key's parts (split at the last "!", so a sheet id may hold one). */
export function parseCellKey(key: CellKey): { sheetId: string; row: number; col: number } {
  const bang = key.lastIndexOf("!");
  const comma = key.indexOf(",", bang);
  return { sheetId: key.slice(0, bang), row: Number(key.slice(bang + 1, comma)), col: Number(key.slice(comma + 1)) };
}

export const BAND_WIDTH = 32;
export const WIDE_RANGE_BANDS = 8;

export interface DependencyGraph {
  /** Static ∪ last dynamic. */
  precedentsOf(key: CellKey): readonly Area[];
  dependentsOf(sheetId: string, row: number, col: number): CellKey[];
  /** Sheet order, row-major. */
  formulaCells(): Iterable<CellKey>;
  readonly volatile: ReadonlySet<CellKey>;
  readonly dynamic: ReadonlySet<CellKey>;
  readonly uncomputable: ReadonlyMap<CellKey, UncomputableReason>;
}

interface Entry {
  area: Area;
  dependent: CellKey;
}

interface SheetIndex {
  bands: Map<number, Entry[]>;
  /** Bands whose entries were appended since they were last sorted. */
  unsorted: Set<number>;
  wide: Entry[];
  /** Array anchors on the sheet: their key and range. */
  arrays: { key: CellKey; area: Area }[];
}

interface GraphNode {
  key: CellKey;
  sheetId: string;
  row: number;
  col: number;
  /** What the node was built from — a cell whose source changed is rebuilt. */
  source: string;
  statics: Area[];
  dynamics: Area[];
  /** Array anchors this node depends on through overlapping areas. */
  anchors: CellKey[];
  /** Whether the node is an array anchor itself. */
  isArray: boolean;
}

/** What a formula cell's node is built from. */
function sourceOf(cell: Cell): string {
  const range = cell.arrayRange;
  return [
    cell.formula ?? "",
    cell.formulaUnparsed ? "u" : "",
    cell.dynamicArray ? "d" : "",
    range ? `${range.top},${range.left},${range.bottom},${range.right}` : "",
  ].join("|");
}

function contains(area: Area, row: number, col: number): boolean {
  return row >= area.top && row <= area.bottom && col >= area.left && col <= area.right;
}

function overlaps(a: Area, b: Area): boolean {
  return a.sheetId === b.sheetId && a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;
}

function isSingleCell(area: Area): boolean {
  return area.top === area.bottom && area.left === area.right;
}

class Graph implements DependencyGraph {
  readonly volatile = new Set<CellKey>();
  readonly dynamic = new Set<CellKey>();
  readonly uncomputable = new Map<CellKey, UncomputableReason>();
  private readonly nodes = new Map<CellKey, GraphNode>();
  private readonly cells = new Map<CellKey, Set<CellKey>>();
  private readonly sheets = new Map<string, SheetIndex>();
  private ordered: CellKey[] | null = null;
  /** Reasons set by evaluation rather than read from the formula. */
  private readonly runtime = new Set<CellKey>();

  constructor(private readonly model: SpreadsheetModel) {}

  /** A cell's place in sheet order, then row-major. */
  rank(key: CellKey): number {
    const { sheetId, row, col } = parseCellKey(key);
    return (sheetIndexOf(this.model, sheetId) * MAX_ROWS + row) * MAX_COLS + col;
  }

  has(key: CellKey): boolean {
    return this.nodes.has(key);
  }

  private sheet(sheetId: string): SheetIndex {
    let index = this.sheets.get(sheetId);
    if (!index) {
      index = { bands: new Map(), unsorted: new Set(), wide: [], arrays: [] };
      this.sheets.set(sheetId, index);
    }
    return index;
  }

  /** Records an array anchor before any node reading its range is added. */
  addArray(key: CellKey, area: Area): void {
    this.sheet(area.sheetId).arrays.push({ key, area });
  }

  private link(area: Area, dependent: CellKey): void {
    if (isSingleCell(area)) {
      const key = cellKey(area.sheetId, area.top, area.left);
      let set = this.cells.get(key);
      if (!set) {
        set = new Set();
        this.cells.set(key, set);
      }
      set.add(dependent);
      return;
    }
    const index = this.sheet(area.sheetId);
    const entry = { area, dependent };
    const first = Math.floor(area.left / BAND_WIDTH);
    const last = Math.floor(area.right / BAND_WIDTH);
    if (last - first + 1 > WIDE_RANGE_BANDS) {
      index.wide.push(entry);
      return;
    }
    for (let band = first; band <= last; band += 1) {
      let entries = index.bands.get(band);
      if (!entries) {
        entries = [];
        index.bands.set(band, entries);
      }
      entries.push(entry);
      index.unsorted.add(band);
    }
  }

  private unlink(area: Area, dependent: CellKey): void {
    if (isSingleCell(area)) {
      this.cells.get(cellKey(area.sheetId, area.top, area.left))?.delete(dependent);
      return;
    }
    const index = this.sheets.get(area.sheetId);
    if (!index) {
      return;
    }
    const keep = (entry: Entry): boolean => entry.area !== area || entry.dependent !== dependent;
    const first = Math.floor(area.left / BAND_WIDTH);
    const last = Math.floor(area.right / BAND_WIDTH);
    if (last - first + 1 > WIDE_RANGE_BANDS) {
      index.wide = index.wide.filter(keep);
      return;
    }
    for (let band = first; band <= last; band += 1) {
      const entries = index.bands.get(band);
      if (entries) {
        index.bands.set(band, entries.filter(keep));
      }
    }
  }

  /** Links a node's areas, and the anchors of the arrays they overlap. */
  private linkAreas(node: GraphNode, areas: readonly Area[]): void {
    for (const area of areas) {
      this.link(area, node.key);
      for (const array of this.sheets.get(area.sheetId)?.arrays ?? []) {
        if (array.key !== node.key && overlaps(area, array.area) && !node.anchors.includes(array.key)) {
          node.anchors.push(array.key);
          this.link({ ...array.area, bottom: array.area.top, right: array.area.left }, node.key);
        }
      }
    }
  }

  /** Adds (or rebuilds) the node of a formula cell. */
  addNode(sheetId: string, row: number, col: number, cell: Cell): void {
    const key = cellKey(sheetId, row, col);
    this.removeNode(key);
    const node: GraphNode = { key, sheetId, row, col, source: sourceOf(cell), statics: [], dynamics: [], anchors: [], isArray: cell.arrayRange !== undefined };
    this.nodes.set(key, node);
    this.ordered = null;
    if (cell.formulaUnparsed || cell.formula === undefined) {
      this.uncomputable.set(key, { kind: "unparsed" });
      return;
    }
    const parsed = parseFormulaAt(cell.formula, { row, col });
    if (parsed instanceof FormulaSyntaxError) {
      this.uncomputable.set(key, { kind: "unparsed" });
      return;
    }
    const facts = formulaFacts(parsed.node);
    const host = { sheetId, row, col };
    for (const ref of facts.refs) {
      const resolved = resolveReference(ref, { model: this.model, host, origin: parsed.origin });
      if (!isError(resolved)) {
        node.statics.push(...resolved.areas);
      }
    }
    // SPEC-133: what its names and table references read, as of now.
    node.statics.push(...referenceAreas(this.model, facts, host));
    this.linkAreas(node, node.statics);
    const range = cell.arrayRange;
    const reason: UncomputableReason | undefined =
      facts.reason ?? (cell.dynamicArray && range && (range.top !== range.bottom || range.left !== range.right) ? { kind: "dynamicArray" } : undefined);
    if (reason) {
      this.uncomputable.set(key, reason);
    }
    if (facts.volatile) {
      this.volatile.add(key);
    }
    if (facts.dynamic) {
      this.dynamic.add(key);
    }
  }

  /** Drops a node and every link it made. */
  removeNode(key: CellKey): void {
    const node = this.nodes.get(key);
    if (!node) {
      return;
    }
    for (const area of [...node.statics, ...node.dynamics]) {
      this.unlink(area, key);
    }
    for (const anchor of node.anchors) {
      this.cells.get(anchor)?.delete(key);
    }
    this.nodes.delete(key);
    this.volatile.delete(key);
    this.dynamic.delete(key);
    this.uncomputable.delete(key);
    this.runtime.delete(key);
    this.ordered = null;
  }

  /** The source a node was built from (undefined: no node). */
  sourceAt(key: CellKey): string | undefined {
    return this.nodes.get(key)?.source;
  }

  /** Whether a node is an array anchor. */
  isArrayAnchor(key: CellKey): boolean {
    return this.nodes.get(key)?.isArray === true;
  }

  /** The anchor of the array whose range holds a cell, if any. */
  arrayAnchorAt(sheetId: string, row: number, col: number): CellKey | undefined {
    const arrays = this.sheets.get(sheetId)?.arrays;
    if (!arrays || arrays.length === 0) {
      return undefined;
    }
    return arrays.find((array) => contains(array.area, row, col))?.key;
  }

  /** Every array anchor and its range. */
  arrays(): { key: CellKey; area: Area }[] {
    return [...this.sheets.values()].flatMap((index) => index.arrays);
  }

  /** Whether a cell is uncomputable whatever it evaluates to (its formula's
   * reason, or an iterating cycle) — a reason set by evaluation is
   * re-checked by evaluating again. */
  hasStaticReason(key: CellKey): boolean {
    return this.uncomputable.has(key) && !this.runtime.has(key);
  }

  /** Sets (or, undefined, clears) a reason found only by evaluating a cell —
   * a single-cell dynamic array that would spill, a regular expression. */
  setRuntimeReason(key: CellKey, reason: UncomputableReason | undefined): void {
    if (!this.nodes.has(key)) {
      return;
    }
    if (reason) {
      this.uncomputable.set(key, reason);
    } else if (this.runtime.has(key)) {
      this.uncomputable.delete(key);
    }
    if (reason) {
      this.runtime.add(key);
    } else {
      this.runtime.delete(key);
    }
  }

  /** Replaces a node's dynamic precedents (INDIRECT/OFFSET results). */
  setDynamic(key: CellKey, areas: readonly Area[]): void {
    const node = this.nodes.get(key);
    if (!node) {
      return;
    }
    for (const area of node.dynamics) {
      this.unlink(area, key);
    }
    node.dynamics = [...areas];
    this.linkAreas(node, node.dynamics);
  }

  /** Marks the cells of an iterating workbook's cycles uncomputable. */
  markIterative(keys: readonly CellKey[]): void {
    for (const key of keys) {
      if (this.nodes.has(key) && !this.uncomputable.has(key)) {
        this.uncomputable.set(key, { kind: "iterative" });
      }
    }
  }

  precedentsOf(key: CellKey): readonly Area[] {
    const node = this.nodes.get(key);
    return node ? [...node.statics, ...node.dynamics] : [];
  }

  dependentsOf(sheetId: string, row: number, col: number): CellKey[] {
    const out = new Set<CellKey>(this.cells.get(cellKey(sheetId, row, col)));
    const index = this.sheets.get(sheetId);
    if (index) {
      const band = Math.floor(col / BAND_WIDTH);
      let entries = index.bands.get(band);
      if (entries) {
        if (index.unsorted.delete(band)) {
          entries = entries.toSorted((a, b) => a.area.top - b.area.top);
          index.bands.set(band, entries);
        }
        for (const entry of entries) {
          if (entry.area.top > row) {
            break;
          }
          if (contains(entry.area, row, col)) {
            out.add(entry.dependent);
          }
        }
      }
      for (const entry of index.wide) {
        if (contains(entry.area, row, col)) {
          out.add(entry.dependent);
        }
      }
    }
    return [...out];
  }

  formulaCells(): CellKey[] {
    if (!this.ordered) {
      const ranked = [...this.nodes.keys()].map((key) => [this.rank(key), key] as const);
      this.ordered = ranked.toSorted((a, b) => a[0] - b[0]).map(([, key]) => key);
    }
    return this.ordered;
  }

  /** A node's position (its key's parts). */
  nodeAt(key: CellKey): { sheetId: string; row: number; col: number } | undefined {
    return this.nodes.get(key);
  }
}

/** Builds a model's graph in steps: every array anchor first (a reader of
 * an array range depends on its anchor), then every formula cell. */
export class GraphBuilder {
  readonly graph: Graph;
  private phase: "arrays" | "nodes" | "done" = "arrays";
  private sheetIndex = 0;
  private rowIndex = 0;

  constructor(private readonly model: SpreadsheetModel) {
    this.graph = new Graph(model);
  }

  get done(): boolean {
    return this.phase === "done";
  }

  /** Works until `overBudget()` says stop (checked every row); true when
   * the graph is complete. */
  step(overBudget: () => boolean): boolean {
    while (this.phase !== "done") {
      const sheet = this.model.sheets[this.sheetIndex];
      if (!sheet) {
        this.phase = this.phase === "arrays" ? "nodes" : "done";
        this.sheetIndex = 0;
        this.rowIndex = 0;
        continue;
      }
      const row = sheet.kind === "grid" ? sheet.rows[this.rowIndex] : undefined;
      if (!row) {
        this.sheetIndex += 1;
        this.rowIndex = 0;
        continue;
      }
      for (const cell of row.cells) {
        if (cell.formula === undefined) {
          continue;
        }
        const key = cellKey(sheet.id, row.index, cell.col);
        if (this.phase === "arrays") {
          if (cell.arrayRange) {
            this.graph.addArray(key, { sheetId: sheet.id, ...cell.arrayRange });
          }
        } else {
          this.graph.addNode(sheet.id, row.index, cell.col, cell);
        }
      }
      this.rowIndex += 1;
      if (overBudget()) {
        break;
      }
    }
    if (this.phase === "done") {
      graphs.set(this.model, this.graph);
    }
    return this.phase === "done";
  }
}

const graphs = new WeakMap<SpreadsheetModel, Graph>();

/** The model's graph, built on first use and cached (WeakMap by model). */
export function graphFor(model: SpreadsheetModel): DependencyGraph {
  return internalGraph(model);
}

/** graphFor with the recalculation's own operations. */
export function internalGraph(model: SpreadsheetModel): Graph {
  let graph = graphs.get(model);
  if (!graph) {
    const builder = new GraphBuilder(model);
    builder.step(() => false);
    graph = builder.graph;
  }
  return graph;
}

/** Whether the model's graph is already built. */
export function hasDependencyGraph(model: SpreadsheetModel): boolean {
  return graphs.has(model);
}

/** Re-reads the given cells into the graph (a cell edit): a cell whose
 * formula changed is rebuilt, one whose formula went is dropped. A change
 * to an array anchor rebuilds the whole graph on next use. */
export function updateGraphCells(model: SpreadsheetModel, keys: readonly CellKey[]): void {
  const graph = graphs.get(model);
  if (!graph) {
    return;
  }
  for (const key of keys) {
    const { sheetId, row, col } = parseCellKey(key);
    const sheet = sheetById(model, sheetId);
    const cell = sheet ? getCell(sheet, row, col) : undefined;
    const before = graph.sourceAt(key);
    const after = cell?.formula !== undefined ? sourceOf(cell) : undefined;
    if (before === after) {
      continue;
    }
    if (cell?.arrayRange || graph.isArrayAnchor(key)) {
      invalidateDependencyGraph(model);
      return;
    }
    if (after === undefined) {
      graph.removeNode(key);
    } else {
      graph.addNode(sheetId, row, col, cell!);
    }
  }
}

/** Rebuilds the given formula cells' nodes even though their formulas did
 * not change (SPEC-133: a name or table they mention did, so what they read
 * moved). Cells without a formula are dropped. */
export function refreshGraphCells(model: SpreadsheetModel, keys: readonly CellKey[]): void {
  const graph = graphs.get(model);
  if (!graph) {
    return;
  }
  for (const key of keys) {
    const { sheetId, row, col } = parseCellKey(key);
    const sheet = sheetById(model, sheetId);
    const cell = sheet ? getCell(sheet, row, col) : undefined;
    if (cell?.arrayRange || graph.isArrayAnchor(key)) {
      invalidateDependencyGraph(model);
      return;
    }
    if (cell?.formula === undefined) {
      graph.removeNode(key);
    } else {
      graph.addNode(sheetId, row, col, cell);
    }
  }
}

/** Drops the model's graph (a structural edit); the next use rebuilds it. */
export function invalidateDependencyGraph(model: SpreadsheetModel): void {
  graphs.delete(model);
}

/** Every cycle of the graph (see the module header). */
export function findCycles(graph: DependencyGraph): CellKey[][] {
  const keys = [...graph.formulaCells()];
  const position = new Map(keys.map((key, i) => [key, i] as const));
  const successors = (i: number): number[] => {
    const { sheetId, row, col } = parseCellKey(keys[i]!);
    return graph
      .dependentsOf(sheetId, row, col)
      .map((key) => position.get(key))
      .filter((j): j is number => j !== undefined)
      .toSorted((a, b) => a - b);
  };
  const index = new Int32Array(keys.length).fill(-1);
  const low = new Int32Array(keys.length);
  const onStack = new Uint8Array(keys.length);
  const stack: number[] = [];
  const cycles: number[][] = [];
  let counter = 0;
  for (let root = 0; root < keys.length; root += 1) {
    if (index[root] !== -1) {
      continue;
    }
    const frames: { node: number; next: number[]; at: number }[] = [];
    const open = (node: number): void => {
      index[node] = counter;
      low[node] = counter;
      counter += 1;
      stack.push(node);
      onStack[node] = 1;
      frames.push({ node, next: successors(node), at: 0 });
    };
    open(root);
    while (frames.length > 0) {
      const frame = frames.at(-1)!;
      if (frame.at < frame.next.length) {
        const next = frame.next[frame.at]!;
        frame.at += 1;
        if (index[next] === -1) {
          open(next);
        } else if (onStack[next]) {
          low[frame.node] = Math.min(low[frame.node]!, index[next]!);
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
        if (component.length > 1 || frame.next.includes(frame.node)) {
          cycles.push(component.toSorted((a, b) => a - b));
        }
      }
    }
  }
  return cycles.toSorted((a, b) => a[0]! - b[0]!).map((cycle) => cycle.map((i) => keys[i]!));
}
