// SPEC-056 in-equation editing engine. Pure, framework-free logic (CLAUDE.md:
// document logic lives in src/lib as plain TypeScript) so every editing rule —
// live linear-input typing, template-palette insertion, slot navigation,
// structural arrow walking and backspace-unwrap — is unit-testable without
// mounting a component. The Vue editor (EquationEditor.vue) is a thin driver
// that turns keystrokes into calls on this module and re-renders from the
// returned body.
//
// The caret is *structural*: a `path` descending from the root body into a
// construct's slot, plus an `offset` index within that slot's row. Typing keeps
// a `pending` linear buffer whose parse (parseLinear) always occupies
// `row[runStart..offset)`; every keystroke re-parses that run, so `\frac{a}{b}`
// becomes a fraction as it is typed, never on blur (AC-4). Any structural action
// seals the run first.
import type { MathNode } from "../documentModel";
import { parseLinear } from "./linearParse";
import { linearForm } from "./linearForm";
import { variantForText } from "./symbols";

/** One descent into a node's slot: the node's position in its parent row and
 * which slot of it the caret entered. */
export interface PathStep {
  index: number;
  slot: string;
}

/** A structural caret: the descent path to a slot's row, and the offset within
 * that row (0..row.length). */
export interface MathCaret {
  path: PathStep[];
  offset: number;
}

export interface EditState {
  body: MathNode[];
  caret: MathCaret;
  /** Raw linear text typed since the current run began (AC-4). Its parse
   * occupies `row[runStart..caret.offset)`. */
  pending: string;
  /** Offset in the caret's row where the current typing run began. */
  runStart: number;
}

// --- slots -----------------------------------------------------------------

/** The ordered slot names of a node (the order Tab and the arrows walk). Only
 * slots that currently exist are listed, so an optional script/limit that is
 * absent is skipped. Atoms (sym/text/opaqueMath) have no slots. */
export function slotNames(node: MathNode): string[] {
  switch (node.kind) {
    case "frac":
      return ["num", "den"];
    case "script": {
      const s = ["base"];
      if (node.sub) s.push("sub");
      if (node.sup) s.push("sup");
      return s;
    }
    case "radical": {
      const s: string[] = [];
      if (node.index) s.push("index");
      s.push("radicand");
      return s;
    }
    case "nary": {
      const s: string[] = [];
      if (node.lower) s.push("lower");
      if (node.upper) s.push("upper");
      s.push("body");
      return s;
    }
    case "delim":
      return ["body"];
    case "accent":
      return ["base"];
    case "func":
      return ["body"];
    case "matrix": {
      const s: string[] = [];
      node.rows.forEach((row, r) => row.forEach((_, c) => s.push(`cell:${r}:${c}`)));
      return s;
    }
    default:
      return [];
  }
}

/** The mutable row backing a node's slot, or null when the slot does not
 * exist. Returns the live array so an in-place splice mutates the node. */
export function getSlot(node: MathNode, slot: string): MathNode[] | null {
  switch (node.kind) {
    case "frac":
      return slot === "num" ? node.num : slot === "den" ? node.den : null;
    case "script":
      return slot === "base" ? node.base : slot === "sub" ? node.sub ?? null : slot === "sup" ? node.sup ?? null : null;
    case "radical":
      return slot === "radicand" ? node.radicand : slot === "index" ? node.index ?? null : null;
    case "nary":
      return slot === "body"
        ? node.body
        : slot === "lower"
          ? node.lower ?? null
          : slot === "upper"
            ? node.upper ?? null
            : null;
    case "delim":
      return slot === "body" ? node.body : null;
    case "accent":
      return slot === "base" ? node.base : null;
    case "func":
      return slot === "body" ? node.body : null;
    case "matrix": {
      const m = /^cell:(\d+):(\d+)$/.exec(slot);
      if (!m) return null;
      return node.rows[Number(m[1])]?.[Number(m[2])] ?? null;
    }
    default:
      return null;
  }
}

/** The row a path points into, walking from the root body. Returns the live
 * array (mutating it mutates `body`), or null when the path is stale. */
export function rowAt(body: MathNode[], path: PathStep[]): MathNode[] | null {
  let row = body;
  for (const step of path) {
    const node = row[step.index];
    if (!node) return null;
    const slot = getSlot(node, step.slot);
    if (!slot) return null;
    row = slot;
  }
  return row;
}

/** The node the last path step descends into (the innermost enclosing
 * construct), plus its parent row — null at the top level. */
function enclosing(body: MathNode[], path: PathStep[]): { parent: MathNode[]; node: MathNode; step: PathStep } | null {
  if (path.length === 0) return null;
  const parentPath = path.slice(0, -1);
  const parent = rowAt(body, parentPath);
  const step = path[path.length - 1];
  if (!parent) return null;
  const node = parent[step.index];
  if (!node) return null;
  return { parent, node, step };
}

// --- construction ----------------------------------------------------------

function clone(state: EditState): EditState {
  return { body: structuredClone(state.body), caret: structuredClone(state.caret), pending: state.pending, runStart: state.runStart };
}

/** A fresh edit state over `body`, caret at the start (default) or end of the
 * top-level row. */
export function createEditState(body: MathNode[], atEnd = false): EditState {
  const offset = atEnd ? body.length : 0;
  return { body: structuredClone(body), caret: { path: [], offset }, pending: "", runStart: offset };
}

/** Seal the current typing run: commit `pending` as final nodes and reset the
 * run to the caret. Every structural op does this first so a half-typed
 * command never leaks into a navigation/insert action. */
export function sealRun(state: EditState): EditState {
  if (state.pending === "" && state.runStart === state.caret.offset) return state;
  return { ...state, pending: "", runStart: state.caret.offset };
}

// --- typing (AC-4) ---------------------------------------------------------

/** Type one character of linear input. Appends to the pending run and
 * re-parses it, so a construct appears live as its closing brace is typed. */
export function typeChar(state: EditState, ch: string): EditState {
  const next = clone(state);
  const row = rowAt(next.body, next.caret.path);
  if (!row) return state;
  const newPending = next.pending + ch;
  const nodes = parseLinear(newPending);
  // Replace the current run's nodes (row[runStart..offset)) with the new parse.
  row.splice(next.runStart, next.caret.offset - next.runStart, ...structuredClone(nodes));
  next.pending = newPending;
  next.caret.offset = next.runStart + nodes.length;
  return next;
}

/** Type a string, character by character, so live parsing runs at each step
 * (AC-4). */
export function typeString(state: EditState, text: string): EditState {
  let out = state;
  for (const ch of text) out = typeChar(out, ch);
  return out;
}

// --- glyph insertion (character picker, AC-14/15) --------------------------

/** Insert a literal glyph as a `sym` node (its variant from the SPEC-056
 * italic/upright rule) at the caret — how SPEC-026's character picker drops a
 * Maths or Greek character into an equation. */
export function insertGlyph(state: EditState, glyph: string): EditState {
  const sealed = sealRun(state);
  const next = clone(sealed);
  const row = rowAt(next.body, next.caret.path);
  if (!row) return state;
  const node: MathNode = { kind: "sym", text: glyph, variant: variantForText(glyph) };
  row.splice(next.caret.offset, 0, node);
  next.caret.offset += 1;
  next.runStart = next.caret.offset;
  return next;
}

// --- template palette (AC-5) -----------------------------------------------

export interface ConstructTemplate {
  id: string;
  label: string;
  make: () => MathNode;
}

/** Every construct in the closed set, with empty slots — the equation
 * toolbar's template palette (AC-5). `opaqueMath` is the preservation case,
 * not a template, so it is deliberately absent. */
export const CONSTRUCT_TEMPLATES: ConstructTemplate[] = [
  { id: "frac", label: "Fraction", make: () => ({ kind: "frac", num: [], den: [], bar: true }) },
  { id: "binom", label: "Binomial", make: () => ({ kind: "frac", num: [], den: [], bar: false }) },
  { id: "sup", label: "Superscript", make: () => ({ kind: "script", base: [], sup: [] }) },
  { id: "sub", label: "Subscript", make: () => ({ kind: "script", base: [], sub: [] }) },
  { id: "subsup", label: "Sub/superscript", make: () => ({ kind: "script", base: [], sub: [], sup: [] }) },
  { id: "sqrt", label: "Square root", make: () => ({ kind: "radical", radicand: [] }) },
  { id: "root", label: "Nth root", make: () => ({ kind: "radical", radicand: [], index: [] }) },
  { id: "sum", label: "Summation", make: () => ({ kind: "nary", op: "∑", lower: [], upper: [], body: [] }) },
  { id: "int", label: "Integral", make: () => ({ kind: "nary", op: "∫", lower: [], upper: [], body: [] }) },
  { id: "delim", label: "Parentheses", make: () => ({ kind: "delim", open: "(", close: ")", body: [] }) },
  { id: "matrix", label: "Matrix", make: () => ({ kind: "matrix", rows: [[[], []], [[], []]] }) },
  { id: "accent", label: "Hat", make: () => ({ kind: "accent", accent: "̂", base: [] }) },
  { id: "func", label: "Function", make: () => ({ kind: "func", name: "sin", body: [] }) },
  { id: "text", label: "Text", make: () => ({ kind: "text", text: "" }) },
];

/** Insert a construct node at the caret and move the caret into its first slot
 * (AC-5). A slotless node (text) leaves the caret after it. */
export function insertConstruct(state: EditState, node: MathNode): EditState {
  const sealed = sealRun(state);
  const next = clone(sealed);
  const row = rowAt(next.body, next.caret.path);
  if (!row) return state;
  const at = next.caret.offset;
  row.splice(at, 0, structuredClone(node));
  const slots = slotNames(node);
  if (slots.length > 0) {
    next.caret.path = [...next.caret.path, { index: at, slot: slots[0] }];
    next.caret.offset = 0;
  } else {
    next.caret.offset = at + 1;
  }
  next.runStart = next.caret.offset;
  return next;
}

/** Insert the template with the given id (AC-5), for the palette buttons. */
export function insertTemplate(state: EditState, id: string): EditState {
  const template = CONSTRUCT_TEMPLATES.find((t) => t.id === id);
  if (!template) return state;
  return insertConstruct(state, template.make());
}

// --- slot navigation (Tab / Shift+Tab, AC-6/7) -----------------------------

/** Tab: move to the next slot of the innermost enclosing construct; from its
 * last slot, step out to just after the whole construct (AC-6). */
export function nextSlot(state: EditState): EditState {
  const sealed = sealRun(state);
  const enc = enclosing(sealed.body, sealed.caret.path);
  if (!enc) return sealed;
  const slots = slotNames(enc.node);
  const i = slots.indexOf(enc.step.slot);
  if (i >= 0 && i < slots.length - 1) {
    const path = sealed.caret.path.slice();
    path[path.length - 1] = { index: enc.step.index, slot: slots[i + 1] };
    return withCaret(sealed, path, 0);
  }
  // Last slot: exit to after the construct in the parent row.
  return withCaret(sealed, sealed.caret.path.slice(0, -1), enc.step.index + 1);
}

/** Shift+Tab: move to the previous slot of the innermost construct (caret at
 * its end); from the first slot, step out to just before the construct
 * (AC-7). */
export function prevSlot(state: EditState): EditState {
  const sealed = sealRun(state);
  const enc = enclosing(sealed.body, sealed.caret.path);
  if (!enc) return sealed;
  const slots = slotNames(enc.node);
  const i = slots.indexOf(enc.step.slot);
  if (i > 0) {
    const path = sealed.caret.path.slice();
    path[path.length - 1] = { index: enc.step.index, slot: slots[i - 1] };
    const row = rowAt(sealed.body, path);
    return withCaret(sealed, path, row ? row.length : 0);
  }
  return withCaret(sealed, sealed.caret.path.slice(0, -1), enc.step.index);
}

function withCaret(state: EditState, path: PathStep[], offset: number): EditState {
  return { body: state.body, caret: { path, offset }, pending: "", runStart: offset };
}

// --- arrow navigation (AC-12) ----------------------------------------------

/** Right arrow: walk one step through the structure. Into a container to the
 * right; over an atom; from a slot's end into the construct's next slot (so
 * Right from the numerator's end enters the denominator) or out past it. */
export function moveRight(state: EditState): EditState {
  const sealed = sealRun(state);
  const row = rowAt(sealed.body, sealed.caret.path);
  if (!row) return sealed;
  if (sealed.caret.offset < row.length) {
    const node = row[sealed.caret.offset];
    const slots = slotNames(node);
    if (slots.length > 0) {
      return withCaret(sealed, [...sealed.caret.path, { index: sealed.caret.offset, slot: slots[0] }], 0);
    }
    return withCaret(sealed, sealed.caret.path, sealed.caret.offset + 1);
  }
  // End of the row: into the next slot of the enclosing construct, or out.
  const enc = enclosing(sealed.body, sealed.caret.path);
  if (!enc) return sealed;
  const slots = slotNames(enc.node);
  const i = slots.indexOf(enc.step.slot);
  if (i >= 0 && i < slots.length - 1) {
    const path = sealed.caret.path.slice();
    path[path.length - 1] = { index: enc.step.index, slot: slots[i + 1] };
    return withCaret(sealed, path, 0);
  }
  return withCaret(sealed, sealed.caret.path.slice(0, -1), enc.step.index + 1);
}

/** Left arrow: the mirror of moveRight — into the last slot of a container to
 * the left (caret at its end), over an atom, or out to before the construct. */
export function moveLeft(state: EditState): EditState {
  const sealed = sealRun(state);
  const row = rowAt(sealed.body, sealed.caret.path);
  if (!row) return sealed;
  if (sealed.caret.offset > 0) {
    const node = row[sealed.caret.offset - 1];
    const slots = slotNames(node);
    if (slots.length > 0) {
      const slot = slots[slots.length - 1];
      const path = [...sealed.caret.path, { index: sealed.caret.offset - 1, slot }];
      const inner = rowAt(sealed.body, path);
      return withCaret(sealed, path, inner ? inner.length : 0);
    }
    return withCaret(sealed, sealed.caret.path, sealed.caret.offset - 1);
  }
  const enc = enclosing(sealed.body, sealed.caret.path);
  if (!enc) return sealed;
  const slots = slotNames(enc.node);
  const i = slots.indexOf(enc.step.slot);
  if (i > 0) {
    const path = sealed.caret.path.slice();
    path[path.length - 1] = { index: enc.step.index, slot: slots[i - 1] };
    const inner = rowAt(sealed.body, path);
    return withCaret(sealed, path, inner ? inner.length : 0);
  }
  return withCaret(sealed, sealed.caret.path.slice(0, -1), enc.step.index);
}

/** Up arrow: from a lower slot to the paired upper one (denominator→numerator,
 * subscript→superscript, an n-ary's lower→upper, a matrix cell→the cell
 * above). A no-op elsewhere. */
export function moveUp(state: EditState): EditState {
  return moveVertical(sealRun(state), "up");
}

/** Down arrow: the mirror of moveUp. */
export function moveDown(state: EditState): EditState {
  return moveVertical(sealRun(state), "down");
}

function moveVertical(state: EditState, dir: "up" | "down"): EditState {
  const enc = enclosing(state.body, state.caret.path);
  if (!enc) return state;
  const path = state.caret.path.slice();
  const setSlot = (slot: string): EditState => {
    path[path.length - 1] = { index: enc.step.index, slot };
    const inner = rowAt(state.body, path);
    const off = Math.min(state.caret.offset, inner ? inner.length : 0);
    return withCaret(state, path, off);
  };
  if (enc.node.kind === "frac") {
    if (dir === "up" && enc.step.slot === "den") return setSlot("num");
    if (dir === "down" && enc.step.slot === "num") return setSlot("den");
  } else if (enc.node.kind === "script") {
    if (dir === "up" && enc.step.slot === "sub" && enc.node.sup) return setSlot("sup");
    if (dir === "down" && enc.step.slot === "sup" && enc.node.sub) return setSlot("sub");
  } else if (enc.node.kind === "nary") {
    if (dir === "up" && enc.step.slot === "lower" && enc.node.upper) return setSlot("upper");
    if (dir === "down" && enc.step.slot === "upper" && enc.node.lower) return setSlot("lower");
  } else if (enc.node.kind === "matrix") {
    const m = /^cell:(\d+):(\d+)$/.exec(enc.step.slot);
    if (m) {
      const r = Number(m[1]) + (dir === "up" ? -1 : 1);
      const c = Number(m[2]);
      if (enc.node.rows[r]?.[c]) return setSlot(`cell:${r}:${c}`);
    }
  }
  return state;
}

// --- backspace (AC-13) -----------------------------------------------------

/** Backspace. While a run is being typed, delete its last character and
 * re-parse (AC-4 in reverse). Otherwise delete the atom to the left; and at
 * the start of a construct's first slot, unwrap the construct, keeping its
 * content (AC-13). */
export function backspace(state: EditState): EditState {
  if (state.pending.length > 0) {
    const next = clone(state);
    const row = rowAt(next.body, next.caret.path);
    if (!row) return state;
    const newPending = next.pending.slice(0, -1);
    const nodes = parseLinear(newPending);
    row.splice(next.runStart, next.caret.offset - next.runStart, ...structuredClone(nodes));
    next.pending = newPending;
    next.caret.offset = next.runStart + nodes.length;
    return next;
  }
  const next = clone(state);
  const row = rowAt(next.body, next.caret.path);
  if (!row) return state;
  if (next.caret.offset > 0) {
    // Delete the atom (or whole construct) immediately to the left.
    row.splice(next.caret.offset - 1, 1);
    next.caret.offset -= 1;
    next.runStart = next.caret.offset;
    return next;
  }
  // Offset 0: unwrap the enclosing construct if this is its first slot,
  // otherwise step to the end of the previous slot / out (AC-13).
  const enc = enclosing(next.body, next.caret.path);
  if (!enc) return state;
  const slots = slotNames(enc.node);
  if (enc.step.slot === slots[0]) {
    return unwrap(next, enc);
  }
  return prevSlot(next);
}

/** Replace an enclosing construct with the concatenation of its slots'
 * contents, caret at the start of the re-inserted content (AC-13). */
function unwrap(state: EditState, enc: { parent: MathNode[]; node: MathNode; step: PathStep }): EditState {
  const content = slotNames(enc.node).flatMap((slot) => getSlot(enc.node, slot) ?? []);
  enc.parent.splice(enc.step.index, 1, ...structuredClone(content));
  return withCaret(state, state.caret.path.slice(0, -1), enc.step.index);
}

// --- readout ---------------------------------------------------------------

/** The linear form of the current body — the default alt text (AC-19) and the
 * source an equation read from a file is edited through (AC-46). */
export function editLinearForm(state: EditState): string {
  return linearForm(state.body);
}
