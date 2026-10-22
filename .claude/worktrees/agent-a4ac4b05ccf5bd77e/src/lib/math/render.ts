// SPEC-056 render layout: turns a MathNode tree into a serialisable RenderBox
// tree the Vue renderer walks. Layout logic (spacing, size progression, rule
// metrics) lives here in framework-free code per CLAUDE.md — only the thin
// EquationView component consumes the result. Rendering is DOM: nested boxes
// and CSS borders for the fraction/radical rules, so selection, text search,
// screen-reader access and PDF text extraction all come for free.
import type { MathNode } from "../documentModel";
import {
  eighteenthsToEm,
  fractionScale,
  MIN_FONT_SIZE_PT,
  RULE_CLEARANCE_EM,
  RULE_THICKNESS_EM,
  scriptScale,
  spacingForRow,
} from "./layout";

export interface RenderItem {
  /** Space to insert before this item, as a CSS em length ("0" when none). */
  spaceBefore: string;
  box: RenderBox;
}

export type RenderBox =
  | { type: "row"; items: RenderItem[] }
  | { type: "glyph"; text: string; variant: "italic" | "upright" }
  | { type: "text"; text: string }
  | { type: "frac"; num: RenderBox; den: RenderBox; bar: boolean; fontScale?: string }
  | { type: "script"; base: RenderBox; sup?: RenderBox; sub?: RenderBox; stackScale?: string }
  | { type: "radical"; radicand: RenderBox; index?: RenderBox }
  | { type: "nary"; op: string; lower?: RenderBox; upper?: RenderBox; body: RenderBox }
  | { type: "delim"; open: string; close: string; body: RenderBox; stretch: number }
  | { type: "matrix"; rows: RenderBox[][] }
  | { type: "accent"; accent: string; base: RenderBox }
  | { type: "func"; name: string; body: RenderBox }
  | { type: "opaque"; linear: string };

/** How deeply a box is nested inside size-shrinking constructs. Script and
 * fraction nesting shrink independently (AC-39/40); both floor at half size,
 * and the combined result is never rendered below MIN_FONT_SIZE_PT (AC-41). */
interface LayoutDepth {
  /** Number of enclosing script stacks (sup/sub) — drives scriptScale. */
  script: number;
  /** Number of enclosing fractions — drives fractionScale. */
  frac: number;
}

const TOP_DEPTH: LayoutDepth = { script: 0, frac: 0 };

/** The CSS `font-size` for a box that shrinks to `factor` of its parent, with
 * the AC-41 5 pt floor expressed as a `max()` so it holds at every nesting
 * level regardless of the equation's base size. Returns undefined when the box
 * does not shrink relative to its parent (factor ≥ 1), so no style is emitted. */
function shrinkFontSize(factor: number): string | undefined {
  if (factor >= 1) return undefined;
  const em = factor.toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
  return `max(${MIN_FONT_SIZE_PT}pt, ${em}em)`;
}

/** Metrics the component reads for CSS, exported so the numbers live in one
 * place with the layout rules. */
export const RENDER_METRICS = {
  ruleThicknessEm: RULE_THICKNESS_EM,
  ruleClearanceEm: RULE_CLEARANCE_EM,
  scriptScale: scriptScale(1),
  scriptScriptScale: scriptScale(2),
};

/** Build the render tree for a row of nodes, computing inter-node spacing from
 * the SPEC-056 spacing table with the AC-11 sign rule. */
export function layoutRow(nodes: MathNode[], depth: LayoutDepth = TOP_DEPTH): RenderBox {
  const spacing = spacingForRow(nodes);
  const items: RenderItem[] = nodes.map((node, i) => ({
    spaceBefore: eighteenthsToEm(spacing[i]),
    box: layoutNode(node, depth),
  }));
  return { type: "row", items };
}

function layoutNode(node: MathNode, depth: LayoutDepth): RenderBox {
  switch (node.kind) {
    case "sym":
      return { type: "glyph", text: node.text, variant: node.variant };
    case "text":
      return { type: "text", text: node.text };
    case "frac": {
      // This fraction sits inside `depth.frac` other fractions; its own size is
      // fractionScale(depth.frac) of the equation base, expressed as a factor
      // relative to its parent so CSS `em` inheritance compounds it (AC-40).
      const factor = fractionScale(depth.frac) / fractionScale(depth.frac - 1 < 0 ? 0 : depth.frac - 1);
      const inner: LayoutDepth = { ...depth, frac: depth.frac + 1 };
      const box: RenderBox = {
        type: "frac",
        num: layoutRow(node.num, inner),
        den: layoutRow(node.den, inner),
        bar: node.bar,
      };
      const fontScale = shrinkFontSize(depth.frac === 0 ? 1 : factor);
      if (fontScale) box.fontScale = fontScale;
      return box;
    }
    case "script": {
      // The base stays at the surrounding size; the sup/sub stack is one script
      // level deeper, shrinking toward the scriptScale floor (AC-39).
      const stackFactor = scriptScale(depth.script + 1) / scriptScale(depth.script);
      const inner: LayoutDepth = { ...depth, script: depth.script + 1 };
      const box: RenderBox = { type: "script", base: layoutRow(node.base, depth) };
      if (node.sup) box.sup = layoutRow(node.sup, inner);
      if (node.sub) box.sub = layoutRow(node.sub, inner);
      const stackScale = shrinkFontSize(stackFactor);
      if (stackScale) box.stackScale = stackScale;
      return box;
    }
    case "radical": {
      const box: RenderBox = { type: "radical", radicand: layoutRow(node.radicand, depth) };
      if (node.index && node.index.length) box.index = layoutRow(node.index, depth);
      return box;
    }
    case "nary": {
      const box: RenderBox = { type: "nary", op: node.op, body: layoutRow(node.body, depth) };
      if (node.lower && node.lower.length) box.lower = layoutRow(node.lower, depth);
      if (node.upper && node.upper.length) box.upper = layoutRow(node.upper, depth);
      return box;
    }
    case "delim":
      return {
        type: "delim",
        open: node.open,
        close: node.close,
        body: layoutRow(node.body, depth),
        stretch: delimStretch(node.body),
      };
    case "matrix":
      return { type: "matrix", rows: node.rows.map((row) => row.map((cell) => layoutRow(cell, depth))) };
    case "accent":
      return { type: "accent", accent: node.accent, base: layoutRow(node.base, depth) };
    case "func":
      return { type: "func", name: node.name, body: layoutRow(node.body, depth) };
    case "opaqueMath":
      return { type: "opaque", linear: node.linear };
  }
}

/** How many rows tall the delimiter must stretch, so the renderer knows when a
 * CSS transform suffices (≤3×) and when to assemble Unicode bracket pieces
 * (AC-44). A matrix drives the height. */
function delimStretch(body: MathNode[]): number {
  let maxRows = 1;
  for (const node of body) {
    if (node.kind === "matrix") maxRows = Math.max(maxRows, node.rows.length);
  }
  return maxRows;
}
