// SPEC-056 Layout rules, made normative and testable. "It looks right" is not
// a test, so the spacing table, the size progression and the rule metrics all
// live here as pure functions the renderer consumes and the unit tests pin
// rule-by-rule.
import type { MathNode } from "../documentModel";

/** The math-class of a node, which drives inter-node spacing. */
export type MathClass = "ord" | "bin" | "rel" | "punct" | "func" | "nary" | "open" | "close";

const BINARY_OPERATORS = new Set(["+", "−", "-", "±", "×", "÷", "∓", "∗", "∙", "·", "⊕", "⊗", "∪", "∩", "∧", "∨", "*"]);
const RELATIONS = new Set(["=", "≠", "<", ">", "≤", "≥", "≈", "≡", "≅", "∼", "∈", "∉", "⊂", "⊃", "⊆", "⊇", "≪", "≫", "∝"]);
const PUNCTUATION = new Set([",", ";"]);

/** The math-class of a node (SPEC-056 spacing table). */
export function classify(node: MathNode): MathClass {
  switch (node.kind) {
    case "sym": {
      const t = node.text;
      if (BINARY_OPERATORS.has(t)) return "bin";
      if (RELATIONS.has(t)) return "rel";
      if (PUNCTUATION.has(t)) return "punct";
      return "ord";
    }
    case "func":
      return "func";
    case "nary":
      return "nary";
    case "delim":
      return "ord";
    default:
      return "ord";
  }
}

/** Inter-node space in eighteenths of an em (SPEC-056 spacing table). `left`
 * is null when `right` is the first node in its row — a binary operator there
 * is a sign, not an operator, and takes no space (AC-11). */
export function spaceEighteenths(left: MathNode | null, right: MathNode): number {
  const rc = classify(right);
  if (left === null) {
    return 0;
  }
  const lc = classify(left);

  // A binary operator with nothing operand-like to its left is a sign (AC-11)
  // and takes no space. (Row-level context — a sign mid-row — is applied by
  // spacingForRow, which the renderer uses.)
  if (rc === "bin" && !isOperand(lc)) return 0;

  // after a comma or semicolon
  if (lc === "punct") return 3;

  // ordinary ↔ binary operator
  if ((lc === "ord" && rc === "bin") || (lc === "bin" && rc === "ord")) return 4;
  // ordinary ↔ relation
  if ((isOperand(lc) && rc === "rel") || (lc === "rel" && isOperand(rc))) return 5;

  // func name → its argument
  if (lc === "func") return 3;
  // either side of an n-ary body
  if (lc === "nary" || rc === "nary") return 3;

  return 0;
}

function isOperand(c: MathClass): boolean {
  return c === "ord" || c === "close" || c === "func" || c === "nary";
}

/** The spacing (in eighteenths of an em) to place *before* each node in a row.
 * Index 0 is always 0. A binary operator that follows nothing operand-like is
 * treated as a sign and contributes no space (AC-11). */
export function spacingForRow(row: MathNode[]): number[] {
  const out: number[] = [];
  for (let i = 0; i < row.length; i++) {
    const left = i === 0 ? null : row[i - 1];
    const right = row[i];
    // Recompute the sign rule with real row context: a binary operator is a
    // sign when the node to its left is absent or not an operand.
    if (classify(right) === "bin") {
      const leftIsOperand = left !== null && isOperand(classify(left));
      out.push(leftIsOperand ? 4 : 0);
      continue;
    }
    if (left !== null && classify(left) === "bin") {
      // The space after a binary operator only applies when that operator was
      // itself a real operator (had an operand to ITS left).
      const beforeLeft = i >= 2 ? row[i - 2] : null;
      const leftWasOperator = beforeLeft !== null && isOperand(classify(beforeLeft));
      out.push(leftWasOperator ? 4 : 0);
      continue;
    }
    out.push(spaceEighteenths(left, right));
  }
  return out;
}

/** em fraction (÷18) → a CSS `em` length string, collapsing 0 to "0". */
export function eighteenthsToEm(n: number): string {
  if (n === 0) return "0";
  return `${(n / 18).toFixed(4).replace(/0+$/, "").replace(/\.$/, "")}em`;
}

// --- Sizes (SPEC-056 "Sizes") ---------------------------------------------

/** Relative scale of a script at a given nesting depth (AC-39): the base is
 * depth 0 (1.0), a script is depth 1 (0.71), a script of a script is depth 2
 * (0.5) and it does not shrink further. */
export function scriptScale(depth: number): number {
  if (depth <= 0) return 1;
  if (depth === 1) return 0.71;
  return 0.5;
}

/** A fraction nested inside another fraction's numerator/denominator is 71% of
 * its parent, and does not compound below the script floor (AC-40). */
export function fractionScale(depth: number): number {
  if (depth <= 0) return 1;
  if (depth === 1) return 0.71;
  return 0.5;
}

/** The minimum rendered type size: nothing renders below 5 pt (AC-41). */
export const MIN_FONT_SIZE_PT = 5;

/** Clamp a computed size to the 5 pt floor (AC-41). */
export function clampFontSize(pt: number): number {
  return Math.max(pt, MIN_FONT_SIZE_PT);
}

// --- Rules (SPEC-056 "Rules") ---------------------------------------------

/** Fraction/radical bar thickness: 1/18 em (AC-42). */
export const RULE_THICKNESS_EM = 1 / 18;
/** Clearance above and below a bar's content: 1/6 em (AC-42). */
export const RULE_CLEARANCE_EM = 1 / 6;
