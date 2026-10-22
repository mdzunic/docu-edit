// SPEC-056 linear-form generator: MathNode[] → the LaTeX-subset string. The
// linear form is regenerated from the tree for editing and for `altText`, so
// an equation read from a .docx (which carries no LaTeX) is editable exactly
// like one typed locally (AC-46), and every equation has a default alt text
// that is its linear form (AC-19).
import type { EquationSpec, MathNode } from "../documentModel";
import { nameForGlyph } from "./symbols";

/** Regenerate the linear form of a row of nodes. */
export function linearForm(nodes: MathNode[]): string {
  return nodes.map(nodeLinear).join("");
}

/** The alt text for an equation: its explicit `altText` when set, otherwise
 * the linear form of its body (AC-19). An explicit "" means "no alt text". */
export function equationAltText(eq: EquationSpec): string {
  return eq.altText !== undefined ? eq.altText : linearForm(eq.body);
}

function nodeLinear(node: MathNode): string {
  switch (node.kind) {
    case "sym": {
      const name = nameForGlyph(node.text);
      return name ? `\\${name}` : node.text;
    }
    case "frac":
      return `${node.bar ? "\\frac" : "\\binom"}{${linearForm(node.num)}}{${linearForm(node.den)}}`;
    case "script": {
      const base = groupIfNeeded(node.base);
      let out = base;
      if (node.sub) out += `_${groupIfNeeded(node.sub)}`;
      if (node.sup) out += `^${groupIfNeeded(node.sup)}`;
      return out;
    }
    case "radical":
      return node.index && node.index.length
        ? `\\sqrt[${linearForm(node.index)}]{${linearForm(node.radicand)}}`
        : `\\sqrt{${linearForm(node.radicand)}}`;
    case "nary": {
      const cmd = naryCommand(node.op);
      let out = `\\${cmd}`;
      if (node.lower) out += `_{${linearForm(node.lower)}}`;
      if (node.upper) out += `^{${linearForm(node.upper)}}`;
      if (node.body.length) out += linearForm(node.body);
      return out;
    }
    case "delim":
      return `\\left${delimToken(node.open)} ${linearForm(node.body)} \\right${delimToken(node.close)}`;
    case "matrix":
      return `\\matrix{${node.rows
        .map((row) => row.map((cell) => linearForm(cell)).join(" & "))
        .join(" \\\\ ")}}`;
    case "accent": {
      const cmd = accentCommand(node.accent);
      return `\\${cmd}{${linearForm(node.base)}}`;
    }
    case "func":
      return node.body.length ? `\\${node.name}{${linearForm(node.body)}}` : `\\${node.name}`;
    case "text":
      return `\\text{${node.text}}`;
    case "opaqueMath":
      return node.linear;
  }
}

/** Wrap a sub-row in braces unless it is a single simple node. */
function groupIfNeeded(nodes: MathNode[]): string {
  if (nodes.length === 1) {
    const only = nodes[0];
    if (only.kind === "sym" && only.text.length === 1 && nameForGlyph(only.text) === null) {
      return only.text;
    }
  }
  return `{${linearForm(nodes)}}`;
}

const NARY_COMMANDS: Record<string, string> = {
  "∑": "sum",
  "∏": "prod",
  "∐": "coprod",
  "∫": "int",
  "∬": "iint",
  "∭": "iiint",
  "∮": "oint",
  "⋃": "bigcup",
  "⋂": "bigcap",
};

function naryCommand(op: string): string {
  return NARY_COMMANDS[op] ?? nameForGlyph(op) ?? "sum";
}

const ACCENT_COMMANDS: Record<string, string> = {
  "̂": "hat",
  "̄": "bar",
  "⃗": "vec",
  "̇": "dot",
  "̈": "ddot",
  "̃": "tilde",
  "̌": "check",
  "̆": "breve",
  "́": "acute",
  "̀": "grave",
};

function accentCommand(accent: string): string {
  return ACCENT_COMMANDS[accent] ?? "hat";
}

function delimToken(delim: string): string {
  if (delim === "") return ".";
  if (delim === "{") return "\\{";
  if (delim === "}") return "\\}";
  if (delim === "‖") return "\\|";
  if (delim === "⟨") return "\\langle";
  if (delim === "⟩") return "\\rangle";
  return delim;
}
