// Function help (SPEC-127 "Function metadata", "The assists") — the
// framework-free half of the three formula-editing assists, over the
// registry's metadata:
//
// - functionSignature: Excel's signature form — required arguments, optional
//   ones bracketed, a variadic function's repeating group shown once,
//   bracketed, then ", …" (`SUMIFS(sum_range, criteria_range1, criteria1,
//   [criteria_range2, criteria2], …)`). signatureParts is the same text in
//   pieces for the argument hint, the repeating group repeated up to the
//   instance holding the current argument, which is marked.
// - argumentAt: an argument index's display name — past the end of the
//   repeating group, `args[f + (i − f) mod g]` with its trailing number
//   raised by ⌊(i − f) ÷ g⌋ (`number7`, `criteria_range3`); null past maxArgs.
// - callAtCaret: the innermost enclosing call of a registered function and
//   the argument the caret is in, from the tokens before the caret with a
//   frame stack — a `(` after a function name opens a call, any other `(` a
//   group, `{` an array constant; a comma advances only a call on top, so
//   commas in nested calls, parentheses, array constants, quoted sheet names
//   and strings never advance the outer argument. Inside an unterminated
//   string there is no call.
// - completionAt: the identifier being typed where a function may appear —
//   after the leading "=", an operator, "(" or "," — outside any string.
// - searchFunctions: the Insert Function dialog's list — name prefix, then
//   name substring, then every query word in the summary, alphabetical
//   within each; the empty query lists the category alphabetically.
// Every string here is en-US whatever the locale (formulas are en-US).
import { getFunction, registeredFunctionNames, type FunctionCategory, type FunctionSpec } from "./registry";
import { tokenize, type Token } from "./tokens";

type Metadata = Pick<FunctionSpec, "name" | "args" | "maxArgs">;

/** The first repeating argument's index, or -1. */
function firstRepeating(spec: Metadata): number {
  return spec.args.findIndex((arg) => arg.repeating === true);
}

/** A display name with its trailing number raised by `by`. */
function renumbered(name: string, by: number): string {
  const match = /^(.*?)(\d+)$/.exec(name);
  return match ? `${match[1]}${Number(match[2]) + by}` : name;
}

/** The display name of argument `index`, or null past maxArgs. */
export function argumentAt(spec: Metadata, index: number): string | null {
  if (index < 0 || index >= spec.maxArgs) {
    return null;
  }
  const f = firstRepeating(spec);
  if (f < 0 || index < f) {
    return spec.args[index]?.name ?? null;
  }
  const g = spec.args.length - f;
  return renumbered(spec.args[f + ((index - f) % g)]!.name, Math.floor((index - f) / g));
}

/** One piece of a signature: text, and whether it is the current argument. */
export interface SignaturePart {
  text: string;
  current?: true;
}

/** A signature in pieces with argument `current` marked (see the module
 * header); no piece is marked past maxArgs. */
export function signatureParts(spec: Metadata, current = -1): SignaturePart[] {
  const parts: SignaturePart[] = [{ text: `${spec.name}(` }];
  const items: { names: { text: string; index: number }[]; bracketed: boolean }[] = [];
  const f = firstRepeating(spec);
  const fixed = f < 0 ? spec.args.length : f;
  for (let i = 0; i < fixed; i += 1) {
    items.push({ names: [{ text: spec.args[i]!.name, index: i }], bracketed: spec.args[i]!.optional === true });
  }
  if (f >= 0) {
    const g = spec.args.length - f;
    const shown = current >= f && current < spec.maxArgs ? Math.floor((current - f) / g) + 1 : 1;
    for (let instance = 0; instance < shown; instance += 1) {
      const names = [];
      for (let j = 0; j < g; j += 1) {
        const index = f + instance * g + j;
        names.push({ text: argumentAt(spec, index) ?? "", index });
      }
      items.push({ names, bracketed: true });
    }
  }
  items.forEach((item, n) => {
    if (n > 0) {
      parts.push({ text: ", " });
    }
    if (item.bracketed) {
      parts.push({ text: "[" });
    }
    item.names.forEach((name, k) => {
      if (k > 0) {
        parts.push({ text: ", " });
      }
      parts.push(name.index === current ? { text: name.text, current: true } : { text: name.text });
    });
    if (item.bracketed) {
      parts.push({ text: "]" });
    }
  });
  parts.push({ text: f >= 0 ? ", …)" : ")" });
  return parts;
}

/** Excel's signature form of a function (see the module header). */
export function functionSignature(spec: Metadata): string {
  return signatureParts(spec)
    .map((part) => part.text)
    .join("");
}

/** The tokens before the caret, or null inside a string (an unterminated
 * string or structured reference). */
function tokensBefore(text: string, caret: number): Token[] | null {
  const tokens = tokenize(text.slice(0, Math.max(0, Math.min(caret, text.length))));
  return tokens.some((token) => token.unterminated) ? null : tokens;
}

/** The innermost enclosing call of a registered function at the caret and
 * the argument index there, or null. */
export function callAtCaret(text: string, caret: number): { name: string; argIndex: number } | null {
  const tokens = tokensBefore(text, caret);
  if (!tokens) {
    return null;
  }
  const frames: { kind: "call" | "group" | "array"; name?: string; argIndex: number }[] = [];
  tokens.forEach((token, i) => {
    switch (token.kind) {
      case "open": {
        const previous = tokens[i - 1];
        frames.push(previous?.kind === "function" && token.lead === "" ? { kind: "call", name: previous.text, argIndex: 0 } : { kind: "group", argIndex: 0 });
        break;
      }
      case "arrayOpen":
        frames.push({ kind: "array", argIndex: 0 });
        break;
      case "argSep": {
        const top = frames.at(-1);
        if (top?.kind === "call") {
          top.argIndex += 1;
        }
        break;
      }
      case "close":
      case "arrayClose":
        frames.pop();
        break;
      default:
        break;
    }
  });
  for (let i = frames.length - 1; i >= 0; i -= 1) {
    const frame = frames[i]!;
    const spec = frame.kind === "call" && frame.name !== undefined ? getFunction(frame.name) : undefined;
    if (spec) {
      return { name: spec.name, argIndex: frame.argIndex };
    }
  }
  return null;
}

/** Kinds after which a function name may be typed. */
const CALL_POSITIONS = new Set<Token["kind"]>(["operator", "open", "argSep", "union"]);

/** An identifier: letters, digits, `_`, `.`, starting with a letter or `_`. */
const IDENTIFIER = /^[A-Za-z_][A-Za-z0-9_.]*$/;

/** The identifier typed at the caret where a function may appear — its
 * text and where it starts — or null (see the module header). */
export function completionAt(text: string, caret: number): { prefix: string; start: number } | null {
  if (!text.startsWith("=")) {
    return null;
  }
  const tokens = tokensBefore(text, caret);
  if (!tokens) {
    return null;
  }
  const end = tokens.at(-1);
  const significant = tokens.filter((token) => token.kind !== "end");
  const last = significant.at(-1);
  const previous = significant.at(-2);
  if (!last || !previous || (end && end.lead !== "") || !["name", "boolean", "ref"].includes(last.kind) || !IDENTIFIER.test(last.text)) {
    return null;
  }
  return CALL_POSITIONS.has(previous.kind) ? { prefix: last.text, start: last.start } : null;
}

/** The registered names starting with a prefix, case-insensitively,
 * alphabetically. */
export function functionsStartingWith(prefix: string): string[] {
  const upper = prefix.toUpperCase();
  return registeredFunctionNames().filter((name) => name.startsWith(upper));
}

/** The Insert Function dialog's list (see the module header); `category`
 * null is "All". */
export function searchFunctions(query: string, category: FunctionCategory | null): string[] {
  const inCategory = registeredFunctionNames().filter((name) => category === null || getFunction(name)?.category === category);
  const q = query.trim().toUpperCase();
  if (q.length === 0) {
    return inCategory;
  }
  const words = q.toLowerCase().split(/\s+/);
  const prefix = inCategory.filter((name) => name.startsWith(q));
  const substring = inCategory.filter((name) => !name.startsWith(q) && name.includes(q));
  const summary = inCategory.filter((name) => {
    if (name.includes(q)) {
      return false;
    }
    const text = (getFunction(name)?.summary ?? "").toLowerCase();
    return words.every((word) => text.includes(word));
  });
  return [...prefix, ...substring, ...summary];
}
