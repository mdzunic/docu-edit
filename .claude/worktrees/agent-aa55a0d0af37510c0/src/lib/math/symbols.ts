// SPEC-056 symbol resolution. The equation feature ships NO second symbol
// table: every `\alpha`-style name resolves against SPEC-026's
// CHARACTER_CATEGORIES catalogue (src/lib/core/specialCharacters.ts). This
// module only derives LaTeX *command names* — the Greek aliases are computed
// from the catalogue's Unicode names, and the operator/relation aliases are a
// name→code-point map whose glyphs still come from the catalogue.
import { CHARACTER_CATEGORIES } from "../core/specialCharacters";

/** Build a `latexName -> "glyph"` map for the Greek letters straight from the
 * catalogue, so there is no hand-maintained Greek table. "GREEK SMALL LETTER
 * ALPHA" -> `alpha`, "GREEK CAPITAL LETTER GAMMA" -> `Gamma`. */
function buildGreekAliases(): Map<string, string> {
  const out = new Map<string, string>();
  const variants: Record<number, string> = {
    0x3d1: "vartheta",
    0x3d5: "varphi",
    0x3d6: "varpi",
    0x3c2: "varsigma",
  };
  // Catalogue spells lambda "LAMDA"; normalise to the LaTeX spelling.
  const rename: Record<string, string> = { lamda: "lambda" };
  for (const cat of CHARACTER_CATEGORIES) {
    if (cat.id !== "greek") continue;
    for (const ch of cat.chars) {
      const glyph = String.fromCodePoint(ch.cp);
      const variant = variants[ch.cp];
      if (variant) {
        out.set(variant, glyph);
        continue;
      }
      const m = /^GREEK (SMALL|CAPITAL) LETTER (?:FINAL )?([A-Z]+)$/.exec(ch.name);
      if (!m) continue;
      let base = m[2].toLowerCase();
      base = rename[base] ?? base;
      const name = m[1] === "CAPITAL" ? base.charAt(0).toUpperCase() + base.slice(1) : base;
      // Keep the first spelling (skip the alternate final sigma for "sigma").
      if (!out.has(name)) out.set(name, glyph);
    }
  }
  return out;
}

/** Operator, relation and arrow command names → code point. These names are
 * the standard LaTeX commands; the glyph rendered is the catalogue's, so this
 * is an alias map over CHARACTER_CATEGORIES, not a parallel character store. */
const OPERATOR_ALIASES: Record<string, number> = {
  times: 0xd7,
  div: 0xf7,
  pm: 0xb1,
  mp: 0x2213,
  cdot: 0x2219,
  ast: 0x2217,
  circ: 0x2218,
  bullet: 0x2219,
  star: 0x2217,
  oplus: 0x2295,
  otimes: 0x2297,
  // relations
  leq: 0x2264,
  le: 0x2264,
  geq: 0x2265,
  ge: 0x2265,
  neq: 0x2260,
  ne: 0x2260,
  equiv: 0x2261,
  approx: 0x2248,
  cong: 0x2245,
  sim: 0x223c,
  propto: 0x221d,
  ll: 0x226a,
  gg: 0x226b,
  subset: 0x2282,
  supset: 0x2283,
  subseteq: 0x2286,
  supseteq: 0x2287,
  parallel: 0x2225,
  // set / logic
  in: 0x2208,
  notin: 0x2209,
  ni: 0x220b,
  cup: 0x222a,
  cap: 0x2229,
  wedge: 0x2227,
  land: 0x2227,
  vee: 0x2228,
  lor: 0x2228,
  forall: 0x2200,
  exists: 0x2203,
  nexists: 0x2204,
  emptyset: 0x2205,
  nabla: 0x2207,
  partial: 0x2202,
  infty: 0x221e,
  angle: 0x2220,
  therefore: 0x2234,
  because: 0x2235,
  perp: 0x22a5,
  // arrows (SPEC-026 Arrows category)
  to: 0x2192,
  rightarrow: 0x2192,
  leftarrow: 0x2190,
  leftrightarrow: 0x2194,
  Rightarrow: 0x21d2,
  Leftarrow: 0x21d0,
  Leftrightarrow: 0x21d4,
  uparrow: 0x2191,
  downarrow: 0x2193,
  mapsto: 0x21a6,
};

/** N-ary operator command names → glyph. Some code points are added to the
 * catalogue's Maths category by this spec (∮ ⋃ ⋂). */
const NARY_ALIASES: Record<string, number> = {
  sum: 0x2211,
  prod: 0x220f,
  coprod: 0x2210,
  int: 0x222b,
  iint: 0x222c,
  iiint: 0x222d,
  oint: 0x222e,
  bigcup: 0x22c3,
  bigcap: 0x22c2,
};

/** Accent command names → the accent glyph placed over the base. */
export const ACCENT_ALIASES: Record<string, string> = {
  hat: "̂",
  bar: "̄",
  vec: "⃗",
  dot: "̇",
  ddot: "̈",
  tilde: "̃",
  check: "̌",
  breve: "̆",
  acute: "́",
  grave: "̀",
};

/** The upright-set function names (SPEC-056 Construct set). */
export const FUNCTION_NAMES = new Set([
  "sin", "cos", "tan", "sec", "csc", "cot",
  "sinh", "cosh", "tanh",
  "arcsin", "arccos", "arctan",
  "exp", "ln", "log", "lim", "max", "min", "sup", "inf",
  "det", "gcd", "deg",
]);

const GREEK_ALIASES = buildGreekAliases();

/** Resolves a `\name` command to a single glyph string, or null when the name
 * is not a known symbol (Greek letter, operator, relation, arrow or n-ary
 * operator). Function names, accents and structural commands are handled by
 * the parser, not here. */
export function resolveSymbolName(name: string): string | null {
  const greek = GREEK_ALIASES.get(name);
  if (greek) return greek;
  const op = OPERATOR_ALIASES[name];
  if (op !== undefined) return String.fromCodePoint(op);
  const nary = NARY_ALIASES[name];
  if (nary !== undefined) return String.fromCodePoint(nary);
  return null;
}

/** Whether a `\name` resolves to an n-ary operator glyph. */
export function naryGlyphForName(name: string): string | null {
  const cp = NARY_ALIASES[name];
  return cp === undefined ? null : String.fromCodePoint(cp);
}

/** The set of n-ary operator glyphs (∑ ∏ ∫ ∬ ∮ ⋃ ⋂ …), used by the reader to
 * classify an OMML/MathML operator as an n-ary. */
export const NARY_GLYPHS = new Set(Object.values(NARY_ALIASES).map((cp) => String.fromCodePoint(cp)));

/** Reverse lookup: a symbol glyph → the shortest LaTeX command that produces
 * it, for regenerating the linear form from a tree read out of a file. */
const GLYPH_TO_NAME = buildGlyphToName();

function buildGlyphToName(): Map<string, string> {
  const out = new Map<string, string>();
  const add = (glyph: string, name: string) => {
    const existing = out.get(glyph);
    if (existing === undefined || name.length < existing.length) out.set(glyph, name);
  };
  for (const [name, glyph] of GREEK_ALIASES) add(glyph, name);
  for (const [name, cp] of Object.entries(OPERATOR_ALIASES)) add(String.fromCodePoint(cp), name);
  for (const [name, cp] of Object.entries(NARY_ALIASES)) add(String.fromCodePoint(cp), name);
  return out;
}

/** The LaTeX command for a glyph, or null when the glyph is an ordinary
 * printable character that needs no backslash. */
export function nameForGlyph(glyph: string): string | null {
  return GLYPH_TO_NAME.get(glyph) ?? null;
}

/** SPEC-056 Layout "Variant" rule, applied by the parser and reader: a single
 * Latin or Greek identifier letter is italic; a number, operator, delimiter or
 * function name is upright. */
export function variantForText(text: string): "italic" | "upright" {
  if ([...text].length !== 1) return "upright";
  const cp = text.codePointAt(0) as number;
  const isLatinLetter = (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a);
  const isGreekLetter = (cp >= 0x391 && cp <= 0x3a9) || (cp >= 0x3b1 && cp <= 0x3d6);
  return isLatinLetter || isGreekLetter ? "italic" : "upright";
}
