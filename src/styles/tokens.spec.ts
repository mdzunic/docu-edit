import { describe, it, expect } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { join, relative } from "node:path";

// SPEC-058's lint-style guard. It reads every .vue file's <style> blocks and
// the token file itself and enforces, mechanically, the rules the spec makes
// normative: no colour literal escapes tokens.css (AC-3/4/5), every
// spacing/radius/border/shadow/type/motion value resolves to a token
// (AC-6..13), the token palette meets WCAG AA contrast computed from the file
// (AC-21/22/23), and focus is never silently removed (AC-45/46/47).

const SRC = join(__dirname, "..");

// ---- What is not a token (the three named exemptions) ---------------------
// 1+2. Document content & the canvas painter: these components render the
//      user's *document* (paper white, model colours, shape/chart fills) or
//      paint geometry that mirrors the model, not app chrome. Their styles are
//      exempt, exactly as src/lib/** is out of scope entirely.
// 3.   Geometry in px that mirrors the model (ruler ticks, table handles,
//      draw previews) stays computed, not tokenised.
const EXEMPT_FILES = new Set([
  "components/DocPage.vue",
  "components/DocPageBlock.vue",
  "components/DocBlockEntries.vue",
  "components/DocTableBlock.vue",
  "components/DocImageBlock.vue",
  "components/DocShapeBlock.vue",
  "components/DocChartBlock.vue",
  "components/DocTocBlock.vue",
  "components/DocTofBlock.vue",
  "components/DocHeadingToggle.vue",
  "components/EquationView.vue",
  "components/EquationEditor.vue",
  "components/MathBox.vue",
  "components/DocumentRuler.vue",
  "components/TableHandlesOverlay.vue",
  "components/ShapeDrawPreview.vue",
  "components/LinkChip.vue",
]);

// A rule whose selector targets the document-content / pages subtree is exempt
// even inside an otherwise-chrome component (e.g. DocumentView hosts both).
// The overlay-highlight selectors (find/comment/spell) annotate the document
// itself — a matched search term, a commented range, a misspelling — with the
// same non-chrome semantic colours the canvas painter uses (search yellow is
// not a UI token), so they belong to the document-content exemption, not the
// design system.
const EXEMPT_SELECTOR =
  /\.document-content|\.pages\b|\.page\b|\.page-|\.doc-|\.ruler|-handle|-marker|::selection|\.paste-preview|find-highlight|comment-highlight|spell-underline/;

function allVueFiles(dir: string = SRC, acc: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      allVueFiles(full, acc);
    } else if (entry.name.endsWith(".vue")) {
      acc.push(full);
    }
  }
  return acc;
}

function styleBlocks(source = "") {
  const blocks = [];
  const re = /<style[^>]*>([\s\S]*?)<\/style>/g;
  let m;
  while ((m = re.exec(source)) !== null) {
    // Offset in lines so violation messages point at the real file line.
    const before = source.slice(0, m.index);
    const startLine = before.split("\n").length;
    blocks.push({ css: m[1], startLine });
  }
  return blocks;
}

function stripComments(css = "") {
  return css.replace(/\/\*[\s\S]*?\*\//g, (c) => c.replace(/[^\n]/g, " "));
}

// Walk declarations, tracking the nearest selector so exempt rules are skipped.
function declarations(css = "", startLine = 1) {
  const clean = stripComments(css);
  const out = [];
  const selectorStack = [];
  let buf = "";
  let line = startLine;
  for (let i = 0; i < clean.length; i++) {
    const ch = clean[i];
    if (ch === "\n") {
      line++;
    }
    if (ch === "{") {
      selectorStack.push(buf.trim().split("\n").pop()?.trim() ?? "");
      buf = "";
    } else if (ch === "}") {
      selectorStack.pop();
      buf = "";
    } else if (ch === ";") {
      const decl = buf.trim();
      buf = "";
      const idx = decl.indexOf(":");
      if (idx > 0) {
        out.push({
          prop: decl.slice(0, idx).trim().toLowerCase(),
          value: decl.slice(idx + 1).trim(),
          selector: selectorStack.join(" "),
          line,
        });
      }
    } else {
      buf += ch;
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Colour literals: hex, rgb()/rgba(), hsl()/hsla(), and a named-colour set.
const HEX = /#[0-9a-fA-F]{3,8}\b/;
const FUNC_COLOUR = /\b(rgb|rgba|hsl|hsla|hwb|lab|lch|color)\s*\(/;
const NAMED = new Set([
  "white", "black", "red", "green", "blue", "yellow", "orange", "purple",
  "gray", "grey", "silver", "maroon", "navy", "teal", "aqua", "fuchsia",
  "lime", "olive", "pink", "brown", "gold", "cyan", "magenta", "violet",
  "indigo", "coral", "salmon", "khaki", "crimson", "tomato", "orchid",
  "lightgray", "lightgrey", "darkgray", "darkgrey", "whitesmoke",
  "gainsboro", "dimgray", "dimgrey", "slategray", "slategrey", "lightblue",
  "steelblue", "royalblue", "dodgerblue", "cornflowerblue", "darkblue",
  "lightgreen", "darkgreen", "forestgreen", "seagreen", "goldenrod",
]);
// These are keywords, not colours — never flagged.
const COLOUR_KEYWORDS = new Set([
  "transparent", "currentcolor", "none", "inherit", "initial", "unset",
  "revert", "auto",
]);

function colourLiteral(value = "") {
  if (HEX.test(value) || FUNC_COLOUR.test(value)) {
    return true;
  }
  for (const word of value.toLowerCase().match(/[a-z]+/g) ?? []) {
    if (NAMED.has(word) && !COLOUR_KEYWORDS.has(word)) {
      return true;
    }
  }
  return false;
}

// ---------------------------------------------------------------------------
// Tokenisation: after removing var() refs and safe keywords, no raw literal of
// the flagged kind may remain for the property.
const LENGTH_PROPS = new Set([
  "padding", "padding-top", "padding-right", "padding-bottom", "padding-left",
  "padding-block", "padding-inline", "margin", "margin-top", "margin-right",
  "margin-bottom", "margin-left", "margin-block", "margin-inline",
  "gap", "row-gap", "column-gap", "border-radius", "border-top-left-radius",
  "border-top-right-radius", "border-bottom-left-radius",
  "border-bottom-right-radius", "font-size",
]);
const WIDTH_PROPS = new Set([
  "border-width", "border-top-width", "border-right-width",
  "border-bottom-width", "border-left-width", "outline-width",
  "border", "border-top", "border-right", "border-bottom", "border-left",
  "outline",
]);

function withoutVarsAndKeywords(value = "") {
  return value
    .replace(/var\([^)]*\)/g, " ")
    .replace(/calc\([^)]*\)/g, " ") // calc math is treated as opaque
    .replace(/\b0(?![.\d])/g, " ") // bare zero
    .replace(/\b\d+%/g, " ") // percentages
    .replace(
      /\b(auto|none|inherit|initial|unset|revert|solid|dashed|dotted|double|hidden|normal|transparent|currentcolor|thin|medium|thick|border-box|content-box)\b/gi,
      " ",
    );
}

const LENGTH_LITERAL = /\b\d*\.?\d+(px|rem|em|ex|ch|vh|vw|vmin|vmax|pt|cm|mm)\b/;
const TIME_LITERAL = /\b\d*\.?\d+m?s\b/;

describe("tokens.css is the sole home of colour literals (AC-3/4/5)", () => {
  const files = allVueFiles();
  for (const file of files) {
    const rel = relative(SRC, file).replace(/\\/g, "/");
    if (EXEMPT_FILES.has(rel)) {
      continue;
    }
    it(`${rel} names no colour literal in a <style> block`, () => {
      const source = readFileSync(file, "utf8");
      const bad: string[] = [];
      for (const block of styleBlocks(source)) {
        for (const d of declarations(block.css, block.startLine)) {
          if (EXEMPT_SELECTOR.test(d.selector)) {
            continue;
          }
          if (colourLiteral(d.value)) {
            bad.push(`  line ${d.line}: ${d.prop}: ${d.value}`);
          }
        }
      }
      expect(bad, `colour literal(s) in ${rel}:\n${bad.join("\n")}`).toEqual([]);
    });
  }
});

describe("spacing/radius/border/shadow/type/motion resolve to tokens (AC-6..13)", () => {
  const files = allVueFiles();
  for (const file of files) {
    const rel = relative(SRC, file).replace(/\\/g, "/");
    if (EXEMPT_FILES.has(rel)) {
      continue;
    }
    it(`${rel} tokenises every flagged declaration`, () => {
      const source = readFileSync(file, "utf8");
      const bad: string[] = [];
      for (const block of styleBlocks(source)) {
        for (const d of declarations(block.css, block.startLine)) {
          if (EXEMPT_SELECTOR.test(d.selector)) {
            continue;
          }
          const residue = withoutVarsAndKeywords(d.value);
          const flag = (why: string) =>
            bad.push(`  line ${d.line}: ${d.prop}: ${d.value}  (${why})`);
          if ((LENGTH_PROPS.has(d.prop) || WIDTH_PROPS.has(d.prop)) && LENGTH_LITERAL.test(residue)) {
            flag("length not a token");
          } else if (d.prop === "box-shadow" && d.value !== "none" && !/var\(--shadow/.test(d.value)) {
            flag("shadow not a token");
          } else if (d.prop === "font-family" && !/var\(--font/.test(d.value) && d.value !== "inherit") {
            flag("font-family not a token");
          } else if (d.prop === "font-weight" && /\b[1-9]00\b/.test(residue)) {
            flag("font-weight not a token");
          } else if (d.prop === "line-height" && /\b\d*\.\d+\b/.test(residue)) {
            flag("line-height not a token");
          } else if (/^transition/.test(d.prop) && TIME_LITERAL.test(residue)) {
            flag("transition duration not a token");
          }
        }
      }
      expect(bad, `non-token value(s) in ${rel}:\n${bad.join("\n")}`).toEqual([]);
    });
  }
});

// ---------------------------------------------------------------------------
// Focus visibility (AC-45/46/47). Every :focus / :focus-visible rule that
// removes the outline must re-establish a --focus-ring outline; a rule that
// only sets `outline: none` with no replacement in the same block fails.
describe("focus is never silently removed (AC-45/46/47)", () => {
  const files = allVueFiles();
  for (const file of files) {
    const rel = relative(SRC, file).replace(/\\/g, "/");
    if (EXEMPT_FILES.has(rel)) {
      continue;
    }
    it(`${rel} keeps a visible focus ring on focus rules`, () => {
      const source = readFileSync(file, "utf8");
      const bad: string[] = [];
      for (const block of styleBlocks(source)) {
        const css = stripComments(block.css);
        // Match each rule block and inspect focus selectors specifically.
        const ruleRe = /([^{}]+)\{([^{}]*)\}/g;
        let m;
        let lineBase = block.startLine;
        while ((m = ruleRe.exec(css)) !== null) {
          const selector = m[1];
          const body = m[2];
          const line = lineBase + css.slice(0, m.index).split("\n").length - 1;
          if (/:focus(-visible)?\b/.test(selector)) {
            const removesOutline = /outline\s*:\s*(none|0)\b/.test(body);
            const hasRing = /outline\s*:\s*[^;]*var\(--focus-ring\)/.test(body);
            if (removesOutline && !hasRing) {
              bad.push(`  line ${line}: "${selector.trim()}" removes outline with no --focus-ring replacement`);
            }
          }
        }
      }
      expect(bad, `focus removed without replacement in ${rel}:\n${bad.join("\n")}`).toEqual([]);
    });
  }
});

// ---------------------------------------------------------------------------
// WCAG AA contrast, computed from tokens.css itself (AC-21/22/23).
function tokenMap(): Record<string, string> {
  const css = readFileSync(join(SRC, "styles/tokens.css"), "utf8");
  const map: Record<string, string> = {};
  const re = /(--[\w-]+):\s*([^;]+);/g;
  let m;
  while ((m = re.exec(css)) !== null) {
    map[m[1]] = m[2].trim();
  }
  return map;
}

function toRgb(value = ""): number[] {
  const hex = value.match(/#([0-9a-fA-F]{6})\b/);
  if (hex) {
    const n = parseInt(hex[1], 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  const short = value.match(/#([0-9a-fA-F]{3})\b/);
  if (short) {
    return short[1].split("").map((c) => parseInt(c + c, 16));
  }
  const rgb = value.match(/rgb[a]?\(\s*(\d+)\s+(\d+)\s+(\d+)/);
  if (rgb) {
    return [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])];
  }
  throw new Error(`cannot parse colour: ${value}`);
}

function luminance(rgb: number[]): number {
  const lin = rgb.map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2];
}

function contrast(a = "", b = ""): number {
  const la = luminance(toRgb(a));
  const lb = luminance(toRgb(b));
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

describe("token palette meets WCAG AA contrast (AC-21/22/23)", () => {
  const t = tokenMap();
  const ratio = (fg: string, bg: string) => contrast(t[fg], t[bg]);

  // Body text — 4.5:1.
  const body = [
    ["--text-primary", "--surface-base"],
    ["--text-primary", "--surface-raised"],
    ["--text-primary", "--surface-canvas"],
    ["--text-primary", "--surface-sunken"],
    ["--text-primary", "--surface-selected"],
    ["--text-secondary", "--surface-base"],
    ["--text-secondary", "--surface-raised"],
    ["--text-secondary", "--surface-canvas"],
    ["--text-secondary", "--surface-sunken"],
    ["--text-on-brand", "--surface-brand"],
    ["--text-on-brand", "--accent"],
    ["--text-danger", "--surface-base"],
  ];
  for (const [fg, bg] of body) {
    it(`${fg} on ${bg} ≥ 4.5:1`, () => {
      expect(ratio(fg, bg)).toBeGreaterThanOrEqual(4.5);
    });
  }

  // Large text / UI-component boundaries — 3:1. --border-subtle and
  // --border-default are decorative separators inside a surface (WCAG 1.4.11
  // exempts non-essential boundaries); --border-strong is the emphasised
  // outline that must stay distinguishable, so it is the one asserted.
  const large = [
    ["--text-accent", "--surface-base"],
    ["--text-accent", "--surface-raised"],
    ["--text-accent", "--surface-selected"],
    ["--accent", "--surface-base"],
    ["--border-strong", "--surface-base"],
    ["--success", "--surface-base"],
    ["--warning", "--surface-base"],
  ];
  for (const [fg, bg] of large) {
    it(`${fg} on ${bg} ≥ 3:1`, () => {
      expect(ratio(fg, bg)).toBeGreaterThanOrEqual(3);
    });
  }
});
