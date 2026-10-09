// SPEC-119 "Layering", asserted by scanning source — alongside
// spreadsheet/boundary.spec.ts, which holds the whole spreadsheet layer to
// spreadsheet siblings and src/lib/core. Every runtime module under
// formula/: imports only spreadsheet siblings and src/lib/core (no Vue, no
// package, no other product), touches no `window` or `document`, and builds
// no code from text — no `eval`, `new Function` or `Function(`. Formula text
// is data: SPEC-120's evaluator over these modules interprets an AST, and
// SPEC-120 extends this scan to the files it adds. And no rewrite is
// print(parse()) — which would discard the user's whitespace: printFormula is
// a test oracle and a builder (SPEC-125, SPEC-140), so no runtime module of
// the spreadsheet layer calls it yet.
import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const formulaDir = dirname(fileURLToPath(import.meta.url));
const spreadsheetDir = resolve(formulaDir, "..");
const coreDir = resolve(spreadsheetDir, "..", "core");

/** Every runtime .ts module under `dir`, recursively (specs excluded). */
function runtimeModules(dir: string): string[] {
  const files: string[] = [];
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) {
      files.push(...runtimeModules(path));
    } else if (name.endsWith(".ts") && !name.endsWith(".spec.ts")) {
      files.push(path);
    }
  }
  return files;
}

/** The source with comments and string literals blanked, so prose and data
 * cannot trip the code checks (their offsets kept). */
function codeOnly(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\/|\/\/[^\n]*|"(?:\\.|[^"\\\n])*"|'(?:\\.|[^'\\\n])*'|`(?:\\.|[^`\\])*`/g, (match) =>
    match.replace(/[^\n]/g, " "),
  );
}

/** Every import/export specifier (core/boundary.spec.ts's patterns). */
function importSpecifiers(source: string): string[] {
  const specifiers: string[] = [];
  for (const pattern of [
    /(?:^|\n)\s*import\s[^;]*?from\s*"([^"]+)"/g,
    /(?:^|\n)\s*import\s*"([^"]+)"/g,
    /(?:^|\n)\s*export\s[^;]*?from\s*"([^"]+)"/g,
    /import\(\s*"([^"]+)"\s*\)/g,
  ]) {
    for (const match of source.matchAll(pattern)) {
      specifiers.push(match[1]!);
    }
  }
  return specifiers;
}

const FORBIDDEN: ReadonlyArray<[RegExp, string]> = [
  [/\bwindow\b/, "window"],
  [/\bdocument\b/, "document"],
  [/\bglobalThis\b/, "globalThis"],
  [/\beval\s*\(/, "eval("],
  [/\bnew\s+Function\b/, "new Function"],
  [/\bFunction\s*\(/, "Function("],
];

describe("the formula layer's purity", () => {
  const modules = runtimeModules(formulaDir);

  it("has the formula modules to check", () => {
    expect(modules.map((path) => path.slice(formulaDir.length + 1)).toSorted()).toEqual(
      expect.arrayContaining(["anchors.ts", "ast.ts", "dialect.ts", "functionNames.ts", "parse.ts", "print.ts", "references.ts", "tokens.ts"]),
    );
  });

  it("imports only spreadsheet siblings and src/lib/core", () => {
    const violations: string[] = [];
    for (const path of modules) {
      for (const specifier of importSpecifiers(readFileSync(path, "utf8"))) {
        const resolved = resolve(dirname(path), specifier);
        const allowed =
          specifier.startsWith(".") && ((resolved + sep).startsWith(spreadsheetDir + sep) || (resolved + sep).startsWith(coreDir + sep));
        if (!allowed) {
          violations.push(`${path} imports "${specifier}"`);
        }
      }
    }
    expect(violations).toEqual([]);
  });

  it("touches no window or document and builds no code from text (no eval, new Function, Function()", () => {
    const violations: string[] = [];
    for (const path of modules) {
      const code = codeOnly(readFileSync(path, "utf8"));
      for (const [pattern, label] of FORBIDDEN) {
        if (pattern.test(code)) {
          violations.push(`${path}: ${label}`);
        }
      }
    }
    expect(violations).toEqual([]);
  });

  it("implements no transformation as print(parse()): no runtime spreadsheet module but print.ts mentions printFormula", () => {
    const callers = runtimeModules(spreadsheetDir)
      .filter((path) => path !== join(formulaDir, "print.ts"))
      .filter((path) => /\bprintFormula\b/.test(codeOnly(readFileSync(path, "utf8"))))
      .map((path) => path.slice(spreadsheetDir.length + 1));
    expect(callers).toEqual([]);
  });

  it("catches what it scans for (the scan itself is not vacuous)", () => {
    expect(codeOnly(`const a = window.x; // window\nconst b = "document";`)).toMatch(/\bwindow\b/);
    expect(codeOnly(`const b = "document"; /* eval( */`)).not.toMatch(/\bdocument\b|\beval\s*\(/);
    expect(FORBIDDEN.some(([pattern]) => pattern.test(codeOnly("const f = new Function('return 1');")))).toBe(true);
    expect(FORBIDDEN.some(([pattern]) => pattern.test(codeOnly("eval (text)")))).toBe(true);
  });
});
