// SPEC-115's layer rule for the deck designer, the same architecture-as-test
// pattern as core/boundary.spec.ts and presentation/boundary.spec.ts: a
// src/lib/design module imports only its siblings, lib/core, lib/presentation
// and lib/generate's deck modules — never documentModel.ts, outline.ts, the
// .docx/.odt/.md layers, lib/ai (the model gateway), the store or Vue. The two
// products still meet only in the app layer, and nothing here can reach the
// network: the app layer asks the model and hands this layer plain data.
//
// Also SPEC-115's parsing rule: a reply is read with JSON.parse inside a try
// and nothing else — no eval, no Function, no dynamic import, no fetch.
import { readdirSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const designDir = dirname(fileURLToPath(import.meta.url));
const libDir = resolve(designDir, "..");

/** lib/generate's deck modules — the description, its slide primitives, the
 * run mapping, the drop report, and the shared validation/error modules. */
const GENERATE_DECK_MODULES = ["deckDescription", "deckDrops", "deckSlides", "slideRuns", "errors", "validate"];

function runtimeModules(): string[] {
  return readdirSync(designDir).filter((name) => name.endsWith(".ts") && !name.endsWith(".spec.ts"));
}

function importSpecifiers(source: string): string[] {
  const specifiers: string[] = [];
  const patterns = [
    /(?:^|\n)\s*import\s[^;]*?from\s*"([^"]+)"/g,
    /(?:^|\n)\s*import\s*"([^"]+)"/g,
    /(?:^|\n)\s*export\s[^;]*?from\s*"([^"]+)"/g,
    /import\(\s*["'`]([^"'`]+)["'`]\s*\)/g,
  ];
  for (const pattern of patterns) {
    for (const match of source.matchAll(pattern)) {
      specifiers.push(match[1]);
    }
  }
  return specifiers;
}

/** Whether `specifier` (from a lib/design module) is an allowed import. */
function allowed(specifier: string): boolean {
  if (!specifier.startsWith(".")) {
    return false; // no packages (Vue included)
  }
  const target = resolve(designDir, specifier).slice(libDir.length + 1).replace(/\\/g, "/");
  if (target.startsWith("design/") || target.startsWith("core/") || target.startsWith("presentation/")) {
    return true;
  }
  return GENERATE_DECK_MODULES.some((name) => target === `generate/${name}`);
}

describe("lib/design boundary (SPEC-115)", () => {
  it("has modules to check", () => {
    expect(runtimeModules()).toEqual(
      expect.arrayContaining(["plan.ts", "planToDescription.ts", "applyDesign.ts", "outlineDesigner.ts", "prompt.ts", "build.ts"]),
    );
  });

  it("imports only lib/core, lib/presentation, lib/generate's deck modules and its siblings", () => {
    const violations = runtimeModules().flatMap((name) =>
      importSpecifiers(readFileSync(join(designDir, name), "utf8"))
        .filter((specifier) => !allowed(specifier))
        .map((specifier) => `${name} imports ${specifier}`),
    );
    expect(violations).toEqual([]);
  });

  it("the rule catches the imports it forbids", () => {
    for (const forbidden of ["../documentModel", "../outline", "../docx/reader", "../odf/writer", "../md/reader", "../ai/modelGateway", "../documentStore", "vue", "../generate/blocks"]) {
      expect(allowed(forbidden), forbidden).toBe(false);
    }
  });

  it("evaluates nothing: no eval, Function, dynamic import or fetch anywhere in lib/design", () => {
    for (const name of runtimeModules()) {
      const code = readFileSync(join(designDir, name), "utf8")
        .replace(/\/\*[\s\S]*?\*\//g, "")
        .replace(/\/\/.*$/gm, "");
      expect(code, name).not.toMatch(/\beval\s*\(|\bnew\s+Function\b|\bFunction\s*\(|\bimport\s*\(|\bfetch\s*\(/);
    }
  });
});
