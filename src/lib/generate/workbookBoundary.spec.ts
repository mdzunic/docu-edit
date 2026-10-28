// SPEC-140 — the workbook product's boundary, the architecture-as-test
// pattern of core/boundary.spec.ts and spreadsheet/boundary.spec.ts:
// workbook.ts, workbookDescription.ts, workbookStyles.ts, readWorkbook.ts
// (and every other workbook*.ts module beside them) import only
// src/lib/spreadsheet/, src/lib/core/ and generate's product-neutral modules
// — errors.ts, validate.ts, nodeDom.ts — and each other: never the text
// builder's runs/blocks/builder/description/read, the deck modules,
// documentModel.ts or openDocument.ts. The reverse direction — nothing under
// src/lib/spreadsheet/ imports generate/ — is SPEC-116's
// spreadsheet/boundary.spec.ts, unchanged; it is restated here for generate/.
import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join, normalize, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const generateDir = dirname(fileURLToPath(import.meta.url));
const libDir = resolve(generateDir, "..");

/** Every import/export specifier in a module's source (the patterns
 * core/boundary.spec.ts uses). */
function importSpecifiers(source: string): string[] {
  const specifiers: string[] = [];
  const patterns = [
    /(?:^|\n)\s*import\s[^;]*?from\s*"([^"]+)"/g,
    /(?:^|\n)\s*import\s*"([^"]+)"/g,
    /(?:^|\n)\s*export\s[^;]*?from\s*"([^"]+)"/g,
    /import\(\s*"([^"]+)"\s*\)/g,
  ];
  for (const pattern of patterns) {
    for (const match of source.matchAll(pattern)) {
      specifiers.push(match[1]!);
    }
  }
  return specifiers;
}

/** The workbook product's runtime modules in generate/. */
const workbookModules = readdirSync(generateDir).filter((name) => /^(workbook\w*|readWorkbook)\.ts$/.test(name) && !name.endsWith(".spec.ts"));

const PRODUCT_NEUTRAL = new Set(["./errors", "./validate", "./nodeDom"]);
const FORBIDDEN = [
  "./runs",
  "./blocks",
  "./builder",
  "./description",
  "./read",
  "./length",
  "./presentation",
  "./deckSlides",
  "./deckDescription",
  "./deckDrops",
  "./deckHtml",
  "./readDeck",
  "./slideMarkdown",
  "./slideRuns",
  "../documentModel",
  "../openDocument",
];

describe("the workbook modules' imports (AC-40)", () => {
  it("covers the four modules the spec names", () => {
    expect(workbookModules).toEqual(expect.arrayContaining(["workbook.ts", "workbookDescription.ts", "workbookStyles.ts", "readWorkbook.ts"]));
  });

  it("import only lib/spreadsheet, lib/core, errors/validate/nodeDom and each other", () => {
    const siblings = new Set(workbookModules.map((name) => `./${name.replace(/\.ts$/, "")}`));
    const violations: string[] = [];
    for (const name of workbookModules) {
      for (const specifier of importSpecifiers(readFileSync(join(generateDir, name), "utf8"))) {
        const ok =
          specifier.startsWith("../spreadsheet/") || specifier.startsWith("../core/") || PRODUCT_NEUTRAL.has(specifier) || siblings.has(specifier);
        if (!ok) {
          violations.push(`${name} imports "${specifier}"`);
        }
      }
    }
    expect(violations).toEqual([]);
  });

  it("never import the text builder, the deck builder, documentModel.ts or openDocument.ts", () => {
    for (const name of workbookModules) {
      const specifiers = importSpecifiers(readFileSync(join(generateDir, name), "utf8"));
      for (const forbidden of FORBIDDEN) {
        expect(specifiers, `${name} must not import ${forbidden}`).not.toContain(forbidden);
      }
    }
  });

  it("index.ts re-exports the workbook surface beside both builders (AC-42)", async () => {
    const index = await import("./index");
    expect(Object.keys(index)).toEqual(
      expect.arrayContaining(["DocumentBuilder", "PresentationBuilder", "WorkbookBuilder", "workbookFromDescription", "readWorkbook", "DocumentGenerationError"]),
    );
  });
});

/** Every runtime .ts module under `dir`, recursively. */
function runtimeModules(dir: string): string[] {
  return readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    return statSync(path).isDirectory() ? runtimeModules(path) : name.endsWith(".ts") && !name.endsWith(".spec.ts") ? [path] : [];
  });
}

describe("nothing under src/lib/spreadsheet imports generate/ (AC-41)", () => {
  it("holds for every spreadsheet module", () => {
    const spreadsheetDir = join(libDir, "spreadsheet");
    const violations: string[] = [];
    for (const path of runtimeModules(spreadsheetDir)) {
      for (const specifier of importSpecifiers(readFileSync(path, "utf8"))) {
        const target = normalize(resolve(dirname(path), specifier));
        if (specifier.startsWith(".") && relative(libDir, target).split(/[\\/]/)[0] === "generate") {
          violations.push(`${relative(libDir, path)} imports "${specifier}"`);
        }
      }
    }
    expect(violations).toEqual([]);
  });
});
