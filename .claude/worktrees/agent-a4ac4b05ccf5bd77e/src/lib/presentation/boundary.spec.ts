// The presentation layer's structural invariant, the same
// architecture-as-test pattern as core/boundary.spec.ts: the two products
// must never entangle. A presentation module may import sibling
// presentation modules and src/lib/core — never documentModel.ts, the
// command/query modules, the .docx/.odt format layers, the store, or Vue.
// (The reverse directions are enforced in core/boundary.spec.ts's
// format-layer block; app-layer modules like openDocument.ts and
// documentStore.ts legitimately import both products — they are the
// composition points, and are not presentation modules.)
import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const presentationDir = dirname(fileURLToPath(import.meta.url));
const coreDir = resolve(presentationDir, "..", "core");

/** Every runtime .ts module under `dir`, recursively (specs excluded — they
 * may import vitest, fixtures of either product, node builtins). */
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

/** Every import/export specifier in a module's source (same patterns as
 * core/boundary.spec.ts). */
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
      specifiers.push(match[1]);
    }
  }
  return specifiers;
}

describe("presentation layer boundary", () => {
  it("has modules to check", () => {
    expect(runtimeModules(presentationDir).length).toBeGreaterThan(0);
  });

  it("presentation modules import only presentation siblings and core", () => {
    const violations: string[] = [];
    for (const filePath of runtimeModules(presentationDir)) {
      const source = readFileSync(filePath, "utf8");
      for (const specifier of importSpecifiers(source)) {
        // Bare packages ("vue") and aliases couple the layer to a framework
        // or product; only relative specifiers can be legal.
        if (!specifier.startsWith(".")) {
          violations.push(`${filePath} imports "${specifier}"`);
          continue;
        }
        const resolved = resolve(dirname(filePath), specifier);
        const inPresentation = (resolved + sep).startsWith(presentationDir + sep);
        const inCore = (resolved + sep).startsWith(coreDir + sep);
        if (!inPresentation && !inCore) {
          violations.push(`${filePath} imports "${specifier}"`);
        }
      }
    }
    expect(violations).toEqual([]);
  });
});
