// The core layer's one structural invariant, enforced as a test so it
// cannot silently erode as specs keep landing: nothing in src/lib/core may
// import from outside src/lib/core. No document model, no commands, no
// docx/odf readers, no Vue — a module that needs any of those is not core
// and belongs a layer up. This is what keeps the layer reusable by another
// product (e.g. Slides) without dragging the word processor in.
import { readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const coreDir = dirname(fileURLToPath(import.meta.url));

/** Every runtime module in core/ (specs excluded — they may import vitest
 * and node builtins, like this one). */
function coreModules(): string[] {
  return readdirSync(coreDir).filter(
    (name) => name.endsWith(".ts") && !name.endsWith(".spec.ts"),
  );
}

/** Every import/export specifier in a module's source: static imports,
 * re-exports and dynamic import() calls alike. */
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

describe("core layer boundary", () => {
  it("has modules to check", () => {
    expect(coreModules().length).toBeGreaterThan(0);
  });

  it("no core module imports from outside src/lib/core", () => {
    const violations: string[] = [];
    for (const name of coreModules()) {
      const source = readFileSync(join(coreDir, name), "utf8");
      for (const specifier of importSpecifiers(source)) {
        // Allowed: sibling core modules ("./x"). Everything else — parent
        // paths ("../documentModel"), aliases ("@/…"), bare packages
        // ("vue") — couples core to a product or a framework.
        if (!specifier.startsWith("./") || specifier.startsWith("./..")) {
          violations.push(`${name} imports "${specifier}"`);
        }
      }
    }
    expect(violations).toEqual([]);
  });
});
