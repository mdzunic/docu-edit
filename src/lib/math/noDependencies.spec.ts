// SPEC-056 AC-26/AC-27: no dependency is added to package.json — the layout
// engine, the linear parser and both serialisers are written for this spec.
// This test pins the manifest's runtime and dev dependency sets so any future
// attempt to pull in KaTeX/MathJax/a math font package fails here.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const manifest = JSON.parse(
  readFileSync(join(dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "package.json"), "utf8"),
) as { dependencies?: Record<string, string>; devDependencies?: Record<string, string> };

describe("SPEC-056 AC-27 — package.json is unchanged", () => {
  it("keeps exactly the two runtime dependencies", () => {
    expect(Object.keys(manifest.dependencies ?? {}).sort()).toEqual(["vue", "vue-router"]);
  });

  it("adds no math-rendering library or font package", () => {
    const all = { ...manifest.dependencies, ...manifest.devDependencies };
    for (const name of Object.keys(all)) {
      expect(name).not.toMatch(/katex|mathjax|mathlive|temml|mathquill/i);
    }
  });

  it("keeps the pinned dev dependency set", () => {
    expect(Object.keys(manifest.devDependencies ?? {}).sort()).toEqual([
      "@playwright/test",
      "@types/node",
      "@vitejs/plugin-vue",
      "@vue/test-utils",
      "eslint",
      "eslint-plugin-vue",
      "jsdom",
      "oxlint",
      "typescript",
      "vite",
      "vitest",
      "vue-tsc",
    ]);
  });
});
