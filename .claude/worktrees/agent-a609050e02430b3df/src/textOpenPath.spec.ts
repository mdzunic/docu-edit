// The text product's open path never statically loads the presentation
// product's codec — the same architecture-as-test pattern as the boundary
// specs, for a property they cannot see: both products may be imported by the
// app layer (openDocument.ts, the library page), so what keeps one product's
// weight off the other's open is WHEN, not WHETHER, it is imported.
//
// SPEC-047's firstPageMs clock starts on the library page and stops at the
// editor's first page with content, so it times the library page's module
// graph, the editor's (HomeView preloads it on mount, while the file is
// parsed) and the parse itself. openDocument.ts once imported the pptx/odp
// readers and the Keynote importer statically, and the library reached both
// presentation writers through its download action and `phKey` — some 35
// modules of deck read/write code fetched and evaluated on every .docx open
// (SPEC-097's gate failed that budget). The codec now loads on demand, when a
// package is a deck. This walks the static import graph from both entry
// points — type-only imports are erased in the build and a dynamic `import()`
// is the on-demand load itself, so neither counts — and asserts no codec
// module is reachable.
import { existsSync, readFileSync, statSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const srcDir = dirname(fileURLToPath(import.meta.url));

/** The presentation codec: the readers, the writers, and the modules that
 * exist to drive them. None of it is needed to open, show or edit a text
 * document. */
const CODEC_MODULES = [
  "lib/presentation/open.ts",
  "lib/presentation/download.ts",
  "lib/presentation/pptx/reader.ts",
  "lib/presentation/pptx/writer.ts",
  "lib/presentation/odp/reader.ts",
  "lib/presentation/odp/writer.ts",
  "lib/presentation/keynote/reader.ts",
];

/** The two graphs inside SPEC-047's firstPageMs window. */
const ENTRY_POINTS = ["views/HomeView.vue", "views/DocumentView.vue"];

/** A module's runtime import specifiers: static `import`/`export … from`
 * declarations that survive type erasure. Dynamic `import()` is excluded
 * on purpose. A `.vue` file contributes its `<script>` blocks only. */
function runtimeSpecifiers(path: string): string[] {
  let source = readFileSync(path, "utf-8");
  if (path.endsWith(".vue")) {
    source = [...source.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map((match) => match[1]).join("\n");
  }
  source = source.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
  const specifiers: string[] = [];
  for (const match of source.matchAll(/(?:^|\n)\s*(import|export)\s+(type\s)?([^;]*?)\s*from\s*"([^"]+)"/g)) {
    const [, , typeOnly, clause, specifier] = match;
    const named = /^\{([\s\S]*)\}$/.exec(clause.trim());
    const allTypes =
      named !== null &&
      named[1]
        .split(",")
        .map((name) => name.trim())
        .filter(Boolean)
        .every((name) => name.startsWith("type "));
    if (!typeOnly && !allTypes) {
      specifiers.push(specifier);
    }
  }
  for (const match of source.matchAll(/(?:^|\n)\s*import\s*"([^"]+)"/g)) {
    specifiers.push(match[1]);
  }
  return specifiers;
}

function resolveModule(from: string, specifier: string): string | null {
  let base: string;
  if (specifier.startsWith("@/")) {
    base = join(srcDir, specifier.slice(2));
  } else if (specifier.startsWith(".")) {
    base = resolve(dirname(from), specifier);
  } else {
    return null; // a package
  }
  for (const candidate of [base, `${base}.ts`, `${base}.vue`, join(base, "index.ts")]) {
    if (existsSync(candidate) && statSync(candidate).isFile()) {
      return candidate;
    }
  }
  return null;
}

/** Every src module reachable from `entry` through runtime static imports,
 * src-relative, each mapped to the chain that first reached it. */
function staticGraph(entry: string): Map<string, string[]> {
  const chains = new Map<string, string[]>();
  const queue: [string, string[]][] = [[join(srcDir, entry), []]];
  while (queue.length > 0) {
    const [path, via] = queue.shift()!;
    const id = relative(srcDir, path);
    if (chains.has(id)) {
      continue;
    }
    const chain = [...via, id];
    chains.set(id, chain);
    for (const specifier of runtimeSpecifiers(path)) {
      const next = resolveModule(path, specifier);
      if (next !== null) {
        queue.push([next, chain]);
      }
    }
  }
  return chains;
}

describe("text open path", () => {
  for (const entry of ENTRY_POINTS) {
    it(`${entry} reaches no presentation codec module through a static import`, () => {
      const graph = staticGraph(entry);
      // The walk itself works: the entry reaches the text open funnel.
      expect(graph.has("lib/openDocument.ts")).toBe(true);
      const reached = CODEC_MODULES.filter((module) => graph.has(module)).map((module) =>
        graph.get(module)!.join(" -> "),
      );
      expect(reached, "load these on demand (dynamic import) instead").toEqual([]);
    });
  }

  it("still finds a codec module the walk is pointed at directly", () => {
    // Guards the guard: a walk that silently stopped following imports would
    // pass the two tests above vacuously.
    const graph = staticGraph("lib/presentation/open.ts");
    expect(graph.has("lib/presentation/pptx/reader.ts")).toBe(true);
    expect(graph.has("lib/presentation/odp/writer.ts")).toBe(true);
  });
});
