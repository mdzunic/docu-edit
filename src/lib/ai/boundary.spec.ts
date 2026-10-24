// SPEC-113's network boundary, enforced as a test so it survives a year of
// feature work: no runtime module under src/ may reference fetch,
// EventSource, WebSocket or XMLHttpRequest except the five below, each with
// the reason it may. The model gateway is the only one whose reason is a
// model; every other entry reaches only the app's own files or a URL the
// user's own input named, and none of them may grow into a model request.
// "Can this app send my words somewhere?" is answered by reading one file.
//
// The same architecture-as-test pattern as core/boundary.spec.ts. The list
// may only shrink: when a site goes away, delete its entry and lower the
// count with it. Never add one — a new network need is either the gateway's
// job or a spec-level decision that changes this file on purpose.
//
// Code, not text, is scanned: each module is parsed with the TypeScript
// compiler (a .vue file's <script> blocks), so a comment or a string that
// mentions "fetch" is not a reference, and an identifier or property access
// that names one is — including `window.fetch` and `globalThis["fetch"]`.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { dirname, join, relative } from "node:path";
import { fileURLToPath } from "node:url";
import ts from "typescript";
import { describe, expect, it } from "vitest";

const SRC = join(dirname(fileURLToPath(import.meta.url)), "..", "..");

/** The network sites, by path under src/, and why each may exist. */
const NETWORK_ALLOWLIST: Record<string, string> = {
  "lib/ai/modelGateway.ts":
    "SPEC-113: the one module permitted to make a model request — consented, origin-pinned, to an endpoint the user named.",
  "lib/openDocumentFromUrl.ts": "SPEC-009: opens a document from a URL the user typed (or a link put in /open?url=).",
  "lib/md/reader.ts": "SPEC-050: fetches a remote image that a Markdown document the user opened references.",
  "views/document/useFonts.ts":
    "The app's own bundled font files under import.meta.env.BASE_URL (same-origin), loaded to lay out and embed text.",
  // The PDF exporters' font loading (usePdfExport.ts, named by SPEC-113)
  // moved here when SPEC-104 shared it with the deck exporter: same files,
  // same reason.
  "lib/pdfAssets.ts":
    "The app's own bundled font files under import.meta.env.BASE_URL (same-origin), embedded into an exported PDF.",
};

const NETWORK_GLOBALS = new Set(["fetch", "EventSource", "WebSocket", "XMLHttpRequest"]);

/** Every runtime module under src/: .ts and .vue, specs and declarations
 * excluded. Paths relative to src/, "/"-separated. */
function runtimeModules(dir = SRC): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      return runtimeModules(full);
    }
    const isTs = entry.name.endsWith(".ts") && !entry.name.endsWith(".spec.ts") && !entry.name.endsWith(".d.ts");
    return isTs || entry.name.endsWith(".vue") ? [relative(SRC, full).replace(/\\/g, "/")] : [];
  });
}

/** A module's code as TypeScript: the file itself, or a .vue file's
 * <script> blocks. */
function codeOf(path: string): string {
  const source = readFileSync(join(SRC, path), "utf8");
  if (!path.endsWith(".vue")) {
    return source;
  }
  return [...source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)].map((match) => match[1]).join("\n");
}

const parsed = new Map<string, ts.SourceFile>();

/** A module parsed once, however many tests ask. */
function sourceOf(path: string): ts.SourceFile {
  let file = parsed.get(path);
  if (!file) {
    file = ts.createSourceFile(path, codeOf(path), ts.ScriptTarget.Latest, false, ts.ScriptKind.TS);
    parsed.set(path, file);
  }
  return file;
}

/** The names from `names` that parsed code references: identifiers
 * (`fetch(…)`, `new WebSocket(…)`, `window.fetch`) and string-keyed
 * element access (`globalThis["fetch"]`). Sorted, each once. */
function namesIn(file: ts.SourceFile, names: Set<string>): string[] {
  const found = new Set<string>();
  const visit = (node: ts.Node): void => {
    if (ts.isIdentifier(node) && names.has(node.text)) {
      found.add(node.text);
    } else if (
      ts.isElementAccessExpression(node) &&
      ts.isStringLiteralLike(node.argumentExpression) &&
      names.has(node.argumentExpression.text)
    ) {
      found.add(node.argumentExpression.text);
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  return [...found].sort();
}

function referencedNames(path: string, names: Set<string>): string[] {
  return namesIn(sourceOf(path), names);
}

const modules = runtimeModules();

describe("the network boundary (SPEC-113)", () => {
  it("scans the whole runtime tree", () => {
    expect(modules.length).toBeGreaterThan(500);
    expect(modules).toContain("lib/ai/modelGateway.ts");
    expect(modules.some((path) => path.endsWith(".vue"))).toBe(true);
  });

  it("has exactly five allowlisted network sites, each with its reason", () => {
    // Exactly five today, and only ever fewer: see the header.
    expect(Object.keys(NETWORK_ALLOWLIST)).toHaveLength(5);
    for (const [path, reason] of Object.entries(NETWORK_ALLOWLIST)) {
      expect(reason.length, `${path} needs its reason`).toBeGreaterThan(20);
    }
  });

  it("no runtime module outside the allowlist references fetch, EventSource, WebSocket or XMLHttpRequest", () => {
    const violations = modules
      .filter((path) => !(path in NETWORK_ALLOWLIST))
      .flatMap((path) => referencedNames(path, NETWORK_GLOBALS).map((name) => `${path} references ${name}`));
    expect(violations).toEqual([]);
  });

  it("every allowlisted module exists and still reaches the network — a dead entry is removed, not kept", () => {
    for (const path of Object.keys(NETWORK_ALLOWLIST)) {
      expect(existsSync(join(SRC, path)), `${path} no longer exists: delete its entry`).toBe(true);
      expect(referencedNames(path, NETWORK_GLOBALS), `${path} makes no request: delete its entry`).not.toEqual([]);
    }
  });

  it("no module opens a socket, an event source or an XHR — not even the gateway, which streams over fetch", () => {
    const others = new Set(["EventSource", "WebSocket", "XMLHttpRequest"]);
    const users = modules.filter((path) => referencedNames(path, others).length > 0);
    expect(users).toEqual([]);
  });

  it("no module sends a beacon (the telemetry primitive)", () => {
    const beacons = modules.filter((path) => referencedNames(path, new Set(["sendBeacon"])).length > 0);
    expect(beacons).toEqual([]);
  });

  it("the scanner sees through comments and strings, and catches property and element access", () => {
    const probe = (code: string) =>
      namesIn(ts.createSourceFile("probe.ts", code, ts.ScriptTarget.Latest, false, ts.ScriptKind.TS), NETWORK_GLOBALS);
    expect(probe('// fetch\nconst a = "fetch the image"; /* WebSocket */')).toEqual([]);
    expect(probe("window.fetch(url)")).toEqual(["fetch"]);
    expect(probe('globalThis["fetch"](url)')).toEqual(["fetch"]);
    expect(probe("new WebSocket(u); new XMLHttpRequest(); new EventSource(u)")).toEqual([
      "EventSource",
      "WebSocket",
      "XMLHttpRequest",
    ]);
  });
});

/** Whether an import/export declaration is erased at build: `import type`,
 * `export type`, or named bindings that are every one `type`. */
function isTypeOnly(node: ts.ImportDeclaration | ts.ExportDeclaration): boolean {
  if (ts.isExportDeclaration(node)) {
    return node.isTypeOnly;
  }
  const clause = node.importClause;
  if (!clause) {
    return false;
  }
  if (clause.isTypeOnly) {
    return true;
  }
  const bindings = clause.namedBindings;
  return (
    clause.name === undefined &&
    bindings !== undefined &&
    ts.isNamedImports(bindings) &&
    bindings.elements.length > 0 &&
    bindings.elements.every((element) => element.isTypeOnly)
  );
}

/** Static and dynamic import specifiers of a module that survive the build
 * (type-only imports load nothing, so they are not edges). */
function importSpecifiers(path: string): string[] {
  const file = sourceOf(path);
  const specifiers: string[] = [];
  const visit = (node: ts.Node): void => {
    if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.moduleSpecifier) {
      if (ts.isStringLiteral(node.moduleSpecifier) && !isTypeOnly(node)) {
        specifiers.push(node.moduleSpecifier.text);
      }
    } else if (
      ts.isCallExpression(node) &&
      node.expression.kind === ts.SyntaxKind.ImportKeyword &&
      node.arguments[0] &&
      ts.isStringLiteralLike(node.arguments[0])
    ) {
      specifiers.push(node.arguments[0].text);
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  return specifiers;
}

/** Runtime modules (outside lib/ai itself) that import `target`. */
function importersOf(target: RegExp): string[] {
  return modules.filter((path) => !path.startsWith("lib/ai/") && importSpecifiers(path).some((spec) => target.test(spec)));
}

describe("who may reach the gateway and the key (SPEC-113)", () => {
  it("only composables/useModelGateway.ts imports the gateway — the app's one door to it", () => {
    expect(importersOf(/(^|\/)ai\/modelGateway$/)).toEqual(["composables/useModelGateway.ts"]);
  });

  it("only the gateway and the provider guard read the API key store", () => {
    const keyReaders = modules.filter((path) =>
      importSpecifiers(path).some((spec) => /(^|\/)apiKey$/.test(spec) || spec === "./apiKey"),
    );
    expect(keyReaders.sort()).toEqual(["lib/ai/modelGateway.ts", "lib/ai/providerGuard.ts"]);
  });

  it("no document reader or writer, and no store, touches lib/ai", () => {
    const formatDirs = ["lib/docx/", "lib/odf/", "lib/md/", "lib/presentation/", "lib/core/", "lib/spreadsheet/", "lib/media/"];
    const offenders = modules.filter(
      (path) =>
        (formatDirs.some((dir) => path.startsWith(dir)) || path === "lib/documentStore.ts") &&
        importSpecifiers(path).some((spec) => /(^|\/)ai\//.test(spec)),
    );
    expect(offenders).toEqual([]);
  });
});
