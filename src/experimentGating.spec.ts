// SPEC-113 — a disabled experiment is absent, not hidden. Each gated entry
// point renders nothing and loads nothing while its flag is off:
// - ExperimentGate never calls its loader (so the surface's import() never
//   runs) and renders nothing; switched on it loads once and renders.
// - guardExperiment sends an experimental route to "/".
// - Labs' model provider panel, and the gateway behind it, are not loaded
//   (views/preferences/LabsSection.spec.ts mounts that one).
// And the code itself is out of the initial bundle: walking the static
// import graph from main.ts and from every route's component, the gateway
// and the surfaces that load it are never reached — only an import() leads
// to them, and each such import() sits behind a flag check.
import { existsSync, readFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import ts from "typescript";
import { flushPromises, mount } from "@vue/test-utils";
import { defineComponent, h, nextTick } from "vue";
import { beforeEach, describe, expect, it, vi } from "vitest";
import ExperimentGate from "./components/ExperimentGate.vue";
import { guardExperiment } from "./router";
import { EXPERIMENTS, setExperimentEnabled } from "./lib/experiments";
import { resetPreferences } from "./lib/preferences";

const SRC = dirname(fileURLToPath(import.meta.url));

beforeEach(() => {
  localStorage.clear();
  resetPreferences();
});

describe("ExperimentGate", () => {
  const Surface = defineComponent({ render: () => h("p", { "data-testid": "surface" }, "experimental surface") });

  it("renders nothing and calls no loader while the flag is off", async () => {
    const load = vi.fn(async () => Surface);
    const wrapper = mount(ExperimentGate, { props: { experiment: "deck-designer", load } });
    await flushPromises();
    expect(load).not.toHaveBeenCalled();
    expect(wrapper.find('[data-testid="surface"]').exists()).toBe(false);
    expect(wrapper.text()).toBe("");
  });

  it("loads and renders once the flag is on, unmounts when it goes off, and never loads twice", async () => {
    const load = vi.fn(async () => Surface);
    const wrapper = mount(ExperimentGate, { props: { experiment: "library-search", load } });
    setExperimentEnabled("library-search", true);
    await flushPromises();
    expect(load).toHaveBeenCalledTimes(1);
    expect(wrapper.find('[data-testid="surface"]').exists()).toBe(true);

    setExperimentEnabled("library-search", false);
    await nextTick();
    expect(wrapper.find('[data-testid="surface"]').exists()).toBe(false);

    setExperimentEnabled("library-search", true);
    await flushPromises();
    expect(wrapper.find('[data-testid="surface"]').exists()).toBe(true);
    expect(load).toHaveBeenCalledTimes(1);
  });

  it("is gated by its own experiment only", async () => {
    const load = vi.fn(async () => Surface);
    mount(ExperimentGate, { props: { experiment: "deck-designer", load } });
    setExperimentEnabled("library-search", true);
    await flushPromises();
    expect(load).not.toHaveBeenCalled();
  });
});

describe("guardExperiment", () => {
  it("redirects an experiment's route to / while the experiment is off, and lets it through when on", () => {
    for (const { id } of EXPERIMENTS) {
      const guard = guardExperiment(id);
      expect(guard()).toBe("/");
      setExperimentEnabled(id, true);
      expect(guard()).toBe(true);
      setExperimentEnabled(id, false);
      expect(guard()).toBe("/");
    }
  });
});

// --- the static import graph -------------------------------------------------

/** A module's code (a .vue file's <script> blocks) parsed as TypeScript. */
function parse(file: string): ts.SourceFile {
  const source = readFileSync(file, "utf8");
  const code = file.endsWith(".vue")
    ? [...source.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)].map((match) => match[1]).join("\n")
    : source;
  return ts.createSourceFile(file, code, ts.ScriptTarget.Latest, false, ts.ScriptKind.TS);
}

/** Whether a declaration is erased at build (`import type`, all-`type`
 * named bindings, `export type`) and so loads nothing. */
function erased(node: ts.ImportDeclaration | ts.ExportDeclaration): boolean {
  if (ts.isExportDeclaration(node)) {
    return node.isTypeOnly;
  }
  const clause = node.importClause;
  if (!clause) {
    return false;
  }
  const bindings = clause.namedBindings;
  return (
    clause.isTypeOnly ||
    (clause.name === undefined &&
      bindings !== undefined &&
      ts.isNamedImports(bindings) &&
      bindings.elements.length > 0 &&
      bindings.elements.every((element) => element.isTypeOnly))
  );
}

/** The specifiers a module imports: `static` edges are bundled with it,
 * `dynamic` ones are import() calls — separate chunks, loaded on demand. */
function edges(file: string): { static: string[]; dynamic: string[] } {
  const found = { static: [] as string[], dynamic: [] as string[] };
  const visit = (node: ts.Node): void => {
    if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.moduleSpecifier) {
      if (ts.isStringLiteral(node.moduleSpecifier) && !erased(node)) {
        found.static.push(node.moduleSpecifier.text);
      }
    } else if (
      ts.isCallExpression(node) &&
      node.expression.kind === ts.SyntaxKind.ImportKeyword &&
      node.arguments[0] &&
      ts.isStringLiteralLike(node.arguments[0])
    ) {
      found.dynamic.push(node.arguments[0].text);
    }
    ts.forEachChild(node, visit);
  };
  visit(parse(file));
  return found;
}

/** Resolves a relative or "@/" specifier to a source file; packages and
 * assets resolve to null. */
function resolveSpecifier(from: string, specifier: string): string | null {
  const bare = specifier.replace(/\?.*$/, "");
  const base = bare.startsWith("@/")
    ? join(SRC, bare.slice(2))
    : bare.startsWith(".")
      ? resolve(dirname(from), bare)
      : null;
  if (base === null) {
    return null;
  }
  for (const candidate of [base, `${base}.ts`, `${base}.vue`, join(base, "index.ts")]) {
    if (existsSync(candidate) && /\.(ts|vue)$/.test(candidate)) {
      return candidate;
    }
  }
  return null;
}

/** Every module bundled statically with `entries` — what loads without any
 * import() being taken. Paths relative to src/. */
function staticClosure(entries: string[]): Set<string> {
  const seen = new Set<string>();
  const queue = [...entries];
  while (queue.length > 0) {
    const file = queue.pop() as string;
    const rel = relative(SRC, file).replace(/\\/g, "/");
    if (seen.has(rel)) {
      continue;
    }
    seen.add(rel);
    for (const specifier of edges(file).static) {
      const target = resolveSpecifier(file, specifier);
      if (target !== null) {
        queue.push(target);
      }
    }
  }
  return seen;
}

/** The modules that import() `target` (relative to src/), anywhere in `scope`. */
function dynamicImportersOf(target: string, scope: Iterable<string>): string[] {
  return [...scope].filter((rel) =>
    edges(join(SRC, rel)).dynamic.some(
      (specifier) => relative(SRC, resolveSpecifier(join(SRC, rel), specifier) ?? "").replace(/\\/g, "/") === target,
    ),
  );
}

describe("a disabled experiment's code is not in the initial bundle", () => {
  const mainTs = join(SRC, "main.ts");
  // The routes' components are themselves import()s, so each is an entry of
  // its own chunk; a module reached statically from any of them loads the
  // moment that route is visited, flag or no flag.
  const routeComponents = edges(join(SRC, "router.ts"))
    .dynamic.map((specifier) => resolveSpecifier(join(SRC, "router.ts"), specifier))
    .filter((file): file is string => file !== null);
  const closure = staticClosure([mainTs, ...routeComponents]);

  /** The modules only an experiment's import() may lead to. */
  const GATED = [
    "lib/ai/modelGateway.ts",
    "composables/useModelGateway.ts",
    "views/preferences/ModelProviderPanel.vue",
    "views/preferences/ModelActivityLog.vue",
    "components/ModelConsentDialog.vue",
    // SPEC-115's deck designer: its entry points, its dialog and session, and
    // the whole of lib/design load only behind the deck-designer gates.
    "views/design/NewDeckFromDescription.vue",
    "views/design/RedesignMenuItem.vue",
    "views/design/DeckDesignerDialog.vue",
    "views/design/useDeckDesigner.ts",
    "views/design/useDesignerBrief.ts",
    "views/design/useDesignerPlan.ts",
    "views/design/designerLibrary.ts",
    "views/design/outlineSource.ts",
    "lib/design/plan.ts",
    "lib/design/reply.ts",
    "lib/design/planToDescription.ts",
    "lib/design/applyDesign.ts",
    "lib/design/outlineDesigner.ts",
    "lib/design/prompt.ts",
    "lib/design/build.ts",
  ];

  it("walks a real graph: the routes, the shell and Labs are in it", () => {
    expect(routeComponents.length).toBeGreaterThan(5);
    for (const expected of [
      "main.ts",
      "router.ts",
      "views/HomeView.vue",
      "components/AppShell.vue",
      "components/PreferencesDialog.vue",
      "views/preferences/LabsSection.vue",
      "components/ExperimentGate.vue",
      "lib/experiments.ts",
    ]) {
      expect(closure.has(expected), expected).toBe(true);
    }
  });

  it("never statically reaches the gateway or the surfaces that load it", () => {
    for (const gated of GATED) {
      expect(existsSync(join(SRC, gated)), `${gated} exists`).toBe(true);
      expect(closure.has(gated), `${gated} is statically reachable from main.ts or a route`).toBe(false);
    }
  });

  it("every surface a gate loads carries the Experimental badge", () => {
    // A component that hosts an ExperimentGate imports its surfaces lazily
    // for it; each of those surfaces must render ExperimentBadge, so a user
    // always knows which part of the product is provisional.
    const hosts = [...closure].filter(
      (rel) => rel.endsWith(".vue") && readFileSync(join(SRC, rel), "utf8").includes("<ExperimentGate"),
    );
    expect(hosts).toContain("views/preferences/LabsSection.vue");
    const surfaces = hosts.flatMap((rel) =>
      edges(join(SRC, rel))
        .dynamic.map((specifier) => resolveSpecifier(join(SRC, rel), specifier))
        .filter((file): file is string => file !== null && file.endsWith(".vue")),
    );
    expect(surfaces.length).toBeGreaterThan(0);
    for (const surface of surfaces) {
      expect(readFileSync(surface, "utf8"), relative(SRC, surface)).toMatch(/<ExperimentBadge\b/);
    }
  });

  it("reaches the provider panel only through Labs' flag-gated import(), and the consent dialog only on demand", () => {
    expect(dynamicImportersOf("views/preferences/ModelProviderPanel.vue", closure)).toEqual([
      "views/preferences/LabsSection.vue",
    ]);
    expect(dynamicImportersOf("components/ModelConsentDialog.vue", closure)).toEqual(["components/AppShell.vue"]);
    // LabsSection hands that import() to ExperimentGate, which calls it only
    // while a model experiment is on (the tests above and LabsSection.spec).
    expect(readFileSync(join(SRC, "views/preferences/LabsSection.vue"), "utf8")).toMatch(
      /<ExperimentGate\s+v-if="modelExperiment"\s+:experiment="modelExperiment"\s+:load="loadProviderPanel"/,
    );
  });
});
