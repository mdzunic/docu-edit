// Architecture-as-test for the e2e timing benchmarks' isolation, in the same
// spirit as lib/core/boundary.spec.ts and the view size budgets: the thing it
// protects is a property of the repo's layout, not of any one module, so
// nothing else in the suite would notice it breaking.
//
// Every `*.perf.spec.ts` budget is calibrated for its benchmark running on an
// otherwise quiet machine (see the file headers of e2e/SPEC-047.perf.spec.ts
// and e2e/SPEC-017.perf.spec.ts). Giving each benchmark its own Playwright
// project was not enough to deliver that: projects at the same dependency
// level are dispatched together the moment workers free up, so SPEC-017's
// find benchmark ran concurrently with SPEC-047's median-of-5 and each timed
// the other. That pairing is what failed the 2026-09-02 main gate
// (mainThreadBlockDuringOpenMs 229 against a 149.5 gate, saveMs 1937 against
// 1380), and it reproduces on demand: pinned to three CPUs, the one SPEC-047
// run overlapping perf-find measured 247ms of open-phase block against
// 103-131ms for the four that did not.
//
// The fix is a dependency chain, one benchmark per level (playwright.config.ts).
// This test pins it: a future project added at the same level as another
// benchmark, or a dropped `dependencies`, silently restores the contention.
import { describe, expect, it } from "vitest";
import config from "../playwright.config";

type ProjectLike = { name?: string; dependencies?: string[] };

const projects = (config.projects ?? []) as ProjectLike[];
const named = (name: string): ProjectLike => {
  const project = projects.find((p) => p.name === name);
  if (!project) {
    throw new Error(`playwright.config.ts has no "${name}" project`);
  }
  return project;
};

/** The dependency level a project runs at: 0 for a project with no
 * dependencies, otherwise one past the deepest project it depends on. This is
 * the number that decides concurrency — Playwright runs one phase per level,
 * and every project at a level runs together. */
function levelOf(project: ProjectLike, seen: ReadonlySet<string> = new Set()): number {
  const deps = project.dependencies ?? [];
  if (deps.length === 0) {
    return 0;
  }
  expect(seen.has(project.name ?? ""), `dependency cycle at ${project.name}`).toBe(false);
  const nextSeen = new Set([...seen, project.name ?? ""]);
  return 1 + Math.max(...deps.map((dep) => levelOf(named(dep), nextSeen)));
}

describe("e2e timing benchmarks run one at a time", () => {
  const benchmarks = ["perf-charts", "perf-find", "perf-large-document"];

  it("gives every *.perf.spec.ts file its own project", () => {
    for (const name of benchmarks) {
      expect(named(name).name).toBe(name);
    }
    // The desktop project must keep ignoring them, or they run twice — once
    // isolated and once mid-suite, which is the contention this guards.
    const chromium = projects.find((p) => p.name === "chromium") as { testIgnore?: RegExp[] } | undefined;
    expect(chromium?.testIgnore?.some((pattern) => pattern.test("SPEC-047.perf.spec.ts"))).toBe(true);
  });

  it("puts each benchmark project at its own dependency level", () => {
    const levels = benchmarks.map((name) => levelOf(named(name)));
    expect(new Set(levels).size, `benchmark levels ${levels.join(", ")} must all differ`).toBe(benchmarks.length);
  });

  it("chains the benchmarks to each other, never to the functional suite", () => {
    // Hanging a benchmark off `chromium`/`compact` would skip its budgets
    // whenever any functional test failed — a red run today, but a green one
    // on a failures-only re-run with the budgets never taken.
    for (const name of benchmarks) {
      for (const dep of named(name).dependencies ?? []) {
        expect(benchmarks, `${name} may only depend on another benchmark`).toContain(dep);
      }
    }
  });

  it("leaves the two baseline-gated benchmarks below the head of the chain", () => {
    // perf-charts is gated on absolute ceilings only, so it tolerates the
    // busy level-0 window; perf-find and perf-large-document carry
    // perf-baseline.json rows and must run after the functional suite has
    // drained, which is what a non-zero level buys them.
    expect(levelOf(named("perf-find"))).toBeGreaterThan(0);
    expect(levelOf(named("perf-large-document"))).toBeGreaterThan(levelOf(named("perf-find")));
  });
});
