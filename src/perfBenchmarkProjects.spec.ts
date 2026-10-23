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
  const benchmarks = ["perf-charts", "perf-find", "perf-large-document", "perf-workbook"];

  it("gives every *.perf.spec.ts file its own project", () => {
    for (const name of benchmarks) {
      expect(named(name).name).toBe(name);
    }
    // The desktop project must keep ignoring them, or they run twice — once
    // isolated and once mid-suite, which is the contention this guards.
    const chromium = projects.find((p) => p.name === "chromium") as { testIgnore?: RegExp[] } | undefined;
    expect(chromium?.testIgnore?.some((pattern) => pattern.test("SPEC-047.perf.spec.ts"))).toBe(true);
    // SPEC-137: the workbook benchmark must be ignored by the desktop
    // project too — testIgnore's PERF_SPEC regex matches any *.perf.spec.ts,
    // so this should already hold without a config change of its own.
    expect(chromium?.testIgnore?.some((pattern) => pattern.test("SPEC-137.perf.spec.ts"))).toBe(true);
  });

  it("puts each benchmark project at its own dependency level", () => {
    const levels = benchmarks.map((name) => levelOf(named(name)));
    expect(new Set(levels).size, `benchmark levels ${levels.join(", ")} must all differ`).toBe(benchmarks.length);
  });

  it("chains the benchmarks to each other or the tests-less head, never to the functional suite", () => {
    // Hanging a benchmark off `chromium`/`compact` would skip its budgets
    // whenever any functional test failed — a red run today, but a green one
    // on a failures-only re-run with the budgets never taken. `perf-quiet`
    // is safe to depend on precisely because it can never fail (next test).
    for (const name of benchmarks) {
      for (const dep of named(name).dependencies ?? []) {
        expect(
          [...benchmarks, "perf-quiet"],
          `${name} may only depend on another benchmark or the perf-quiet head`,
        ).toContain(dep);
      }
    }
  });

  it("heads the chain with a project that matches no tests and so cannot fail", () => {
    const quiet = named("perf-quiet") as ProjectLike & { testMatch?: RegExp };
    expect(quiet.dependencies ?? []).toEqual([]);
    expect(quiet.testMatch).toBeInstanceOf(RegExp);
    for (const file of ["smoke.spec.ts", "SPEC-052.perf.spec.ts", "SPEC-068.compact.spec.ts"]) {
      expect(quiet.testMatch?.test(file), `perf-quiet must not own ${file}`).toBe(false);
    }
  });

  it("keeps every benchmark past the functional suite's level", () => {
    // All three carry ceilings calibrated for a quiet machine. perf-charts
    // originally ran at level 0 on the theory that its absolute ceilings
    // tolerate the busy window; the 2026-09-02 gate disproved that (its 2s
    // first-paint ceiling failed twice mid-suite), so the whole chain now
    // starts after the functional suite has drained.
    expect(levelOf(named("perf-charts"))).toBeGreaterThan(0);
    expect(levelOf(named("perf-find"))).toBeGreaterThan(levelOf(named("perf-charts")));
    expect(levelOf(named("perf-large-document"))).toBeGreaterThan(levelOf(named("perf-find")));
    // SPEC-137: perf-workbook is the chain's new tail, one level past
    // perf-large-document (today's tail) — so it too runs alone, after
    // every other benchmark has drained.
    expect(levelOf(named("perf-workbook"))).toBeGreaterThan(levelOf(named("perf-large-document")));
  });
});
