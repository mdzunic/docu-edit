import { defineConfig, devices } from "@playwright/test";

// The factory runs e2e in two situations:
//  - regression gate: no dev server running  → webServer starts one
//  - QA stage: container entrypoint already started it → reuseExistingServer
// reuseExistingServer must stay true for both paths to work.
//
// Playwright dispatches test groups to workers FIFO, in project order, one
// PHASE per dependency level (see runner/tasks.js + dispatcher.js): every
// project at one level finishes before any project at the next level starts.
// The timing benchmarks below (`*.perf.spec.ts`) need that, because their
// baselines are calibrated for the benchmark running on an otherwise quiet
// machine (see the SPEC-047 file header): inside one `chromium` project they
// were dispatched mid-suite and spent their whole 2-minute run sharing the
// CPU with a sibling worker, and the single-shot open/save rows crept over
// the 15% relative gate as the suite and the factory host got busier.
//
// Separate projects alone were not enough. With every benchmark project at
// dependency level 0, the three of them are dispatched together the moment
// workers free up — so SPEC-017's find benchmark ran *concurrently with*
// SPEC-047's median-of-5, and each timed the other. That is exactly the pair
// that failed the 2026-09-02 gate (perf-find and perf-large-document, with
// perf-charts green), and it reproduces on demand: pinned to three CPUs, the
// SPEC-047 run that overlaps perf-find measures mainThreadBlockDuringOpenMs
// at 247ms against 103-131ms for the four that do not — the gate's 229.
//
// So the benchmarks are chained into their own levels instead:
// perf-charts (level 1) → perf-find (level 2) → perf-large-document
// (level 3) → perf-workbook (level 4, SPEC-137). Level 1 begins only once EVERY level-0 project — the whole
// functional suite — has finished, so each benchmark gets the whole
// machine, one at a time, after the suite has drained. Deliberately still
// NOT `dependencies: ["chromium"]`: a dependency also SKIPS its dependents
// when it fails, and hanging the benchmarks off the functional suite would
// leave the budgets unmeasured whenever any functional test failed —
// including on a failures-only re-run (`--last-failed`), where the gate
// would then go green with the budgets never taken. The head of the chain
// is instead `perf-quiet`, a project whose testMatch matches nothing: it
// finishes instantly, cannot fail, and exists only to place the first
// benchmark one level past the functional suite. (perf-charts originally
// sat at level 0 itself on the theory that its absolute ceilings tolerate
// a busy machine; the 2026-09-02 gate disproved that — its 2s first-paint
// ceiling failed twice while sharing the workers and the dev server with
// the functional suite.)
const PERF_SPEC = /\.perf\.spec\.ts$/;
// SPEC-068: the compact touch project owns the *.compact.spec.ts files and runs
// them at a 390×844 phone viewport with touch emulation. The desktop `chromium`
// project ignores them so the full suite keeps running unchanged at the desktop
// viewport (AC-48) — these specs assert the compact layout, which the desktop
// project would not exercise.
const COMPACT_SPEC = /\.compact\.spec\.ts$/;

export default defineConfig({
  testDir: "./e2e",
  fullyParallel: true,
  forbidOnly: false,
  retries: 0,
  // Playwright's 5s default assumes a warm dev server. The factory runs this
  // suite against a clean clone inside a container: Vite transforms every
  // module on first request while every worker asks at once, so the first
  // wait on a lazily-loaded route can legitimately outlast 5s. Failing tests
  // still fail — they just get a budget that reflects where this runs.
  //
  // 10s was still not that budget. On the 2026-09-04 main gate every test in
  // the suite ran 5-10x slower than the same test does on a quiet container:
  // SPEC-047's *five-block* document — the trivial half of that file — took
  // 13.9s to open and assert, against ~2s here, and its ceiling sibling blew
  // this timeout waiting for a warning bar that appears in 0.6s (1.5s cold)
  // when the machine is not oversubscribed. Four Playwright workers, each
  // driving a Chromium, plus the single-threaded Vite dev server they all
  // request every module from, is more than the container's CPU quota, and
  // the whole suite dilates together. Nothing about a passing run costs more
  // for these being larger — a timeout is only ever spent on a wait that was
  // going to fail — so they are set where a loaded factory host still lands
  // inside them.
  timeout: 60_000,
  expect: { timeout: 20_000 },
  reporter: [["list"]],
  use: {
    baseURL: "http://localhost:5173",
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
  },
  projects: [
    { name: "chromium", use: { ...devices["Desktop Chrome"] }, testIgnore: [PERF_SPEC, COMPACT_SPEC] },
    // A phone-sized, touch-enabled project for the compact tier (SPEC-068
    // AC-47). 390×844 is the reference compact viewport; hasTouch/isMobile turn
    // on touch emulation and the mobile input model.
    {
      name: "compact",
      use: {
        ...devices["Desktop Chrome"],
        viewport: { width: 390, height: 844 },
        hasTouch: true,
        isMobile: true,
      },
      testMatch: COMPACT_SPEC,
    },
    // The benchmark chain (see the header): one per dependency level, each
    // starting only after the functional suite and every earlier benchmark
    // have finished, so nothing shares the machine with the run being timed.
    //
    // Every benchmark project sets `bypassCSP` (SEC-6). The three perf specs
    // are the only ones that time a paint with `page.waitForFunction`, which
    // Playwright evaluates in the page's main world — and index.html's
    // `script-src 'self' 'sha256-...'` refuses it, since a CSP cannot
    // distinguish the harness's predicate from injected script. Scoped to
    // these projects on purpose: the functional suite keeps running under the
    // real policy, so a change that breaks the app under CSP still fails the
    // suite. Nothing in the app itself needs eval — and the alternative,
    // rewriting the predicate to run Playwright-side, is the exact
    // measurement mistake SPEC-052.perf's own header documents at length.
    // perf-quiet is the chain's tests-less head — it matches no files, so it
    // completes instantly and cannot fail, and its only job is to hold the
    // first real benchmark at level 1. The cost of the chain is that a
    // benchmark failure skips the ones below it; that run is already red, so
    // no unmeasured budget can pass as green.
    { name: "perf-quiet", testMatch: /^$/ },
    {
      name: "perf-charts",
      use: { ...devices["Desktop Chrome"], bypassCSP: true },
      testMatch: /SPEC-052\.perf\.spec\.ts$/,
      dependencies: ["perf-quiet"],
    },
    {
      name: "perf-find",
      use: { ...devices["Desktop Chrome"], bypassCSP: true },
      testMatch: /SPEC-017\.perf\.spec\.ts$/,
      dependencies: ["perf-charts"],
    },
    {
      name: "perf-large-document",
      use: { ...devices["Desktop Chrome"], bypassCSP: true },
      testMatch: /SPEC-047\.perf\.spec\.ts$/,
      dependencies: ["perf-find"],
    },
    // SPEC-137: the workbook performance benchmark — chained after
    // perf-large-document (today's chain tail), so it too gets the whole
    // machine to itself once everything above it has drained.
    {
      name: "perf-workbook",
      use: { ...devices["Desktop Chrome"], bypassCSP: true },
      testMatch: /SPEC-137\.perf\.spec\.ts$/,
      dependencies: ["perf-large-document"],
    },
  ],
  webServer: {
    command: "npm run dev",
    url: "http://localhost:5173",
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
