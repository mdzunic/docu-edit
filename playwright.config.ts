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
// perf-charts (level 0, alongside the functional suite as before) →
// perf-find (level 1) → perf-large-document (level 2). Level 1 begins only
// once EVERY level-0 project has finished, so the two benchmarks that need
// the quietest window get the whole machine, one at a time, after the
// functional suite has drained. Deliberately still NOT
// `dependencies: ["chromium"]`: a dependency also SKIPS its dependents when
// it fails, and hanging the benchmarks off the functional suite would leave
// the budgets unmeasured whenever any functional test failed — including on
// a failures-only re-run (`--last-failed`), where the gate would then go
// green with the budgets never taken. Chaining benchmark-to-benchmark keeps
// that risk off the functional suite: the only thing that can skip a later
// benchmark is an EARLIER BENCHMARK FAILING, which is already a red run.
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
  expect: { timeout: 10_000 },
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
    // The benchmark chain (see the header): one per dependency level, so no
    // two of them — and, from perf-find on, nothing else in the suite — ever
    // share the machine with the run being timed. SPEC-052's two samples stay
    // at level 0, alongside the functional suite as before: they are gated on
    // absolute ceilings only (2s first paint, 100ms re-render), with no
    // baseline-relative row to drift, so they are the one benchmark that
    // tolerates a busy machine — which is what makes them a safe head for the
    // chain. The cost of the chain is that a perf-charts failure skips the two
    // below it; that run is already red, so no unmeasured budget can pass as
    // green.
    { name: "perf-charts", use: { ...devices["Desktop Chrome"] }, testMatch: /SPEC-052\.perf\.spec\.ts$/ },
    {
      name: "perf-find",
      use: { ...devices["Desktop Chrome"] },
      testMatch: /SPEC-017\.perf\.spec\.ts$/,
      dependencies: ["perf-charts"],
    },
    {
      name: "perf-large-document",
      use: { ...devices["Desktop Chrome"] },
      testMatch: /SPEC-047\.perf\.spec\.ts$/,
      dependencies: ["perf-find"],
    },
  ],
  webServer: {
    command: "npm run dev",
    url: "http://localhost:5173",
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
