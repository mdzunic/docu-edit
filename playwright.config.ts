import { defineConfig, devices } from "@playwright/test";

// The factory runs e2e in two situations:
//  - regression gate: no dev server running  → webServer starts one
//  - QA stage: container entrypoint already started it → reuseExistingServer
// reuseExistingServer must stay true for both paths to work.
//
// Playwright dispatches test groups to workers FIFO, in project order (see
// runner/tasks.js + dispatcher.js: one phase per dependency level, each
// project's groups appended in config order). The timing benchmarks below
// (`*.perf.spec.ts`) exploit that: their baselines are calibrated for the
// benchmark running on an otherwise quiet machine (see the SPEC-047 file
// header), yet inside one `chromium` project they were dispatched mid-suite
// and spent their whole 2-minute run sharing the CPU with a sibling worker —
// and the single-shot open/save rows crept over the 15% relative gate as the
// suite and the factory host got busier. Listing them in their own projects
// after `chromium` means they start only once every functional test has
// been dispatched; what still overlaps (the last in-flight functional tests,
// and SPEC-052's two short samples ahead of SPEC-047) lands inside the
// first of SPEC-047's five runs, which the median-of-5 absorbs. Deliberately
// NOT `dependencies: ["chromium"]`: a dependency skips the benchmarks whenever
// any functional test fails, and a failures-only re-run (`--last-failed`)
// would then never run them at all — the gate would go green with the
// budgets unmeasured.
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
    // SPEC-052's two single-shot samples and SPEC-017's find-first-match
    // benchmark go ahead of SPEC-047's median-of-5 so the long one inherits
    // the quietest window. SPEC-017's timing test lives in its own project
    // (rather than the mid-suite `chromium` run it used to share) for the same
    // reason every benchmark here does: its perf-baseline row is calibrated
    // for a quiet machine, so full-suite CPU contention was timing the sibling
    // workers, not the search.
    { name: "perf-charts", use: { ...devices["Desktop Chrome"] }, testMatch: /SPEC-052\.perf\.spec\.ts$/ },
    { name: "perf-find", use: { ...devices["Desktop Chrome"] }, testMatch: /SPEC-017\.perf\.spec\.ts$/ },
    { name: "perf-large-document", use: { ...devices["Desktop Chrome"] }, testMatch: /SPEC-047\.perf\.spec\.ts$/ },
  ],
  webServer: {
    command: "npm run dev",
    url: "http://localhost:5173",
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
