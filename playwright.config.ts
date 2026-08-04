import { defineConfig, devices } from "@playwright/test";

// The factory runs e2e in two situations:
//  - regression gate: no dev server running  → webServer starts one
//  - QA stage: container entrypoint already started it → reuseExistingServer
// reuseExistingServer must stay true for both paths to work.
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
  projects: [{ name: "chromium", use: { ...devices["Desktop Chrome"] } }],
  webServer: {
    command: "npm run dev",
    url: "http://localhost:5173",
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
