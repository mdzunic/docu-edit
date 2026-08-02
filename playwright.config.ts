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
