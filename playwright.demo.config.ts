import { defineConfig } from "@playwright/test";

// Headed, watchable product tour — NOT part of the regression suite.
//
// Separation works in both directions: the main playwright.config.ts only
// matches `*.spec.ts`/`*.test.ts`, so `e2e/demo/demo.tour.ts` is invisible to
// it (and to CI); this config only matches `*.tour.ts` under e2e/demo, so it
// never runs the regular specs. Run the tour with:
//
//   npx playwright test --config=playwright.demo.config.ts
//
// The browser window opens on screen (headless: false) and the run is also
// recorded to video under demo-results/.
//
// DEMO_PORT moves the whole run (baseURL + the server it starts) off 5173 —
// use it when another project's dev server already holds 5173, which the
// reuse below would otherwise happily point the tour at.
const port = Number(process.env.DEMO_PORT ?? 5173);

export default defineConfig({
  testDir: "./e2e/demo",
  testMatch: /\.tour\.ts$/,
  outputDir: "./demo-results",
  workers: 1,
  retries: 0,
  // The tour itself paces to ~2 minutes; leave slack for a cold dev server.
  timeout: 240_000,
  expect: { timeout: 10_000 },
  reporter: [["list"]],
  use: {
    baseURL: `http://localhost:${port}`,
    headless: false,
    viewport: { width: 1440, height: 900 },
    // Software compositing + no occlusion/background throttling: with GPU
    // compositing, the screencast recorder can drop a promoted layer (the
    // zoom-transformed page went blank in one recording while the app —
    // and a Playwright screenshot — rendered it fine), and covering the
    // demo window with another app degrades captured frames. Recording
    // fidelity beats raster speed here.
    launchOptions: {
      args: [
        "--window-position=40,40",
        "--disable-gpu",
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding",
        "--disable-background-timer-throttling",
      ],
    },
    video: { mode: "on", size: { width: 1440, height: 900 } },
    trace: "off",
  },
  // Same contract as the main config: reuse a dev server if one is already
  // listening on the port, otherwise start one there.
  webServer: {
    command: port === 5173 ? "npm run dev" : `npm run dev -- --port ${port} --strictPort`,
    url: `http://localhost:${port}`,
    reuseExistingServer: true,
    timeout: 120_000,
  },
});
