import { test, expect } from "@playwright/test";

// Baseline e2e check: the dev server serves the app and Vue mounts.
// Deliberately asserts nothing about layout or copy, so it stays green as
// features land. Feature-specific e2e specs are written by the QA agent as
// e2e/SPEC-NNN.spec.ts.
test("app boots and mounts", async ({ page }) => {
  const response = await page.goto("/");
  expect(response?.status()).toBe(200);
  await expect(page.locator("#app")).toBeAttached();
});
