import { test, expect } from "@playwright/test";

// Repair-main guard: the app boots to its landing screen and the core
// word-processor flow (create a document, type into it, and have that content
// survive a reload) works end to end. This is the product's primary flow; if a
// future change breaks app load or document persistence, this fails fast.

test("app renders its main landing screen", async ({ page }) => {
  const res = await page.goto("/");
  expect(res?.status()).toBe(200);
  await expect(page.locator("#app")).toBeAttached();
  // The landing screen offers the primary entry points, so it rendered.
  await expect(page.getByTestId("new-document")).toBeVisible();
});

test("core flow: create a document, edit it, and edits persist across reload", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  const content = page.getByTestId("document-content");
  await expect(content).toBeVisible();

  await content.click();
  const marker = "Hello from the core-flow guard.";
  await page.keyboard.type(marker);
  await expect(content).toContainText(marker);

  const url = page.url();
  await page.waitForTimeout(600); // let autosave settle before reloading
  await page.reload();
  await expect(page).toHaveURL(url);
  await expect(page.getByTestId("document-content")).toContainText(marker);
});
