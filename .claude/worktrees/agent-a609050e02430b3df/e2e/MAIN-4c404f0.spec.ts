import { test, expect } from "@playwright/test";

// MAIN-4c404f0 — repair main: the product check suite (install, unit,
// typecheck, lint, e2e) was failing on a clean clone because three e2e
// timing benchmarks shared the machine with the functional suite and each
// other (playwright.config.ts's dependency chain fixes the scheduling; see
// its header and src/perfBenchmarkProjects.spec.ts). No product code
// changed. This spec is the QA-stage codification of the repair's own
// acceptance criteria: the app still boots and the core edit-and-export
// flow still works end to end.

test("the application loads and renders its main screen (AC-1)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Welcome to docu-edit" })).toBeVisible();
  await expect(page.getByTestId("new-document")).toBeVisible();
});

test("create, type, autosave and export a document round-trips (AC-2)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  const content = page.getByTestId("document-content");
  await content.click();
  await page.keyboard.type("Hello from QA, verifying the primary editing flow after the repair.");

  await expect(page.getByTestId("word-count-button")).toHaveText("11 words");
  await expect(page.getByText("All changes saved")).toBeVisible();

  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const download = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]).then(([d]) => d);
  expect(download.suggestedFilename()).toMatch(/\.docx$/);
});
