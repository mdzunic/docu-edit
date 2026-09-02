import { test, expect } from "@playwright/test";

// MAIN-0d4e899 — repair main: the product check suite (install, unit,
// typecheck, lint, e2e) was failing on a clean clone because AC-39's
// data-sheet re-render benchmark (e2e/SPEC-052.perf.spec.ts) timed itself
// from the Playwright side, so Playwright's own actionability/CDP overhead
// around locator.click() was counted against the product's 100ms budget.
// The fix moves the stopwatch inside the page — an in-page MutationObserver
// on the chart canvas — so the reading covers only click -> model update ->
// repaint -> frame on screen. No other product code changed. This spec is
// the QA-stage codification of the repair's own acceptance criteria: the
// app still boots and the core edit flow still works end to end.

test("the application loads and renders its main screen (AC-1)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Welcome to docu-edit" })).toBeVisible();
  await expect(page.getByTestId("new-document")).toBeVisible();
});

test("create a document and type in it: the core edit flow works end to end (AC-2)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  const content = page.getByTestId("document-content");
  await content.click();
  await page.keyboard.type("Hello from QA, verifying the primary editing flow after the repair.");

  await expect(content).toHaveText("Hello from QA, verifying the primary editing flow after the repair.");
  await expect(page.getByTestId("word-count-button")).toHaveText("11 words");
  await expect(page.getByText("All changes saved")).toBeVisible();
});
