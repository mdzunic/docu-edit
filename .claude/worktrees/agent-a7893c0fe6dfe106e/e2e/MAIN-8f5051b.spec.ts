import { test, expect } from "@playwright/test";

// MAIN-8f5051b — repair main: the product check suite (install, unit,
// typecheck, lint, e2e) was failing on a clean clone with three e2e failures
// — SPEC-047's ceiling warning bar, SPEC-051's text-wrap regression test and
// SPEC-052's 20-chart paint budget. None of the three was a product defect:
// each passes in isolation here, and the paint budget measures the same
// 451-1045ms at HEAD as it does at dbd2ae7, the last commit before SPEC-077
// widened the document model. What the gate's own log shows instead is a
// container whose every test ran 5-10x slower than it does on a quiet one
// (SPEC-047's five-block document: 13.9s to open, against ~2s here), so the
// two Playwright timeouts and a single-sample perf reading all landed on the
// wrong side of the line. The repair raises those timeouts to a budget that
// reflects where this runs and takes the paint budget from the median of
// five opens instead of one, the instrument SPEC-047.perf.spec.ts already
// uses. No product code changed and no assertion was relaxed.
//
// This spec is the QA-stage codification of the repair's own acceptance
// criteria: the app still boots and the core edit flow still works end to
// end.

test("the application loads and renders its main screen (AC-1)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Welcome to docu-edit" })).toBeVisible();
  await expect(page.getByTestId("new-document")).toBeVisible();
});

test("create a document, type in it and save: the core flow works end to end (AC-2)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  const content = page.getByTestId("document-content");
  await content.click();
  await page.keyboard.type("The primary editing flow still works after the main repair.");

  await expect(content).toHaveText("The primary editing flow still works after the main repair.");
  await expect(page.getByTestId("word-count-button")).toHaveText("10 words");
  await expect(page.getByTestId("save-state")).toHaveText("All changes saved");

  // The document it just created is in the library, under the title the
  // editor gave it — the open path, not only the edit path.
  const title = (await page.getByTestId("document-title").innerText()).trim();
  await page.goto("/");
  await expect(page.getByTestId("row-title").filter({ hasText: title }).first()).toBeVisible();
});
