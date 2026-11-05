import { test, expect } from "@playwright/test";

// MAIN-b0d62ff — repair main: the product check suite was failing on a clean
// clone because e2e/SPEC-137.spec.ts:147 reloaded the page right after typing
// into a cell, before SPEC-138's draft/autosave cadence (introduced after
// SPEC-137 was written) had written the edit. The reload could land before
// the draft was persisted and read back the pre-edit value. The fix makes
// that test wait for the `save-state` badge's `data-save-state` attribute to
// read "saved" before reloading — no product code changed.
//
// This spec is the QA-stage codification of the repair's own acceptance
// criteria (AC-1, AC-2): the app still boots, and the edit → autosave →
// reload flow the regression hit still works end to end, for the word
// processor (the product this draft/autosave cadence, SPEC-010, is shared
// by).

test("the application loads and renders its main screen (AC-1)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Welcome to docu-edit" })).toBeVisible();
  await expect(page.getByTestId("new-document")).toBeVisible();
});

test("a document edit is saved on the draft cadence and survives a reload (AC-2)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  const content = page.getByTestId("document-content");
  await content.click();
  await page.keyboard.type("Repair verification MAIN-b0d62ff", { delay: 20 });
  await expect(content).toContainText("Repair verification MAIN-b0d62ff");

  // Wait for the draft cadence to actually write before reloading — the
  // exact ordering bug this repair fixed in e2e/SPEC-137.spec.ts.
  await expect(page.getByTestId("save-state")).toHaveAttribute("data-save-state", "saved", { timeout: 15_000 });

  await page.reload();
  await expect(page.getByTestId("document-content")).toContainText("Repair verification MAIN-b0d62ff");
});
