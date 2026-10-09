import { test, expect } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// MAIN-f8da022 — repair main: the product check suite (install, unit,
// typecheck, lint, e2e) was failing on a clean clone because SPEC-004 and
// SPEC-005's picker "accept" assertions still held the pre-SPEC-116 value,
// before SPEC-116's AC-1 added the six workbook extensions to
// OPENABLE_EXTENSIONS. The fix only updates those two e2e expectations to the
// value the picker already produces. No product code changed. This spec is
// the QA-stage codification of the repair's own acceptance criteria: the app
// still boots, and the open-document flow those two tests guard still works
// end to end.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "src", "lib", "docx", "fixtures");
const wordFixture = join(fixturesDir, "word-authored.docx");

test("the application loads and renders its main screen (AC-1)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Welcome to docu-edit" })).toBeVisible();
  await expect(page.getByTestId("new-document")).toBeVisible();
});

test("opening a .docx through the file picker renders it at /doc/:id (AC-2)", async ({ page }) => {
  await page.goto("/");
  const input = page.getByTestId("open-document-input");
  // The repaired assertion's own value: SPEC-116's workbook extensions sit
  // alongside every earlier spec's additions, and .docx/.odt are still there.
  await expect(input).toHaveAttribute(
    "accept",
    ".docx,.odt,.md,.markdown,.pptx,.odp,.key,.xlsx,.xlsm,.xltx,.xltm,.ods,.ots,.pdf,.png,.jpg,.jpeg,.gif,.svg,.webp",
  );

  await input.setInputFiles(wordFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("document-title")).toHaveText("word-authored");
  await expect(page.getByTestId("format-badge")).toHaveText("DOCX");
  await expect(page.getByRole("heading", { name: "Quarterly Report" })).toBeVisible();
});
