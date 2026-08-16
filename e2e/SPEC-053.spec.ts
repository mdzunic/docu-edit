import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-053 — the unified field system. Unit/component tests cover the evaluator,
// the readers/writers, the dialogs and the field commands in isolation; this
// file exercises the seams only a real browser can confirm and that a prior QA
// round found broken: inserting a field through the Date and time dialog (the
// selection the handler reads was always null), and a field being a clickable,
// shadeable one-character atom in the actual canvas.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const fieldAtoms = (page: Page) => page.locator('[data-testid="field"]');

async function openMenu(page: Page, menuId: string): Promise<void> {
  await page.getByTestId(`menu-title-${menuId}`).click();
}

async function insertDateField(page: Page): Promise<void> {
  await openMenu(page, "insert");
  await page.getByTestId("insert-date-time-button").click();
  await expect(page.getByTestId("date-time-dialog")).toBeVisible();
  await page.getByTestId("date-time-format-d MMMM yyyy").check();
  await page.getByTestId("date-time-update-automatically").check();
  await page.getByTestId("date-time-insert").click();
  await expect(page.getByTestId("date-time-dialog")).not.toBeVisible();
}

test("Insert → Date and time inserts plain text with the toggle off and a date field with it on (AC-4, AC-5)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Signed ");

  // Toggle off → plain text, never a field.
  await openMenu(page, "insert");
  await page.getByTestId("insert-date-time-button").click();
  await expect(page.getByTestId("date-time-dialog")).toBeVisible();
  await page.getByTestId("date-time-format-d MMMM yyyy").check();
  await page.getByTestId("date-time-insert").click();
  await expect(page.getByTestId("date-time-dialog")).not.toBeVisible();
  await expect(fieldAtoms(page)).toHaveCount(0);
  await expect(content(page)).toContainText("Signed");

  // Toggle on → exactly one date field atom is inserted.
  await insertDateField(page);
  await expect(fieldAtoms(page)).toHaveCount(1);
});

test("clicking a field selects its whole atom — it shades and shows the toolbar chip — and typing over it replaces it (AC-7, AC-8, AC-10, AC-14)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await insertDateField(page);

  const field = fieldAtoms(page).first();
  await expect(field).toHaveCount(1);

  // A plain click lands the selection on the field: it shades (AC-10) and the
  // field toolbar chip appears (AC-14) — both were broken when the caret never
  // reached the field.
  await field.click();
  await expect(field).toHaveClass(/doc-field-shaded/);
  await expect(page.getByTestId("field-toolbar")).toBeVisible();

  // Typing over the selected field replaces the whole field (AC-8).
  await page.keyboard.type("Z");
  await expect(fieldAtoms(page)).toHaveCount(0);
  await expect(content(page)).toContainText("Z");
});

test("View → Field shading shades a field regardless of the caret (AC-11)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Draft ");
  await insertDateField(page);

  // Move the caret well away from the field so any shading is the View toggle's
  // doing, not caret proximity.
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.type(" copy");
  const field = fieldAtoms(page).first();
  await expect(field).not.toHaveClass(/doc-field-shaded/);

  await openMenu(page, "view");
  await page.getByTestId("view-menu-field-shading").click();
  await expect(field).toHaveClass(/doc-field-shaded/);
});
