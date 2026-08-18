import { test, expect, type Page } from "@playwright/test";

// SPEC-056 equations. The round-3 QA blocking finding was AC-1: after inserting
// an equation, typing immediately (no click) went to the surrounding document,
// not the equation, because DocumentView's post-edit caret restore reclaimed
// focus from the just-mounted equation surface. These regressions drive the
// exact unassisted-user path — insert, then type straight away — and assert the
// keystrokes build the equation and focus stays on its surface.

async function newDocument(page: Page) {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page.getByTestId("document-content")).toBeVisible();
}

const content = (page: Page) => page.getByTestId("document-content");
const linear = (page: Page) => page.getByTestId("equation-editor-linear");

test.describe("SPEC-056 equation insert focus", () => {
  test("Insert → Equation: typing immediately builds the equation (AC-1)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();
    await page.keyboard.type("Notes: ");

    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-equation-button").click();
    await expect(page.getByTestId("equation-editor-surface")).toBeVisible();

    // No click into the surface — AC-1 promises the caret is already inside it.
    await page.keyboard.type("a+b");

    // The engine received the keystrokes (the linear form is empty in the bug,
    // where the document swallowed them).
    await expect(linear(page)).toHaveText("a+b");
    // And focus is on the equation surface, not the document body.
    const focused = await page.evaluate(() =>
      document.activeElement?.getAttribute("data-testid"),
    );
    expect(focused).toBe("equation-editor-surface");
  });

  test("Alt+= inserts an inline equation that captures typing (AC-3)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();

    await page.keyboard.press("Alt+Equal");
    await expect(page.getByTestId("equation-editor-surface")).toBeVisible();

    await page.keyboard.type("x^2");
    await expect(linear(page)).toHaveText("x^2");
  });

  test("Insert → Display equation: typing immediately builds the equation (AC-2)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();

    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-display-equation-button").click();
    await expect(page.getByTestId("equation-editor-surface")).toBeVisible();

    await page.keyboard.type("y");
    await expect(linear(page)).toHaveText("y");
  });
});
