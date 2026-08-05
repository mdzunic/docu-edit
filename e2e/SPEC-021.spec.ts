import { test, expect, type Page } from "@playwright/test";

// SPEC-021 — the consolidated keyboard-shortcut registry: the searchable
// "Keyboard shortcuts" dialog (Ctrl/Cmd+/), toolbar tooltips read from the
// registry, and shortcuts not firing while focus is outside the document
// (a dialog, an input). createDocument/content follow e2e/SPEC-017.spec.ts.

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

const content = (page: Page) => page.getByTestId("document-content");
const shortcutsDialog = (page: Page) => page.getByTestId("shortcuts-dialog");

async function seedDocument(page: Page, text = "Hello world"): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
}

test("Ctrl/Cmd+/ opens the searchable Keyboard shortcuts dialog, listing every registered shortcut by category (AC-6, AC-7)", async ({
  page,
}) => {
  await createDocument(page);

  await page.keyboard.press("Control+/");
  await expect(shortcutsDialog(page)).toBeVisible();
  await expect(page.getByTestId("shortcuts-search")).toBeFocused();

  await expect(page.getByTestId("shortcut-row-format.bold")).toContainText("Ctrl+B");
  await expect(page.getByTestId("shortcut-row-edit.undo")).toContainText("Ctrl+Z");
  await expect(shortcutsDialog(page)).toContainText("Formatting");
  await expect(shortcutsDialog(page)).toContainText("Navigation");

  await page.getByTestId("shortcuts-search").fill("bold");
  await expect(page.getByTestId("shortcut-row-format.bold")).toBeVisible();
  await expect(page.getByTestId("shortcut-row-edit.undo")).toBeHidden();
});

test("the dialog documents the browser shortcuts it deliberately overrides (AC-10) and the ones it never binds (AC-11)", async ({
  page,
}) => {
  await createDocument(page);

  await page.keyboard.press("Control+/");
  await expect(page.getByTestId("shortcut-row-format.bold")).toContainText("overrides browser default");
  await expect(page.getByTestId("shortcut-row-edit.undo")).not.toContainText("overrides browser default");
  await expect(page.getByTestId("shortcuts-reserved-note")).toContainText("Ctrl/Cmd+T");
  await expect(page.getByTestId("shortcuts-reserved-note")).toContainText("Ctrl/Cmd+W");
  await expect(page.getByTestId("shortcuts-reserved-note")).toContainText("Ctrl/Cmd+N");
});

test("the dialog is keyboard-navigable and closes with Escape (AC-8, AC-9)", async ({ page }) => {
  await createDocument(page);

  await page.keyboard.press("Control+/");
  await expect(shortcutsDialog(page)).toBeVisible();

  const firstRow = page.locator('[role="option"]').first();
  await expect(firstRow).toHaveClass(/active/);
  await page.keyboard.press("ArrowDown");
  await expect(firstRow).not.toHaveClass(/active/);

  await page.keyboard.press("Escape");
  await expect(shortcutsDialog(page)).toBeHidden();
});

test("toolbar tooltips show the shortcut read from the registry (AC-15, AC-16)", async ({ page }) => {
  await createDocument(page);

  await expect(page.getByTestId("format-bold")).toHaveAttribute("title", "Bold (Ctrl+B)");
  await expect(page.getByTestId("align-left")).toHaveAttribute("title", "Align left (Ctrl+L)");
  await expect(page.getByTestId("undo-button")).toHaveAttribute("title", "Undo (Ctrl+Z)");
});

test("registry shortcuts do not fire while focus is in a dialog's text input (AC-12, AC-14)", async ({ page }) => {
  await createDocument(page);
  await seedDocument(page);
  await content(page).click();
  await page.keyboard.press("Control+a");

  await page.getByTestId("line-spacing-button").click();
  await page.getByTestId("line-spacing-custom").click();
  await expect(page.getByTestId("custom-spacing-dialog")).toBeVisible();
  await expect(page.getByTestId("spacing-value")).toBeFocused();

  // Ctrl+B and Ctrl+/ target the dialog's input, not the document — neither
  // should act while it's focused.
  await page.keyboard.press("Control+b");
  await page.keyboard.press("Control+/");
  await expect(shortcutsDialog(page)).toBeHidden();

  await page.keyboard.press("Escape");
  await expect(page.getByTestId("custom-spacing-dialog")).toBeHidden();
  await expect(page.getByTestId("format-bold")).toHaveAttribute("aria-pressed", "false");
});

test("registry shortcuts do not fire while focus is in the find bar (AC-13)", async ({ page }) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press("Control+f");
  await expect(page.getByTestId("find-bar")).toBeVisible();
  await expect(page.getByTestId("find-query")).toBeFocused();

  await page.keyboard.press("Control+/");
  await expect(shortcutsDialog(page)).toBeHidden();
  await expect(page.getByTestId("find-bar")).toBeVisible();
});

test("Ctrl/Cmd+P prints instead of opening the browser's print dialog (AC-10)", async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { __printCalls: number }).__printCalls = 0;
    window.print = () => {
      (window as unknown as { __printCalls: number }).__printCalls += 1;
    };
  });
  await createDocument(page);
  await seedDocument(page);
  await content(page).click();

  await page.keyboard.press("Control+p");
  await expect
    .poll(() => page.evaluate(() => (window as unknown as { __printCalls: number }).__printCalls))
    .toBe(1);
});
