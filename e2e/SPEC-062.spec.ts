import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-062 — command palette. A keyboard-first search over every registered
// command (and the open document's styles), in one accessible combobox/listbox
// dialog shared by the editor and the library. Opened from Tools → Search
// commands or Ctrl/Cmd+Shift+K in the editor, and Ctrl/Cmd+K in the library.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("document-content")).toBeVisible();
}

test.describe("SPEC-062 command palette", () => {
  test("opens from Tools → Search commands, filters, and runs a command (AC-4, AC-9, AC-12, AC-27)", async ({ page }) => {
    await createDocument(page);

    await page.getByTestId("menu-title-tools").click();
    await page.getByTestId("tools-menu-command-palette").click();

    await expect(page.getByTestId("command-palette")).toBeVisible();
    const input = page.getByTestId("command-palette-input");
    await expect(input).toBeFocused();
    await expect(input).toHaveAttribute("aria-expanded", "true");
    await expect(input).toHaveAttribute("aria-controls", "command-palette-listbox");
    await expect(page.locator('#command-palette-listbox[role="listbox"]')).toBeVisible();

    await input.fill("word count");
    const option = page.locator('[data-command-id="tools.word-count"]');
    await expect(option).toHaveAttribute("role", "option");
    await expect(option).toBeVisible();

    // Enter runs the highlighted command — identical to the menu item.
    await page.keyboard.press("Enter");
    await expect(page.getByTestId("statistics-dialog")).toBeVisible();
  });

  test("opens in the editor from the registered app.commandPalette shortcut (AC-2, AC-35)", async ({ page }) => {
    // SPEC-025 owns Ctrl/Cmd+K (insert link); app.commandPalette registers on
    // Ctrl/Cmd+Shift+K, the closest free key, so opening stays keyboard-only.
    await createDocument(page);
    await page.getByTestId("document-content").click();

    await page.keyboard.press(`${MOD}+Shift+k`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await expect(page.getByTestId("command-palette-input")).toBeFocused();
  });

  test("Up/Down move the highlight and Escape closes (AC-13, AC-14)", async ({ page }) => {
    await createDocument(page);
    await page.getByTestId("menu-title-tools").click();
    await page.getByTestId("tools-menu-command-palette").click();

    const input = page.getByTestId("command-palette-input");
    await expect(input).toHaveAttribute("aria-activedescendant", "command-palette-option-0");
    await page.keyboard.press("ArrowDown");
    await expect(input).toHaveAttribute("aria-activedescendant", "command-palette-option-1");

    await page.keyboard.press("Escape");
    await expect(page.getByTestId("command-palette")).toBeHidden();
  });

  test("is available in the library via Ctrl/Cmd+K and runs a library command (AC-33, AC-35)", async ({ page }) => {
    await page.goto("/");
    await expect(page.getByTestId("empty-state")).toBeVisible();

    await page.keyboard.press(`${MOD}+k`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("new document");
    await expect(page.locator('[data-command-id="library.new-document"]')).toBeVisible();
    await page.keyboard.press("Enter");
    await expect(page).toHaveURL(/\/doc\/.+/);
  });
});
