import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-062 — command palette. A keyboard-first search over every registered
// command (and the open document's styles), in one accessible combobox/listbox
// dialog shared by the editor and the library. Opened from Ctrl/Cmd+K anywhere
// in the editor (AC-1), from Tools → Search commands (AC-4), and from
// Ctrl/Cmd+K in the library (AC-33). Insert link moved to Ctrl/Cmd+Alt+K to
// free the chord (SPEC-025 regression covered below).

async function newDoc(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("document-content")).toBeVisible();
}

async function openViaTools(page: Page): Promise<void> {
  await page.getByTestId("menu-title-tools").click();
  await page.getByTestId("tools-menu-command-palette").click();
  await expect(page.getByTestId("command-palette")).toBeVisible();
}

test.describe("SPEC-062 command palette", () => {
  test("Ctrl/Cmd+K opens the palette anywhere in the editor, not the link dialog (AC-1, AC-2, AC-35)", async ({ page }) => {
    await newDoc(page);
    // Focus the document canvas — opening must work with the editor focused.
    await page.getByTestId("document-content").click();

    await page.keyboard.press(`${MOD}+k`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await expect(page.getByTestId("command-palette-input")).toBeFocused();
    // The previously-bound SPEC-025 link dialog must NOT be what Ctrl+K opens.
    await expect(page.getByTestId("link-dialog")).toBeHidden();
  });

  test("Insert link moved to Ctrl/Cmd+Alt+K and still opens (SPEC-025 regression)", async ({ page }) => {
    await newDoc(page);
    await page.getByTestId("document-content").click();
    await page.keyboard.press(`${MOD}+Alt+k`);
    await expect(page.getByTestId("link-dialog")).toBeVisible();
  });

  test("the shortcut is registered under View and shown in the reference dialog (AC-2, AC-3)", async ({ page }) => {
    await newDoc(page);
    await page.keyboard.press(`${MOD}+/`);
    await expect(page.getByTestId("shortcuts-dialog")).toBeVisible();
    await page.getByTestId("shortcuts-search").fill("Search commands");
    await expect(page.getByTestId("shortcuts-dialog")).toContainText(/view/i);
    await expect(page.getByTestId("shortcuts-list")).toContainText("Search commands");
  });

  test("Tools → Search commands opens the palette (AC-4)", async ({ page }) => {
    await newDoc(page);
    await openViaTools(page);
    await expect(page.getByTestId("command-palette-input")).toBeFocused();
  });

  test("rows show menu path, icon and shortcut label (AC-6, AC-7, AC-8)", async ({ page }) => {
    await newDoc(page);
    await openViaTools(page);
    const input = page.getByTestId("command-palette-input");

    await input.fill("bold");
    const bold = page.locator('[data-command-id="format.bold"]');
    await expect(bold.locator(".palette-path")).toContainText("Format");
    await expect(bold.locator(".palette-shortcut")).toContainText("B");

    await input.fill("text colour");
    await expect(
      page.locator('[data-command-id="format.text-color"] .palette-icon svg'),
    ).toBeVisible();
  });

  test("subsequence filter, ranking and match highlight (AC-9, AC-10, AC-11)", async ({ page }) => {
    await newDoc(page);
    await openViaTools(page);
    await page.getByTestId("command-palette-input").fill("ftnt");
    const first = page.getByTestId("command-palette-option").first();
    await expect(first).toContainText("Footnote");
    await expect(first.locator("mark.palette-hit").first()).toBeVisible();
  });

  test("arrows move the highlight, Enter runs the highlighted command (AC-12, AC-13)", async ({ page }) => {
    await newDoc(page);
    await openViaTools(page);
    const input = page.getByTestId("command-palette-input");

    await input.fill("format");
    await expect(input).toHaveAttribute("aria-activedescendant", "command-palette-option-0");
    await page.keyboard.press("ArrowDown");
    await expect(input).toHaveAttribute("aria-activedescendant", "command-palette-option-1");
    await page.keyboard.press("ArrowUp");
    await expect(input).toHaveAttribute("aria-activedescendant", "command-palette-option-0");

    await input.fill("word count");
    await page.keyboard.press("Enter");
    await expect(page.getByTestId("statistics-dialog")).toBeVisible();
  });

  test("Escape closes; reopening keeps selection-dependent commands enabled (AC-14, AC-15, AC-26)", async ({ page }) => {
    await newDoc(page);
    const content = page.getByTestId("document-content");
    await content.click();
    await content.dblclick(); // select a word

    await page.keyboard.press(`${MOD}+k`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("command-palette")).toBeHidden();

    await page.keyboard.press(`${MOD}+k`);
    await page.getByTestId("command-palette-input").fill("bold");
    await expect(page.locator('[data-command-id="format.bold"]')).not.toHaveAttribute(
      "aria-disabled",
      "true",
    );
  });

  test("inapplicable commands are listed and disabled with a reason (AC-18, AC-19)", async ({ page }) => {
    await newDoc(page);
    await openViaTools(page);
    await page.getByTestId("command-palette-input").fill("undo");
    const undo = page.locator('[data-command-id="edit.undo"]');
    await expect(undo).toBeVisible();
    await expect(undo).toHaveAttribute("aria-disabled", "true");
    await expect(undo.locator(".palette-reason")).toBeVisible();
  });

  test("empty query lists recents first, persisted under docu-edit:recent-commands (AC-20, AC-21)", async ({ page }) => {
    await newDoc(page);
    await openViaTools(page);
    await page.getByTestId("command-palette-input").fill("word count");
    await page.keyboard.press("Enter");
    await expect(page.getByTestId("statistics-dialog")).toBeVisible();
    await page.keyboard.press("Escape");

    const stored = await page.evaluate(() =>
      localStorage.getItem("docu-edit:recent-commands"),
    );
    expect(stored).toContain("tools.word-count");

    await openViaTools(page);
    await expect(page.getByTestId("command-palette-option").first()).toHaveAttribute(
      "data-command-id",
      "tools.word-count",
    );
  });

  test("styles are searchable and offered as Apply-style rows (AC-22, AC-23)", async ({ page }) => {
    await newDoc(page);
    await page.getByTestId("document-content").click();
    await openViaTools(page);
    await page.getByTestId("command-palette-input").fill("Heading 1");
    const styleRow = page
      .getByTestId("command-palette-option")
      .filter({ hasText: "Apply style" })
      .first();
    await expect(styleRow).toBeVisible();
    await styleRow.click();
    await expect(page.getByTestId("command-palette")).toBeHidden();
  });

  test("clicking outside dismisses the palette (AC-25)", async ({ page }) => {
    await newDoc(page);
    await openViaTools(page);
    await page.mouse.click(5, 5);
    await expect(page.getByTestId("command-palette")).toBeHidden();
  });

  test("accessible combobox/listbox pattern with a live result count (AC-27, AC-28, AC-29, AC-30)", async ({ page }) => {
    await newDoc(page);
    await openViaTools(page);
    await expect(page.getByTestId("command-palette").locator('[role="dialog"]')).toBeVisible();

    const input = page.getByTestId("command-palette-input");
    await expect(input).toHaveAttribute("role", "combobox");
    await expect(input).toHaveAttribute("aria-expanded", "true");
    await expect(input).toHaveAttribute("aria-controls", "command-palette-listbox");
    await expect(input).toHaveAttribute("aria-activedescendant", /command-palette-option-\d+/);
    await expect(page.locator('#command-palette-listbox[role="listbox"]')).toBeVisible();
    await expect(page.getByTestId("command-palette-option").first()).toHaveAttribute("role", "option");

    const live = page.getByTestId("command-palette-count");
    await expect(live).toHaveAttribute("aria-live", "polite");
    await input.fill("word count");
    await expect(live).toContainText(/result/);
  });

  test("the same palette works in the library, keyboard-only (AC-33, AC-34, AC-35)", async ({ page }) => {
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
