import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-068 AC-47 — the compact (390×844, touch) project runs the smoke path
// plus open/edit/comment/download: open a document, read it at fit-width, place
// the caret, type, apply bold from the essential band, open the menu sheet and
// reach a Tools command, add a comment, and download. Every desktop suite keeps
// running unchanged at the desktop viewport (AC-48) — these live only here.

const content = (page: Page) => page.getByTestId("document-content");

async function newDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page)).toBeVisible();
}

test("the library is usable and collapses its header actions at compact (AC-37/38/39)", async ({ page }) => {
  await page.goto("/");
  // One primary action + the search box stay visible; the rest collapse behind
  // the overflow menu.
  await expect(page.getByTestId("new-document")).toBeVisible();
  await expect(page.getByTestId("library-search")).toBeVisible();
  await expect(page.getByTestId("library-actions-overflow")).toBeVisible();
  await expect(page.getByTestId("new-from-template")).toHaveCount(0);
  await page.getByTestId("library-actions-overflow").click();
  await expect(page.getByTestId("new-from-template")).toBeVisible();
});

test("compact editor: fit-width web layout, essential toolbar, menu sheet, edit, comment, download", async ({
  page,
}) => {
  await newDocument(page);

  // Reads at fit-width in Web layout (AC-35).
  await expect(page.getByTestId("app-shell")).toHaveClass(/is-web-layout/);

  // The menu bar has collapsed to a single Menu button (AC-7); the docked bar
  // is gone.
  await expect(page.getByTestId("menu-button")).toBeVisible();
  await expect(page.getByTestId("menu-bar")).toHaveCount(1); // the compact wrapper, not a role=menubar
  await expect(page.locator('[role="menubar"]')).toHaveCount(0);

  // The essential toolbar band: undo/redo, style picker, bold, italic, overflow —
  // and NOT the font family select (which moved to the overflow). AC-9.
  await expect(page.getByTestId("format-bold")).toBeVisible();
  await expect(page.getByTestId("format-italic")).toBeVisible();
  await expect(page.getByTestId("toolbar-overflow-button")).toBeVisible();
  // Underline lives in the (closed, hidden) overflow menu, not the essential
  // band; opening the overflow reveals it — nothing was silently dropped.
  await expect(page.getByTestId("format-underline")).toBeHidden();
  await page.getByTestId("toolbar-overflow-button").click();
  await expect(page.getByTestId("format-underline")).toBeVisible();
  await page.getByTestId("toolbar-overflow-button").click();

  // Place the caret and type (AC-25 tap places caret).
  await content(page).click();
  await page.keyboard.type("Hello compact world");
  await expect(content(page)).toContainText("Hello compact world");

  // Apply bold from the essential band to a selection.
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTestId("format-bold").click();
  await expect(page.getByTestId("format-bold")).toHaveAttribute("aria-pressed", "true");

  // Open the menu sheet and reach a Tools command (AC-8, AC-10).
  await page.getByTestId("menu-button").click();
  await expect(page.getByTestId("menu-sheet")).toBeVisible();
  await expect(page.getByTestId("menu-sheet-section-tools")).toBeVisible();
  await page.getByTestId("tools-menu-word-count").click();
  await expect(page.getByTestId("statistics-dialog")).toBeVisible();
  // The sheet closed when the command ran.
  await expect(page.getByTestId("menu-sheet")).toHaveCount(0);
  // Close the statistics dialog.
  await page.keyboard.press("Escape");

  // Add a comment over a word (AC-47 comment step).
  await content(page).click();
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press(`${MOD}+Alt+m`);
  const composer = page.getByTestId("comments-pane").getByTestId("comment-composer-input");
  await expect(composer).toBeVisible();
  await composer.fill("Looks good on mobile");
  await page.getByTestId("comments-pane").getByTestId("comment-composer-submit").click();
  await expect(page.locator('[data-testid^="comment-thread-"]')).toHaveCount(1);

  // Download the document as .docx (AC-47 download step).
  await expect(page.getByTestId("download-menu")).toBeVisible();
  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toMatch(/\.docx$/);
});

test("never disables browser zoom (AC-45/46)", async ({ page }) => {
  await page.goto("/");
  const viewport = await page.locator('meta[name="viewport"]').getAttribute("content");
  expect(viewport ?? "").not.toContain("user-scalable=no");
  expect(viewport ?? "").not.toMatch(/maximum-scale/);
});
