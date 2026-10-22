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
  // Seed a document so the populated library (with its header row) renders,
  // rather than the empty state's own action buttons.
  await newDocument(page);
  await page.goto("/");
  await expect(page.getByTestId("library")).toBeVisible();
  // One primary action + the search box stay visible; the rest collapse behind
  // the overflow menu.
  await expect(page.getByTestId("new-document")).toBeVisible();
  await expect(page.getByTestId("library-search")).toBeVisible();
  await expect(page.getByTestId("library-actions-overflow")).toBeVisible();
  await expect(page.getByTestId("new-from-template")).toHaveCount(0);
  await page.getByTestId("library-actions-overflow").click();
  await expect(page.getByTestId("new-from-template")).toBeVisible();
});

test("the library list fits the compact width and keeps every row action reachable", async ({ page }) => {
  // The regression this guards: the list is the DEFAULT view, and it had no
  // width breakpoint at all while the grid did. Seven columns could not share
  // a phone's width, so the table overflowed a parent that does not scroll —
  // the row's overflow menu sat off-screen entirely, and the title column,
  // squeezed to ~66px, broke mid-word. The grid view was fine, which is
  // exactly why nothing caught it.
  await newDocument(page);
  await page.goto("/");
  const row = page.getByTestId("library-row").first();
  await expect(row).toBeVisible();

  // Nothing may push the page sideways.
  const overflow = await page.evaluate(
    () => document.documentElement.scrollWidth - window.innerWidth,
  );
  expect(overflow).toBeLessThanOrEqual(0);

  // Every cell stays inside the viewport — the row menu above all, since a
  // clipped one is unreachable rather than merely awkward.
  const viewport = page.viewportSize()!.width;
  for (const testId of ["row-title", "row-modified", "row-format", "row-size"]) {
    const box = (await row.getByTestId(testId).boundingBox())!;
    expect(box.x + box.width, `${testId} runs past the viewport`).toBeLessThanOrEqual(viewport);
  }
  const menuButton = row.getByRole("button", { name: /^Actions for / });
  const menuBox = (await menuButton.boundingBox())!;
  expect(menuBox.x + menuBox.width).toBeLessThanOrEqual(viewport);

  // The title gets a readable share of the row rather than a shredded column.
  const titleBox = (await row.getByTestId("row-title").boundingBox())!;
  expect(titleBox.width).toBeGreaterThan(viewport / 3);

  // And the menu it opens is on-screen too, with its actions.
  await menuButton.click();
  const rename = page.getByRole("menuitem", { name: "Rename" });
  await expect(rename).toBeVisible();
  const renameBox = (await rename.boundingBox())!;
  expect(renameBox.x).toBeGreaterThanOrEqual(0);
  expect(renameBox.x + renameBox.width).toBeLessThanOrEqual(viewport);
});

test("the app bar keeps the document's name legible at compact", async ({ page }) => {
  // The regression this guards: `overflow: hidden` zeroes a flex item's
  // automatic minimum, so the title had nothing holding it open. A crowded
  // bar shrank it to exactly the pencil's 16px and the name rendered at zero
  // width — not even an ellipsis, just a context-free pencil icon where the
  // document's identity should be.
  await newDocument(page);
  const title = page.getByTestId("document-title");
  await expect(title).toBeVisible();

  const text = title.locator(".document-title-text");
  const box = (await text.boundingBox())!;
  expect(box.width, "the title collapsed to nothing").toBeGreaterThan(40);

  // The full name stays in the DOM even when the visible text is ellipsised,
  // so it is still what assistive technology and a rename read.
  await expect(text).toHaveText("Untitled document");

  // It also has to beat the thing that was taking its space: at compact the
  // resting save status is a dot, not ~110px of "All changes saved".
  const save = page.getByTestId("save-state");
  await expect(save).toHaveText("All changes saved");
  const saveBox = (await save.boundingBox())!;
  expect(saveBox.width).toBeLessThan(box.width);

  // And none of it pushes the page sideways.
  const overflow = await page.evaluate(
    () => document.documentElement.scrollWidth - window.innerWidth,
  );
  expect(overflow).toBeLessThanOrEqual(0);
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

  // Close the comments overlay so it no longer covers the app bar (AC-4: the
  // panel overlays at compact).
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("comments-pane")).toHaveCount(0);

  // Download the document as .docx (AC-47 download step) — Download lives in
  // File -> Export (the app-bar button was removed), reached through the
  // sheet; the Export submenu carries data-menu-keep-open, so the sheet
  // stays up while the flyout is used.
  await page.getByTestId("menu-button").click();
  await expect(page.getByTestId("menu-sheet")).toBeVisible();
  await page.getByTestId("file-menu-export").click();
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
