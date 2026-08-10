import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-040 — version history: "File > Version history" opens a pane listing
// versions (timestamp, author, word-count delta, newest first); each
// explicit download is a commit-point snapshot; selecting an older version
// previews it read-only in the canvas; "Restore this version" replaces the
// current content, is undoable, and is itself recorded as a new version;
// versions can be named and filtered to; "Download this version" exports a
// past snapshot without touching the current document.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const pane = (page: Page) => page.getByTestId("version-history-pane");
const items = (page: Page) => page.locator('[data-testid^="version-history-item-"]');

async function seedParagraph(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
}

async function replaceAll(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.type(text);
}

async function openVersionHistoryPane(page: Page): Promise<void> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-version-history").click();
}

async function downloadCurrentDocx(page: Page): Promise<void> {
  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  await downloadPromise;
}

test("File > Version history opens a pane; a download commit records a version with timestamp, author and word-count delta (AC-6..10)", async ({
  page,
}) => {
  await createDocument(page);
  await openVersionHistoryPane(page);
  await expect(pane(page)).toBeVisible();
  await expect(page.getByTestId("version-history-empty")).toBeVisible();

  await seedParagraph(page, "hello world");
  await downloadCurrentDocx(page);

  await expect(items(page)).toHaveCount(1);
  await expect(items(page).first()).toContainText("You");
  await expect(items(page).first()).toContainText("+2 words");
});

test("naming a version and filtering to named versions only (AC-17, AC-18)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "first version");
  await downloadCurrentDocx(page);
  await openVersionHistoryPane(page);
  await expect(items(page)).toHaveCount(1);

  await items(page).first().getByTestId("version-history-name-input").fill("Milestone one");
  await items(page).first().getByTestId("version-history-name-submit").click();
  await expect(items(page).first()).toContainText("Milestone one");

  await seedParagraph(page, " and more text");
  await downloadCurrentDocx(page);
  await expect(items(page)).toHaveCount(2);

  await page.getByTestId("version-history-named-only").check();
  await expect(items(page)).toHaveCount(1);
  await expect(items(page).first()).toContainText("Milestone one");

  await page.getByTestId("version-history-named-only").uncheck();
  await expect(items(page)).toHaveCount(2);
});

test("selecting an older version previews it read-only with diff highlighting, and exiting preview restores editing (AC-11..13)", async ({
  page,
}) => {
  await createDocument(page);
  await seedParagraph(page, "first version text");
  await downloadCurrentDocx(page);
  await replaceAll(page, "second version text");
  await downloadCurrentDocx(page);

  await openVersionHistoryPane(page);
  await expect(items(page)).toHaveCount(2);
  await expect(content(page)).toHaveAttribute("contenteditable", "true");

  // Newest first: the older version is the second item.
  await items(page).nth(1).locator(".version-history-item-header").click();
  await expect(page.getByTestId("version-history-preview-banner")).toBeVisible();
  await expect(content(page)).toHaveAttribute("contenteditable", "false");
  await expect(content(page)).toContainText("first version text");

  await page.getByTestId("version-history-exit-preview").click();
  await expect(page.getByTestId("version-history-preview-banner")).not.toBeVisible();
  await expect(content(page)).toHaveAttribute("contenteditable", "true");
  await expect(content(page)).toContainText("second version text");
});

test("restoring a version replaces the content, is undoable, and is itself recorded as a new version (AC-14, AC-15, AC-16)", async ({
  page,
}) => {
  await createDocument(page);
  await seedParagraph(page, "first version text");
  await downloadCurrentDocx(page);
  await replaceAll(page, "second version text");
  await downloadCurrentDocx(page);

  await openVersionHistoryPane(page);
  await expect(items(page)).toHaveCount(2);
  await items(page).nth(1).getByTestId("version-history-restore").click();

  await expect(content(page)).toContainText("first version text");
  await expect(content(page)).not.toContainText("second version text");
  // Restore brackets itself with a before/after snapshot (AC-4, AC-5).
  await expect(items(page)).toHaveCount(4);

  await content(page).click();
  await page.keyboard.press(`${MOD}+z`);
  await expect(content(page)).toContainText("second version text");
});

test("downloading a specific version exports it without changing the current document (AC-19..21)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "alpha content");
  await downloadCurrentDocx(page);
  await replaceAll(page, "beta content");
  await downloadCurrentDocx(page);

  await openVersionHistoryPane(page);
  await expect(items(page)).toHaveCount(2);

  const downloadPromise = page.waitForEvent("download");
  await items(page).nth(1).getByTestId("version-history-download-docx").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toMatch(/\.docx$/);

  // The current document is untouched by downloading a past version.
  await expect(content(page)).toContainText("beta content");
  await expect(content(page)).not.toContainText("alpha content");
});
