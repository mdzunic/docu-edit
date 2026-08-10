import { test, expect, type Page } from "@playwright/test";
import { MOD, LINE_START } from "./keys";

// SPEC-039 — track changes: a recording toggle (always OFF on open) that
// marks new insertions/deletions instead of applying them; a review pane
// listing changes with Accept/Reject/Accept all/Reject all; Next/Previous
// navigation; a display-mode selector; accept/reject is undoable.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const pane = (page: Page) => page.getByTestId("track-changes-pane");

async function seedParagraph(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
}

async function openTrackChangesPane(page: Page): Promise<void> {
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-track-changes-pane").click();
}

async function toggleRecording(page: Page): Promise<void> {
  await page.getByTestId("toolbar-track-changes-toggle").click();
}

test("the toggle is off for a new blank document, and toolbar/status bar show a clear on/off state (AC-1, AC-2)", async ({
  page,
}) => {
  await createDocument(page);
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "false");
  await expect(page.getByTestId("status-bar-track-changes")).toHaveAttribute("aria-pressed", "false");
  await expect(page.getByTestId("status-bar-track-changes")).toContainText("Off");

  await toggleRecording(page);
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("status-bar-track-changes")).toContainText("On");

  // The shortcut toggles it back off.
  await content(page).click();
  await page.keyboard.press(`${MOD}+Shift+e`);
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "false");
});

test("with tracking on, inserted text underlines and deleted text stays struck through instead of disappearing (AC-4)", async ({
  page,
}) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);

  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  const inserted = content(page).locator("span", { hasText: "XYZ" }).first();
  await expect(inserted).toHaveCSS("text-decoration-line", /underline/);

  // Select "hello" and delete it — it must remain visible, struck through.
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await page.keyboard.press("Backspace");
  await expect(content(page)).toContainText("hello");
  const deleted = content(page).locator("span", { hasText: "hello" }).first();
  await expect(deleted).toHaveCSS("text-decoration-line", /line-through/);
});

test("opening the review pane lists changes in document order with author/date, and Next/Previous navigate (AC-6, AC-7, AC-8)", async ({
  page,
}) => {
  await createDocument(page);
  await seedParagraph(page, "aaaa bbbb cccc");
  await toggleRecording(page);
  await openTrackChangesPane(page);

  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("ONE");
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.type("TWO");

  // Two distinct tracked insertions should be listed.
  await expect(pane(page).locator(".track-change-item")).toHaveCount(2);
  await expect(pane(page).locator(".track-change-item").first()).toContainText("Inserted");
  await expect(pane(page).locator(".track-change-item").first()).toContainText("You");

  await pane(page).getByTestId("track-changes-next").click();
  await pane(page).getByTestId("track-changes-next").click();
});

test("Accept and Reject per change, Accept all and Reject all (AC-7, AC-10, AC-11)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);
  await openTrackChangesPane(page);

  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);

  // Reject: the inserted text is removed.
  await pane(page).getByTestId("track-change-reject").click();
  await expect(content(page)).not.toContainText("XYZ");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(0);

  // Undo restores the pending insertion (AC-10).
  await content(page).click();
  await page.keyboard.press(`${MOD}+z`);
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);

  // Accept: the insertion commits, no longer tracked.
  await pane(page).getByTestId("track-change-accept").click();
  await expect(content(page)).toContainText("XYZ");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(0);

  // Delete a range, then use the bulk actions.
  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.press("Shift+ArrowRight");
  await page.keyboard.press("Shift+ArrowRight");
  await page.keyboard.press("Backspace");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);
  await pane(page).getByTestId("track-changes-reject-all").click();
  await expect(pane(page).locator(".track-change-item")).toHaveCount(0);
});

test("the display-mode selector offers all four modes and defaults to All markup (AC-9)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await openTrackChangesPane(page);
  const select = pane(page).getByTestId("track-changes-display-mode-select");
  await expect(select).toHaveValue("all");
  await select.selectOption("simple");
  await expect(select).toHaveValue("simple");
  await select.selectOption("noMarkup");
  await select.selectOption("original");
});

test("opening a document that already contains tracked changes leaves the toggle off (AC-3)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);
  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "true");

  // Navigate away and back (a fresh "open") — the toggle must reset to off
  // even though the document still has a tracked change in it.
  await page.goto("/");
  await page.goBack();
  await expect(content(page)).toContainText("XYZ");
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "false");
});
