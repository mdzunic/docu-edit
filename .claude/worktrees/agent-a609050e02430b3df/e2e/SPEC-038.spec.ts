import { test, expect, type Page } from "@playwright/test";
import { MOD, LINE_START } from "./keys";

// SPEC-038 — comments: select text and Ctrl/Cmd+Alt+M highlights the range
// and opens a composer in the right-hand margin pane; submitting adds the
// comment with an author and timestamp; threads support replies; a thread
// can be resolved (hiding its highlight, moving it to "Resolved") and
// reopened; a comment's author can edit/delete their own comments; deleting
// a thread removes its anchor; the comments-only view lists every thread in
// document order with open/resolved and author filters; comment text always
// renders as plain text.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const pane = (page: Page) => page.getByTestId("comments-pane");

/** Selects a substring of the current (single-paragraph) line by character
 * offsets, via Home + arrow keys — the same technique SPEC-012's e2e specs
 * use for precise selections. */
async function selectRange(page: Page, from: number, to: number): Promise<void> {
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < from; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.down("Shift");
  for (let i = from; i < to; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.up("Shift");
}

async function seedParagraph(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
}

/** Selects `word` (the first occurrence in `text`) and opens the comment
 * composer over it via the Ctrl/Cmd+Alt+M shortcut. */
async function commentOn(page: Page, text: string, word: string, commentText: string): Promise<void> {
  const from = text.indexOf(word);
  // Focus may currently sit in the pane (an identity input, a filter select)
  // rather than the document — the keyboard selection below needs it back.
  await content(page).click();
  await selectRange(page, from, from + word.length);
  await page.keyboard.press(`${MOD}+Alt+m`);
  await pane(page).getByTestId("comment-composer-input").fill(commentText);
  await pane(page).getByTestId("comment-composer-submit").click();
}

function threadCards(page: Page) {
  return pane(page).locator('[data-testid^="comment-thread-"]');
}

function highlightRects(page: Page) {
  return page.locator('[data-testid^="comment-highlight-"]');
}

test("selecting text and Ctrl/Cmd+Alt+M highlights the range and opens the composer in the margin pane (AC-1, AC-2)", async ({
  page,
}) => {
  await createDocument(page);
  const text = "Please review this paragraph carefully.";
  await seedParagraph(page, text);
  const from = text.indexOf("paragraph");
  await selectRange(page, from, from + "paragraph".length);
  await page.keyboard.press(`${MOD}+Alt+m`);
  await expect(pane(page)).toBeVisible();
  await expect(pane(page).getByTestId("comment-composer")).toBeVisible();
  // The selected range stays highlighted (the browser's own selection
  // rendering — nothing here collapses it) while the composer is open.
  expect(await page.evaluate(() => window.getSelection()?.toString())).toBe("paragraph");
  // Submitting turns the selection into a permanently anchored highlight.
  await pane(page).getByTestId("comment-composer-input").fill("Please expand");
  await pane(page).getByTestId("comment-composer-submit").click();
  await expect(highlightRects(page)).toHaveCount(1);
});

test("submitting adds the comment with the settable author identity (default 'You') and a timestamp (AC-3, AC-4)", async ({
  page,
}) => {
  await createDocument(page);
  const text = "This sentence needs a citation.";
  await seedParagraph(page, text);
  await commentOn(page, text, "citation", "Please add a source");
  const card = threadCards(page).filter({ hasText: "Please add a source" });
  await expect(card).toBeVisible();
  await expect(card).toContainText("You");
});

test("a settable identity changes the author of new comments and replies (AC-4)", async ({ page }) => {
  await createDocument(page);
  const text = "One more pass on this section.";
  await seedParagraph(page, text);
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-comments-pane").click();
  await pane(page).getByTestId("comment-author-name-input").fill("Ada Lovelace");
  await commentOn(page, text, "section", "Looks good");
  const card = threadCards(page).filter({ hasText: "Looks good" });
  await expect(card).toContainText("Ada Lovelace");
});

test("comment threads support replies, anchored beside the highlighted range (AC-5, AC-6)", async ({ page }) => {
  await createDocument(page);
  const text = "The introduction is too long.";
  await seedParagraph(page, text);
  await commentOn(page, text, "introduction", "Trim this down");
  const card = threadCards(page).filter({ hasText: "Trim this down" });
  await card.getByTestId("comment-reply-input").fill("Agreed, will shorten");
  await card.getByRole("button", { name: "Reply" }).click();
  await expect(card).toContainText("Agreed, will shorten");
});

test("clicking a highlighted range focuses its thread, and clicking a thread scrolls to and highlights its range (AC-7, AC-8)", async ({
  page,
}) => {
  await createDocument(page);
  const text = "Check the numbers in this table.";
  await seedParagraph(page, text);
  await commentOn(page, text, "numbers", "Verify these");
  const card = threadCards(page).filter({ hasText: "Verify these" });
  const highlight = highlightRects(page).first();

  // Body -> pane.
  await highlight.click();
  await expect(card).toHaveClass(/is-active/);

  // Pane -> body: clicking the (now inactive-looking) card re-affirms focus
  // and moves the caret to the anchor.
  await content(page).click(); // deselect / move focus away first
  await card.click();
  await expect(card).toHaveClass(/is-active/);
});

test("resolving a thread removes its highlight and moves it to Resolved; it can be reopened (AC-9..12)", async ({
  page,
}) => {
  await createDocument(page);
  const text = "This paragraph is unclear.";
  await seedParagraph(page, text);
  await commentOn(page, text, "unclear", "Please clarify");
  const card = threadCards(page).filter({ hasText: "Please clarify" });
  await expect(highlightRects(page)).toHaveCount(1);

  await card.getByRole("button", { name: "Resolve" }).click();
  await expect(highlightRects(page)).toHaveCount(0);
  const resolvedCard = threadCards(page).filter({ hasText: "Please clarify" });
  await expect(resolvedCard).toHaveClass(/is-resolved/);

  await resolvedCard.getByRole("button", { name: "Reopen" }).click();
  await expect(highlightRects(page)).toHaveCount(1);
});

test("the comment's author can edit or delete their own comments (AC-13, AC-14)", async ({ page }) => {
  await createDocument(page);
  const text = "The summary needs work.";
  await seedParagraph(page, text);
  await commentOn(page, text, "summary", "Original comment text");
  // A stable handle: once editing starts the text moves into a <textarea>
  // VALUE, which a `hasText` filter (matching on textContent) can no longer
  // see, so re-evaluating that filter after this point would find nothing.
  const testId = await threadCards(page).filter({ hasText: "Original comment text" }).getAttribute("data-testid");
  const card = page.getByTestId(testId!);

  await card.getByRole("button", { name: "Edit" }).click();
  await card.locator(".comment-edit-input").fill("Edited comment text");
  await card.getByRole("button", { name: "Save" }).click();
  await expect(card).toContainText("Edited comment text");
  await expect(card).not.toContainText("Original comment text");

  await card.getByRole("button", { name: "Delete", exact: true }).first().click();
  await expect(threadCards(page)).toHaveCount(0);
});

test("deleting a thread removes its anchor (AC-15)", async ({ page }) => {
  await createDocument(page);
  const text = "Remove this whole comment thread please.";
  await seedParagraph(page, text);
  await commentOn(page, text, "thread", "To be deleted");
  await expect(highlightRects(page)).toHaveCount(1);
  const card = threadCards(page).filter({ hasText: "To be deleted" });
  await card.getByRole("button", { name: "Delete", exact: true }).last().click();
  await expect(threadCards(page)).toHaveCount(0);
  await expect(highlightRects(page)).toHaveCount(0);
});

test("editing text inside a commented range keeps the comment anchored; deleting the whole range deletes the comment with an undoable warning (AC-16, AC-17)", async ({
  page,
}) => {
  await createDocument(page);
  const text = "Keep an eye on this phrase over here.";
  await seedParagraph(page, text);
  await commentOn(page, text, "phrase", "Watching this");
  await expect(highlightRects(page)).toHaveCount(1);

  // Type inside the commented range — the highlight survives.
  const phraseStart = text.indexOf("phrase");
  await selectRange(page, phraseStart + 2, phraseStart + 2);
  await page.keyboard.type("XY");
  await expect(highlightRects(page)).toHaveCount(1);

  // Now delete the entire (now-extended) commented range.
  const updatedFrom = phraseStart;
  const updatedTo = phraseStart + "phrase".length + 2;
  await selectRange(page, updatedFrom, updatedTo);
  await page.keyboard.press("Backspace");
  await expect(highlightRects(page)).toHaveCount(0);
  await expect(page.getByTestId("toast")).toContainText(/undo/i);
});

test("the comments-only view lists every thread in document order with open/resolved and author filters (AC-22..24)", async ({
  page,
}) => {
  await createDocument(page);
  const text = "Alpha bravo charlie delta echo foxtrot.";
  await seedParagraph(page, text);
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-comments-pane").click();
  await pane(page).getByTestId("comment-author-name-input").fill("Ada");
  await commentOn(page, text, "bravo", "Second comment target first in text");
  await commentOn(page, text, "delta", "First comment target second in text");

  await pane(page).getByTestId("comments-only-toggle").click();
  const cards = threadCards(page);
  await expect(cards).toHaveCount(2);
  // Document order: "bravo" (earlier in the text) before "delta".
  await expect(cards.nth(0)).toContainText("Second comment target first in text");
  await expect(cards.nth(1)).toContainText("First comment target second in text");

  // Resolve one, then filter to open only.
  await cards.nth(0).getByRole("button", { name: "Resolve" }).click();
  await pane(page).getByTestId("comment-filter-select").selectOption("open");
  await expect(threadCards(page)).toHaveCount(1);
  await expect(threadCards(page)).toContainText("First comment target second in text");

  await pane(page).getByTestId("comment-filter-select").selectOption("resolved");
  await expect(threadCards(page)).toHaveCount(1);
  await expect(threadCards(page)).toContainText("Second comment target first in text");

  await pane(page).getByTestId("comment-filter-select").selectOption("all");
  await expect(threadCards(page)).toHaveCount(2);
});

test("comment text renders as plain text, never as markup (AC-25, AC-26)", async ({ page }) => {
  await createDocument(page);
  const text = "Sanitize this input carefully.";
  await seedParagraph(page, text);
  await commentOn(page, text, "input", "<script>window.__pwned = true</script>");
  const card = threadCards(page).filter({ hasText: "script" });
  await expect(card).toContainText("<script>window.__pwned = true</script>");
  expect(await page.locator("script", { hasText: "__pwned" }).count()).toBe(0);
  expect(await page.evaluate(() => (window as unknown as { __pwned?: boolean }).__pwned)).toBeUndefined();
});

test("the View menu toggles the comments pane and the comments-only view", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "Just some text.");
  await expect(pane(page)).toHaveCount(0);
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-comments-pane").click();
  await expect(pane(page)).toBeVisible();
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-comments-only").click();
  await expect(pane(page).getByTestId("comments-only-toggle")).toHaveAttribute("aria-pressed", "true");
});

// The comments dock (right-hand margin pane) and the Download menu both
// occupy the top-right of the canvas; adding a comment auto-opens the dock
// and leaves it open, which is exactly the state a real user is in when they
// then reach for Download (AC-18, AC-19, AC-20 all round-trip through it).
// A stacking-order regression here doesn't fail any existing test (the menu
// still opens, its items still exist in the DOM) — only a click actually
// landing on the intended item catches it.
test("the Download menu stays clickable with the comments pane open (AC-18, AC-19)", async ({ page }) => {
  await createDocument(page);
  const text = "One more pass on this section.";
  await seedParagraph(page, text);
  await commentOn(page, text, "section", "Looks good");
  await expect(pane(page)).toBeVisible();

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const docxDownload = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  await docxDownload;

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const odtDownload = page.waitForEvent("download");
  await page.getByTestId("download-odt").click();
  await odtDownload;
});
