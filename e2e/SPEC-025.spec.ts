import { test, expect, type Page } from "@playwright/test";
import { LINE_START, MOD } from "./keys";

// SPEC-025 — hyperlinks and bookmarks: the Ctrl/Cmd+K link dialog, autolink
// on space/Enter with single-step undo, the link chip (open/edit/copy/
// remove), dangerous-scheme rejection, heading/bookmark link targets, the
// bookmark manager, and .docx/.odt round-trips through download+reopen.
// createDocument/content follow e2e/SPEC-021.spec.ts; the synthetic paste
// dispatch follows e2e/SPEC-016.spec.ts.


async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const links = (page: Page) => content(page).locator("a[data-doc-link]");
const linkDialog = (page: Page) => page.getByTestId("link-dialog");
const linkChip = (page: Page) => page.getByTestId("link-chip");

/** Selects [from, to) of the first paragraph with the keyboard. */
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

/** Dispatches a paste ClipboardEvent carrying the given flavours. */
async function dispatchPaste(page: Page, flavours: Record<string, string>): Promise<void> {
  await page.evaluate((data) => {
    const transfer = new DataTransfer();
    for (const [type, value] of Object.entries(data)) {
      transfer.setData(type, value);
    }
    const target = document.querySelector('[data-testid="document-content"]');
    target?.dispatchEvent(
      new ClipboardEvent("paste", {
        clipboardData: transfer,
        bubbles: true,
        cancelable: true,
      }),
    );
  }, flavours);
}

/** Creates a link over the whole first-paragraph text via the dialog. */
async function linkWholeText(page: Page, text: string, url: string): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press(`${MOD}+k`);
  await expect(linkDialog(page)).toBeVisible();
  await page.getByTestId("link-url-input").fill(url);
  await page.getByTestId("link-save").click();
  await expect(linkDialog(page)).toBeHidden();
}

// --- the dialog (AC-1..4) ----------------------------------------------------

test("Ctrl/Cmd+K with a selection opens the dialog prefilled with the selected text and an empty URL (AC-1)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("linkme");
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press(`${MOD}+k`);
  await expect(linkDialog(page)).toBeVisible();
  await expect(page.getByTestId("link-text-input")).toHaveValue("linkme");
  await expect(page.getByTestId("link-url-input")).toHaveValue("");
});

test("Ctrl/Cmd+K with no selection opens with both fields empty (AC-2)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("hello");
  await page.keyboard.press(`${MOD}+k`);
  await expect(linkDialog(page)).toBeVisible();
  await expect(page.getByTestId("link-text-input")).toHaveValue("");
  await expect(page.getByTestId("link-url-input")).toHaveValue("");
});

test("saving wraps the text in a hyperlink rendered underlined in the link colour (AC-3, AC-4)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Visit example now");
  await selectRange(page, 6, 13);
  await page.keyboard.press(`${MOD}+k`);
  await page.getByTestId("link-url-input").fill("https://example.com/");
  await page.getByTestId("link-save").click();

  const link = links(page);
  await expect(link).toHaveCount(1);
  await expect(link).toHaveText("example");
  await expect(link).toHaveAttribute("href", "https://example.com/");
  await expect(link).toHaveCSS("text-decoration-line", "underline");
  await expect(link).toHaveCSS("color", "rgb(26, 115, 232)");
  await expect(content(page)).toHaveText("Visit example now");
});

// --- autolink (AC-5..7) ------------------------------------------------------

test("typing a bare URL followed by space autolinks it, and Ctrl/Cmd+Z reverts to plain text (AC-5, AC-7)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("See http://example.com");
  await page.keyboard.type(" ");

  const link = links(page);
  await expect(link).toHaveCount(1);
  await expect(link).toHaveText("http://example.com");
  await expect(link).toHaveAttribute("href", "http://example.com");

  await page.keyboard.press(`${MOD}+z`);
  await expect(links(page)).toHaveCount(0);
  await expect(content(page)).toContainText("See http://example.com");
});

test("Enter after a bare URL autolinks it and the caret stays on the new line (AC-5)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("See http://example.com");
  await page.keyboard.press("Enter");

  const link = links(page);
  await expect(link).toHaveCount(1);
  await expect(link).toHaveText("http://example.com");

  // The regression this guards: the caret must be on the paragraph the
  // Enter created, not back at the end of the linked line — and typing
  // there must not extend the link.
  await page.keyboard.type("next line");
  await expect(content(page).locator('[data-block="1"]')).toHaveText("next line");
  await expect(content(page).locator('[data-block="0"]')).not.toContainText("next line");
  await expect(link).toHaveText("http://example.com");
});

test("pasting a bare URL followed by space autolinks it (AC-6)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await dispatchPaste(page, { "text/plain": "http://example.com" });
  await expect(content(page)).toContainText("http://example.com");
  await page.keyboard.type(" ");
  const link = links(page);
  await expect(link).toHaveCount(1);
  await expect(link).toHaveAttribute("href", "http://example.com");
});

// --- the link chip (AC-8..11, AC-33, AC-34) ---------------------------------

test("clicking a link with the caret shows a chip with the target and Edit/Copy/Remove actions (AC-8..10)", async ({
  page,
  context,
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("go http://example.com ");
  await links(page).click();

  await expect(linkChip(page)).toBeVisible();
  await expect(page.getByTestId("link-chip-target")).toHaveText("http://example.com");
  await expect(page.getByTestId("link-chip-edit")).toBeVisible();
  await expect(page.getByTestId("link-chip-remove")).toBeVisible();

  // Copy link (AC-10) puts the target on the clipboard.
  await page.getByTestId("link-chip-copy").click();
  await expect
    .poll(() => page.evaluate(() => navigator.clipboard.readText()))
    .toBe("http://example.com");

  // Edit (AC-9) reopens the dialog prefilled with the link.
  await page.getByTestId("link-chip-edit").click();
  await expect(linkDialog(page)).toBeVisible();
  await expect(page.getByTestId("link-text-input")).toHaveValue("http://example.com");
  await expect(page.getByTestId("link-url-input")).toHaveValue("http://example.com");
});

test("Remove link leaves the text and its other formatting intact (AC-11, AC-33, AC-34)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("boldlink");
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.press(`${MOD}+k`);
  await page.getByTestId("link-url-input").fill("https://example.com/");
  await page.getByTestId("link-save").click();
  await expect(links(page)).toHaveCount(1);

  await links(page).click();
  await expect(linkChip(page)).toBeVisible();
  await page.getByTestId("link-chip-remove").click();

  await expect(links(page)).toHaveCount(0);
  await expect(content(page)).toHaveText("boldlink");
  await expect(content(page).locator("span", { hasText: "boldlink" }).first()).toHaveCSS(
    "font-weight",
    "700",
  );
});

// --- opening links (AC-12, AC-13) -------------------------------------------

test("Ctrl/Cmd+click opens the target in a new tab with rel noopener noreferrer (AC-12)", async ({
  page,
}) => {
  await createDocument(page);
  const origin = new URL(page.url()).origin;
  await linkWholeText(page, "open me", `${origin}/`);

  const link = links(page);
  await expect(link).toHaveAttribute("rel", "noopener noreferrer");
  await expect(link).toHaveAttribute("target", "_blank");

  const [popup] = await Promise.all([
    page.context().waitForEvent("page"),
    link.click({ modifiers: [MOD] }),
  ]);
  await popup.waitForLoadState("domcontentloaded");
  expect(popup.url().startsWith(origin)).toBe(true);
});

test("the chip's open action opens the target in a new tab (AC-13)", async ({ page }) => {
  await createDocument(page);
  const origin = new URL(page.url()).origin;
  await linkWholeText(page, "open me", `${origin}/`);

  await links(page).click();
  await expect(linkChip(page)).toBeVisible();
  const [popup] = await Promise.all([
    page.context().waitForEvent("page"),
    page.getByTestId("link-chip-open").click(),
  ]);
  await popup.waitForLoadState("domcontentloaded");
  expect(popup.url().startsWith(origin)).toBe(true);
});

// --- dangerous schemes (AC-14..18) ------------------------------------------

test("javascript:, data: and vbscript: URLs are rejected on entry (AC-14..16, AC-18)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("target");
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press(`${MOD}+k`);
  await expect(linkDialog(page)).toBeVisible();

  for (const url of ["javascript:alert(1)", "data:text/html,x", "vbscript:msgbox(1)"]) {
    await page.getByTestId("link-url-input").fill(url);
    await expect(page.getByTestId("link-error")).toBeVisible();
    await expect(page.getByTestId("link-save")).toBeDisabled();
  }
  await page.getByTestId("link-cancel").click();
  await expect(links(page)).toHaveCount(0);
});

test("dangerous URLs are stripped on paste — never autolinked (AC-17, AC-18)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await dispatchPaste(page, { "text/plain": "javascript:alert(1)" });
  await page.keyboard.type(" ");
  await expect(links(page)).toHaveCount(0);
  await expect(content(page)).toContainText("javascript:alert(1)");
});

// --- heading and bookmark targets (AC-19..21) --------------------------------

test("the dialog lists headings and bookmarks, and a heading link jumps to the heading (AC-19..21)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Chapter One");
  await page.keyboard.press(`${MOD}+Alt+1`);
  await page.keyboard.press("Enter");
  await page.keyboard.press(`${MOD}+Alt+0`);
  await page.keyboard.type("body text goes here");

  // A bookmark at the current caret, so the dialog has one of each kind.
  await page.getByTestId("toolbar-insert-bookmark").click();
  await page.getByTestId("bookmark-name-input").fill("notes");
  await page.getByTestId("bookmark-save").click();

  await page.keyboard.press(`${MOD}+k`);
  await expect(linkDialog(page)).toBeVisible();
  const targetList = page.getByTestId("link-target-list");
  await expect(targetList).toContainText("Headings");
  await expect(targetList).toContainText("Chapter One");
  await expect(targetList).toContainText("Bookmarks");
  await expect(targetList).toContainText("notes");

  // Choosing the heading creates an internal link labelled after it.
  await page.getByTestId("link-target-_h0_Chapter_One").click();
  await page.getByTestId("link-save").click();
  const link = links(page);
  await expect(link).toHaveCount(1);
  await expect(link).toHaveText("Chapter One");

  // Its open action jumps to the heading: typing lands at its start.
  await link.click();
  await expect(linkChip(page)).toBeVisible();
  await page.getByTestId("link-chip-open").click();
  await page.keyboard.type("X");
  await expect(content(page).locator("h1")).toHaveText("XChapter One");
});

// --- bookmarks (AC-22..28) ---------------------------------------------------

test("insert bookmark validates names: no spaces, unique, inline error (AC-22, AC-26..28)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Hello world");

  await page.getByTestId("toolbar-insert-bookmark").click();
  await page.getByTestId("bookmark-name-input").fill("has space");
  await expect(page.getByTestId("bookmark-error")).toContainText("spaces");
  await expect(page.getByTestId("bookmark-save")).toBeDisabled();

  await page.getByTestId("bookmark-name-input").fill("mark_1");
  await expect(page.getByTestId("bookmark-error")).toBeHidden();
  await page.getByTestId("bookmark-save").click();
  await expect(page.getByTestId("bookmark-dialog")).toBeHidden();

  // A second bookmark with the same name conflicts (AC-27, AC-28).
  await page.getByTestId("toolbar-insert-bookmark").click();
  await page.getByTestId("bookmark-name-input").fill("mark_1");
  await expect(page.getByTestId("bookmark-error")).toContainText("already exists");
  await expect(page.getByTestId("bookmark-save")).toBeDisabled();
  await page.getByTestId("bookmark-cancel").click();
});

test("the bookmark manager lists, jumps to and deletes bookmarks (AC-23..25)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Hello world");
  // Bookmark after "Hello" (offset 5).
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.getByTestId("toolbar-insert-bookmark").click();
  await page.getByTestId("bookmark-name-input").fill("middle");
  await page.getByTestId("bookmark-save").click();

  await page.keyboard.press(`${MOD}+End`);
  await page.getByTestId("toolbar-bookmarks").click();
  await expect(page.getByTestId("bookmark-manager")).toBeVisible();
  await expect(page.getByTestId("bookmark-row")).toHaveCount(1);
  await expect(page.getByTestId("bookmark-row")).toContainText("middle");

  // Jump (AC-24): the caret lands at the bookmark, so typing splits the
  // text exactly there.
  await page.getByTestId("bookmark-jump-middle").click();
  await expect(page.getByTestId("bookmark-manager")).toBeHidden();
  await page.keyboard.type("X");
  await expect(content(page)).toHaveText("HelloX world");

  // Delete (AC-25).
  await page.getByTestId("toolbar-bookmarks").click();
  await page.getByTestId("bookmark-delete-middle").click();
  await expect(page.getByTestId("bookmark-manager-empty")).toBeVisible();
  await page.getByTestId("bookmark-manager-close").click();
  await expect(content(page)).toHaveText("HelloX world");
});

// --- round trips (AC-29..32) -------------------------------------------------

async function seedLinkAndBookmark(page: Page): Promise<void> {
  await content(page).click();
  await page.keyboard.type("go http://example.com ");
  await page.getByTestId("toolbar-insert-bookmark").click();
  await page.getByTestId("bookmark-name-input").fill("rt_mark");
  await page.getByTestId("bookmark-save").click();
}

async function downloadAs(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("download-menu").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId(`download-${format}`).click(),
  ]);
  const path = await download.path();
  expect(path).toBeTruthy();
  return path!;
}

test("links and bookmarks round-trip through .docx (AC-29, AC-30)", async ({ page }) => {
  await createDocument(page);
  await seedLinkAndBookmark(page);
  const path = await downloadAs(page, "docx");

  await openFixture(page, path);
  const link = links(page);
  await expect(link).toHaveCount(1);
  await expect(link).toHaveAttribute("href", "http://example.com");
  await expect(link).toHaveText("http://example.com");
  await page.getByTestId("toolbar-bookmarks").click();
  await expect(page.getByTestId("bookmark-row")).toContainText("rt_mark");
});

test("links and bookmarks round-trip through .odt (AC-31, AC-32)", async ({ page }) => {
  await createDocument(page);
  await seedLinkAndBookmark(page);
  const path = await downloadAs(page, "odt");

  await openFixture(page, path);
  const link = links(page);
  await expect(link).toHaveCount(1);
  await expect(link).toHaveAttribute("href", "http://example.com");
  await expect(link).toHaveText("http://example.com");
  await page.getByTestId("toolbar-bookmarks").click();
  await expect(page.getByTestId("bookmark-row")).toContainText("rt_mark");
});
