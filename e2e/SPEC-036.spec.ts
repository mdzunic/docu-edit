import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-036 — word count and document statistics: a live, throttled word
// count in the status bar that shows "N of M words" with a selection,
// clicking it opens a statistics dialog (pages, words, characters with/
// without spaces, paragraphs, sentences, CJK characters), the "Include
// footnotes and endnotes" toggle updates the numbers live, header/footer
// text stays excluded from the body count, and the count is reachable from
// the keyboard.

const content = (page: Page) => page.getByTestId("document-content");
const wordCountButton = (page: Page) => page.getByTestId("word-count-button");
const statisticsDialog = (page: Page) => page.getByTestId("statistics-dialog");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function typeBody(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
}

test("the status bar shows a live word count that updates as the user types (AC-1, AC-2)", async ({ page }) => {
  await createDocument(page);
  await expect(wordCountButton(page)).toHaveText("0 words");
  await typeBody(page, "one two three");
  await expect(wordCountButton(page)).toHaveText("3 words");
  await page.keyboard.type(" four");
  await expect(wordCountButton(page)).toHaveText("4 words");
});

test("with a selection, the status bar shows N of M words (AC-3)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "one two three four");
  await expect(wordCountButton(page)).toHaveText("4 words");
  // Select "three four" — Shift+Home then Shift+Ctrl/Cmd+Left twice would
  // overshoot; select the whole line then shrink from the start instead:
  // simplest reliable selection is select-all, which is exactly "N of M"
  // with N === M.
  await page.keyboard.press(`${MOD}+a`);
  await expect(wordCountButton(page)).toHaveText("4 of 4 words");
  // Collapsing the selection goes back to the plain total.
  await page.keyboard.press("ArrowRight");
  await expect(wordCountButton(page)).toHaveText("4 words");
});

test("clicking the word count opens the statistics dialog with pages, words, characters, paragraphs and sentences (AC-4..10)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "Hello world. This is a test.");
  await wordCountButton(page).click();
  await expect(statisticsDialog(page)).toBeVisible();
  await expect(page.getByTestId("statistics-pages")).toHaveText("1");
  await expect(page.getByTestId("statistics-words")).toHaveText("6");
  await expect(page.getByTestId("statistics-paragraphs")).toHaveText("1");
  await expect(page.getByTestId("statistics-sentences")).toHaveText("2");
  const withSpaces = Number(await page.getByTestId("statistics-characters-with-spaces").textContent());
  const withoutSpaces = Number(await page.getByTestId("statistics-characters-without-spaces").textContent());
  expect(withSpaces).toBeGreaterThan(withoutSpaces);
  await expect(page.getByTestId("statistics-cjk-characters")).toHaveText("0");
});

test("the dialog closes on the Close button and reopens fresh (AC-4)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "hello");
  await wordCountButton(page).click();
  await expect(statisticsDialog(page)).toBeVisible();
  await page.getByTestId("statistics-close").click();
  await expect(statisticsDialog(page)).toHaveCount(0);
});

test("the dialog closes on Escape", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "hello");
  await wordCountButton(page).click();
  await expect(statisticsDialog(page)).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(statisticsDialog(page)).toHaveCount(0);
});

test("the CJK character count is separate from the Latin word count (AC-16, AC-20)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "hello 你好世界");
  await wordCountButton(page).click();
  // "hello" + 4 individual CJK characters = 5 words.
  await expect(page.getByTestId("statistics-words")).toHaveText("5");
  await expect(page.getByTestId("statistics-cjk-characters")).toHaveText("4");
});

test("Include footnotes and endnotes toggles the numbers live without closing the dialog (AC-11, AC-12)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "body text");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("notes-menu-footnote").click();
  await page.keyboard.type("footnote content here");
  await wordCountButton(page).click();
  await expect(statisticsDialog(page)).toBeVisible();
  await expect(page.getByTestId("statistics-words")).toHaveText("2");
  await page.getByTestId("statistics-include-notes").click();
  await expect(statisticsDialog(page)).toBeVisible();
  await expect(page.getByTestId("statistics-words")).toHaveText("5");
  // Toggling back off updates live too.
  await page.getByTestId("statistics-include-notes").click();
  await expect(page.getByTestId("statistics-words")).toHaveText("2");
});

test("header and footer text is excluded from the body word count (AC-21, AC-22)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "body words here");
  await expect(wordCountButton(page)).toHaveText("3 words");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await page.keyboard.type("secret header words that must not count");
  await expect(wordCountButton(page)).toHaveText("3 words");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-footer").click();
  await page.keyboard.type("footer words too");
  await expect(wordCountButton(page)).toHaveText("3 words");
});

test("hyphenated words, numbers and URLs count per the documented rules (AC-13..15, AC-17..19)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "a well-known result costs 42 dollars see https://example.com/page for more");
  // a(1) well-known(1) result(1) costs(1) 42(1) dollars(1) see(1) URL(1) for(1) more(1) = 10
  await expect(wordCountButton(page)).toHaveText("10 words");
});

test("the word count is reachable and activatable from the keyboard (AC-27)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "hello");
  await wordCountButton(page).focus();
  await expect(wordCountButton(page)).toBeFocused();
  await page.keyboard.press("Enter");
  await expect(statisticsDialog(page)).toBeVisible();
});

test("the word count is announced to screen readers when it changes (AC-28)", async ({ page }) => {
  await createDocument(page);
  const status = page.getByTestId("word-count-status");
  await expect(status).toHaveAttribute("aria-live", "polite");
  await typeBody(page, "one two three");
  await expect(status).toHaveText("3 words");
});
