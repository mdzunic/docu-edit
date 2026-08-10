import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-017 — find and replace: the find bar (open/close shortcuts, match
// count, navigation, wrap hint, highlight overlay), the replace row (expand
// toggle, single-undo-step replace all, formatting-preserving replace),
// search options (match case / whole words / regex, session persistence,
// invalid-regex handling), scope (body text only) and the large-document
// performance contract. createDocument/content follow e2e/SPEC-015.spec.ts.

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

const content = (page: Page) => page.getByTestId("document-content");
const findBar = (page: Page) => page.getByTestId("find-bar");
const findQuery = (page: Page) => page.getByTestId("find-query");
const findCount = (page: Page) => page.getByTestId("find-count");
const findStatus = (page: Page) => page.getByTestId("find-status");
const highlights = (page: Page) => page.locator('[data-testid="find-highlights"] .find-highlight');

const SAMPLE_TEXT =
  "The quick brown fox jumps over the lazy dog. The dog barks loudly. THE cat sleeps there quietly. foo123 bar456 foo789.";

async function seedDocument(page: Page, text: string = SAMPLE_TEXT): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
}

test("Ctrl/Cmd+F opens the find bar with a focused query field, previous/next and close buttons (AC-1, AC-3, AC-4)", async ({
  page,
}) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press(`${MOD}+f`);
  await expect(findBar(page)).toBeVisible();
  await expect(findQuery(page)).toBeFocused();
  await expect(page.getByTestId("find-prev")).toBeVisible();
  await expect(page.getByTestId("find-next")).toBeVisible();
  await expect(page.getByTestId("find-close")).toBeVisible();
});

test("typing highlights every match, the current one distinctly, with a live count (AC-2, AC-5, AC-6, AC-7)", async ({
  page,
}) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press(`${MOD}+f`);
  await findQuery(page).fill("the");
  await expect(findCount(page)).toHaveText("1 of 5");
  await expect(highlights(page)).toHaveCount(5);
  await expect(page.locator('[data-testid="find-highlights"] .find-highlight[data-current="true"]')).toHaveCount(1);
});

test("Enter/Shift+Enter step through matches and wrap with a 'Wrapped to top' hint (AC-8, AC-9, AC-10)", async ({
  page,
}) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press(`${MOD}+f`);
  await findQuery(page).fill("the");
  await expect(findCount(page)).toHaveText("1 of 5");

  await page.keyboard.press("Enter");
  await expect(findCount(page)).toHaveText("2 of 5");

  await page.keyboard.press("Shift+Enter");
  await expect(findCount(page)).toHaveText("1 of 5");

  await page.keyboard.press("Shift+Enter");
  await expect(findCount(page)).toHaveText("5 of 5");
  await expect(findStatus(page)).toHaveText("Wrapped to bottom");

  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("Enter");
  }
  await expect(findCount(page)).toHaveText("5 of 5");
  await expect(findStatus(page)).toHaveText("Wrapped to top");
});

test("Escape closes the bar, clears highlights and restores the caret (AC-11, AC-12, AC-13)", async ({ page }) => {
  await createDocument(page);
  await seedDocument(page);

  await content(page).click();
  await page.keyboard.press(`${MOD}+Home`);
  const caretBefore = await page.getByTestId("caret-status").textContent();

  await page.keyboard.press(`${MOD}+f`);
  await findQuery(page).fill("dog");
  await expect(findCount(page)).toHaveText("1 of 2");

  await page.keyboard.press("Escape");
  await expect(findBar(page)).not.toBeVisible();
  await expect(page.locator('[data-testid="find-highlights"]')).toHaveCount(0);
  await expect(page.getByTestId("caret-status")).toHaveText(caretBefore ?? "");
});

test("Ctrl/Cmd+H and the expand toggle both reveal the replace row with Replace and Replace all (AC-14, AC-15, AC-16, AC-17)", async ({
  page,
}) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press(`${MOD}+h`);
  await expect(page.getByTestId("find-replace-row")).toBeVisible();
  await expect(page.getByTestId("find-replace-input")).toBeVisible();
  await expect(page.getByTestId("find-replace")).toBeVisible();
  await expect(page.getByTestId("find-replace-all")).toBeVisible();

  await page.getByTestId("find-expand").click();
  await expect(page.getByTestId("find-replace-row")).not.toBeVisible();
  await page.getByTestId("find-expand").click();
  await expect(page.getByTestId("find-replace-row")).toBeVisible();
});

test("Replace all reports the count via a toast and is a single undo step (AC-18, AC-19)", async ({ page }) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press(`${MOD}+h`);
  await findQuery(page).fill("dog");
  await expect(findCount(page)).toHaveText("1 of 2");
  await page.getByTestId("find-replace-input").fill("wolf");
  await page.getByTestId("find-replace-all").click();

  await expect(page.getByTestId("toast").last()).toHaveText("Replaced 2 occurrences");
  await expect(content(page)).toContainText("wolf");
  await expect(content(page)).not.toContainText(/\bdog\b/);

  await content(page).click();
  await page.keyboard.press(`${MOD}+z`);
  await expect(content(page)).toContainText("dog");
  await expect(content(page)).not.toContainText("wolf");
});

test("Match case, whole words and regex narrow the results, and persist for the session (AC-20, AC-21, AC-22, AC-23)", async ({
  page,
}) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press(`${MOD}+f`);

  // Match case: "THE" case-insensitively matches all 5 "the"-like substrings;
  // case-sensitively only the literal "THE" ("THE cat sleeps...").
  await findQuery(page).fill("THE");
  await expect(findCount(page)).toHaveText(/ of 5$/);
  await page.getByTestId("find-match-case").click();
  await expect(page.getByTestId("find-match-case")).toHaveAttribute("aria-pressed", "true");
  await expect(findCount(page)).toHaveText("1 of 1");
  await page.getByTestId("find-match-case").click();

  // Whole words: "the" alone matches 5 (including inside "there"); whole-words drops to 4.
  await findQuery(page).fill("the");
  await expect(findCount(page)).toHaveText(/ of 5$/);
  await page.getByTestId("find-whole-words").click();
  await expect(page.getByTestId("find-whole-words")).toHaveAttribute("aria-pressed", "true");
  await expect(findCount(page)).toHaveText(/ of 4$/);
  await page.getByTestId("find-whole-words").click();

  // Regex: `foo\d+` matches the two numbered tokens.
  await page.getByTestId("find-regex").click();
  await expect(page.getByTestId("find-regex")).toHaveAttribute("aria-pressed", "true");
  await findQuery(page).fill("foo\\d+");
  await expect(findCount(page)).toHaveText("1 of 2");

  // Persist for the session: closing and reopening the bar keeps regex mode on.
  await page.keyboard.press("Escape");
  await page.keyboard.press(`${MOD}+f`);
  await expect(page.getByTestId("find-regex")).toHaveAttribute("aria-pressed", "true");
});

test("an invalid regular expression shows an inline error and does not run the search (AC-24, AC-25)", async ({
  page,
}) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press(`${MOD}+f`);
  await page.getByTestId("find-regex").click();
  await findQuery(page).fill("(unterminated");

  await expect(findStatus(page)).toHaveText("Invalid regular expression");
  await expect(findCount(page)).toHaveText("0 of 0");
  await expect(highlights(page)).toHaveCount(0);
});

test("a replacement takes the formatting of the first run it replaced (AC-26)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("plain text then ");
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type("boldword");
  await page.keyboard.press(`${MOD}+b`);

  await page.keyboard.press(`${MOD}+h`);
  await findQuery(page).fill("boldword");
  await page.getByTestId("find-replace-input").fill("shinyword");
  await page.getByTestId("find-replace").click();

  const replaced = content(page).locator("span", { hasText: "shinyword" });
  await expect(replaced).toHaveCount(1);
  await expect(replaced).toHaveCSS("font-weight", "700");
});

test("search covers only body text and states its scope (AC-27, AC-28, AC-29)", async ({ page }) => {
  await createDocument(page);
  await seedDocument(page);

  await page.keyboard.press(`${MOD}+f`);
  await findQuery(page).fill("fox");
  await expect(findCount(page)).toHaveText("1 of 1");

  const scope = page.getByTestId("find-scope");
  await expect(scope).toBeVisible();
  await expect(scope).toHaveText("Body text only");
  // Headers, footers and footnotes are not part of the document model yet
  // (no such feature exists in the app), so there is nothing outside body
  // text a match could leak in from; findMatches() in src/lib/findReplace.ts
  // only ever iterates model.sections[].blocks[].
});

test("a 200-page document returns first results within 300ms and typing stays responsive (AC-30, AC-31)", async ({
  page,
}) => {
  const docId = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("seed");
  await expect(page.getByTestId("save-state")).toHaveText("All changes saved");

  // Navigate off the document first so its autosave (beforeunload/pagehide)
  // can't race the direct IndexedDB seed below and reintroduce the small
  // in-memory model; the document-list route shares the origin/IndexedDB
  // but has no autosave tied to this document.
  await page.goto("/");
  const seeded = await page.evaluate(async (id) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const req = indexedDB.open("docu-edit"); // no explicit version: the app has already opened the DB at its current schema version by this point
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    const readTx = db.transaction("documents", "readonly");
    const getRq = readTx.objectStore("documents").get(id);
    const committed: { model: { sections: { blocks: unknown[] }[] } } = await new Promise((resolve, reject) => {
      getRq.onsuccess = () => resolve(getRq.result);
      getRq.onerror = () => reject(getRq.error);
    });
    const paragraphs = [];
    // ~45 lines/page * 200 pages of body text.
    for (let i = 0; i < 9000; i++) {
      paragraphs.push({
        runs: [{ text: `Paragraph number ${i} with some target needle words scattered about.`, formatting: {} }],
      });
    }
    committed.model.sections[0].blocks = paragraphs;
    const writeTx = db.transaction("documents", "readwrite");
    const store = writeTx.objectStore("documents");
    store.put({ ...committed, lastModified: Date.now() });
    store.delete("draft:" + id);
    await new Promise((resolve, reject) => {
      writeTx.oncomplete = () => resolve(undefined);
      writeTx.onerror = () => reject(writeTx.error);
    });
    const verifyRq = db.transaction("documents", "readonly").objectStore("documents").get(id);
    const verified: { model: { sections: { blocks: unknown[] }[] } } = await new Promise((resolve, reject) => {
      verifyRq.onsuccess = () => resolve(verifyRq.result);
      verifyRq.onerror = () => reject(verifyRq.error);
    });
    db.close();
    return verified.model.sections[0].blocks.length;
  }, docId);
  expect(seeded).toBe(9000);

  await page.goto(`/doc/${docId}`);
  await expect(content(page).locator(".doc-paragraph").first()).toBeVisible();
  await content(page).click();

  await page.keyboard.press(`${MOD}+f`);
  const start = Date.now();
  await findQuery(page).fill("needle");
  await expect(findCount(page)).not.toHaveText("0 of 0", { timeout: 2000 });
  expect(Date.now() - start).toBeLessThanOrEqual(300);

  // Typing stays responsive: a burst of keystrokes into the query field is
  // never dropped while the (much more expensive) search/highlight work for
  // a 9000-paragraph document runs behind it.
  await findQuery(page).fill("");
  const typed = "abcdefghijklmnopqrstuvwxyz";
  await findQuery(page).pressSequentially(typed, { delay: 10 });
  await expect(findQuery(page)).toHaveValue(typed);
});
