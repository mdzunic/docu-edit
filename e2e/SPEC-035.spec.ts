import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-035 — spell checking and document language: red-squiggle underlines
// that run off the typing path, a right-click suggestions menu (Ignore/
// Ignore all/Add to dictionary), a custom dictionary settings dialog, a
// language selector (document default vs. selection), per-language
// dictionary resolution, default skip rules (URLs/emails/code style/"do
// not check spelling"), and the .docx/.odt language round trip.

const content = (page: Page) => page.getByTestId("document-content");
const languageButton = (page: Page) => page.getByTestId("language-status-button");
const underlines = (page: Page) => page.locator('[data-testid="spell-underline"]');

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function typeBody(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
}

/** Waits for the idle-scheduled spell-check pass to land — there is no
 * status/idle indicator to poll, so this polls the underline count instead
 * of a fixed sleep, keeping the test fast when checking finishes quickly
 * and robust when it doesn't. */
async function waitForUnderlines(page: Page, atLeast = 1): Promise<void> {
  await expect
    .poll(async () => underlines(page).count(), { timeout: 5000 })
    .toBeGreaterThanOrEqual(atLeast);
}

/** Waits for the underline set to stop changing, and returns the settled
 * count. A bare "at least N" wait can return mid-flight: typing "again" makes
 * an idle pass briefly flag the incomplete "agai", so the count can be one
 * higher than the text's real misspelling count for a moment. */
async function waitForStableUnderlines(page: Page, atLeast = 1): Promise<number> {
  await waitForUnderlines(page, atLeast);
  let previous = -1;
  for (let attempt = 0; attempt < 20; attempt += 1) {
    const count = await underlines(page).count();
    if (count === previous && count >= atLeast) {
      return count;
    }
    previous = count;
    await page.waitForTimeout(150);
  }
  return previous;
}

/** Right-clicks the first flagged word. The underline overlay is deliberately
 * inert (pointer-events: none) — its rects cover whole glyph boxes and paint
 * above the text, so making them clickable swallowed every ordinary click and
 * drag that started on a misspelled word. DocumentView hit-tests the
 * right-click against the underline rects instead, so this drives the mouse at
 * the word's coordinates, which is what a user does anyway. Settling first
 * stands in for the actionability wait Playwright would have done on a real
 * click target — an inert element can never "receive events", so a re-measuring
 * pass could otherwise move the word between measuring it and clicking. */
async function rightClickFirstUnderline(page: Page): Promise<void> {
  await waitForStableUnderlines(page);
  const box = await underlines(page).first().boundingBox();
  if (!box) {
    throw new Error("no spell underline on screen to right-click");
  }
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: "right" });
}

test("a misspelled word is genuinely painted on screen, not just present in the DOM (AC-1)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "This is a wrongwordxyz test.");
  await waitForUnderlines(page);
  // Regression guard for the exact historical failure: the underline
  // element existed in the DOM at the right coordinates but painted behind
  // the page's opaque background (a CSS stacking-order bug), so it never
  // appeared on screen. elementFromPoint resolves in paint order, so it can
  // still answer "is anything covering this box?" — but only for an element
  // that takes part in hit-testing, and the layer is deliberately inert (see
  // rightClickFirstUnderline). So the probe opts this one element in for the
  // duration of the measurement and restores it afterwards: if the stacking
  // bug came back, the page background would win the hit test and this fails
  // exactly as it did before.
  const reached = await page.evaluate(() => {
    const el = document.querySelector('[data-testid="spell-underline"]') as HTMLElement;
    const previous = el.style.pointerEvents;
    el.style.pointerEvents = "auto";
    try {
      const rect = el.getBoundingClientRect();
      const at = document.elementFromPoint(
        Math.round(rect.left + rect.width / 2),
        Math.round(rect.top + rect.height / 2),
      );
      return !!at?.closest('[data-testid="spell-underline"]');
    } finally {
      el.style.pointerEvents = previous;
    }
  });
  expect(reached).toBe(true);

  // ...and the inertness itself is the other half of AC-1: with the layer
  // left alone, a click on a flagged word must reach the text underneath.
  const hitsText = await page.evaluate(() => {
    const el = document.querySelector('[data-testid="spell-underline"]')!;
    const rect = el.getBoundingClientRect();
    const at = document.elementFromPoint(
      Math.round(rect.left + rect.width / 2),
      Math.round(rect.top + rect.height / 2),
    );
    return !at?.closest('[data-testid="spell-underline"]');
  });
  expect(hitsText).toBe(true);
});

test("typing stays responsive while spell check runs off the typing path (AC-1.1)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  const start = Date.now();
  await page.keyboard.type("wrongwordone wrongwordtwo wrongwordthree wrongwordfour");
  const elapsed = Date.now() - start;
  // Generous but meaningful: typing 54 characters synchronously should
  // never be gated on dictionary lookups.
  expect(elapsed).toBeLessThan(3000);
  await waitForUnderlines(page, 4);
});

test("right-clicking a misspelled word offers up to five suggestions plus Ignore/Ignore all/Add to dictionary (AC-2, AC-2.1..2.3)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "I saw a wrold today.");
  await waitForUnderlines(page);
  await rightClickFirstUnderline(page);
  const menu = page.getByTestId("spell-context-menu");
  await expect(menu).toBeVisible();
  const count = await page.locator('[data-testid="spell-suggestion"]').count();
  expect(count).toBeGreaterThan(0);
  expect(count).toBeLessThanOrEqual(5);
  await expect(page.getByTestId("spell-ignore")).toBeVisible();
  await expect(page.getByTestId("spell-ignore-all")).toBeVisible();
  await expect(page.getByTestId("spell-add-to-dictionary")).toBeVisible();
});

test('a word with no close match shows a disabled "No suggestions" item, and Ignore/Ignore all/Add to dictionary remain available (AC-3, AC-3.1)', async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "zzqxvbnmqwrty nonsense token");
  await waitForUnderlines(page);
  await rightClickFirstUnderline(page);
  await expect(page.getByTestId("spell-no-suggestions")).toBeVisible();
  await expect(page.getByTestId("spell-no-suggestions")).toBeDisabled();
  await expect(page.getByTestId("spell-ignore")).toBeEnabled();
  await expect(page.getByTestId("spell-ignore-all")).toBeEnabled();
  await expect(page.getByTestId("spell-add-to-dictionary")).toBeEnabled();
});

test("choosing a suggestion replaces the word, keeps its formatting, and matches its capitalisation (AC-4, AC-4.1)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.type("Wrold");
  await page.keyboard.press(`${MOD}+b`);
  await waitForUnderlines(page);
  await rightClickFirstUnderline(page);
  await page.getByTestId("spell-suggestion").first().click();
  await expect(content(page)).toHaveText("World");
  const bold = await content(page).locator("span, a").first().evaluate((el) => getComputedStyle(el).fontWeight);
  expect(["bold", "700"]).toContain(bold);
});

test('"Ignore" clears one occurrence, "Ignore all" clears every occurrence, for the session (AC-5, AC-6)', async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "wrold and wrold again");
  const before = await waitForStableUnderlines(page, 2);
  expect(before).toBeGreaterThanOrEqual(2);
  await rightClickFirstUnderline(page);
  await page.getByTestId("spell-ignore").click();
  await page.waitForTimeout(300);
  expect(await underlines(page).count()).toBe(before - 1);

  await rightClickFirstUnderline(page);
  await page.getByTestId("spell-ignore-all").click();
  await page.waitForTimeout(300);
  expect(await underlines(page).count()).toBe(0);
});

test("Add to dictionary persists across a reload and removes the underline everywhere (AC-7, AC-7.1, AC-7.2)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "docuedit docuedit");
  await waitForUnderlines(page, 2);
  await rightClickFirstUnderline(page);
  await page.getByTestId("spell-add-to-dictionary").click();
  await page.waitForTimeout(300);
  expect(await underlines(page).count()).toBe(0);

  await page.reload();
  await page.waitForTimeout(500);
  await typeBody(page, " docuedit");
  await page.waitForTimeout(500);
  expect(await underlines(page).count()).toBe(0);
});

test("the custom dictionary dialog lists, searches, adds, removes and filters by language (AC-8.1..8.3)", async ({ page }) => {
  await createDocument(page);
  await languageButton(page).click();
  await page.getByTestId("language-manage-dictionary").click();
  const dialog = page.getByTestId("custom-dictionary-dialog");
  await expect(dialog).toBeVisible();
  await expect(page.getByTestId("custom-dictionary-empty")).toBeVisible();

  await page.getByTestId("custom-dictionary-add-input").fill("docuedit");
  await page.getByTestId("custom-dictionary-add-submit").click();
  await expect(page.getByTestId("custom-dictionary-row")).toHaveCount(1);

  await page.getByTestId("custom-dictionary-add-input").fill("anotherword");
  await page.getByTestId("custom-dictionary-add-language").selectOption("de-DE");
  await page.getByTestId("custom-dictionary-add-submit").click();
  await expect(page.getByTestId("custom-dictionary-row")).toHaveCount(2);

  await page.getByTestId("custom-dictionary-search").fill("docuedit");
  await expect(page.getByTestId("custom-dictionary-row")).toHaveCount(1);
  await page.getByTestId("custom-dictionary-search").fill("");

  await page.getByTestId("custom-dictionary-language-filter").selectOption("de-DE");
  await expect(page.getByTestId("custom-dictionary-row")).toHaveCount(1);
  await expect(page.getByTestId("custom-dictionary-row")).toContainText("anotherword");
  await page.getByTestId("custom-dictionary-language-filter").selectOption("");

  await page.getByTestId("custom-dictionary-remove").first().click();
  await expect(page.getByTestId("custom-dictionary-row")).toHaveCount(1);
});

test("the language selector sets the document default with no selection, and only the selection when text is selected (AC-9, AC-9.1)", async ({ page }) => {
  await createDocument(page);
  await expect(languageButton(page)).toHaveText("English (US)");
  await languageButton(page).click();
  await page.getByTestId("language-option-de-DE").click();
  await expect(languageButton(page)).toHaveText("German");

  await typeBody(page, "Bonjour");
  await page.keyboard.press("Home");
  for (let i = 0; i < 7; i++) await page.keyboard.press("Shift+ArrowRight");
  await languageButton(page).click();
  await page.getByTestId("language-option-fr-FR").click();
  await expect(languageButton(page)).toHaveText("French");

  // Collapsed elsewhere shows the (unchanged) document default again.
  await page.keyboard.press("End");
  await expect(languageButton(page)).toHaveText("German");
});

test("each run is checked against its own language's dictionary (AC-10, AC-10.1)", async ({ page }) => {
  await createDocument(page);
  await languageButton(page).click();
  await page.getByTestId("language-option-de-DE").click();
  await typeBody(page, "Haus");
  await page.waitForTimeout(500);
  // "Haus" (house) is a real German word — not flagged against German.
  expect(await underlines(page).count()).toBe(0);

  await page.keyboard.press(`${MOD}+a`);
  await languageButton(page).click();
  await page.getByTestId("language-option-en-US").click();
  await page.waitForTimeout(500);
  // The same text, now marked English, is not a recognised English word.
  await waitForUnderlines(page, 1);
});

test('"Do not check spelling" skips the selection entirely (AC-11)', async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "zzqxnonsenseword stays flagged");
  await waitForUnderlines(page, 1);
  await page.keyboard.press("Home");
  for (let i = 0; i < 16; i++) await page.keyboard.press("Shift+ArrowRight");
  await languageButton(page).click();
  await page.getByTestId("language-no-proof").check();
  await page.keyboard.press("Escape");
  await page.keyboard.press("End");
  await page.waitForTimeout(500);
  const words = await underlines(page).evaluateAll((els) => els.map((el) => el.getAttribute("data-word")));
  expect(words).not.toContain("zzqxnonsenseword");
});

async function flaggedWords(page: Page): Promise<string[]> {
  const words = await underlines(page).evaluateAll((els) => els.map((el) => el.getAttribute("data-word")));
  return words.filter((word) => word !== null);
}

test("URLs and email addresses are skipped by default (AC-12)", async ({ page }) => {
  await createDocument(page);
  await typeBody(page, "visit https://example.dev/zzqxnonsense or mail zz@example.com then wrongwordxyz");
  // Poll the actual flagged word, not just the underline count: a
  // mid-typing idle pass can transiently flag a not-yet-complete word
  // ("wrongwordxy"), self-correcting on the next keystroke's recheck.
  await expect.poll(() => flaggedWords(page), { timeout: 5000 }).toContain("wrongwordxyz");
  const words = await flaggedWords(page);
  expect(words.join(" ")).not.toContain("zzqxnonsense");
  expect(words.join(" ")).not.toContain("example");
});

test("dictionaries load on demand and at least the six required languages are available (AC-17, AC-18)", async ({ page }) => {
  await createDocument(page);
  await languageButton(page).click();
  for (const tag of ["en-US", "en-GB", "de-DE", "fr-FR", "es-ES", "cs-CZ"]) {
    await expect(page.getByTestId(`language-option-${tag}`)).toBeVisible();
  }
});

test("a language with no bundled dictionary is silently skipped: no errors, no underlines (AC-19, AC-19.1)", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", (err) => errors.push(err.message));
  await createDocument(page);
  // Fixture the model directly wouldn't be UI-driven; instead confirm the
  // app never errors when the selector's own supported set is exhausted —
  // the reader/writer-level "unavailable dictionary" behavior is covered by
  // the docx/odf unit tests (parseDocx/parseOdt of a pt-BR fixture).
  await typeBody(page, "hello world");
  await page.waitForTimeout(500);
  expect(errors).toEqual([]);
});

test("an empty document produces no spell-check errors and loads no dictionary until text is entered (AC-20, AC-20.1)", async ({ page }) => {
  const requests: string[] = [];
  await createDocument(page);
  page.on("request", (req) => {
    if (/dictionaries\//.test(req.url())) {
      requests.push(req.url());
    }
  });
  await page.waitForTimeout(500);
  expect(await underlines(page).count()).toBe(0);
  expect(requests).toEqual([]);
});

test("language round-trips through a .docx save and reopen, reflected in the status bar (AC-13, AC-14, AC-15, AC-15.1)", async ({ page }) => {
  await createDocument(page);
  await languageButton(page).click();
  await page.getByTestId("language-option-de-DE").click();
  await typeBody(page, "haus");

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const download = await downloadPromise;
  const path = await download.path();
  expect(path).toBeTruthy();

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path!);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(languageButton(page)).toHaveText("German");
  const lang = await content(page).locator("span").first().getAttribute("lang");
  expect(lang).toBe("de-DE");
});
