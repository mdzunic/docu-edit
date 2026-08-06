import { test, expect, type Page } from "@playwright/test";
import { LINE_END } from "./keys";

// SPEC-026 — the special-character picker, search, the Recent row and its
// persistence, the autocorrect toggle, undo of a substitution, and non-BMP
// insert/render/delete as single grapheme units.

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

const content = (page: Page) => page.getByTestId("document-content");

/** The exact text content of the first paragraph (no whitespace collapsing,
 * so NBSP and friends are observable). */
async function paraText(page: Page): Promise<string> {
  return page.evaluate(() => {
    const p = document.querySelector('[data-testid="document-content"] [data-section]');
    return p?.textContent ?? "";
  });
}

async function focusCaret(page: Page): Promise<void> {
  await content(page).click();
}

test.describe("special-character picker", () => {
  test("opens a searchable grid, inserts on click and stays open (AC-1, AC-2, AC-5, AC-6)", async ({
    page,
  }) => {
    await createDocument(page);
    await focusCaret(page);
    await page.getByTestId("insert-special-character-button").click();
    await expect(page.getByTestId("special-character-dialog")).toBeVisible();
    // All eight categories present.
    for (const id of ["punctuation", "currency", "maths", "arrows", "latin", "greek", "symbols", "emoji"]) {
      await expect(page.getByTestId(`special-character-category-${id}`)).toBeVisible();
    }
    // Insert the em dash, twice — the dialog stays open between inserts.
    await page.getByTestId("special-char-2014").click();
    await page.getByTestId("special-char-2014").click();
    await expect(page.getByTestId("special-character-dialog")).toBeVisible();
    expect(await paraText(page)).toBe("——");
  });

  test("search filters by name and by hex code point (AC-2, AC-4)", async ({ page }) => {
    await createDocument(page);
    await focusCaret(page);
    await page.getByTestId("insert-special-character-button").click();
    const search = page.getByTestId("special-character-search");
    await search.fill("euro sign");
    await expect(page.getByTestId("special-char-20AC")).toBeVisible();
    await expect(page.getByTestId("special-char-2014")).toHaveCount(0);
    await search.fill("u+2014");
    await expect(page.getByTestId("special-char-2014")).toBeVisible();
    await expect(page.getByTestId("special-char-20AC")).toHaveCount(0);
  });

  test("tracks a Recent row that survives a reload (AC-7, AC-10)", async ({ page }) => {
    await createDocument(page);
    await focusCaret(page);
    await page.getByTestId("insert-special-character-button").click();
    await page.getByTestId("special-char-2014").click();
    await page.getByTestId("special-character-close").click();

    await page.reload();
    await focusCaret(page);
    await page.getByTestId("insert-special-character-button").click();
    await expect(page.getByTestId("special-character-recent")).toBeVisible();
    await expect(page.getByTestId("recent-char-2014")).toBeVisible();
  });
});

test.describe("autocorrect", () => {
  test("substitutes dashes and smart quotes while typing (AC-16, AC-18)", async ({ page }) => {
    await createDocument(page);
    await focusCaret(page);
    await page.keyboard.type("a--");
    expect(await paraText(page)).toBe("a–");
    await page.keyboard.type("-");
    expect(await paraText(page)).toBe("a—");
    await page.keyboard.press("Enter");
    await page.keyboard.type("'");
    // Opening single quote at the new paragraph's start.
    const all = await page.evaluate(
      () => document.querySelector('[data-testid="document-content"]')?.textContent ?? "",
    );
    expect(all).toContain("‘");
  });

  test("one undo restores the literal dashes (AC-20, AC-21)", async ({ page }) => {
    await createDocument(page);
    await focusCaret(page);
    await page.keyboard.type("a--");
    expect(await paraText(page)).toBe("a–");
    await page.keyboard.press("Control+z");
    expect(await paraText(page)).toBe("a--");
  });

  test("turning dash substitution off leaves -- literal (AC-24)", async ({ page }) => {
    await createDocument(page);
    await focusCaret(page);
    await page.getByTestId("autocorrect-menu").click();
    await page.getByTestId("autocorrect-dashes").click();
    // Close the menu, then type.
    await focusCaret(page);
    await page.keyboard.type("a--");
    expect(await paraText(page)).toBe("a--");
  });
});

test.describe("non-BMP characters", () => {
  test("insert, render and delete as single units (AC-32, AC-33, AC-35)", async ({ page }) => {
    await createDocument(page);
    await focusCaret(page);
    await page.keyboard.type("a");
    await page.getByTestId("insert-special-character-button").click();
    await page.getByTestId("special-character-search").fill("grinning face");
    await page.getByTestId("special-char-1F600").first().click();
    await page.getByTestId("special-character-close").click();
    // Rendered as a single grinning face after the "a".
    expect(await paraText(page)).toBe("a😀");
    // One Backspace removes the whole surrogate pair, not half of it.
    await focusCaret(page);
    await page.keyboard.press("End");
    await page.keyboard.press("Backspace");
    expect(await paraText(page)).toBe("a");
  });

  test("caret motion and selection step over a surrogate pair as one unit (AC-36, AC-37, AC-40)", async ({
    page,
  }) => {
    await createDocument(page);
    await focusCaret(page);
    await page.keyboard.type("a");
    await page.getByTestId("insert-special-character-button").click();
    await page.getByTestId("special-character-search").fill("grinning face");
    await page.getByTestId("special-char-1F600").first().click();
    await page.getByTestId("special-character-close").click();
    await focusCaret(page);
    await page.keyboard.type("b"); // "a😀b"
    // From the line end, one Left steps past "b", the next past the whole
    // emoji (2 code units) in a single press.
    await page.keyboard.press(LINE_END);
    await page.keyboard.press("ArrowLeft");
    await page.keyboard.press("ArrowLeft");
    // caret-status counts characters 1-based, so offset 1 (immediately before
    // the emoji) reads as "character 2" — crucially not "character 3", which
    // would mean the caret split the surrogate pair.
    await expect(page.getByTestId("caret-status")).toHaveText(
      "Paragraph 1 of 1, character 2",
    );
    // Shift-extending right selects the emoji whole, never a lone surrogate.
    await page.keyboard.press("Shift+ArrowRight");
    const selected = await page.evaluate(() => window.getSelection()?.toString() ?? "");
    expect(selected).toBe("😀");
  });
});
