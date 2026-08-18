import { test, expect, type Page } from "@playwright/test";

// SPEC-056 equations. The round-3 QA blocking finding was AC-1: after inserting
// an equation, typing immediately (no click) went to the surrounding document,
// not the equation, because DocumentView's post-edit caret restore reclaimed
// focus from the just-mounted equation surface. These regressions drive the
// exact unassisted-user path — insert, then type straight away — and assert the
// keystrokes build the equation and focus stays on its surface.

async function newDocument(page: Page) {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page.getByTestId("document-content")).toBeVisible();
}

const content = (page: Page) => page.getByTestId("document-content");
const linear = (page: Page) => page.getByTestId("equation-editor-linear");
const surface = (page: Page) => page.getByTestId("equation-editor-surface");

test.describe("SPEC-056 equation insert focus", () => {
  test("Insert → Equation: typing immediately builds the equation (AC-1)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();
    await page.keyboard.type("Notes: ");

    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-equation-button").click();
    await expect(page.getByTestId("equation-editor-surface")).toBeVisible();

    // No click into the surface — AC-1 promises the caret is already inside it.
    await page.keyboard.type("a+b");

    // The engine received the keystrokes (the linear form is empty in the bug,
    // where the document swallowed them).
    await expect(linear(page)).toHaveText("a+b");
    // And focus is on the equation surface, not the document body.
    const focused = await page.evaluate(() =>
      document.activeElement?.getAttribute("data-testid"),
    );
    expect(focused).toBe("equation-editor-surface");
  });

  test("Alt+= inserts an inline equation that captures typing (AC-3)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();

    await page.keyboard.press("Alt+Equal");
    await expect(page.getByTestId("equation-editor-surface")).toBeVisible();

    await page.keyboard.type("x^2");
    await expect(linear(page)).toHaveText("x^2");
  });

  test("Insert → Display equation: typing immediately builds the equation (AC-2)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();

    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-display-equation-button").click();
    await expect(page.getByTestId("equation-editor-surface")).toBeVisible();

    await page.keyboard.type("y");
    await expect(linear(page)).toHaveText("y");
  });
});

test.describe("SPEC-056 equation editing surface", () => {
  test("template palette inserts a construct with empty slots, Tab/Shift+Tab move between them (AC-5/6/7)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();
    await page.keyboard.press("Alt+Equal");
    await expect(surface(page)).toBeVisible();

    await page.getByTestId("equation-template-frac").click();
    await expect(linear(page)).toHaveText("\\frac{}{}");

    await page.keyboard.type("a");
    await page.keyboard.press("Tab");
    await page.keyboard.type("b");
    await expect(linear(page)).toHaveText("\\frac{a}{b}");

    await page.keyboard.press("Shift+Tab");
    await page.keyboard.type("A");
    await expect(linear(page)).toHaveText("\\frac{aA}{b}");
  });

  test("Right arrow leaves the numerator into the denominator; Backspace at a slot start unwraps the construct (AC-12/13)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();
    await page.keyboard.press("Alt+Equal");
    await page.getByTestId("equation-template-frac").click();
    await page.keyboard.type("a");
    await page.keyboard.press("ArrowRight");
    await page.keyboard.type("Z");
    await expect(linear(page)).toHaveText("\\frac{a}{Z}");

    await page.keyboard.press("ArrowLeft");
    await page.keyboard.press("ArrowLeft");
    await page.keyboard.press("ArrowLeft");
    await page.keyboard.press("Backspace");
    // Unwrapping keeps the construct's content — numerator "a" and denominator
    // "Z" flatten into the row in order.
    await expect(linear(page)).toHaveText("aZ");
  });

  test("the special character picker inserts Maths and Greek glyphs into an open equation (AC-14/15)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();
    await page.keyboard.press("Alt+Equal");
    await expect(surface(page)).toBeVisible();

    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-special-character-button").click();
    await page.getByTestId("special-character-category-maths").scrollIntoViewIfNeeded();
    await page.getByTestId("special-char-AC").click(); // ¬ NOT SIGN
    await expect(surface(page)).toContainText("¬");

    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-special-character-button").click();
    await page.getByTestId("special-character-category-greek").scrollIntoViewIfNeeded();
    await page.getByTestId("special-char-391").click(); // Α GREEK CAPITAL ALPHA
    await expect(surface(page)).toContainText("Α");
  });

  test("\\alpha resolves through SPEC-026's catalogue (AC-16)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();
    await page.keyboard.press("Alt+Equal");
    await page.keyboard.type("\\alpha");
    await expect(surface(page)).toContainText("α");
  });

  test("Insert → Numbered equation gets a caption number and is a valid cross-reference target (AC-17/18)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-numbered-equation-button").click();
    await expect(surface(page)).toBeVisible();
    await page.keyboard.type("y=x");
    await expect(content(page)).toContainText("(1)");

    await page.getByTestId("equation-editor-done").click();
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("insert-cross-reference-button").click();
    await expect(page.getByTestId("cross-reference-targets")).toContainText("Equation");
  });

  test("alt text defaults to the linear form, is editable, and the committed equation re-renders live (AC-19/20/21)", async ({ page }) => {
    await newDocument(page);
    await content(page).click();
    await page.keyboard.press("Alt+Equal");
    await page.keyboard.type("a+b");

    const altInput = page.getByTestId("equation-alt-text");
    await expect(altInput).toHaveValue("");
    await altInput.fill("custom description");
    await page.getByTestId("equation-editor-done").click();

    const committed = content(page).locator("[aria-label='custom description']");
    await expect(committed).toBeVisible();
    await expect(committed).toContainText("a+b");
  });

  test("input outside the linear grammar is inserted as literal glyphs, not an error (AC-45)", async ({ page }) => {
    const errors: string[] = [];
    await newDocument(page);
    page.on("pageerror", (e) => errors.push(String(e)));
    await content(page).click();
    await page.keyboard.press("Alt+Equal");
    await page.keyboard.type("\\notarealcommand");
    await expect(surface(page)).toContainText("notarealcommand");
    expect(errors).toEqual([]);
  });
});
