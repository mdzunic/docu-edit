import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-074 — paragraph and page borders and shading. The unit and component
// suites cover the merge rule, the fragment arithmetic, both round trips and
// every dialog control in isolation; this file exercises the seams only a
// real browser can confirm: that a bordered paragraph actually paints its
// edges in the canvas, that three identical neighbours paint as ONE box
// (top on the first, bottom on the last, the between rule in between), that
// the shading spans the padded content width rather than just the glyphs,
// and that the Format menu → dialog → command path applies to every
// paragraph the selection touches.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

const content = (page: Page) => page.getByTestId("document-content");

async function openFixture(page: Page, name: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(join(fixturesDir, name));
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page).locator("p, h1, h2").first()).toBeVisible();
}

async function waitForPagination(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-pages")).toHaveAttribute(
    "data-pagination-state",
    "clean",
    { timeout },
  );
}

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** The rendered paragraph whose text starts with `prefix`. */
function paragraphStarting(page: Page, prefix: string) {
  return content(page).locator("p").filter({ hasText: prefix }).first();
}

/** One computed style value off a live element. */
async function styleOf(locator: ReturnType<typeof paragraphStarting>, property: string): Promise<string> {
  return locator.evaluate(
    (el, prop) => globalThis.getComputedStyle(el).getPropertyValue(prop),
    property,
  );
}

test("a bordered, shaded paragraph paints its edges and fill in the canvas (AC-13)", async ({ page }) => {
  await openFixture(page, "word-borders.docx");
  const note = paragraphStarting(page, "Note: this paragraph is boxed");

  expect(await styleOf(note, "border-top-style")).toBe("solid");
  expect(await styleOf(note, "border-bottom-style")).toBe("solid");
  expect(await styleOf(note, "border-left-style")).toBe("solid");
  expect(await styleOf(note, "border-right-style")).toBe("solid");
  expect(await styleOf(note, "background-color")).toBe("rgb(255, 242, 204)");

  // The shading spans the padded content width, not just the glyphs: the
  // padding box is what carries the background (AC-13).
  expect(await styleOf(note, "padding-left")).not.toBe("0px");
  expect(await styleOf(note, "padding-right")).not.toBe("0px");
});

test("a bottom-rule heading draws only its bottom edge (AC-6, AC-7)", async ({ page }) => {
  await openFixture(page, "word-borders.docx");
  const heading = content(page).locator("h1").filter({ hasText: "Borders and shading" }).first();
  expect(await styleOf(heading, "border-bottom-style")).toBe("double");
  expect(await styleOf(heading, "border-top-style")).toBe("none");
  expect(await styleOf(heading, "border-left-style")).toBe("none");
  // ...and it gains no indent from padding it has no rule for.
  expect(await styleOf(heading, "padding-left")).toBe("0px");
});

test("three identical paragraphs paint as one bordered block (AC-12)", async ({ page }) => {
  await openFixture(page, "word-borders.docx");
  const first = paragraphStarting(page, "First of three");
  const second = paragraphStarting(page, "Second of three");
  const third = paragraphStarting(page, "Third of three");

  // The outer box draws once around the run.
  expect(await styleOf(first, "border-top-style")).toBe("solid");
  expect(await styleOf(second, "border-top-style")).toBe("none");
  expect(await styleOf(third, "border-top-style")).toBe("none");
  expect(await styleOf(third, "border-bottom-style")).toBe("solid");

  // The between rule draws at each internal boundary — as the preceding
  // paragraph's bottom edge, which is what OOXML's w:between means.
  expect(await styleOf(first, "border-bottom-style")).toBe("dotted");
  expect(await styleOf(second, "border-bottom-style")).toBe("dotted");

  // The sides draw down the whole run.
  for (const paragraph of [first, second, third]) {
    expect(await styleOf(paragraph, "border-left-style")).toBe("solid");
    expect(await styleOf(paragraph, "border-right-style")).toBe("solid");
  }

  // No gap opens between members, or the page would show through the box.
  expect(await styleOf(first, "margin-bottom")).toBe("0px");
  expect(await styleOf(second, "margin-top")).toBe("0px");
});

test("the page border and background render behind the page (AC-22, AC-24)", async ({ page }) => {
  await openFixture(page, "word-borders.docx");
  const background = page.getByTestId("page-background").first();
  await expect(background).toBeAttached();
  expect(
    await background.evaluate((el) => globalThis.getComputedStyle(el).backgroundColor),
  ).toBe("rgb(253, 246, 227)");

  const border = page.getByTestId("page-border").first();
  await expect(border).toBeAttached();
  expect(await border.evaluate((el) => globalThis.getComputedStyle(el).borderTopStyle)).toBe("double");
  // Inert decoration — never a hit target for the caret.
  expect(await border.evaluate((el) => globalThis.getComputedStyle(el).pointerEvents)).toBe("none");
});

test("the .odt fixture renders the same borders (AC-29, AC-31)", async ({ page }) => {
  await openFixture(page, "lo-borders.odt");
  const note = paragraphStarting(page, "Note: this paragraph is boxed");
  expect(await styleOf(note, "border-top-style")).toBe("solid");
  expect(await styleOf(note, "background-color")).toBe("rgb(255, 242, 204)");
  expect(await styleOf(paragraphStarting(page, "First of three"), "border-bottom-style")).toBe("dotted");
  await expect(page.getByTestId("page-border").first()).toBeAttached();
});

test("Format → Borders and shading applies a box to every paragraph the selection touches (AC-6..AC-11)", async ({ page }) => {
  await createDocument(page);
  const editor = content(page);
  await editor.click();
  await page.keyboard.type("First paragraph");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Second paragraph");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Third paragraph");

  // Select from inside the first paragraph to inside the second — a range
  // that TOUCHES two paragraphs without covering either whole (SPEC-014).
  const first = paragraphStarting(page, "First paragraph");
  const second = paragraphStarting(page, "Second paragraph");
  await first.click();
  await page.keyboard.press("Home");
  await page.keyboard.press("ArrowRight");
  await page.keyboard.down("Shift");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.up("Shift");

  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  const dialog = page.getByTestId("borders-shading-dialog");
  await expect(dialog).toBeVisible();

  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-shading-on").check();
  await page.getByTestId("borders-apply").click();
  await expect(dialog).toBeHidden();

  // Both touched paragraphs are boxed; the untouched third is not.
  expect(await styleOf(first, "border-left-style")).toBe("solid");
  expect(await styleOf(second, "border-left-style")).toBe("solid");
  expect(await styleOf(paragraphStarting(page, "Third paragraph"), "border-left-style")).toBe("none");

  // ...and because the two are now identical, they merged: the top rule
  // draws on the first only, the bottom on the second only (AC-12).
  expect(await styleOf(first, "border-top-style")).toBe("solid");
  expect(await styleOf(second, "border-top-style")).toBe("none");
  expect(await styleOf(first, "border-bottom-style")).toBe("none");
  expect(await styleOf(second, "border-bottom-style")).toBe("solid");
});

test("Page setup's Borders tab sets a page border and background (AC-19..AC-23)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("A certificate");

  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("page-setup-button").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeVisible();

  await page.getByTestId("page-setup-tab-borders").click();
  for (const edge of ["top", "bottom", "left", "right"]) {
    await page.getByTestId(`page-border-edge-${edge}`).check();
  }
  await page.getByTestId("page-border-style").selectOption("double");
  await page.getByTestId("page-border-measure-from").selectOption("text");
  await page.getByTestId("page-background-on").check();
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeHidden();

  const border = page.getByTestId("page-border").first();
  await expect(border).toBeAttached();
  expect(await border.evaluate((el) => globalThis.getComputedStyle(el).borderTopStyle)).toBe("double");
  await expect(page.getByTestId("page-background").first()).toBeAttached();
});

test("the print preview offers 'print background colours', off by default (AC-25)", async ({ page }) => {
  await openFixture(page, "word-borders.docx");
  await page.keyboard.press(`${MOD}+p`);
  await expect(page.getByTestId("print-preview")).toBeVisible();

  const toggle = page.getByTestId("print-background");
  await expect(toggle).not.toBeChecked();
  // Off means the background layer is excluded from print — the body carries
  // no marker class, so App.vue's @media print rule hides it.
  expect(await page.evaluate(() => document.body.classList.contains("prints-background"))).toBe(false);

  await toggle.check();
  expect(await page.evaluate(() => document.body.classList.contains("prints-background"))).toBe(true);
});

test("a bordered paragraph split across a page break keeps its sides, top and bottom right (AC-14..AC-16)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();

  // One paragraph long enough to run well past a page boundary. Inserted in
  // one go rather than typed, so the test spends its time on the assertion
  // rather than on thousands of keystrokes.
  const sentence = "This bordered paragraph is deliberately long enough to run past the bottom of the first page. ";
  await page.keyboard.insertText(sentence.repeat(90));
  await waitForPagination(page);

  const target = paragraphStarting(page, "This bordered paragraph is deliberately long");
  await target.click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-apply").click();
  await waitForPagination(page);

  // The paragraph really did split: one fragment continues, one does not.
  const fragments = content(page).locator("p[data-block]").filter({ hasText: "This bordered paragraph is deliberately long" });
  expect(await fragments.count()).toBeGreaterThanOrEqual(2);
  const firstFragment = fragments.first();
  const lastFragment = fragments.last();
  await expect(firstFragment).toHaveAttribute("data-slice-more", "1");
  await expect(lastFragment).not.toHaveAttribute("data-slice-more", "1");

  // AC-14: the side borders draw on BOTH fragments.
  for (const fragment of [firstFragment, lastFragment]) {
    expect(await styleOf(fragment, "border-left-style")).toBe("solid");
    expect(await styleOf(fragment, "border-right-style")).toBe("solid");
  }
  // AC-15: the top border only on the fragment holding the first line.
  expect(await styleOf(firstFragment, "border-top-style")).toBe("solid");
  expect(await styleOf(lastFragment, "border-top-style")).toBe("none");
  // AC-16: the bottom border only on the fragment holding the last line.
  expect(await styleOf(firstFragment, "border-bottom-style")).toBe("none");
  expect(await styleOf(lastFragment, "border-bottom-style")).toBe("solid");
});
