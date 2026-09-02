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

  // The tint really is on screen, and it covers the whole sheet (AC-24).
  await expect(background).toBeVisible();
  const layer = await background.boundingBox();
  const sheet = await page.locator(".page").first().boundingBox();
  expect(layer).not.toBeNull();
  expect(sheet).not.toBeNull();
  expect(Math.abs(layer!.width - sheet!.width)).toBeLessThan(1);
  expect(Math.abs(layer!.height - sheet!.height)).toBeLessThan(1);

  // ...and it paints ABOVE the page's own white paper rather than behind it.
  // The layer sits at z-index -1, which only stays inside the sheet because
  // .page is its own stacking context; drop that and the colour vanishes
  // behind the paper while every computed-style assertion above still
  // passes. Assert the condition the browser actually decides this on.
  const stacking = await page.locator(".page").first().evaluate((el) => {
    const style = globalThis.getComputedStyle(el);
    return { position: style.position, zIndex: style.zIndex, background: style.backgroundColor };
  });
  expect(stacking.position).not.toBe("static");
  expect(stacking.zIndex).not.toBe("auto");
  expect(await background.evaluate((el) => globalThis.getComputedStyle(el).zIndex)).toBe("-1");
  // The paper underneath stays opaque white, so a translucent fill still
  // reads against paper rather than against the workspace.
  expect(stacking.background).toBe("rgb(255, 255, 255)");
});

test("a merged box with no between rule has no gap down its sides (AC-12)", async ({ page }) => {
  await createDocument(page);
  const editor = content(page);
  await editor.click();
  // Three paragraphs in a style that carries paragraph spacing — the case a
  // box has to survive: with space-after left standing between two members,
  // the left and right rules break at every internal boundary and the page
  // shows through the side of the box.
  for (const [index, text] of ["Box first", "Box second", "Box third"].entries()) {
    if (index > 0) {
      await page.keyboard.press("Enter");
    }
    await page.keyboard.type(text);
  }
  await page.keyboard.down("Shift");
  await page.keyboard.press(`${MOD}+Home`);
  await page.keyboard.up("Shift");

  // "Add space before paragraph" puts an explicit 12pt on each of the three
  // (SPEC-028's space-before toggle).
  await page.getByTestId("line-spacing-button").click();
  await page.getByTestId("space-before-toggle").click();
  await waitForPagination(page);

  // The "Box" preset, NOT "All": no between edge anywhere in this run.
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-apply").click();
  await waitForPagination(page);

  const first = paragraphStarting(page, "Box first");
  const second = paragraphStarting(page, "Box second");
  const third = paragraphStarting(page, "Box third");

  // The spacing really is on the paragraphs — otherwise this test would pass
  // for the wrong reason. It survives on the run's outer edge (12pt = 16px)
  // and is collapsed at both internal boundaries.
  expect(await styleOf(first, "margin-top")).toBe("16px");
  expect(await styleOf(second, "margin-top")).toBe("0px");
  expect(await styleOf(third, "margin-top")).toBe("0px");
  // ...and no internal rule is drawn, so nothing but the sides joins them.
  expect(await styleOf(first, "border-bottom-style")).toBe("none");
  expect(await styleOf(second, "border-top-style")).toBe("none");

  // The box is continuous: each member's border box ends exactly where the
  // next one begins.
  const boxes = [];
  for (const paragraph of [first, second, third]) {
    const box = await paragraph.boundingBox();
    expect(box).not.toBeNull();
    boxes.push(box!);
  }
  for (let i = 0; i + 1 < boxes.length; i += 1) {
    const gap = boxes[i + 1].y - (boxes[i].y + boxes[i].height);
    expect(Math.abs(gap)).toBeLessThan(1);
    // ...and the sides line up, so the rules are one continuous pair.
    expect(Math.abs(boxes[i + 1].x - boxes[i].x)).toBeLessThan(1);
    expect(Math.abs(boxes[i + 1].width - boxes[i].width)).toBeLessThan(1);
  }
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

test("the style editor sets borders and shading on a style (AC-17, AC-18)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Styled by a callout");

  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-styles-pane").click();
  await expect(page.getByTestId("styles-pane")).toBeVisible();

  await page.getByTestId("new-style-paragraph").click();
  await expect(page.getByTestId("style-editor-dialog")).toBeVisible();
  await page.getByTestId("style-name-input").fill("Callout");

  // The same Borders and shading dialog a paragraph's are edited by, mounted
  // inside the style editor's own dialog. Escape must close the inner dialog
  // and leave the editor standing — two nested modals, one dismissal.
  await page.getByTestId("style-borders-button").click();
  const borders = page.getByTestId("borders-shading-dialog");
  await expect(borders).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(borders).toBeHidden();
  await expect(page.getByTestId("style-editor-dialog")).toBeVisible();

  await page.getByTestId("style-borders-button").click();
  await expect(borders).toBeVisible();
  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-shading-on").check();
  await page.getByTestId("borders-apply").click();
  await expect(borders).toBeHidden();
  await page.getByTestId("style-editor-save").click();
  await expect(page.getByTestId("style-editor-dialog")).not.toBeVisible();

  // The paragraph inherits the style's borders and shading until it sets its
  // own (AC-17).
  const paragraph = paragraphStarting(page, "Styled by a callout");
  await paragraph.click();
  await page.locator('[data-testid^="style-apply-"]').filter({ hasText: "Callout" }).first().click();
  await waitForPagination(page);
  expect(await styleOf(paragraph, "border-left-style")).toBe("solid");
  expect(await styleOf(paragraph, "border-top-style")).toBe("solid");
  expect(await styleOf(paragraph, "background-color")).not.toBe("rgba(0, 0, 0, 0)");
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

test("page backgrounds are hidden on paper unless the switch is on (AC-25)", async ({ page }) => {
  await openFixture(page, "word-borders.docx");
  const background = page.getByTestId("page-background").first();
  await expect(background).toBeAttached();
  const displayOf = () => background.evaluate((el) => globalThis.getComputedStyle(el).display);

  // On screen the layer is painted; under print media it is not. This is the
  // assertion the body-class check cannot make: it observes the @media print
  // rule resolving, not just the class that keys it.
  await page.emulateMedia({ media: "print" });
  expect(await displayOf()).toBe("none");
  // The page BORDER has no such switch — it always prints.
  expect(
    await page.getByTestId("page-border").first().evaluate((el) => globalThis.getComputedStyle(el).display),
  ).not.toBe("none");

  await page.emulateMedia({ media: "screen" });
  expect(await displayOf()).not.toBe("none");

  // Turning "print background colours" on brings it back on paper.
  await page.keyboard.press(`${MOD}+p`);
  await expect(page.getByTestId("print-preview")).toBeVisible();
  await page.getByTestId("print-background").check();
  await page.emulateMedia({ media: "print" });
  expect(await displayOf()).toBe("block");
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

test("reopening Borders and shading after only moving the caret prefills from the NEW paragraph (AC-6, AC-11)", async ({ page }) => {
  await createDocument(page);
  const editor = content(page);
  await editor.click();
  await page.keyboard.type("Boxed paragraph");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Plain paragraph");

  const boxed = paragraphStarting(page, "Boxed paragraph");
  const plain = paragraphStarting(page, "Plain paragraph");
  const dialog = page.getByTestId("borders-shading-dialog");

  async function openDialog(): Promise<void> {
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("format-menu-borders").click();
    await expect(dialog).toBeVisible();
  }

  // Box and fill the first paragraph only.
  await boxed.click();
  await openDialog();
  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-shading-on").check();
  await page.getByTestId("borders-apply").click();
  await expect(dialog).toBeHidden();
  expect(await styleOf(boxed, "border-left-style")).toBe("solid");

  // Inspect the boxed paragraph and cancel — the seed now reads "box".
  await boxed.click();
  await openDialog();
  await expect(page.getByTestId("borders-preset-box")).toHaveAttribute("aria-pressed", "true");
  await page.getByTestId("borders-cancel").click();
  await expect(dialog).toBeHidden();

  // The regression: move the caret to the plain paragraph with NO edit in
  // between and reopen. The dialog must show what THIS paragraph has —
  // nothing — not the previous paragraph's box and fill.
  await plain.click();
  await openDialog();
  await expect(page.getByTestId("borders-preset-none")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("borders-shading-on")).not.toBeChecked();

  // ...so an OK that only adds a fill does not also stamp a box onto it.
  await page.getByTestId("borders-shading-on").check();
  await page.getByTestId("borders-apply").click();
  await expect(dialog).toBeHidden();
  expect(await styleOf(plain, "border-left-style")).toBe("none");
  expect(await styleOf(plain, "background-color")).not.toBe("rgba(0, 0, 0, 0)");
  expect(await styleOf(boxed, "border-left-style")).toBe("solid");
});
