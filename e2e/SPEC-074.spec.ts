import { test, expect, type Download, type Page } from "@playwright/test";
import { MOD } from "./keys";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
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
  // The URL changes the moment the document is stored — before the parse, and
  // well before layout: without a barrier here a test reads state the app has
  // not built yet. Same pair SPEC-030/051 wait on.
  await expect(page.getByTestId("document-opening")).toBeHidden();
  await waitForPagination(page);
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

async function waitForSaved(page: Page): Promise<void> {
  await expect(page.getByTestId("save-state")).toHaveText("All changes saved", { timeout: 15_000 });
}

/** File → Export → the chosen format, returning the download. */
async function exportAs(page: Page, format: "docx" | "odt"): Promise<Download> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const pending = page.waitForEvent("download");
  await page.getByTestId(`download-${format}`).click();
  return pending;
}

/** Minimal central-directory walk over a downloaded archive, so a test can
 * read the OOXML/ODF it just wrote (same shape as e2e/SPEC-008.spec.ts). */
function unzipParts(buffer: Buffer): Map<string, string> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd, "end-of-central-directory record should exist").toBeGreaterThanOrEqual(0);
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const parts = new Map<string, string>();
  for (let i = 0; i < totalEntries; i++) {
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.toString("utf8", offset + 46, offset + 46 + nameLength);
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const data = buffer.subarray(dataStart, dataStart + compressedSize);
    parts.set(name, (method === 8 ? inflateRawSync(data) : Buffer.from(data)).toString("utf8"));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return parts;
}

async function partsOf(download: Download): Promise<Map<string, string>> {
  const path = await download.path();
  expect(path).toBeTruthy();
  return unzipParts(readFileSync(path!));
}

/** Every block's painted borders, shading, padding and indent, plus the page
 * decoration — the observable stand-in for `equalStructural` (AC-33). */
async function paintSnapshot(page: Page): Promise<{ blocks: string[]; page: string }> {
  return page.evaluate(() => {
    const host = document.querySelector('[data-testid="document-content"]')!;
    const blocks = Array.from(host.querySelectorAll("p, h1, h2, h3")).map((block) => {
      const style = globalThis.getComputedStyle(block);
      return [
        (block.textContent ?? "").slice(0, 20),
        [style.borderTopStyle, style.borderRightStyle, style.borderBottomStyle, style.borderLeftStyle].join(),
        [style.borderTopWidth, style.borderRightWidth, style.borderBottomWidth, style.borderLeftWidth].join(),
        [style.borderTopColor, style.borderRightColor, style.borderBottomColor, style.borderLeftColor].join(),
        style.backgroundColor,
        [style.paddingTop, style.paddingRight, style.paddingBottom, style.paddingLeft].join(),
        style.marginInlineStart,
      ].join(" | ");
    });
    const background = document.querySelector('[data-testid="page-background"]');
    const border = document.querySelector('[data-testid="page-border"]');
    const borderStyle = border && globalThis.getComputedStyle(border);
    return {
      blocks,
      page: [
        background ? globalThis.getComputedStyle(background).backgroundColor : "no background",
        borderStyle
          ? [borderStyle.borderTopStyle, borderStyle.borderTopWidth, borderStyle.borderTopColor, borderStyle.top].join()
          : "no border",
      ].join(" | "),
    };
  });
}

/** Reopen a file the app just wrote, from the library's file input. */
async function reopen(page: Page, download: Download): Promise<void> {
  const path = await download.path();
  expect(path).toBeTruthy();
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path!);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page).locator("p, h1").first()).toBeVisible();
  await waitForPagination(page);
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

test("a paragraph inherits its style borders and shading until it sets its own (AC-17, AC-18)", async ({
  page,
}) => {
  const consoleErrors: string[] = [];
  page.on("console", (message) => {
    if (message.type() === "error") {
      consoleErrors.push(message.text());
    }
  });

  await createDocument(page);
  const editor = content(page);
  await editor.click();
  await page.keyboard.type("Inherits the paragraph style box");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Sets its own box");

  const inherits = paragraphStarting(page, "Inherits the paragraph style box");
  const own = paragraphStarting(page, "Sets its own box");

  // The second paragraph gets a box of its own first, so the inherited one
  // below cannot be confused for it.
  await own.click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await page.getByTestId("borders-preset-topAndBottom").click();
  await page.getByTestId("borders-style-top").selectOption("dotted");
  await page.getByTestId("borders-apply").click();
  await waitForPagination(page);
  expect(await styleOf(inherits, "border-top-style")).toBe("none");

  // AC-18: set Normal's own borders and shading through the style editor.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-styles-pane").click();
  await expect(page.getByTestId("styles-pane")).toBeVisible();
  await page.getByTestId("style-modify-Normal").click();
  await expect(page.getByTestId("style-editor-dialog")).toBeVisible();
  await page.getByTestId("style-borders-button").click();
  await expect(page.getByTestId("borders-shading-dialog")).toBeVisible();
  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-shading-on").check();
  await page.getByTestId("borders-shading-color").fill("#fff2cc");
  await page.getByTestId("borders-apply").click();
  await page.getByTestId("style-editor-save").click();
  await expect(page.getByTestId("style-editor-dialog")).not.toBeVisible();
  await waitForPagination(page);

  // AC-17: the paragraph that set nothing of its own now paints Normal's box
  // and fill — this is a paragraph with no styleId at all, the implicit
  // Normal every blank document's paragraphs carry.
  expect(await styleOf(inherits, "border-top-style")).toBe("solid");
  expect(await styleOf(inherits, "border-left-style")).toBe("solid");
  expect(await styleOf(inherits, "background-color")).toBe("rgb(255, 242, 204)");

  // ...and the paragraph that set its own keeps them: direct formatting wins.
  expect(await styleOf(own, "border-top-style")).toBe("dotted");
  expect(await styleOf(own, "border-left-style")).toBe("none");
  // Shading is a separate property, so it still inherits the style's fill.
  expect(await styleOf(own, "background-color")).toBe("rgb(255, 242, 204)");

  // Saving a style with borders used to put a Vue reactive proxy into the
  // document model, which the live word-count worker could not clone.
  expect(consoleErrors.join("\n")).not.toContain("DataCloneError");
  expect(consoleErrors).toEqual([]);
});

test("a paragraph can switch its style's box and fill back off (AC-17)", async ({ page }) => {
  await createDocument(page);
  const editor = content(page);
  await editor.click();
  await page.keyboard.type("Turns the box off");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Keeps the inherited box");

  const off = paragraphStarting(page, "Turns the box off");
  const kept = paragraphStarting(page, "Keeps the inherited box");
  const dialog = page.getByTestId("borders-shading-dialog");

  // Normal draws a box and a fill, so every paragraph in this document
  // inherits both — no paragraph sets anything of its own.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-styles-pane").click();
  await page.getByTestId("style-modify-Normal").click();
  await page.getByTestId("style-borders-button").click();
  await expect(dialog).toBeVisible();
  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-shading-on").check();
  await page.getByTestId("borders-shading-color").fill("#ffd9d9");
  await page.getByTestId("borders-apply").click();
  await page.getByTestId("style-editor-save").click();
  await waitForPagination(page);
  expect(await styleOf(off, "border-left-style")).toBe("solid");
  expect(await styleOf(off, "background-color")).toBe("rgb(255, 217, 217)");

  // The dialog opens showing what the paragraph PAINTS, box included — were
  // it to show "None" here, OK would strip a box the user never touched.
  await off.click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await expect(dialog).toBeVisible();
  await expect(page.getByTestId("borders-preset-box")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("borders-shading-on")).toBeChecked();

  // The failure this test exists for: None + cleared Fill + OK used to leave
  // the inherited box and fill painted, because dropping the paragraph's
  // (absent) value simply re-inherited the style's.
  await page.getByTestId("borders-preset-none").click();
  await page.getByTestId("borders-shading-on").uncheck();
  await page.getByTestId("borders-apply").click();
  await expect(dialog).toBeHidden();
  await waitForPagination(page);

  for (const edge of ["top", "bottom", "left", "right"]) {
    expect(await styleOf(off, `border-${edge}-style`)).toBe("none");
  }
  expect(await styleOf(off, "background-color")).toBe("rgba(0, 0, 0, 0)");

  // It is an override on ONE paragraph, not a change to the style: the
  // neighbour still inherits.
  expect(await styleOf(kept, "border-left-style")).toBe("solid");
  expect(await styleOf(kept, "background-color")).toBe("rgb(255, 217, 217)");

  // Reopening shows the paragraph as unbordered and unfilled...
  await off.click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await expect(page.getByTestId("borders-preset-none")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("borders-shading-on")).not.toBeChecked();
  await page.getByTestId("borders-cancel").click();

  // ...and the override survives the trip through storage, so it is a real
  // value in the model rather than a transient render decision. Wait for the
  // save indicator rather than a wall-clock sleep: the debounce is what this
  // reload has to land after, and the indicator is what reports it.
  await waitForSaved(page);
  await page.reload();
  await expect(content(page).locator("p").first()).toBeVisible();
  await waitForPagination(page);
  expect(await styleOf(paragraphStarting(page, "Turns the box off"), "border-left-style")).toBe("none");
  expect(await styleOf(paragraphStarting(page, "Turns the box off"), "background-color")).toBe(
    "rgba(0, 0, 0, 0)",
  );
  expect(await styleOf(paragraphStarting(page, "Keeps the inherited box"), "border-left-style")).toBe(
    "solid",
  );
});

test("paragraph and page borders round-trip through .docx (AC-28, AC-30, AC-33)", async ({ page }) => {
  await openFixture(page, "word-borders.docx");
  await waitForPagination(page);
  const before = await paintSnapshot(page);
  // The fixture really does carry all five elements this AC names.
  expect(before.blocks).toHaveLength(9);
  expect(before.page).toContain("rgb(253, 246, 227)");

  const saved = await exportAs(page, "docx");
  const xml = (await partsOf(saved)).get("word/document.xml") ?? "";

  // AC-28: w:pBdr with the five edges, @w:val, @w:sz in EIGHTHS of a point,
  // @w:space in POINTS and @w:color; w:shd carrying @w:fill.
  expect(xml).toContain('<w:top w:val="single" w:sz="8" w:space="4" w:color="7F7F7F"/>');
  expect(xml).toContain('<w:left w:val="single" w:sz="8" w:space="4" w:color="7F7F7F"/>');
  expect(xml).toContain('<w:bottom w:val="single" w:sz="8" w:space="4" w:color="7F7F7F"/>');
  expect(xml).toContain('<w:right w:val="single" w:sz="8" w:space="4" w:color="7F7F7F"/>');
  expect(xml).toContain('<w:between w:val="dotted" w:sz="4" w:space="0" w:color="7F7F7F"/>');
  expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>');
  // sz=8 eighths is the 1px/1pt edge the note paints; sz=24 the 3pt heading rule.
  expect(xml).toContain('<w:bottom w:val="double" w:sz="24" w:space="2" w:color="C00000"/>');

  // AC-30: w:pgBorders inside w:sectPr, w:background on the document.
  expect(xml).toMatch(
    /<w:sectPr[ >][\s\S]*<w:pgBorders w:offsetFrom="page">[\s\S]*<\/w:pgBorders>[\s\S]*<\/w:sectPr>/,
  );
  expect(xml).toContain('<w:background w:color="FDF6E3"/>');

  // AC-33: reopening renders identically, block for block.
  await reopen(page, saved);
  expect(await paintSnapshot(page)).toEqual(before);
});

test("paragraph and page borders round-trip through .odt (AC-29, AC-31, AC-33)", async ({ page }) => {
  await openFixture(page, "word-borders.docx");
  await waitForPagination(page);
  const before = await paintSnapshot(page);

  const saved = await exportAs(page, "odt");
  const parts = await partsOf(saved);
  const content = parts.get("content.xml") ?? "";
  const styles = parts.get("styles.xml") ?? "";

  // AC-29: fo:border*, fo:padding*, fo:background-color and style:join-border.
  expect(content).toContain('fo:border-top="1pt solid #7f7f7f"');
  expect(content).toContain('fo:padding-top="4pt"');
  expect(content).toContain('fo:background-color="#fff2cc"');
  expect(content).toContain('style:join-border="true"');
  // ODF has no between edge, so it is preserved rather than dropped.
  expect(content).toContain('de:border-between="0.5pt dotted #7f7f7f"');

  // AC-31: the page border and background live on the page layout the master
  // page points at.
  expect(styles).toContain('fo:border-top="3pt double #1f4e79"');
  expect(styles).toContain('fo:background-color="#fdf6e3"');
  expect(styles).toMatch(/<style:master-page style:name="Standard" style:page-layout-name="pm1"\/>/);

  // AC-33: and it comes back looking the same.
  await reopen(page, saved);
  expect(await paintSnapshot(page)).toEqual(before);
});

test("a dash-dot border saved as .odt maps to its nearest equivalent and is reported (AC-32)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("A dash-dot ruled paragraph");

  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await page.getByTestId("borders-style-top").selectOption("dashDot");
  await page.getByTestId("borders-apply").click();
  await waitForPagination(page);

  const saved = await exportAs(page, "odt");
  // ODF's keyword set has no dash-dot, so the writer picks the nearest one...
  expect((await partsOf(saved)).get("content.xml") ?? "").toContain('fo:border-top="1pt dashed #000000"');
  // ...and says so, rather than silently changing the document (DropRecord).
  const report = page.getByTestId("conversion-report");
  await expect(report).toBeVisible();
  await expect(report).toContainText("Paragraph formatting");
  await expect(report).toContainText("nearest equivalent, dashed");

  // The .docx writer, which HAS the style, keeps it: dash-dot is w:val="dotDash".
  await page.getByTestId("conversion-report-close").click();
  expect((await partsOf(await exportAs(page, "docx"))).get("word/document.xml") ?? "").toContain(
    'w:val="dotDash"',
  );
});

test("formatting marks and the accessibility checker read the shading, not the page (AC-26, AC-27)", async ({
  page,
}) => {
  await openFixture(page, "word-borders.docx");
  await waitForPagination(page);
  const note = paragraphStarting(page, "Note: this paragraph is boxed");
  const plain = paragraphStarting(page, "An ordinary paragraph");

  // AC-26: with marks on, only a shaded paragraph gets a mark colour of its
  // own, and that colour is computed from ITS fill.
  await page.keyboard.press(`${MOD}+Shift+p`);
  await expect(content(page)).toHaveClass(/marks-on/);
  const markColor = (locator: ReturnType<typeof paragraphStarting>) =>
    locator.evaluate((el) => globalThis.getComputedStyle(el).getPropertyValue("--mark-color").trim());
  expect(await markColor(plain)).toBe("");
  const lightFillMark = await markColor(note);
  expect(lightFillMark).not.toBe("");

  // Flip the fill to near-black: the marks must flip to a light colour, which
  // they can only do by measuring against the shading.
  await note.click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await page.getByTestId("borders-shading-color").fill("#1f1f1f");
  await page.getByTestId("borders-apply").click();
  await waitForPagination(page);
  expect(await styleOf(note, "background-color")).toBe("rgb(31, 31, 31)");
  const darkFillMark = await markColor(note);
  expect(darkFillMark).not.toBe(lightFillMark);
  const luminance = (hex: string) => parseInt(hex.slice(1, 3), 16) + parseInt(hex.slice(3, 5), 16);
  expect(luminance(darkFillMark)).toBeGreaterThan(luminance(lightFillMark));

  // AC-27: the checker measures the text against that same near-black fill —
  // against the cream page it would be ~16:1 and clean.
  await page.getByTestId("menu-title-tools").click();
  await page.getByTestId("tools-menu-accessibility").click();
  await expect(page.getByTestId("accessibility-pane")).toBeVisible();
  await expect(page.getByTestId("accessibility-summary")).toHaveText(/1 issue/);
  await expect(page.getByTestId("accessibility-group-error")).toContainText("Low contrast text");

  // Negative control: put the light fill back and the report clears.
  await page.getByTestId("accessibility-pane-close").click();
  await note.click();
  await page.keyboard.press(`${MOD}+z`);
  await expect(async () => {
    expect(await styleOf(note, "background-color")).toBe("rgb(255, 242, 204)");
  }).toPass();
  await page.getByTestId("menu-title-tools").click();
  await page.getByTestId("tools-menu-accessibility").click();
  await expect(page.getByTestId("accessibility-summary")).toHaveText("No accessibility issues found");
});

test("every edge takes its own style, width and colour, in the document's unit (AC-3, AC-7, AC-8, AC-9, AC-10)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Four different edges");
  const target = paragraphStarting(page, "Four different edges");

  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  const dialog = page.getByTestId("borders-shading-dialog");
  await expect(dialog).toBeVisible();

  // AC-3: the three line styles this spec adds are offered on every edge.
  for (const edge of ["top", "bottom", "left", "right", "between"]) {
    const values = await page
      .getByTestId(`borders-style-${edge}`)
      .evaluate((el) => Array.from((el as HTMLSelectElement).options).map((option) => option.value));
    expect(values).toEqual(["none", "single", "thick", "double", "dashed", "dotted", "dashDot", "dashDotDot"]);
  }

  // AC-9: padding is per edge and defaults to 4pt — 0.14cm in the cm document.
  await expect(dialog).toContainText("Padding (cm)");
  for (const edge of ["top", "bottom", "left", "right"]) {
    await expect(page.getByTestId(`borders-padding-${edge}`)).toHaveValue("0.14");
  }

  // AC-7: four edges, four different styles, widths and colours.
  const edges = [
    { edge: "top", style: "thick", width: "3", color: "#ff0000" },
    { edge: "bottom", style: "dashDotDot", width: "2", color: "#0000ff" },
    { edge: "left", style: "dashDot", width: "1", color: "#00aa00" },
    { edge: "right", style: "double", width: "4", color: "#ff00ff" },
  ];
  for (const { edge, style, width, color } of edges) {
    await page.getByTestId(`borders-style-${edge}`).selectOption(style);
    await page.getByTestId(`borders-width-${edge}`).fill(width);
    await page.getByTestId(`borders-color-${edge}`).fill(color);
  }
  // AC-8: the edges no longer match any preset, so "Custom" takes over.
  await expect(page.getByTestId("borders-preset-custom")).toHaveAttribute("aria-pressed", "true");
  for (const preset of ["none", "box", "all", "topAndBottom"]) {
    await expect(page.getByTestId(`borders-preset-${preset}`)).toHaveAttribute("aria-pressed", "false");
  }

  // AC-9: a padding set on one edge only.
  await page.getByTestId("borders-padding-top").fill("0.5");
  // AC-10: a white fill — a real colour, not "no fill".
  await page.getByTestId("borders-shading-on").check();
  await page.getByTestId("borders-shading-color").fill("#ffffff");
  await page.getByTestId("borders-apply").click();
  await expect(dialog).toBeHidden();
  await waitForPagination(page);

  expect(await styleOf(target, "border-top-color")).toBe("rgb(255, 0, 0)");
  expect(await styleOf(target, "border-bottom-color")).toBe("rgb(0, 0, 255)");
  expect(await styleOf(target, "border-left-color")).toBe("rgb(0, 170, 0)");
  expect(await styleOf(target, "border-right-color")).toBe("rgb(255, 0, 255)");
  expect(await styleOf(target, "border-right-style")).toBe("double");
  expect(await styleOf(target, "border-left-width")).toBe("1px");
  expect(await styleOf(target, "border-bottom-width")).toBe("2px");
  // 0.5cm = 14pt on top; the untouched edges keep the 4pt default.
  expect(await styleOf(target, "padding-top")).toBe("18.6667px");
  expect(await styleOf(target, "padding-left")).toBe("5.33333px");
  expect(await styleOf(target, "background-color")).toBe("rgb(255, 255, 255)");

  // AC-9: switch the document's unit and the padding fields follow it — the
  // same 4pt default, now stated in points.
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("page-setup-button").click();
  await page.getByTestId("page-setup-unit").selectOption("pt");
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeHidden();

  await target.click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await expect(dialog).toContainText("Padding (pt)");
  await expect(page.getByTestId("borders-padding-top")).toHaveValue("14");
  await expect(page.getByTestId("borders-padding-left")).toHaveValue("4");

  // AC-10: clearing the checkbox is a THIRD state, distinct from white.
  await page.getByTestId("borders-shading-on").uncheck();
  await page.getByTestId("borders-apply").click();
  await expect(dialog).toBeHidden();
  expect(await styleOf(target, "background-color")).toBe("rgba(0, 0, 0, 0)");
  // ...and it did not take the borders with it.
  expect(await styleOf(target, "border-top-color")).toBe("rgb(255, 0, 0)");
});

test("page borders choose their edges and measure from the page or the text (AC-20, AC-21)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("A certificate");

  const openBordersTab = async () => {
    await page.getByTestId("menu-title-file").click();
    await page.getByTestId("page-setup-button").click();
    await expect(page.getByTestId("page-setup-dialog")).toBeVisible();
    await page.getByTestId("page-setup-tab-borders").click();
  };
  const border = page.getByTestId("page-border").first();
  const insetOf = () =>
    border.evaluate((el) => {
      const style = globalThis.getComputedStyle(el);
      return { top: style.top, left: style.left, edges: [style.borderTopStyle, style.borderRightStyle, style.borderBottomStyle, style.borderLeftStyle].join() };
    });

  await openBordersTab();
  for (const edge of ["top", "bottom", "left", "right"]) {
    await page.getByTestId(`page-border-edge-${edge}`).check();
  }
  await page.getByTestId("page-border-style").selectOption("dashed");
  await page.getByTestId("page-border-measure-from").selectOption("page");
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeHidden();
  await expect(border).toBeAttached();

  // Measured from the page edge: the offset is all that separates them.
  const fromPage = await insetOf();
  expect(fromPage.edges).toBe("dashed,dashed,dashed,dashed");
  expect(fromPage.top).toBe("32px");

  // AC-21: measured from the text, the border moves in to the text area —
  // the margin (72pt) less the offset (24pt).
  await openBordersTab();
  await page.getByTestId("page-border-measure-from").selectOption("text");
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeHidden();
  await expect(async () => {
    expect((await insetOf()).top).toBe("64px");
  }).toPass();

  // AC-20: the edge checkboxes decide which sides are drawn at all.
  await openBordersTab();
  await page.getByTestId("page-border-edge-left").uncheck();
  await page.getByTestId("page-border-edge-right").uncheck();
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeHidden();
  await expect(async () => {
    expect((await insetOf()).edges).toBe("dashed,none,dashed,none");
  }).toPass();
});

test("switching an inherited box off survives both file formats, and a plain document brings back no husks (AC-17)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Turns the box off");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Keeps the inherited box");

  // Normal draws a box and a fill; every paragraph inherits both.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-styles-pane").click();
  await page.getByTestId("style-modify-Normal").click();
  await page.getByTestId("style-borders-button").click();
  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-shading-on").check();
  await page.getByTestId("borders-shading-color").fill("#ffd9d9");
  await page.getByTestId("borders-apply").click();
  await page.getByTestId("style-editor-save").click();
  await waitForPagination(page);

  // One paragraph switches both back off.
  await paragraphStarting(page, "Turns the box off").click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-borders").click();
  await page.getByTestId("borders-preset-none").click();
  await page.getByTestId("borders-shading-on").uncheck();
  await page.getByTestId("borders-apply").click();
  await waitForPagination(page);

  const assertOverrideHeld = async () => {
    const off = paragraphStarting(page, "Turns the box off");
    const kept = paragraphStarting(page, "Keeps the inherited box");
    for (const edge of ["top", "bottom", "left", "right"]) {
      expect(await styleOf(off, `border-${edge}-style`)).toBe("none");
    }
    expect(await styleOf(off, "background-color")).toBe("rgba(0, 0, 0, 0)");
    expect(await styleOf(kept, "border-left-style")).toBe("solid");
    expect(await styleOf(kept, "background-color")).toBe("rgb(255, 217, 217)");
  };
  await assertOverrideHeld();

  // The "off" is a value in the file, not a render-time decision: .docx says
  // so with an all-nil w:pBdr and a fill-less w:shd.
  const asDocx = await exportAs(page, "docx");
  const docxXml = (await partsOf(asDocx)).get("word/document.xml") ?? "";
  expect(docxXml).toContain('<w:top w:val="nil" w:sz="0" w:space="0" w:color="auto"/>');
  expect(docxXml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
  await reopen(page, asDocx);
  await assertOverrideHeld();

  const asOdt = await exportAs(page, "odt");
  await reopen(page, asOdt);
  await assertOverrideHeld();

  // The other half of the same judgement: a document that never had borders
  // must not come back carrying "off" husks, or it could never pick a style's
  // box up afterwards.
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Plain before the round trip");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Plain after it too");
  const plainSave = await exportAs(page, "docx");
  expect((await partsOf(plainSave)).get("word/document.xml") ?? "").not.toContain("<w:pBdr>");

  await reopen(page, plainSave);
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-styles-pane").click();
  await page.getByTestId("style-modify-Normal").click();
  await page.getByTestId("style-borders-button").click();
  await page.getByTestId("borders-preset-box").click();
  await page.getByTestId("borders-apply").click();
  await page.getByTestId("style-editor-save").click();
  await waitForPagination(page);
  expect(await styleOf(paragraphStarting(page, "Plain before the round trip"), "border-left-style")).toBe(
    "solid",
  );
  expect(await styleOf(paragraphStarting(page, "Plain after it too"), "border-left-style")).toBe("solid");
});
