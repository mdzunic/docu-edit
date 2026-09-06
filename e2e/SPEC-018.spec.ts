import { test, expect, type Download, type Page } from "@playwright/test";
import { MOD, LINE_END, LINE_START } from "./keys";
import { inflateRawSync } from "node:zlib";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-018 — bulleted and numbered lists: toolbar toggles and marker
// dropdowns, Tab/Shift+Tab levels, Enter/Backspace behaviours, autoformat,
// derived numbering with restart-on-interruption, and .docx/.odt round
// trips. Markers are asserted through data-list-marker — the attribute the
// canvas paints the marker from — and indents through the computed style,
// so the assertions observe exactly what renders.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxFixture = join(fixturesDir, "lists.docx");
const odtFixture = join(fixturesDir, "lists.odt");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const paragraphAt = (page: Page, index: number) =>
  content(page).locator(".doc-paragraph").nth(index);

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
  // The URL changes the moment the document is stored — before the parse, and
  // well before layout: without a barrier here a test reads state the app has
  // not built yet. Same pair SPEC-030/051 wait on.
  await expect(page.getByTestId("document-opening")).toBeHidden();
  await expect(page.getByTestId("document-pages")).toHaveAttribute("data-pagination-state", "clean");
}

/** Downloads the document in the given format and returns the file path. */
async function download(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const file: Download = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId(`download-${format}`).click(),
  ]).then(([d]) => d);
  const path = await file.path();
  expect(path).toBeTruthy();
  return path!;
}

/** Minimal central-directory walk over a downloaded archive — the same
 * helper as e2e/SPEC-006.spec.ts. */
function unzipParts(buffer: Buffer): Map<string, Buffer> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd).toBeGreaterThanOrEqual(0);
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const parts = new Map<string, Buffer>();
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
    parts.set(name, method === 8 ? inflateRawSync(data) : Buffer.from(data));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return parts;
}

/** The computed physical start margin — how the per-level indent renders. */
async function marginPx(page: Page, index: number): Promise<number> {
  return paragraphAt(page, index).evaluate((el) =>
    Number.parseFloat(getComputedStyle(el).marginInlineStart || "0"),
  );
}

// --- toolbar toggles (AC-1, AC-2) ------------------------------------------

test("the bulleted toggle applies to every paragraph in the selection and toggles off when all are bulleted (AC-1)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("one");
  await page.keyboard.press("Enter");
  await page.keyboard.type("two");
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTestId("list-bulleted").click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "•");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "•");
  await expect(page.getByTestId("list-bulleted")).toHaveAttribute("aria-pressed", "true");

  // All bulleted -> toggling removes the formatting from all of them.
  await page.getByTestId("list-bulleted").click();
  await expect(paragraphAt(page, 0)).not.toHaveAttribute("data-list-marker");
  await expect(paragraphAt(page, 1)).not.toHaveAttribute("data-list-marker");
  await expect(page.getByTestId("list-bulleted")).toHaveAttribute("aria-pressed", "false");
});

test("toggling with only some paragraphs bulleted applies to all (AC-1)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("one");
  await page.keyboard.press("Enter");
  await page.keyboard.type("two");
  // Only the second paragraph (the caret's) becomes bulleted first.
  await page.getByTestId("list-bulleted").click();
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "•");
  await expect(paragraphAt(page, 0)).not.toHaveAttribute("data-list-marker");
  // Some-but-not-all: the toggle applies to the whole selection.
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTestId("list-bulleted").click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "•");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "•");
});

test("the numbered toggle numbers the selection and toggles off with the same semantics (AC-2, AC-10)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("first");
  await page.keyboard.press("Enter");
  await page.keyboard.type("second");
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTestId("list-numbered").click();
  // Continuous numbering across consecutive items (AC-10).
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "2.");
  await page.getByTestId("list-numbered").click();
  await expect(paragraphAt(page, 0)).not.toHaveAttribute("data-list-marker");
  await expect(paragraphAt(page, 1)).not.toHaveAttribute("data-list-marker");
});

// --- Tab / Shift+Tab (AC-3, AC-4, AC-5) ------------------------------------
//
// AC-3's "inside a list item" is the caret collapsed at offset 0 of the item:
// SPEC-054's Tab precedence table gives the demote rule 2 (offset 0) and a
// literal tab rule 4 (anywhere else in the paragraph), so a Tab pressed
// mid-text advances to a tab stop instead of demoting. Clicking a paragraph
// lands the caret wherever the click was, so these tests press LINE_START
// first to sit where the demote binding lives. The levels, markers, clamping
// and round trips asserted below are unchanged.

test("Tab demotes and Shift+Tab promotes a list item, clamped to levels 1..9 (AC-3, AC-4)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("item");
  await page.getByTestId("list-bulleted").click();
  await content(page).click();
  await page.keyboard.press(LINE_START);

  await page.keyboard.press("Tab");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-level", "2");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "◦");

  await page.keyboard.press("Shift+Tab");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-level", "1");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "•");

  // Shift+Tab at level 1 has no effect, and the item keeps its text.
  await page.keyboard.press("Shift+Tab");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-level", "1");

  // Tab at level 9 has no effect.
  for (let i = 0; i < 10; i++) {
    await page.keyboard.press("Tab");
  }
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-level", "9");
  await expect(paragraphAt(page, 0)).toHaveText("item");
});

test("each level shows its distinct marker, cycling every three levels (AC-5)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("bullets");
  await page.getByTestId("list-bulleted").click();
  await content(page).click();
  await page.keyboard.press(LINE_START);
  const bulletCycle = ["•", "◦", "▪", "•", "◦", "▪", "•", "◦", "▪"];
  for (let level = 1; level <= 9; level++) {
    await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", bulletCycle[level - 1]);
    await page.keyboard.press("Tab");
  }
  // Numbered: levels 1..3 then the cycle repeats at level 4.
  await page.getByTestId("list-numbered").click();
  await content(page).click();
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 9; i++) {
    await page.keyboard.press("Shift+Tab");
  }
  const numberCycle = ["1.", "a.", "i.", "1."];
  for (let level = 1; level <= 4; level++) {
    await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", numberCycle[level - 1]);
    await page.keyboard.press("Tab");
  }
});

// --- Enter and Backspace (AC-6..AC-9) --------------------------------------

test("Enter on a non-empty item creates the next item at the same level with the caret in it (AC-6)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Apples");
  await page.getByTestId("list-numbered").click();
  await content(page).click();
  await page.keyboard.press(LINE_END);
  await page.keyboard.press("Enter");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "2.");
  await expect(paragraphAt(page, 1)).toHaveText("");
  // The caret sits in the new empty item.
  await page.keyboard.type("Bananas");
  await expect(paragraphAt(page, 1)).toHaveText("Bananas");
  await expect(paragraphAt(page, 0)).toHaveText("Apples");
});

test("Enter on an empty item outdents it at level > 1 and exits the list at level 1 (AC-7, AC-8)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("parent");
  await page.getByTestId("list-bulleted").click();
  await content(page).click();
  await page.keyboard.press(LINE_END);
  await page.keyboard.press("Enter");
  await page.keyboard.press("Tab");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-level", "2");

  // Empty item at level 2: Enter outdents, still a list item (AC-7).
  await page.keyboard.press("Enter");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-level", "1");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "•");
  await expect(content(page).locator(".doc-paragraph")).toHaveCount(2);

  // Empty item at level 1: Enter converts it to a normal paragraph (AC-8).
  await page.keyboard.press("Enter");
  await expect(paragraphAt(page, 1)).not.toHaveAttribute("data-list-marker");
  await expect(content(page).locator(".doc-paragraph")).toHaveCount(2);
  // The cursor stays in the now-normal paragraph.
  await page.keyboard.type("outside");
  await expect(paragraphAt(page, 1)).toHaveText("outside");
  await expect(paragraphAt(page, 1)).not.toHaveAttribute("data-list-marker");
});

test("Backspace at the start of an item removes the list formatting without deleting text (AC-9)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("before");
  await page.keyboard.press("Enter");
  await page.keyboard.type("keep me");
  await page.getByTestId("list-bulleted").click();
  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.press("Backspace");
  await expect(paragraphAt(page, 1)).not.toHaveAttribute("data-list-marker");
  await expect(paragraphAt(page, 1)).toHaveText("keep me");
  // Now a normal paragraph: Backspace at its start merges as usual.
  await page.keyboard.press("Backspace");
  await expect(paragraphAt(page, 0)).toHaveText("beforekeep me");
});

// --- numbering restart (AC-11) ---------------------------------------------

test("a non-list paragraph between items restarts the numbering after it (AC-11)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("one");
  await page.keyboard.press("Enter");
  await page.keyboard.type("middle");
  await page.keyboard.press("Enter");
  await page.keyboard.type("three");
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTestId("list-numbered").click();
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "3.");

  // Turn the middle item into a plain paragraph: the list is interrupted
  // and the item after it restarts at the default start value 1.
  await content(page).click();
  await paragraphAt(page, 1).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.press("Backspace");
  await expect(paragraphAt(page, 1)).not.toHaveAttribute("data-list-marker");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "1.");
});

// --- marker dropdowns (AC-12, AC-13) ---------------------------------------

test("each list button has a dropdown with the spec's marker styles (AC-12)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("list-bulleted-menu-button").click();
  const bulletMenu = page.getByTestId("bullet-style-menu");
  await expect(bulletMenu.getByTestId("bullet-style-disc")).toHaveText("•");
  await expect(bulletMenu.getByTestId("bullet-style-circle")).toHaveText("◦");
  await expect(bulletMenu.getByTestId("bullet-style-square")).toHaveText("▪");
  await page.keyboard.press("Escape");

  await page.getByTestId("list-numbered-menu-button").click();
  const numberMenu = page.getByTestId("number-style-menu");
  await expect(numberMenu.getByTestId("number-style-decimal")).toHaveText("1.");
  await expect(numberMenu.getByTestId("number-style-lowerAlpha")).toHaveText("a.");
  await expect(numberMenu.getByTestId("number-style-upperAlpha")).toHaveText("A.");
  await expect(numberMenu.getByTestId("number-style-lowerRoman")).toHaveText("i.");
  await expect(numberMenu.getByTestId("number-style-upperRoman")).toHaveText("I.");
});

test("picking a marker style re-renders every item at the current level immediately (AC-13)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("first");
  await page.keyboard.press("Enter");
  await page.keyboard.type("second");
  await page.keyboard.press("Enter");
  await page.keyboard.type("sub");
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTestId("list-numbered").click();
  await content(page).click();
  await paragraphAt(page, 2).click();
  // Offset 0 is where the demote lives (see the Tab section above).
  await page.keyboard.press(LINE_START);
  await page.keyboard.press("Tab");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "a.");

  // With the caret in a level-1 item, pick upper-roman: both level-1 items
  // re-render; the level-2 item keeps its own marker.
  await paragraphAt(page, 0).click();
  await page.getByTestId("list-numbered-menu-button").click();
  await page.getByTestId("number-style-upperRoman").click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "I.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "II.");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "a.");

  // Bullet styles work the same way through the bullet dropdown.
  await page.getByTestId("list-bulleted-menu-button").click();
  await page.getByTestId("bullet-style-square").click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "▪");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "▪");
});

// --- autoformat (AC-14..AC-17) ---------------------------------------------

test("typing '- ' or '* ' at the start of an empty paragraph makes a level-1 bullet item (AC-14)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("- ");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "•");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-level", "1");
  // The prefix is consumed, not kept as text.
  await page.keyboard.type("milk");
  await expect(paragraphAt(page, 0)).toHaveText("milk");

  await page.keyboard.press(LINE_END);
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter"); // exit the list on the empty item
  await page.keyboard.type("* ");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "•");
  await expect(paragraphAt(page, 1)).toHaveText("");
});

test("typing '1. ' or '1) ' at the start of an empty paragraph makes a decimal numbered item (AC-15)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("1. ");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await page.keyboard.type("steps");
  await expect(paragraphAt(page, 0)).toHaveText("steps");

  await page.keyboard.press(LINE_END);
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter");
  // A plain paragraph between the lists, so the second autoformat starts
  // its own list at 1 instead of continuing the first (AC-11 adjacency).
  await page.keyboard.type("between");
  await page.keyboard.press("Enter");
  await page.keyboard.type("1) ");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 2)).toHaveText("");
});

test("Ctrl+Z immediately after an autoformat restores the literal typed prefix (AC-16)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("- ");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "•");
  await page.keyboard.press(`${MOD}+z`);
  await expect(paragraphAt(page, 0)).not.toHaveAttribute("data-list-marker");
  await expect(paragraphAt(page, 0)).toHaveText("- ");

  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press("Delete");
  await page.keyboard.type("1. ");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await page.keyboard.press(`${MOD}+z`);
  await expect(paragraphAt(page, 0)).not.toHaveAttribute("data-list-marker");
  await expect(paragraphAt(page, 0)).toHaveText("1. ");
});

test("autoformat does not trigger when anything precedes the prefix (AC-17)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("x- ");
  await expect(paragraphAt(page, 0)).not.toHaveAttribute("data-list-marker");
  await expect(paragraphAt(page, 0)).toHaveText("x- ");

  await page.keyboard.press("Enter");
  await page.keyboard.type(" 1. ");
  await expect(paragraphAt(page, 1)).not.toHaveAttribute("data-list-marker");
});

// --- export (AC-18, AC-19) -------------------------------------------------

/** An editor-built two-level numbered list plus a bullet list. */
async function buildListsDocument(page: Page): Promise<void> {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("First");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Nested");
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTestId("list-numbered").click();
  await content(page).click();
  await paragraphAt(page, 1).click();
  // Offset 0 is where the demote lives (see the Tab section above).
  await page.keyboard.press(LINE_START);
  await page.keyboard.press("Tab");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "a.");
}

test("saving to .docx writes numbering.xml with an abstractNum and a referencing num, and w:numPr on the paragraphs (AC-18)", async ({ page }) => {
  await buildListsDocument(page);
  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  const numbering = parts.get("word/numbering.xml")?.toString("utf8") ?? "";
  expect(numbering).toContain("<w:abstractNum ");
  expect(numbering).toMatch(/<w:num w:numId="\d+"><w:abstractNumId w:val="\d+"\/><\/w:num>/);
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
  expect(documentXml).toContain('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>');
  expect(parts.get("[Content_Types].xml")!.toString("utf8")).toContain("/word/numbering.xml");
});

test("saving to .odt writes nested text:list markup with a matching text:list-style (AC-19)", async ({ page }) => {
  await buildListsDocument(page);
  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const contentXml = parts.get("content.xml")!.toString("utf8");
  expect(contentXml).toContain('<text:list text:style-name="L1">');
  // The level-2 item nests a list inside the level-1 item.
  expect(contentXml).toMatch(/<text:list-item>.*<text:list><text:list-item>/);
  expect(contentXml).toContain('<text:list-style style:name="L1">');
  expect(contentXml).toContain('style:num-format="1"');
  expect(contentXml).toContain('style:num-format="a"');
});

// --- import (AC-20, AC-21, AC-22) ------------------------------------------

/** Both fixtures encode the same document, so the render assertions are
 * shared: markers, levels and the per-level indents the definitions carry
 * (level 1 of the bullet list has a non-default 1.905 cm indent = 72 px;
 * the level-2 default is 2.54 cm = 96 px). */
async function expectFixtureRender(page: Page): Promise<void> {
  await expect(paragraphAt(page, 0)).toHaveText("Groceries");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "•");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "•");
  await expect(paragraphAt(page, 3)).toHaveAttribute("data-list-marker", "◦");
  await expect(paragraphAt(page, 3)).toHaveAttribute("data-list-level", "2");
  await expect(paragraphAt(page, 4)).toHaveAttribute("data-list-marker", "▪");
  await expect(paragraphAt(page, 4)).toHaveAttribute("data-list-level", "3");
  await expect(paragraphAt(page, 6)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 7)).toHaveAttribute("data-list-marker", "2.");
  await expect(paragraphAt(page, 8)).toHaveAttribute("data-list-marker", "a.");
  // The custom level-1 indent from the list definition, and the default
  // level-2 indent.
  expect(await marginPx(page, 1)).toBeCloseTo(72, 1);
  expect(await marginPx(page, 3)).toBeCloseTo(96, 1);
  // The list starting at 5 (AC-22).
  await expect(paragraphAt(page, 10)).toHaveAttribute("data-list-marker", "5.");
  await expect(paragraphAt(page, 11)).toHaveAttribute("data-list-marker", "6.");
}

test("a .docx fixture's lists render with the markers, levels and indents of its numbering.xml (AC-20, AC-22)", async ({ page }) => {
  await openFixture(page, docxFixture);
  await expectFixtureRender(page);
});

test("a .odt fixture's lists render with the markers, levels and indents of its text:list-style (AC-21, AC-22)", async ({ page }) => {
  await openFixture(page, odtFixture);
  await expectFixtureRender(page);
});

// --- re-export fidelity (AC-23) --------------------------------------------

test("re-exporting the imported .docx preserves markers, levels, indents and start values (AC-23)", async ({ page }) => {
  await openFixture(page, docxFixture);
  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  const numbering = parts.get("word/numbering.xml")!.toString("utf8");
  // The original definitions survive verbatim; the model's lists are
  // re-declared with the same values above them.
  expect(numbering).toContain('<w:start w:val="5"/>');
  expect(numbering).toContain('w:left="1080"');
  expect(numbering).toContain('<w:numFmt w:val="lowerLetter"/>');
  // Re-opening the export renders identically — the user-visible fidelity.
  await openFixture(page, path);
  await expectFixtureRender(page);
});

test("re-exporting the imported .odt preserves markers, levels, indents and start values (AC-23)", async ({ page }) => {
  await openFixture(page, odtFixture);
  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const contentXml = parts.get("content.xml")!.toString("utf8");
  expect(contentXml).toContain('text:start-value="5"');
  expect(contentXml).toContain('fo:margin-left="1.905cm"');
  await openFixture(page, path);
  await expectFixtureRender(page);
});

// --- cross-format ----------------------------------------------------------

test("lists survive a cross-format save: .docx fixture exported as .odt renders the same", async ({ page }) => {
  await openFixture(page, docxFixture);
  const path = await download(page, "odt");
  await openFixture(page, path);
  await expectFixtureRender(page);
});
