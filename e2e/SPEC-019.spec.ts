import { test, expect, type Download, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { extend, WORD_LEFT } from "./keys";

// SPEC-019 — paragraph styles and the style picker: the current style name
// in the toolbar, the built-in and document style lists with previews,
// applying styles by picker and by Ctrl/Cmd+Alt+0..6, direct formatting
// layered on styles, the update/reset context-menu actions, and real
// heading semantics in the DOM and in saved .docx/.odt output.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxFixture = join(fixturesDir, "styles.docx");
const odtFixture = join(fixturesDir, "styles.odt");

// Ctrl on Windows/Linux, Cmd on macOS — the platform running the browser.
const mod = process.platform === "darwin" ? "Meta" : "Control";

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
const paragraphAt = (page: Page, index: number) =>
  content(page).locator(".doc-paragraph").nth(index);
const pickerCurrent = (page: Page) => page.getByTestId("style-picker-current");

async function tagNameAt(page: Page, index: number): Promise<string> {
  return paragraphAt(page, index).evaluate((el) => el.tagName);
}

async function openPicker(page: Page): Promise<void> {
  await page.getByTestId("style-picker-button").click();
  await expect(page.getByTestId("style-picker-menu")).toBeVisible();
}

async function applyStyle(page: Page, styleId: string): Promise<void> {
  await openPicker(page);
  await page.getByTestId(`style-option-${styleId}`).click();
}

async function fontSizePx(page: Page, index: number): Promise<number> {
  return paragraphAt(page, index)
    .locator("span")
    .first()
    .evaluate((el) => Number.parseFloat(getComputedStyle(el).fontSize));
}

/** Downloads the document in the given format and returns the file path —
 * the same helper as e2e/SPEC-018.spec.ts. */
async function download(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("download-menu").click();
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

async function downloadedPart(page: Page, format: "docx" | "odt", part: string): Promise<string> {
  const { readFileSync } = await import("node:fs");
  const path = await download(page, format);
  return unzipParts(readFileSync(path)).get(part)!.toString("utf8");
}

const BUILT_IN_NAMES = [
  "Normal",
  "Title",
  "Subtitle",
  "Heading 1",
  "Heading 2",
  "Heading 3",
  "Heading 4",
  "Heading 5",
  "Heading 6",
  "Quote",
  "Code",
];

test("the toolbar shows the current paragraph's style name (AC-1)", async ({ page }) => {
  await createDocument(page);
  await expect(pickerCurrent(page)).toHaveText("Normal");

  await openFixture(page, docxFixture);
  await paragraphAt(page, 1).click();
  await expect(pickerCurrent(page)).toHaveText("Heading 1");
  await paragraphAt(page, 2).click();
  await expect(pickerCurrent(page)).toHaveText("Fancy Title");
});

test("a blank document's picker lists the built-in set, previewed in each style (AC-2)", async ({
  page,
}) => {
  await createDocument(page);
  await openPicker(page);
  const options = page.getByTestId("style-picker-menu").getByRole("menuitemradio");
  await expect(options).toHaveText(BUILT_IN_NAMES);

  const heading1 = page.getByTestId("style-option-Heading1").locator("span");
  await expect(heading1).toHaveCSS("font-weight", "700");
  expect(await heading1.evaluate((el) => Number.parseFloat(getComputedStyle(el).fontSize))).toBeCloseTo(
    (20 * 96) / 72,
    1,
  );
  await expect(page.getByTestId("style-option-Code").locator("span")).toHaveCSS(
    "font-family",
    /Courier/,
  );
});

test("choosing a style applies it to every paragraph in the selection, immediately (AC-3, AC-4)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("first");
  await page.keyboard.press("Enter");
  await page.keyboard.type("second");
  await page.keyboard.press("Enter");
  await page.keyboard.type("third");
  // Select from inside the first paragraph into the second.
  await paragraphAt(page, 0).click();
  await page.keyboard.press("Shift+ArrowDown");
  await applyStyle(page, "Heading2");

  expect(await tagNameAt(page, 0)).toBe("H2");
  expect(await tagNameAt(page, 1)).toBe("H2");
  expect(await tagNameAt(page, 2)).toBe("P");
  await expect(paragraphAt(page, 0).locator("span").first()).toHaveCSS("font-weight", "700");
  expect(await fontSizePx(page, 0)).toBeCloseTo((16 * 96) / 72, 1);
  await expect(pickerCurrent(page)).toHaveText("Heading 2");
});

test("Ctrl/Cmd+Alt+1..6 apply Heading 1-6 (AC-5)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("shortcut styled");
  for (const level of [1, 3, 6]) {
    await page.keyboard.press(`${mod}+Alt+${level}`);
    expect(await tagNameAt(page, 0)).toBe(`H${level}`);
    await expect(pickerCurrent(page)).toHaveText(`Heading ${level}`);
  }
});

test("Ctrl/Cmd+Alt+0 applies Normal (AC-6)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("back to body");
  await page.keyboard.press(`${mod}+Alt+1`);
  expect(await tagNameAt(page, 0)).toBe("H1");

  await page.keyboard.press(`${mod}+Alt+0`);
  expect(await tagNameAt(page, 0)).toBe("P");
  await expect(pickerCurrent(page)).toHaveText("Normal");
  await expect(paragraphAt(page, 0).locator("span").first()).toHaveCSS("font-weight", "400");
});

test("an opened document's own styles are listed instead of the built-in set, with inherited previews (AC-7, AC-8)", async ({
  page,
}) => {
  await openFixture(page, docxFixture);
  await openPicker(page);
  const options = page.getByTestId("style-picker-menu").getByRole("menuitemradio");
  await expect(options).toHaveText(["Normal", "Fancy Base", "Fancy Title", "Heading 1"]);
  await expect(page.getByTestId("style-option-Code")).toHaveCount(0);
  await expect(page.getByTestId("style-option-Title")).toHaveCount(0);

  // Fancy Title inherits italic from Fancy Base (w:basedOn) and adds bold
  // of its own — the preview resolves the whole chain.
  const fancy = page.getByTestId("style-option-FancyTitle").locator("span");
  await expect(fancy).toHaveCSS("font-style", "italic");
  await expect(fancy).toHaveCSS("font-weight", "700");
});

test("applying a document style resolves its inheritance chain (AC-3, AC-7)", async ({ page }) => {
  await openFixture(page, docxFixture);
  await paragraphAt(page, 0).click();
  await applyStyle(page, "FancyTitle");
  const span = paragraphAt(page, 0).locator("span").first();
  await expect(span).toHaveCSS("font-style", "italic");
  await expect(span).toHaveCSS("font-weight", "700");
  expect(await fontSizePx(page, 0)).toBeCloseTo((24 * 96) / 72, 1);
});

test("direct formatting layered on a styled paragraph wins in rendering (AC-9, AC-10)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Chapter one");
  await applyStyle(page, "Heading1");
  // Select the last word and italicise it directly.
  await page.keyboard.press("End");
  await page.keyboard.press(extend(WORD_LEFT));
  await page.keyboard.press(`${mod}+i`);

  const spans = paragraphAt(page, 0).locator("span");
  await expect(spans.last()).toHaveCSS("font-style", "italic");
  // The style's own formatting still applies underneath.
  await expect(spans.last()).toHaveCSS("font-weight", "700");
  await expect(spans.first()).toHaveCSS("font-style", "normal");
});

test("style and layered direct formatting both survive a .docx save (AC-11, AC-15)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Chapter one");
  await applyStyle(page, "Heading1");
  await page.keyboard.press("End");
  await page.keyboard.press(extend(WORD_LEFT));
  await page.keyboard.press(`${mod}+i`);

  const documentXml = await downloadedPart(page, "docx", "word/document.xml");
  // A real heading: w:pStyle referencing the built-in heading style id.
  expect(documentXml).toContain('<w:pStyle w:val="Heading1"/>');
  expect(documentXml).toContain("<w:i/>");

  const stylesXml = await downloadedPart(page, "docx", "word/styles.xml");
  expect(stylesXml).toContain('w:styleId="Heading1"');
  expect(stylesXml).toContain('<w:outlineLvl w:val="0"/>');
});

test("an applied heading saves to .odt as text:h with the outline level (AC-11, AC-16)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Chapter one");
  await applyStyle(page, "Heading3");
  await page.keyboard.press("End");
  await page.keyboard.press(extend(WORD_LEFT));
  await page.keyboard.press(`${mod}+i`);

  const contentXml = await downloadedPart(page, "odt", "content.xml");
  expect(contentXml).toContain("<text:h");
  expect(contentXml).toContain('text:outline-level="3"');
  expect(contentXml).toContain('fo:font-style="italic"');
});

test("Update <style> to match redefines the style and re-renders every user (AC-12, AC-13)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("alpha");
  await page.keyboard.press("Enter");
  await page.keyboard.type("beta");
  // Make both paragraphs Heading 1.
  await page.keyboard.press(`${mod}+a`);
  await applyStyle(page, "Heading1");
  expect(await tagNameAt(page, 0)).toBe("H1");
  expect(await tagNameAt(page, 1)).toBe("H1");

  // Italicise the whole first paragraph directly.
  await paragraphAt(page, 0).click();
  await page.keyboard.press("Home");
  await page.keyboard.press("Shift+End");
  await page.keyboard.press(`${mod}+i`);
  await expect(paragraphAt(page, 0).locator("span").first()).toHaveCSS("font-style", "italic");
  await expect(paragraphAt(page, 1).locator("span").first()).toHaveCSS("font-style", "normal");

  // Update Heading 1 from the current (first) paragraph via the entry's
  // context menu: the second heading re-renders italic too.
  await openPicker(page);
  await page.getByTestId("style-option-Heading1").click({ button: "right" });
  const contextMenu = page.getByTestId("style-context-menu");
  await expect(contextMenu).toBeVisible();
  await expect(contextMenu.getByTestId("style-update")).toHaveText("Update Heading 1 to match");
  await contextMenu.getByTestId("style-update").click();

  await expect(paragraphAt(page, 1).locator("span").first()).toHaveCSS("font-style", "italic");
  expect(await tagNameAt(page, 1)).toBe("H1");
});

test("Reset style removes direct formatting, leaving the style's own (AC-14)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("quote me");
  await applyStyle(page, "Quote");
  await expect(paragraphAt(page, 0).locator("span").first()).toHaveCSS("font-style", "italic");

  // Layer direct bold over the whole paragraph.
  await page.keyboard.press("Home");
  await page.keyboard.press("Shift+End");
  await page.keyboard.press(`${mod}+b`);
  await expect(paragraphAt(page, 0).locator("span").first()).toHaveCSS("font-weight", "700");

  await openPicker(page);
  await page.getByTestId("style-option-Quote").click({ button: "right" });
  await page.getByTestId("style-reset").click();

  const span = paragraphAt(page, 0).locator("span").first();
  await expect(span).toHaveCSS("font-weight", "400");
  await expect(span).toHaveCSS("font-style", "italic");
  await expect(pickerCurrent(page)).toHaveText("Quote");
});

test("rendered headings expose real h1-h6 semantics to assistive technology (AC-17)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Accessible heading");
  await page.keyboard.press(`${mod}+Alt+2`);
  await expect(page.getByRole("heading", { level: 2 })).toHaveText("Accessible heading");

  await openFixture(page, docxFixture);
  await expect(page.getByRole("heading", { level: 1 })).toHaveText("Chapter heading");
});

test("an opened .odt lists its named styles with display names and inheritance (AC-7, AC-8)", async ({
  page,
}) => {
  await openFixture(page, odtFixture);
  await expect(page.getByRole("heading", { level: 1 })).toHaveText("Odt chapter heading");
  await paragraphAt(page, 1).click();
  await expect(pickerCurrent(page)).toHaveText("Heading 1");

  await openPicker(page);
  const options = page.getByTestId("style-picker-menu").getByRole("menuitemradio");
  // The document defines no Normal, so a fallback Normal entry is offered
  // alongside its own styles.
  await expect(options).toHaveText(["Normal", "Standard", "Heading 1", "My Quote"]);

  // Applying the document's quote style resolves style:parent-style-name.
  await paragraphAt(page, 0).click();
  await applyStyle(page, "My_20_Quote");
  await expect(paragraphAt(page, 0).locator("span").first()).toHaveCSS("font-style", "italic");
  await expect(pickerCurrent(page)).toHaveText("My Quote");
});

test("a heading applied to an opened .odt saves back as text:h (AC-15, AC-16)", async ({ page }) => {
  await openFixture(page, odtFixture);
  await paragraphAt(page, 0).click();
  await page.keyboard.press(`${mod}+Alt+2`);
  expect(await tagNameAt(page, 0)).toBe("H2");

  const contentXml = await downloadedPart(page, "odt", "content.xml");
  expect(contentXml).toContain('text:outline-level="2"');
});
