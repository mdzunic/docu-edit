import { test, expect, type Download, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";

// SPEC-031 — footnotes and endnotes: Insert menu + Ctrl/Cmd+Alt+F/D, the
// caret landing in the new note's own area, the separator line, automatic
// numbering that updates on insert/delete, reference/note integrity, hover
// tooltips and click-to-jump both ways, the settings dialog (format, start
// number, per-page/per-section restart), split-on-overflow, round-tripping
// through .docx/.odt, and the 24-footnote fixture's per-page placement.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

const content = (page: Page) => page.getByTestId("document-content");
const pages = (page: Page) => page.getByTestId("document-page");
const marks = (page: Page) => page.locator(".doc-note-ref");
const footnoteAreas = (page: Page) => page.getByTestId("footnote-area");
const endnoteAreas = (page: Page) => page.getByTestId("endnote-area");
const activeNoteBody = (page: Page) => page.locator("[data-note-editable]");

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

async function waitForPagination(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-pages")).toHaveAttribute(
    "data-pagination-state",
    "clean",
    { timeout },
  );
}

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

/** Minimal central-directory walk over a downloaded archive — same helper
 * as e2e/SPEC-029.spec.ts. */
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

/** Types body text, then inserts a footnote through the Notes menu. */
async function insertFootnoteViaMenu(page: Page): Promise<void> {
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("notes-menu-footnote").click();
}

test("Insert -> Footnote adds a superscript mark and puts the caret in the footnote area (AC-1, AC-2, AC-5, AC-6)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Body text");
  await insertFootnoteViaMenu(page);

  await expect(marks(page)).toHaveCount(1);
  await expect(marks(page).first()).toHaveText("1");
  // The mark renders superscript.
  await expect(marks(page).first()).toHaveCSS("vertical-align", "super");
  // A footnote area with a separator line appears at the bottom of the page.
  await expect(footnoteAreas(page)).toHaveCount(1);
  await expect(footnoteAreas(page).first().getByTestId("note-separator")).toBeVisible();
  // The caret is in the note's editable body: typing lands in the note.
  await expect(activeNoteBody(page)).toHaveCount(1);
  await page.keyboard.type("The note text");
  await expect(footnoteAreas(page).first()).toContainText("The note text");
  // The body kept its own text unchanged.
  await expect(content(page)).toContainText("Body text");
});

test("Ctrl/Cmd+Alt+F inserts a footnote; formatting works inside the note (AC-1, AC-6)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Shortcut body");
  await page.keyboard.press(`${MOD}+Alt+f`);
  await expect(marks(page)).toHaveCount(1);
  await expect(activeNoteBody(page)).toHaveCount(1);
  await page.keyboard.type("bold note");
  // Select the note's text and bold it from the toolbar (AC-6: the note
  // area holds fully formattable content).
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTestId("format-bold").click();
  await expect(
    footnoteAreas(page).first().locator("span[style*='font-weight']").first(),
  ).toContainText("bold note");
});

test("Ctrl/Cmd+Alt+D inserts an endnote whose area sits at the end of the document (AC-3, AC-4)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Endnote body");
  await page.keyboard.press(`${MOD}+Alt+d`);
  await expect(marks(page)).toHaveCount(1);
  await expect(endnoteAreas(page)).toHaveCount(1);
  await expect(activeNoteBody(page)).toHaveCount(1);
  await page.keyboard.type("The endnote content");
  await expect(endnoteAreas(page).first()).toContainText("The endnote content");
  // Endnotes number in their own stream (roman by default).
  await expect(endnoteAreas(page).first().getByTestId("note-number")).toHaveText("i");
});

test("footnotes renumber immediately when one is inserted before another (AC-7, AC-8)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("First sentence. Second sentence.");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("note at the end");
  await page.keyboard.press("Escape");

  // Put the caret just after "First sentence." and insert another footnote.
  await content(page).click();
  await page.keyboard.press(`${MOD}+Home`);
  for (let i = 0; i < "First sentence.".length; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("note in the middle");
  await page.keyboard.press("Escape");

  await expect(marks(page)).toHaveCount(2);
  await expect(marks(page).nth(0)).toHaveText("1");
  await expect(marks(page).nth(1)).toHaveText("2");
  // The note area lists them in reference order.
  const numbers = footnoteAreas(page).first().getByTestId("note-number");
  await expect(numbers).toHaveCount(2);
  await expect(footnoteAreas(page).first()).toContainText("note in the middle");
  await expect(footnoteAreas(page).first()).toContainText("note at the end");
});

test("deleting a reference mark deletes its note and renumbers the rest (AC-9, AC-10)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Alpha");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("first note");
  await page.keyboard.press("Escape");
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press("Enter");
  await page.keyboard.type("Beta");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("second note");
  await page.keyboard.press("Escape");
  await expect(marks(page)).toHaveCount(2);

  // Backspace over the FIRST mark: End on the first paragraph puts the
  // caret right after it (deterministic — no native stepping over the
  // non-editable mark element).
  await content(page).click();
  await page.keyboard.press(`${MOD}+Home`);
  await page.keyboard.press("End");
  await page.keyboard.press("Backspace");

  await expect(marks(page)).toHaveCount(1);
  await expect(marks(page).first()).toHaveText("1");
  await expect(footnoteAreas(page).first()).toContainText("second note");
  await expect(footnoteAreas(page).first()).not.toContainText("first note");
});

test("deleting all of a note's text keeps the reference mark (AC-11)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Body");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("wiped");
  for (let i = 0; i < "wiped".length; i++) {
    await page.keyboard.press("Backspace");
  }
  await page.keyboard.press("Escape");
  await expect(marks(page)).toHaveCount(1);
  await expect(footnoteAreas(page)).toHaveCount(1);
});

test("hovering a mark shows the note text; clicking jumps to the note and back (AC-14, AC-15, AC-16)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Jump body");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("tooltip text");
  await page.keyboard.press("Escape");
  await expect(activeNoteBody(page)).toHaveCount(0);

  // The tooltip carries the note's text (AC-14).
  await expect(marks(page).first()).toHaveAttribute("title", "tooltip text");

  // Clicking the mark jumps into the note (AC-15).
  await marks(page).first().click();
  await expect(activeNoteBody(page)).toHaveCount(1);
  await page.keyboard.type(" extended");
  await expect(footnoteAreas(page).first()).toContainText("tooltip text extended");

  // Clicking the note's number jumps back to the reference (AC-16): the
  // caret lands right after the mark, so typing continues there ("1" is
  // the rendered mark between the body text and the typed "!").
  await footnoteAreas(page).first().getByTestId("note-number").click();
  await expect(activeNoteBody(page)).toHaveCount(0);
  await page.keyboard.type("!");
  await expect(content(page)).toContainText("Jump body1!");
});

test("the settings dialog changes format, start number and restart (AC-17..20)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Settings body");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("the note");
  await page.keyboard.press("Escape");
  await expect(marks(page).first()).toHaveText("1");

  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("notes-menu-settings").click();
  await expect(page.getByTestId("note-settings-dialog")).toBeVisible();
  // Restart options are offered (AC-19, AC-20).
  await expect(page.getByTestId("note-settings-restart-page")).toBeVisible();
  await expect(page.getByTestId("note-settings-restart-section")).toBeVisible();

  await page.getByTestId("note-settings-format").selectOption("lowerRoman");
  await page.getByTestId("note-settings-start-at").fill("3");
  await page.getByTestId("note-settings-apply").click();

  // 1 -> iii: format and start number both applied (AC-17, AC-18).
  await expect(marks(page).first()).toHaveText("iii");
  await expect(footnoteAreas(page).first().getByTestId("note-number")).toHaveText("iii");
});

test("a footnote too long for the page splits across pages (AC-13)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Short body");
  await insertFootnoteViaMenu(page);
  await expect(activeNoteBody(page)).toHaveCount(1);
  // ~4000 characters — more than the footnote area's share of one page.
  await page.keyboard.insertText("lorem ipsum dolor sit amet ".repeat(150).trim());
  await page.keyboard.press("Escape");
  await waitForPagination(page);

  // The note continues on later pages instead of overflowing the first.
  const pageCount = await pages(page).count();
  expect(pageCount).toBeGreaterThanOrEqual(2);
  const areas = footnoteAreas(page);
  const areaCount = await areas.count();
  expect(areaCount).toBeGreaterThanOrEqual(2);
  await expect(areas.nth(0)).toContainText("lorem ipsum");
  await expect(areas.nth(1)).toContainText("lorem ipsum");
  // Only the start of the note carries its number.
  await expect(areas.nth(0).getByTestId("note-number")).toHaveCount(1);
  await expect(areas.last().getByTestId("note-number")).toHaveCount(0);
});

test("footnotes and endnotes round-trip through .docx (AC-21, AC-22)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Round trip body");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("docx footnote text");
  await page.keyboard.press("Escape");
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press(`${MOD}+Alt+d`);
  await page.keyboard.type("docx endnote text");
  await page.keyboard.press("Escape");

  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  const footnotesXml = parts.get("word/footnotes.xml")?.toString("utf8") ?? "";
  expect(footnotesXml).toContain("docx footnote text");
  const endnotesXml = parts.get("word/endnotes.xml")?.toString("utf8") ?? "";
  expect(endnotesXml).toContain("docx endnote text");
  const documentXml = parts.get("word/document.xml")?.toString("utf8") ?? "";
  expect(documentXml).toContain("w:footnoteReference");
  expect(documentXml).toContain("w:endnoteReference");

  // Re-open the download: both notes and their marks survive.
  await openFixture(page, path);
  await waitForPagination(page);
  await expect(marks(page)).toHaveCount(2);
  await expect(footnoteAreas(page).first()).toContainText("docx footnote text");
  await expect(endnoteAreas(page).first()).toContainText("docx endnote text");
});

test("footnotes and endnotes round-trip through .odt (AC-23, AC-24)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Odt body");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("odt footnote text");
  await page.keyboard.press("Escape");
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press(`${MOD}+Alt+d`);
  await page.keyboard.type("odt endnote text");
  await page.keyboard.press("Escape");

  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const contentXml = parts.get("content.xml")?.toString("utf8") ?? "";
  expect(contentXml).toContain('text:note-class="footnote"');
  expect(contentXml).toContain('text:note-class="endnote"');
  expect(contentXml).toContain("text:note-citation");
  expect(contentXml).toContain("text:note-body");
  expect(contentXml).toContain("odt footnote text");
  expect(contentXml).toContain("odt endnote text");

  await openFixture(page, path);
  await waitForPagination(page);
  await expect(marks(page)).toHaveCount(2);
  await expect(footnoteAreas(page).first()).toContainText("odt footnote text");
  await expect(endnoteAreas(page).first()).toContainText("odt endnote text");
});

test("a fixture with 24 footnotes places every note on the page of its reference (AC-12, AC-25)", async ({ page }) => {
  await openFixture(page, join(fixturesDir, "spec031-footnotes.docx"));
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(4);
  await expect(marks(page)).toHaveCount(24);
  await expect(footnoteAreas(page)).toHaveCount(4);

  for (let pageIndex = 0; pageIndex < 4; pageIndex += 1) {
    const area = footnoteAreas(page).nth(pageIndex);
    const firstNote = pageIndex * 6 + 1;
    const lastNote = firstNote + 5;
    await expect(area).toContainText(`Note ${firstNote} content`);
    await expect(area).toContainText(`Note ${lastNote} content`);
    // Nothing from a neighbouring page's notes leaks in.
    if (pageIndex > 0) {
      await expect(area).not.toContainText(`Note ${firstNote - 1} content`);
    }
    if (pageIndex < 3) {
      await expect(area).not.toContainText(`Note ${lastNote + 1} content`);
    }
  }
});

test("a footnote moves with its reference when re-pagination moves the reference (AC-12)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Anchor paragraph");
  await insertFootnoteViaMenu(page);
  await page.keyboard.type("moving note");
  await page.keyboard.press("Escape");
  await waitForPagination(page);
  await expect(footnoteAreas(page)).toHaveCount(1);

  // Push the reference onto page 2 with a page break before it.
  await content(page).click();
  await page.keyboard.press(`${MOD}+Home`);
  await page.keyboard.press(`${MOD}+Enter`);
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(2);
  // The note now sits on page 2 with its reference, not page 1.
  await expect(pages(page).nth(1).getByTestId("footnote-area")).toHaveCount(1);
  await expect(pages(page).nth(0).getByTestId("footnote-area")).toHaveCount(0);
  await expect(pages(page).nth(1).getByTestId("footnote-area")).toContainText("moving note");
});
