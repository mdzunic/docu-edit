import { test, expect, type Download, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

// SPEC-033 — table of contents: "Insert -> Table of contents" builds a TOC
// from Heading 1-6 paragraphs with a dot leader and page numbers, the
// options dialog (levels, page numbers, hyperlinks-only), clickable entries
// that jump to their heading, the "Update table of contents" control that
// appears when headings or pagination drift, the not-directly-editable
// contract (typing shows a hint instead), deleting the block leaving
// headings untouched, the empty-document placeholder, and round-tripping
// through .docx (a TOC field with a cached result) and .odt
// (text:table-of-content).

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
const tocBlock = (page: Page) => page.getByTestId("toc-block");
const tocEntries = (page: Page) => page.getByTestId("toc-entry");

/** Types a heading paragraph via the SPEC-021 shortcut, then a normal body
 * paragraph — mirroring e2e/SPEC-025.spec.ts's own heading helper. */
async function typeHeadingThenBody(page: Page, headingText: string, level: number, bodyText: string): Promise<void> {
  await page.keyboard.type(headingText);
  await page.keyboard.press(`${MOD}+Alt+${level}`);
  await page.keyboard.press("Enter");
  await page.keyboard.press(`${MOD}+Alt+0`);
  await page.keyboard.type(bodyText);
  await page.keyboard.press("Enter");
}

async function insertToc(
  page: Page,
  options: { levels?: number[]; showPageNumbers?: boolean; hyperlinksOnly?: boolean } = {},
): Promise<void> {
  await page.getByTestId("insert-toc-button").click();
  await expect(page.getByTestId("toc-options-dialog")).toBeVisible();
  if (options.levels) {
    for (const level of [1, 2, 3, 4, 5, 6]) {
      const checkbox = page.getByTestId(`toc-level-${level}`);
      const shouldBeChecked = options.levels.includes(level);
      if ((await checkbox.isChecked()) !== shouldBeChecked) {
        await checkbox.click();
      }
    }
  }
  if (options.showPageNumbers === false) {
    await page.getByTestId("toc-show-page-numbers").click();
  }
  if (options.hyperlinksOnly) {
    await page.getByTestId("toc-hyperlinks-only").click();
  }
  await page.getByTestId("toc-options-apply").click();
}

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

/** Minimal central-directory walk over a downloaded archive — same helper
 * as e2e/SPEC-028.spec.ts and SPEC-029.spec.ts. */
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

test("Insert -> Table of contents builds one from Heading 1-6 paragraphs, with text, a dot leader and page numbers (AC-1..4)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "Body under chapter one.");
  await typeHeadingThenBody(page, "Section A", 2, "Body under section A.");
  await page.keyboard.press(`${MOD}+Home`);

  await insertToc(page);
  await expect(tocBlock(page)).toBeVisible();
  await expect(tocEntries(page)).toHaveCount(2);
  await expect(tocEntries(page).nth(0)).toContainText("Chapter One");
  await expect(tocEntries(page).nth(1)).toContainText("Section A");
  // A dot leader between the text and the page number.
  await expect(tocBlock(page).locator(".toc-leader")).toHaveCount(2);
  await expect(tocBlock(page).locator(".toc-page-number").first()).toHaveText("1");
});

test("the options dialog controls which levels are included, page numbers, and hyperlinks-only (AC-5..7)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "x");
  await typeHeadingThenBody(page, "Section A", 2, "y");
  await page.keyboard.press(`${MOD}+Home`);

  await insertToc(page, { levels: [1], showPageNumbers: true, hyperlinksOnly: false });
  await expect(tocEntries(page)).toHaveCount(1);
  await expect(tocEntries(page)).toContainText("Chapter One");

  // Delete it and reinsert as hyperlinks-only, both levels.
  const box = (await tocBlock(page).boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height - 2);
  await page.keyboard.press("Delete");
  await expect(tocBlock(page)).toHaveCount(0);

  await page.keyboard.press(`${MOD}+Home`);
  await insertToc(page, { levels: [1, 2], hyperlinksOnly: true });
  await expect(tocEntries(page)).toHaveCount(2);
  await expect(tocBlock(page).locator(".toc-leader")).toHaveCount(0);
  await expect(tocBlock(page).locator(".toc-page-number")).toHaveCount(0);
});

test("clicking an entry scrolls the caret to its heading (AC-8, AC-9)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "x");
  await typeHeadingThenBody(page, "Section A", 2, "y");
  await page.keyboard.press(`${MOD}+Home`);
  await insertToc(page);

  await tocEntries(page).nth(1).locator(".toc-entry-link").click();
  // The caret landed at the start of "Section A": typing prepends there.
  await page.keyboard.type("X");
  await expect(content(page).locator("h2")).toHaveText("XSection A");
});

test("the update control appears when a heading changes, and offers page-numbers-only vs. entire-table updates (AC-10..13)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "x");
  await page.keyboard.press(`${MOD}+Home`);
  await insertToc(page);
  await expect(page.getByTestId("toc-update-bar")).toHaveCount(0);

  // Edit the heading text: the TOC is now stale against the document.
  await content(page).locator("h1").click();
  await page.keyboard.press("End");
  await page.keyboard.type(" Renamed");
  await expect(page.getByTestId("toc-update-bar")).toBeVisible();

  await page.getByTestId("toc-update-entire").click();
  await expect(page.getByTestId("toc-update-bar")).toHaveCount(0);
  await expect(tocEntries(page)).toHaveText(["Chapter One Renamed1"]);
});

test("update page numbers only keeps entry text/order, refreshing only page numbers (AC-12)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "x");
  await page.keyboard.press(`${MOD}+Home`);
  await insertToc(page);

  // Add a second heading elsewhere without regenerating: "update page
  // numbers only" must not pick it up as a new entry.
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press("Enter");
  await typeHeadingThenBody(page, "Chapter Two", 1, "y");
  await expect(page.getByTestId("toc-update-bar")).toBeVisible();

  await page.getByTestId("toc-update-page-numbers").click();
  await expect(tocEntries(page)).toHaveCount(1);
  await expect(page.getByTestId("toc-update-bar")).toBeVisible();
});

test("the TOC is not directly editable: typing after selecting it shows the update hint instead (AC-14, AC-15)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "x");
  await page.keyboard.press(`${MOD}+Home`);
  await insertToc(page);

  const box = (await tocBlock(page).boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height - 2);
  await page.keyboard.type("z");
  await expect(page.getByText('This table of contents isn')).toBeVisible();
  await expect(content(page)).not.toContainText("z");
});

test("deleting the TOC leaves the headings untouched (AC-16)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "x");
  await page.keyboard.press(`${MOD}+Home`);
  await insertToc(page);
  await expect(tocBlock(page)).toBeVisible();

  const box = (await tocBlock(page).boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height - 2);
  await page.keyboard.press("Delete");

  await expect(tocBlock(page)).toHaveCount(0);
  await expect(content(page).locator("h1")).toHaveText("Chapter One");
  await expect(content(page)).toContainText("x");
});

test("when the document has no headings, the block still inserts with a placeholder, and can be updated and deleted (AC-22..27)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Just a plain paragraph, no headings.");
  await page.keyboard.press(`${MOD}+Home`);

  await insertToc(page);
  await expect(tocBlock(page)).toBeVisible();
  await expect(page.getByTestId("toc-placeholder")).toHaveText("No table of contents entries found.");
  await expect(page.getByTestId("toc-update-bar")).toHaveCount(0);

  // Add a heading: now the block is stale and can be updated to show it.
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press("Enter");
  await typeHeadingThenBody(page, "New Heading", 1, "more text");
  await expect(page.getByTestId("toc-update-bar")).toBeVisible();
  await page.getByTestId("toc-update-entire").click();
  await expect(page.getByTestId("toc-placeholder")).toHaveCount(0);
  await expect(tocEntries(page)).toHaveText(["New Heading1"]);

  const box = (await tocBlock(page).boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + box.height - 2);
  await page.keyboard.press("Delete");
  await expect(tocBlock(page)).toHaveCount(0);
});

test("round-trips through .docx as a TOC field with a cached result (AC-17)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "x");
  await page.keyboard.press(`${MOD}+Home`);
  await insertToc(page);

  const path = await download(page, "docx");
  const { readFileSync } = await import("node:fs");
  const parts = unzipParts(readFileSync(path));
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain("TOC");
  expect(documentXml).toContain('w:fldCharType="begin"');
  expect(documentXml).toContain('w:fldCharType="separate"');
  expect(documentXml).toContain('w:fldCharType="end"');
  // The cached result text is present, not just the field code — Word shows
  // it without an update prompt.
  expect(documentXml).toContain("Chapter One");

  // Re-opening the saved file renders the same entries.
  await openFixture(page, path);
  await expect(tocEntries(page)).toHaveCount(1);
  await expect(tocEntries(page)).toContainText("Chapter One");
});

test("opening a Word-authored fixture renders the existing TOC entries and allows regeneration (AC-19, AC-20)", async ({
  page,
}) => {
  await openFixture(page, join(fixturesDir, "word-toc.docx"));
  await expect(tocBlock(page)).toBeVisible();
  await expect(tocEntries(page)).toHaveCount(2);
  await expect(tocEntries(page).nth(0)).toContainText("Introduction");
  await expect(tocEntries(page).nth(1)).toContainText("Background");
  await expect(page.getByTestId("toc-update-bar")).toHaveCount(0);

  // Rename a heading and regenerate — the fixture's TOC is a real, live
  // block, not inert imported content.
  await content(page).locator("h2").click();
  await page.keyboard.press("End");
  await page.keyboard.type(" Renamed");
  await expect(page.getByTestId("toc-update-bar")).toBeVisible();
  await page.getByTestId("toc-update-entire").click();
  await expect(tocEntries(page).nth(1)).toContainText("Background Renamed");
});

test("round-trips through .odt as text:table-of-content (AC-18)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "x");
  await page.keyboard.press(`${MOD}+Home`);
  await insertToc(page);

  const path = await download(page, "odt");
  const { readFileSync } = await import("node:fs");
  const parts = unzipParts(readFileSync(path));
  const contentXml = parts.get("content.xml")!.toString("utf8");
  expect(contentXml).toContain("text:table-of-content");
  expect(contentXml).toContain("text:index-body");
  expect(contentXml).toContain("Chapter One");

  await openFixture(page, path);
  await expect(tocEntries(page)).toHaveCount(1);
  await expect(tocEntries(page)).toContainText("Chapter One");
});
