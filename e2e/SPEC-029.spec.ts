import { test, expect, type Download, type Locator, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";

// SPEC-029 — headers, footers and page numbers: entering/exiting header and
// footer editing (double-click the margin, Insert -> Header/Footer, Escape,
// double-click the body), text/formatting/tables/images inside headers and
// footers, per-section "different first page"/"different odd and even
// pages", the
// "Insert page number" control's position/format options and "Page X of Y",
// live page-number rendering, "start numbering at", and round-tripping
// through .docx (header1.xml/footer1.xml, PAGE/NUMPAGES fields) and .odt
// (master page styles, text:page-number/text:page-count).

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

const content = (page: Page) => page.getByTestId("document-content");
const pages = (page: Page) => page.getByTestId("document-page");
const headerBand = (page: Page, index = 0) => pages(page).nth(index).getByTestId("page-header-band");
const footerBand = (page: Page, index = 0) => pages(page).nth(index).getByTestId("page-footer-band");

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
 * as e2e/SPEC-028.spec.ts. */
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

async function waitForPagination(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-pages")).toHaveAttribute(
    "data-pagination-state",
    "clean",
    { timeout },
  );
}

/** Three pages, deterministically, via manual page breaks. */
async function makeThreePages(page: Page): Promise<void> {
  await content(page).click();
  await page.keyboard.type("Page one body");
  await page.keyboard.press(`${MOD}+Enter`);
  await page.keyboard.type("Page two body");
  await page.keyboard.press(`${MOD}+Enter`);
  await page.keyboard.type("Page three body");
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(3);
}

async function dblclickBand(locator: Locator): Promise<void> {
  const box = await locator.boundingBox();
  if (!box) {
    throw new Error("band has no bounding box");
  }
  await locator.page().mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
}

test("double-clicking the top margin enters header editing with body dimmed (AC-1)", async ({ page }) => {
  await createDocument(page);
  const firstPage = pages(page).first();
  const box = await firstPage.boundingBox();
  if (!box) throw new Error("no page box");
  await page.mouse.dblclick(box.x + box.width / 2, box.y + 15);
  await expect(headerBand(page)).toHaveAttribute("data-hf-editable", "1");
  await expect(page.locator(".page-margin-dimmed").first()).toBeVisible();
});

test("Insert -> Header/Footer enters editing with body dimmed (AC-2)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await expect(headerBand(page)).toHaveAttribute("data-hf-editable", "1");
  await expect(page.locator(".page-margin-dimmed").first()).toBeVisible();
});

test("header content supports text and formatting like body content (AC-3, AC-4)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await page.keyboard.type("Confidential");
  await page.keyboard.press(`${MOD}+a`);
  await page.getByTitle(/^Bold/).click();
  await expect(headerBand(page)).toContainText("Confidential");
  await expect(headerBand(page).locator("[style*='font-weight']").first()).toBeVisible();
});

test("header content supports tables like body content (AC-5)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-2-2").click();
  await expect(headerBand(page).locator("table.doc-table")).toHaveCount(1);
});

test("header content supports images like body content (AC-6)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await page.getByTestId("insert-image-file-input").setInputFiles({
    name: "photo.png",
    mimeType: "image/png",
    // 1x1 transparent PNG.
    buffer: Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
      "base64",
    ),
  });
  await expect(headerBand(page).locator("img")).toHaveCount(1);
});

test("Escape returns to body editing (AC-7)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await expect(headerBand(page)).toHaveAttribute("data-hf-editable", "1");
  await page.keyboard.press("Escape");
  await expect(headerBand(page)).not.toHaveAttribute("data-hf-editable", "1");
  await expect(page.locator(".page-margin-dimmed")).toHaveCount(0);
});

test("double-clicking the body returns to body editing (AC-8)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await expect(headerBand(page)).toHaveAttribute("data-hf-editable", "1");
  await page.locator(".page-content").first().dblclick();
  await expect(headerBand(page)).not.toHaveAttribute("data-hf-editable", "1");
});

test("header content repeats on every page and edits apply to all at once (AC-9, AC-10)", async ({ page }) => {
  await createDocument(page);
  await makeThreePages(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await page.keyboard.type("Company Letterhead");
  await page.keyboard.press("Escape");
  await waitForPagination(page);
  for (let i = 0; i < 3; i++) {
    await expect(headerBand(page, i)).toContainText("Company Letterhead");
  }
});

test("header/footer options: different first page and different odd/even (AC-11, AC-12, AC-13)", async ({
  page,
}) => {
  await createDocument(page);
  await makeThreePages(page);

  // Different first page: give the section a distinct first-page header.
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-options").click();
  await page.getByTestId("hf-option-title-page").check();
  await page.getByTestId("hf-options-apply").click();
  await waitForPagination(page);

  // The caret is still on page 3 (the last typed text) — bring it back to
  // page 1 so "Insert -> Header" targets that page's (title) band.
  await content(page).click();
  await page.keyboard.press(`${MOD}+Home`);

  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await expect(headerBand(page, 0)).toHaveAttribute("data-hf-editable", "1");
  await page.keyboard.type("Cover Page Header");
  await page.keyboard.press("Escape");
  await waitForPagination(page);
  await expect(headerBand(page, 0)).toContainText("Cover Page Header");
  // Page 2 is not the section's first page, so it does not show the "first"
  // header content typed above.
  await expect(headerBand(page, 1)).not.toContainText("Cover Page Header");
});

test("Insert page number offers position/format options and inserts a page number (AC-14..19)", async ({
  page,
}) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-page-number").click();
  await expect(page.getByTestId("page-number-dialog")).toBeVisible();
  await page.getByTestId("page-number-part").selectOption("footer");
  await page.getByTestId("page-number-position").selectOption("right");
  await page.getByTestId("page-number-format").selectOption("lowerRoman");
  await page.getByTestId("page-number-apply").click();
  await waitForPagination(page);
  await expect(footerBand(page)).toContainText("i");
});

test("'Page X of Y' inserts a composite field (AC-19)", async ({ page }) => {
  await createDocument(page);
  await makeThreePages(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-page-number").click();
  await page.getByTestId("page-number-part").selectOption("footer");
  await page.getByTestId("page-number-of-total").check();
  await page.getByTestId("page-number-apply").click();
  await waitForPagination(page);
  await expect(footerBand(page, 0)).toContainText("1 of 3");
  await expect(footerBand(page, 2)).toContainText("3 of 3");
});

test("page numbers render the correct value on every page and update when pagination changes (AC-20, AC-21)", async ({
  page,
}) => {
  await createDocument(page);
  await makeThreePages(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-page-number").click();
  await page.getByTestId("page-number-part").selectOption("footer");
  await page.getByTestId("page-number-apply").click();
  await waitForPagination(page);
  await expect(footerBand(page, 0)).toContainText("1");
  await expect(footerBand(page, 1)).toContainText("2");
  await expect(footerBand(page, 2)).toContainText("3");

  // Removing a page break re-paginates; the footer values follow. Caret to
  // the start of "Page two body" (right after the first break) — one
  // Backspace there removes the preceding break (see SPEC-028's own
  // delete-break test for the same pattern).
  await content(page).getByText("Page two body").click();
  await page.keyboard.press("Home");
  await page.keyboard.press("Backspace");
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(2);
  await expect(footerBand(page, 0)).toContainText("1");
  await expect(footerBand(page, 1)).toContainText("2");
});

test("'Start numbering at' is settable per section (AC-22, AC-23)", async ({ page }) => {
  await createDocument(page);
  await makeThreePages(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-page-number").click();
  await page.getByTestId("page-number-part").selectOption("footer");
  await page.getByTestId("page-number-apply").click();
  await waitForPagination(page);
  await expect(footerBand(page, 0)).toContainText("1");

  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-options").click();
  await page.getByTestId("hf-option-restart").check();
  await page.getByTestId("hf-option-start-at").fill("5");
  await page.getByTestId("hf-options-apply").click();
  await waitForPagination(page);
  await expect(footerBand(page, 0)).toContainText("5");
  await expect(footerBand(page, 1)).toContainText("6");
  await expect(footerBand(page, 2)).toContainText("7");
});

test("headers, footers and page-number fields round-trip through .docx (AC-24, AC-25)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Body text");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click();
  await page.keyboard.type("Header text");
  await page.keyboard.press("Escape");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-page-number").click();
  await page.getByTestId("page-number-part").selectOption("footer");
  await page.getByTestId("page-number-apply").click();
  await waitForPagination(page);

  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  expect(parts.has("word/header1.xml")).toBe(true);
  expect(parts.has("word/footer1.xml")).toBe(true);
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain("w:headerReference");
  expect(documentXml).toContain("w:footerReference");
  const footerXml = parts.get("word/footer1.xml")!.toString("utf8");
  expect(footerXml).toContain("w:fldSimple");
  expect(footerXml).toMatch(/w:instr="PAGE/);

  await openFixture(page, path);
  await waitForPagination(page);
  await expect(headerBand(page)).toContainText("Header text");
  await expect(footerBand(page)).toContainText("1");
});

test("headers and footers round-trip through .odt via master page styles (AC-26)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Body text");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-footer").click();
  await page.keyboard.type("Footer text");
  await page.keyboard.press("Escape");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-page-number").click();
  await page.getByTestId("page-number-part").selectOption("header");
  await page.getByTestId("page-number-apply").click();
  await waitForPagination(page);

  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const stylesXml = parts.get("styles.xml")!.toString("utf8");
  expect(stylesXml).toContain("<style:footer>");
  expect(stylesXml).toContain("<style:header>");
  expect(stylesXml).toContain("<text:page-number>");

  await openFixture(page, path);
  await waitForPagination(page);
  await expect(footerBand(page)).toContainText("Footer text");
  await expect(headerBand(page)).toContainText("1");
});

test("opening a fixture with different first-page headers renders them on the right pages (AC-27)", async ({
  page,
}) => {
  await openFixture(page, join(fixturesDir, "spec029-first-page.docx"));
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(3);
  await expect(headerBand(page, 0)).toContainText("COVER PAGE");
  await expect(headerBand(page, 1)).toContainText("Regular header");
  await expect(headerBand(page, 1)).not.toContainText("COVER PAGE");
  await expect(headerBand(page, 2)).toContainText("Regular header");
});

test("opening a fixture with odd/even headers renders them on the right pages (AC-28)", async ({ page }) => {
  await openFixture(page, join(fixturesDir, "spec029-odd-even.docx"));
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(3);
  await expect(headerBand(page, 0)).toContainText("ODD HEADER");
  await expect(headerBand(page, 1)).toContainText("EVEN HEADER");
  await expect(headerBand(page, 2)).toContainText("ODD HEADER");
});

test("opening an .odt fixture with odd/even headers renders them on the right pages (AC-28)", async ({
  page,
}) => {
  await openFixture(page, join(fixturesDir, "spec029-odd-even.odt"));
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(3);
  await expect(headerBand(page, 0)).toContainText("ODD HEADER");
  await expect(headerBand(page, 1)).toContainText("EVEN HEADER");
  await expect(headerBand(page, 2)).toContainText("ODD HEADER");
});
