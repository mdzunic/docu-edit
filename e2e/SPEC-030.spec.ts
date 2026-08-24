import { test, expect, type Page } from "@playwright/test";

// SPEC-030 — multi-column layout: the Format → Columns dialog (presets,
// count 1-6, spacing, equal-width toggle, line-between, apply-to), per-column
// text flow and column breaks, column-width containment for tables/lists,
// ruler boundary handles, cross-column caret navigation, and .docx/.odt
// round-trip. The dialog and rendering are the parts the QA agent exercises.

const content = (page: Page) => page.getByTestId("document-content");
const docColumns = (page: Page) => page.getByTestId("doc-column");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** Opens a document through the real picker and waits until it is laid out.
 * The URL changes as soon as the document is stored, which is *before* the
 * parse finishes and well before pagination has run — and columns are a
 * pagination product, so a one-shot `docColumns(page).count()` taken at that
 * moment reads 0 and the assertion fails with no bug behind it. Waiting for
 * the "Opening…" indicator to clear and for pagination to report `clean` is
 * the same barrier SPEC-047's `waitForFullyInteractive` uses. */
async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("document-opening")).toBeHidden();
  await expect(page.getByTestId("document-pages")).toHaveAttribute("data-pagination-state", "clean");
}

/** The column index (and rendered page index) the caret currently sits in,
 * or nulls when it is outside any column. Used by the cross-column/cross-page
 * navigation and flow assertions (AC-9, AC-11, AC-20, AC-21). */
async function caretLocation(page: Page): Promise<{ column: string | null; pageIndex: number | null }> {
  return page.evaluate(() => {
    const node = getSelection()?.anchorNode ?? null;
    const el = node instanceof Element ? node : (node?.parentElement ?? null);
    const col = el?.closest?.(".doc-column") ?? null;
    const pageEl = el?.closest?.('[data-testid="document-page"]') ?? null;
    const pages = Array.from(document.querySelectorAll('[data-testid="document-page"]'));
    return {
      column: col?.getAttribute("data-column") ?? null,
      pageIndex: pageEl ? pages.indexOf(pageEl) : null,
    };
  });
}

/** Which rendered page/column a block of text landed in, by its own text
 * content — used to check flow order (AC-9) without relying on the caret. */
async function locate(page: Page, text: string): Promise<{ pageIndex: number; column: number } | null> {
  return page.evaluate((needle) => {
    const pages = Array.from(document.querySelectorAll('[data-testid="document-page"]'));
    for (let p = 0; p < pages.length; p++) {
      const cols = Array.from(pages[p].querySelectorAll('[data-testid="doc-column"]'));
      for (let c = 0; c < cols.length; c++) {
        if (cols[c].textContent?.includes(needle)) {
          return { pageIndex: p, column: Number(cols[c].getAttribute("data-column")) };
        }
      }
    }
    return null;
  }, text);
}

async function typeParagraphs(page: Page, count: number): Promise<void> {
  await content(page).click();
  for (let i = 1; i <= count; i++) {
    await page.keyboard.type(`Paragraph ${i} of the document body text.`);
    if (i < count) {
      await page.keyboard.press("Enter");
    }
  }
}

/** Opens the Columns dialog and applies `count` equal columns to the section. */
async function applyColumns(page: Page, count: number): Promise<void> {
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("columns-button").click();
  await expect(page.getByTestId("columns-dialog")).toBeVisible();
  await page.getByTestId(`columns-preset-${count}`).click();
  await page.getByTestId("columns-confirm").click();
  await expect(page.getByTestId("columns-dialog")).toHaveCount(0);
  // The dialog closes on the model edit; the column boxes appear one
  // pagination pass later. Same barrier as openFixture, for the same reason.
  await expect(page.getByTestId("document-pages")).toHaveAttribute("data-pagination-state", "clean");
}

test.describe("SPEC-030 multi-column layout", () => {
  test("the dialog offers 1/2/3 presets and a 1-6 custom count (AC-1, AC-2)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("columns-button").click();
    await expect(page.getByTestId("columns-dialog")).toBeVisible();
    await expect(page.getByTestId("columns-preset-1")).toBeVisible();
    await expect(page.getByTestId("columns-preset-2")).toBeVisible();
    await expect(page.getByTestId("columns-preset-3")).toBeVisible();

    // An out-of-range count is rejected and reverts to the previous valid value.
    await page.getByTestId("columns-preset-2").click();
    const countField = page.getByTestId("columns-count");
    await countField.fill("9");
    await countField.blur();
    await expect(page.getByTestId("columns-count-error")).toBeVisible();
    await expect(countField).toHaveValue("2");
    // A valid count within range commits.
    await countField.fill("6");
    await countField.blur();
    await expect(page.getByTestId("columns-count-error")).toHaveCount(0);
    await expect(countField).toHaveValue("6");
  });

  test("equal-width disables the per-column width fields (AC-3, AC-3a)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("columns-button").click();
    await page.getByTestId("columns-preset-2").click();
    // Equal widths on: the width fields are disabled.
    await expect(page.getByTestId("columns-width-0")).toBeDisabled();
    // Turning equal off enables them.
    await page.getByTestId("columns-equal").click();
    await expect(page.getByTestId("columns-width-0")).toBeEnabled();
    // Spacing and line-between and apply-to are all present.
    await expect(page.getByTestId("columns-spacing")).toBeVisible();
    await expect(page.getByTestId("columns-separator")).toBeVisible();
    await expect(page.getByTestId("columns-apply-to")).toBeVisible();
  });

  test("applying two columns renders two column boxes (AC-6)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 6);
    await applyColumns(page, 2);
    await expect(docColumns(page).first()).toBeVisible();
    expect(await docColumns(page).count()).toBeGreaterThanOrEqual(2);
  });

  test("a column inside the section is narrower than the page text area (AC-15)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 6);
    await applyColumns(page, 2);
    const columnWidth = (await docColumns(page).first().boundingBox())!.width;
    const contentWidth = (await page.getByTestId("page-margin").first().boundingBox())!.width;
    // Two columns plus a gutter: each column is well under half the text area
    // plus a comfortable margin — certainly under the full width.
    expect(columnWidth).toBeLessThan(contentWidth * 0.6);
  });

  test("a table inserted into a column fits within the column width (AC-13/AC-15)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 4);
    await applyColumns(page, 2);
    // Put the caret in the first column, then insert a 2×3 table.
    await content(page).getByText("Paragraph 1", { exact: false }).first().click();
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("table-grid-cell-2-3").click();
    const table = content(page).locator("table.doc-table").first();
    await expect(table).toBeVisible();
    const tableWidth = (await table.boundingBox())!.width;
    const column = table.locator("xpath=ancestor::*[@data-testid='doc-column']");
    const columnWidth = (await column.boundingBox())!.width;
    // The table lays out within its column, not overflowing the column box.
    expect(tableWidth).toBeLessThanOrEqual(columnWidth + 1);
  });

  test("a column break moves following text to the next column (AC-10)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.keyboard.type("Left column text");
    await applyColumns(page, 2);
    // Caret at end of the paragraph, insert a column break, then more text.
    await content(page).getByText("Left column text").click();
    await page.keyboard.press("End");
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("break-column").click();
    await page.keyboard.type("Right column text");
    // "Left column text" is in column 0; "Right column text" in column 1.
    const col0 = docColumns(page).nth(0);
    const col1 = docColumns(page).nth(1);
    await expect(col0).toContainText("Left column text");
    await expect(col1).toContainText("Right column text");
  });

  test("the ruler shows a column boundary handle for a two-column section (AC-14)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 4);
    // No boundary handle for a single column.
    await expect(page.getByTestId("ruler-column-boundary-0")).toHaveCount(0);
    await applyColumns(page, 2);
    await expect(page.getByTestId("ruler-column-boundary-0")).toBeVisible();
  });

  test("columns round-trip through .docx (AC-16)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 6);
    await applyColumns(page, 3);
    await page.getByTestId("menu-title-file").click();
    await page.getByTestId("file-menu-export").click();
    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-docx").click(),
    ]);
    const path = await download.path();
    await openFixture(page, path!);
    expect(await docColumns(page).count()).toBeGreaterThanOrEqual(3);
  });

  test("columns round-trip through .odt (AC-17)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 6);
    await applyColumns(page, 2);
    await page.getByTestId("menu-title-file").click();
    await page.getByTestId("file-menu-export").click();
    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-odt").click(),
    ]);
    const path = await download.path();
    await openFixture(page, path!);
    expect(await docColumns(page).count()).toBeGreaterThanOrEqual(2);
  });

  test("arrow keys move the caret between columns (AC-20/AC-21)", async ({ page }) => {
    await createDocument(page);
    // Enough short paragraphs that the flow fills column 0 and spills into
    // column 1, so there is a genuine column boundary to cross.
    await content(page).click();
    await page.keyboard.type(Array.from({ length: 80 }, (_, i) => `Para${i}`).join("\n"));
    await applyColumns(page, 2);
    await expect(docColumns(page).nth(1)).not.toBeEmpty();
    const columnOfCaret = () =>
      page.evaluate(() => {
        const node = getSelection()?.anchorNode ?? null;
        const el = node instanceof Element ? node : node?.parentElement ?? null;
        return el?.closest?.(".doc-column")?.getAttribute("data-column") ?? null;
      });

    // Down at the bottom of column 0 crosses into column 1 (AC-20).
    await content(page).locator('.doc-column[data-column="0"] .doc-paragraph').last().click();
    await page.keyboard.press("End");
    expect(await columnOfCaret()).toBe("0");
    await page.keyboard.press("ArrowDown");
    expect(await columnOfCaret()).toBe("1");

    // Up at the top of column 1 crosses back into column 0 (AC-20).
    await content(page).locator('.doc-column[data-column="1"] .doc-paragraph').first().click();
    await page.keyboard.press("Home");
    expect(await columnOfCaret()).toBe("1");
    await page.keyboard.press("ArrowUp");
    expect(await columnOfCaret()).toBe("0");
  });

  test("opens a two-column fixture and renders both sections (AC-19)", async ({ page }) => {
    await openFixture(page, "e2e/fixtures/two-column-section.docx");
    // The two-column body renders as column boxes...
    expect(await docColumns(page).count()).toBeGreaterThanOrEqual(2);
    // ...and the single-column heading section renders too.
    await expect(content(page)).toContainText("Newsletter Title");
    await expect(content(page)).toContainText("Column paragraph 1");
  });

  test("the 'Line between columns' toggle renders a vertical rule in the gutter (AC-4)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 4);
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("columns-button").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "visible" });
    await page.getByTestId("columns-preset-2").click();
    await page.getByTestId("columns-separator").click();
    await page.getByTestId("columns-confirm").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "detached" });
    const rule = page.locator(".doc-column-rule");
    await expect(rule).toHaveCount(1);
    await expect(rule.first()).toBeVisible();
  });

  test("the apply-to selector offers 'Current section' and 'Selection' (AC-5)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("columns-button").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "visible" });
    const options = await page.getByTestId("columns-apply-to").locator("option").allTextContents();
    expect(options.some((o) => /current section/i.test(o))).toBe(true);
    expect(options.some((o) => /selection/i.test(o))).toBe(true);
  });

  test("applying to a non-empty selection wraps only the touched whole paragraphs (AC-7, AC-7a, AC-7b)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    const paragraphs = ["Alpha one two three.", "Bravo four five six.", "Charlie seven eight nine.", "Delta ten.", "Echo eleven."];
    for (let i = 0; i < paragraphs.length; i++) {
      await page.keyboard.type(paragraphs[i]);
      if (i < paragraphs.length - 1) await page.keyboard.press("Enter");
    }
    // Select from mid-way into "Bravo" through mid-way into "Delta" — a partial
    // selection that AC-7 must expand to the full paragraphs it touches.
    await page.evaluate(() => {
      const root = document.querySelector('[data-testid="document-content"]')!;
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      let bravoNode: Text | null = null;
      let deltaNode: Text | null = null;
      let n: Node | null;
      // eslint-disable-next-line no-cond-assign
      while ((n = walker.nextNode())) {
        if (n.textContent?.includes("Bravo")) bravoNode = n as Text;
        if (n.textContent?.includes("Delta")) deltaNode = n as Text;
      }
      const range = document.createRange();
      range.setStart(bravoNode!, 6);
      range.setEnd(deltaNode!, 6);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
    });
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("columns-button").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "visible" });
    await page.getByTestId("columns-preset-2").click();
    await page.getByTestId("columns-apply-to").selectOption("selection");
    await page.getByTestId("columns-confirm").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "detached" });

    // No paragraph text was lost or split (AC-7b).
    const bodyText = await content(page).innerText();
    for (const p of paragraphs) {
      expect(bodyText).toContain(p);
    }
    // Only Bravo/Charlie/Delta (the touched paragraphs) sit inside the new
    // column section; Alpha and Echo stay outside it (AC-7, AC-7a).
    const colText = (await docColumns(page).allInnerTexts()).join(" | ");
    expect(colText).not.toContain("Alpha");
    expect(colText).not.toContain("Echo");
    expect(colText).toContain("Bravo");
    expect(colText).toContain("Charlie");
    expect(colText).toContain("Delta");
  });

  test("applying to an empty (collapsed) selection wraps just the caret's paragraph (AC-8)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 3);
    await content(page).getByText("Paragraph 2", { exact: false }).click();
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("columns-button").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "visible" });
    await page.getByTestId("columns-preset-2").click();
    await page.getByTestId("columns-apply-to").selectOption("selection");
    await page.getByTestId("columns-confirm").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "detached" });
    const colText = (await docColumns(page).allInnerTexts()).join(" | ");
    expect(colText).toContain("Paragraph 2");
    expect(colText).not.toContain("Paragraph 1");
    expect(colText).not.toContain("Paragraph 3");
  });

  test("text flows column-to-column then to the next page's first column (AC-9)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.keyboard.type(Array.from({ length: 140 }, (_, i) => `Line${i}`).join("\n"));
    await applyColumns(page, 2);
    const first = await locate(page, "Line0");
    const last = await locate(page, "Line139");
    expect(first).not.toBeNull();
    expect(last).not.toBeNull();
    // Line139 landed strictly after Line0 in the flow: a later page, or the
    // same page's later column.
    const advanced =
      last!.pageIndex > first!.pageIndex || (last!.pageIndex === first!.pageIndex && last!.column > first!.column);
    expect(advanced).toBe(true);
  });

  test("a column break in the last column of a page continues at the first column of the next page (AC-11)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.keyboard.type("FirstColumnFiller");
    await applyColumns(page, 2);
    await content(page).getByText("FirstColumnFiller").click();
    await page.keyboard.press("End");
    // First break: column 0 -> column 1 (still the same page).
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("break-column").click();
    await page.keyboard.type("SecondColumnMarker");
    await page.keyboard.press("End");
    // Second break: column 1 is the page's last column, so this one must
    // continue on the next page's first column (AC-11).
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("break-column").click();
    await page.keyboard.type("NextPageMarker");
    const location = await locate(page, "NextPageMarker");
    expect(location).toEqual({ pageIndex: 1, column: 0 });
  });

  test("reducing to a single column makes an existing column break a true no-op, and restoring columns resumes it (AC-12, AC-13)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.keyboard.type("BeforeBreak");
    await applyColumns(page, 3);
    await content(page).getByText("BeforeBreak").click();
    await page.keyboard.press("End");
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("break-column").click();
    await page.keyboard.type("AfterBreak");
    expect(await docColumns(page).count()).toBeGreaterThanOrEqual(2);

    // AC-12: collapse to 1 column — the break becomes inert: single page,
    // both paragraphs present, no column boxes at all.
    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("columns-button").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "visible" });
    await page.getByTestId("columns-preset-1").click();
    await page.getByTestId("columns-confirm").click();
    await page.getByTestId("columns-dialog").waitFor({ state: "detached" });
    await expect(docColumns(page)).toHaveCount(0);
    await expect(page.getByTestId("document-page")).toHaveCount(1);
    await expect(content(page)).toContainText("BeforeBreak");
    await expect(content(page)).toContainText("AfterBreak");

    // AC-13: restoring columns resumes the break's normal behaviour.
    await applyColumns(page, 2);
    const colTexts = await docColumns(page).allInnerTexts();
    expect(colTexts[0]).toContain("BeforeBreak");
    expect(colTexts[0]).not.toContain("AfterBreak");
    expect(colTexts.some((t) => t.includes("AfterBreak"))).toBe(true);
  });

  test("equal-width columns round-trip through .docx with no w:col children (AC-16a)", async ({ page }) => {
    await createDocument(page);
    await typeParagraphs(page, 4);
    await applyColumns(page, 2);
    await page.getByTestId("menu-title-file").click();
    await page.getByTestId("file-menu-export").click();
    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-docx").click(),
    ]);
    const path = await download.path();
    // Read the .docx (a ZIP) via the platform's DecompressionStream, the same
    // approach lib/core/zip.ts uses to parse packages — no new dependency.
    const fs = await import("node:fs");
    const bytes = new Uint8Array(fs.readFileSync(path!));
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocdOffset = -1;
    for (let o = bytes.byteLength - 22; o >= Math.max(0, bytes.byteLength - 22 - 0xffff); o--) {
      if (view.getUint32(o, true) === 0x06054b50) {
        eocdOffset = o;
        break;
      }
    }
    expect(eocdOffset).toBeGreaterThanOrEqual(0);
    const totalEntries = view.getUint16(eocdOffset + 10, true);
    let offset = view.getUint32(eocdOffset + 16, true);
    const decoder = new TextDecoder("utf-8");
    let documentXmlBytes: Uint8Array | null = null;
    for (let i = 0; i < totalEntries; i++) {
      const compressionMethod = view.getUint16(offset + 10, true);
      const compressedSize = view.getUint32(offset + 20, true);
      const fileNameLength = view.getUint16(offset + 28, true);
      const extraFieldLength = view.getUint16(offset + 30, true);
      const fileCommentLength = view.getUint16(offset + 32, true);
      const localHeaderOffset = view.getUint32(offset + 42, true);
      const nameStart = offset + 46;
      const fileName = decoder.decode(bytes.subarray(nameStart, nameStart + fileNameLength));
      if (fileName === "word/document.xml") {
        const lo = localHeaderOffset;
        const fnLen = view.getUint16(lo + 26, true);
        const efLen = view.getUint16(lo + 28, true);
        const dataStart = lo + 30 + fnLen + efLen;
        const compressed = bytes.subarray(dataStart, dataStart + compressedSize);
        if (compressionMethod === 0) {
          documentXmlBytes = compressed;
        } else {
          const stream = new Blob([compressed as BlobPart]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
          const chunks: Uint8Array[] = [];
          const reader = stream.getReader();
          for (;;) {
            // eslint-disable-next-line no-await-in-loop
            const { done, value } = await reader.read();
            if (done) break;
            if (value) chunks.push(value);
          }
          const total = chunks.reduce((s, c) => s + c.length, 0);
          const out = new Uint8Array(total);
          let pos = 0;
          for (const c of chunks) {
            out.set(c, pos);
            pos += c.length;
          }
          documentXmlBytes = out;
        }
        break;
      }
      offset = nameStart + fileNameLength + extraFieldLength + fileCommentLength;
    }
    expect(documentXmlBytes).not.toBeNull();
    const xml = new TextDecoder("utf-8").decode(documentXmlBytes!);
    const colsMatch = xml.match(/<w:cols\b[^>]*>[\s\S]*?<\/w:cols>|<w:cols\b[^>]*\/>/);
    expect(colsMatch).not.toBeNull();
    expect(colsMatch![0]).toMatch(/w:num="2"/);
    expect(colsMatch![0]).not.toContain("<w:col ");
  });

  test("a column break survives a .docx round-trip as w:br type=column (AC-18)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.keyboard.type("BeforeBreak");
    await applyColumns(page, 2);
    await content(page).getByText("BeforeBreak").click();
    await page.keyboard.press("End");
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId("break-column").click();
    await page.keyboard.type("AfterBreak");
    await page.getByTestId("menu-title-file").click();
    await page.getByTestId("file-menu-export").click();
    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-docx").click(),
    ]);
    const path = await download.path();
    await openFixture(page, path!);
    const colTexts = await docColumns(page).allInnerTexts();
    expect(colTexts[0]).toContain("BeforeBreak");
    expect(colTexts[0]).not.toContain("AfterBreak");
    expect(colTexts.some((t) => t.includes("AfterBreak"))).toBe(true);
  });

  test("arrow keys cross a page boundary between columns (AC-21)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
    await page.keyboard.type(Array.from({ length: 160 }, (_, i) => `Row${i}`).join("\n"));
    await applyColumns(page, 2);
    expect(await page.getByTestId("document-page").count()).toBeGreaterThan(1);

    // Down at the last column of page 0 continues at the first column of page 1.
    const page0Col1Last = page
      .getByTestId("document-page")
      .nth(0)
      .locator('.doc-column[data-column="1"] .doc-paragraph')
      .last();
    await page0Col1Last.click();
    await page.keyboard.press("End");
    expect(await caretLocation(page)).toEqual({ column: "1", pageIndex: 0 });
    await page.keyboard.press("ArrowDown");
    expect(await caretLocation(page)).toEqual({ column: "0", pageIndex: 1 });

    // Up at the first column of page 1 continues at the last column of page 0.
    const page1Col0First = page
      .getByTestId("document-page")
      .nth(1)
      .locator('.doc-column[data-column="0"] .doc-paragraph')
      .first();
    await page1Col0First.click();
    await page.keyboard.press("Home");
    expect(await caretLocation(page)).toEqual({ column: "0", pageIndex: 1 });
    await page.keyboard.press("ArrowUp");
    expect(await caretLocation(page)).toEqual({ column: "1", pageIndex: 0 });
  });
});
