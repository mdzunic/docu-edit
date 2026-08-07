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

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
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
  await page.getByTestId("columns-button").click();
  await expect(page.getByTestId("columns-dialog")).toBeVisible();
  await page.getByTestId(`columns-preset-${count}`).click();
  await page.getByTestId("columns-confirm").click();
  await expect(page.getByTestId("columns-dialog")).toHaveCount(0);
}

test.describe("SPEC-030 multi-column layout", () => {
  test("the dialog offers 1/2/3 presets and a 1-6 custom count (AC-1, AC-2)", async ({ page }) => {
    await createDocument(page);
    await content(page).click();
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
    await page.getByTestId("insert-table-button").click();
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
    await page.getByTestId("insert-break-button").click();
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
    await page.getByTestId("download-menu").click();
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
    await page.getByTestId("download-menu").click();
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
});
