import { test, expect, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzip } from "../src/lib/core/zip";

// SPEC-136 — page setup, printing and PDF export for sheets (AC-77–AC-83),
// with the real pointer and keyboard, against the Excel- and
// LibreOffice-shaped spec136-print fixtures: "Budget" A1:H70 (row 1 the print
// titles, A30:B31 merged across the natural break, column D hidden), A4
// landscape, fit 1 wide × automatic tall, a manual break above row 36 that
// fit-to-page ignores; "Notes" a one-page sheet.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 } });

async function open(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
}

/** Clicks the centre of a cell through its column and row headers. */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
}

async function openPageSetup(page: Page): Promise<void> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("sheet-menu-page-setup").click();
  await expect(page.getByTestId("sheet-page-setup-dialog")).toBeVisible();
}

async function openPreview(page: Page): Promise<void> {
  await page.keyboard.press("ControlOrMeta+p");
  await expect(page.getByTestId("print-preview")).toBeVisible();
}

const pdfText = (bytes: Buffer): string => bytes.toString("latin1");

/** AC-77 and its .ods twin: the dialog shows A4, landscape, fit 1 × automatic and the header. */
async function expectFixtureSetup(page: Page, header: { left: string; center: string }): Promise<void> {
  await openPageSetup(page);
  await expect(page.getByTestId("sheet-page-paper")).toHaveValue("A4");
  await expect(page.getByTestId("sheet-page-landscape")).toBeChecked();
  await expect(page.getByTestId("sheet-page-scale-mode-fit")).toBeChecked();
  await expect(page.getByTestId("sheet-page-fit-wide")).toHaveValue("1");
  await expect(page.getByTestId("sheet-page-fit-tall")).toHaveValue("");
  await page.getByTestId("sheet-page-setup-tab-headerFooter").click();
  await expect(page.getByTestId("sheet-header-left")).toHaveValue(header.left);
  await expect(page.getByTestId("sheet-header-center")).toHaveValue(header.center);
  await page.getByTestId("sheet-page-setup-cancel").click();
  await expect(page.getByTestId("sheet-page-setup-dialog")).toBeHidden();
}

/** AC-78 and its .ods twin: three pages, the title row atop page 2, the merge whole on page 2. */
async function expectThreePages(page: Page): Promise<void> {
  await openPreview(page);
  await expect(page.getByTestId("print-page-count")).toHaveText("3 pages");
  const pages = page.getByTestId("sheet-print-page");
  await expect(pages).toHaveCount(3);
  await expect(pages.nth(1)).toHaveAttribute("aria-label", "Page 2 of 3");
  const second = pages.nth(1).getByTestId("sheet-print-cell");
  await expect(second.first()).toHaveAttribute("data-ref", "A1");
  await expect(second.first()).toHaveText("Account");
  await expect(pages.nth(1).locator('[data-testid="sheet-print-cell"][data-ref="A30"]')).toHaveAttribute("data-rows", "2");
  await expect(pages.nth(1).locator('[data-testid="sheet-print-cell"][data-ref="A30"]')).toHaveText("Merged total");
  await expect(pages.nth(0).locator('[data-ref="A30"], [data-ref="A31"]')).toHaveCount(0);
  await expect(pages.nth(1).getByTestId("sheet-print-heading").first()).toBeVisible();
  // Closing the preview turns the page-break lines on, as Excel does.
  await page.getByTestId("print-preview-close").click();
  await expect(page.getByTestId("print-preview")).toBeHidden();
  await expect(page.getByTestId("sheet-page-break-line").first()).toBeAttached();
  await expect(page.getByTestId("sheet-menu-page-breaks")).toHaveCount(1);
}

test("spec136-print.xlsx: page setup, the preview, a break, PDF, the workbook scope, 50% and a round trip (AC-77–AC-82)", async ({ page }) => {
  await open(page, join(fixtures, "spec136-print.xlsx"));

  // AC-77
  await expectFixtureSetup(page, { left: '&"Arial,Bold"&12Budget', center: "&P of &N" });

  // AC-78
  await expectThreePages(page);
  await expect(page.locator('[data-testid="sheet-page-break-line"][data-axis="row"][data-index="29"]')).toHaveAttribute("data-kind", "auto");

  // AC-79: a page break at row 30 (A30: a row break above it) — the sheet leaves fit-to-page.
  await clickCell(page, "A30");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("sheet-menu-insert-page-break").click();
  await expect(page.locator('[data-testid="sheet-page-break-line"][data-axis="row"][data-index="29"]')).toHaveAttribute("data-kind", "manual");
  await openPreview(page);
  await expect(page.getByTestId("print-page-count")).toHaveText("4 pages");
  await expect(page.getByTestId("sheet-print-page")).toHaveCount(4);

  // AC-80: Download PDF.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("sheet-print-download-pdf").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("spec136-print.pdf");
  const pdf = pdfText(readFileSync(await download.path()));
  expect(pdf.startsWith("%PDF-1.7")).toBe(true);
  expect(/\/Type \/Pages \/Kids \[[^\]]*\] \/Count (\d+)/.exec(pdf)?.[1]).toBe("4");
  expect(/\/MediaBox \[([^\]]*)\]/.exec(pdf)?.[1]).toBe("0 0 841.89 595.28");
  expect(pdf).toContain("/Title (spec136-print)");

  // AC-81: the entire workbook sums both sheets' pages; 50% gives fewer.
  await page.getByTestId("sheet-print-scope").selectOption("workbook");
  await expect(page.getByTestId("print-page-count")).toHaveText("5 pages");
  await page.getByTestId("sheet-print-page-setup").click();
  await page.getByTestId("sheet-page-scale-mode-percent").check();
  await page.getByTestId("sheet-page-scale").fill("50");
  await page.getByTestId("sheet-page-setup-ok").click();
  await expect(page.getByTestId("sheet-page-setup-dialog")).toBeHidden();
  await expect(page.getByTestId("print-page-count")).toHaveText("4 pages");
  await page.getByTestId("print-preview-close").click();

  // AC-82: download the workbook, reopen — the setup, the break and the print area come back.
  const savedPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const saved = await savedPromise;
  const path = join(tmpdir(), `spec136-${Date.now()}-print.xlsx`);
  await saved.saveAs(path);
  await open(page, path);
  await openPageSetup(page);
  await expect(page.getByTestId("sheet-page-paper")).toHaveValue("A4");
  await expect(page.getByTestId("sheet-page-landscape")).toBeChecked();
  await expect(page.getByTestId("sheet-page-scale-mode-percent")).toBeChecked();
  await expect(page.getByTestId("sheet-page-scale")).toHaveValue("50");
  await page.getByTestId("sheet-page-setup-tab-sheet").click();
  await expect(page.getByTestId("sheet-print-area")).toHaveValue("$A$1:$H$70");
  await expect(page.getByTestId("sheet-print-title-rows")).toHaveValue("$1:$1");
  await page.getByTestId("sheet-page-setup-cancel").click();
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("sheet-menu-page-breaks").click();
  await expect(page.locator('[data-testid="sheet-page-break-line"][data-axis="row"][data-index="29"]')).toHaveAttribute("data-kind", "manual");
  await expect(page.locator('[data-testid="sheet-page-break-line"][data-axis="row"][data-index="35"]')).toHaveAttribute("data-kind", "manual");
});

test("spec136-print.xlsx in one session: each preview shows the sheet on show and the model as it is now — a link typed after the first one reaches the PDF — the page-break lines follow a cell, and an empty sheet has nothing to print (AC-31, AC-36, AC-38, AC-49, AC-56)", async ({ page }) => {
  await open(page, join(fixtures, "spec136-print.xlsx"));
  const tab = (name: string) => page.getByTestId("sheet-tab").filter({ hasText: name });
  const count = page.getByTestId("print-page-count");
  const printCell = (ref: string) => page.locator(`[data-testid="sheet-print-cell"][data-ref="${ref}"]`);

  // AC-31: a tab switch between previews changes the sheet they print.
  await openPreview(page);
  await expect(count).toHaveText("3 pages");
  await page.getByTestId("print-preview-close").click();
  await tab("Notes").click();
  await openPreview(page);
  await expect(count).toHaveText("1 page");
  await expect(printCell("A1")).toHaveText("Prepared by");
  await page.getByTestId("print-preview-close").click();
  await tab("Budget").click();
  await openPreview(page);
  await expect(count).toHaveText("3 pages");
  await expect(printCell("A1").first()).toHaveText("Account");
  await expect(page.getByTestId("sheet-print-cell").filter({ hasText: "Prepared by" })).toHaveCount(0);
  await page.getByTestId("print-preview-close").click();

  // AC-38, AC-49: a HYPERLINK() typed after those previews prints as a link and becomes a /Link annotation.
  await clickCell(page, "H2");
  await page.keyboard.type('=HYPERLINK("https://example.com/budget","Site")');
  await page.keyboard.press("Enter");
  await expect(page.locator('[data-testid="sheet-cell"][data-ref="H2"]')).toHaveText("Site");
  await openPreview(page);
  const link = printCell("H2").locator("a[data-doc-link]");
  await expect(link).toHaveText("Site");
  await expect(link).toHaveAttribute("href", "https://example.com/budget");
  // The preview never follows it.
  await link.dispatchEvent("click");
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("print-preview")).toBeVisible();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("sheet-print-download-pdf").click();
  const pdf = pdfText(readFileSync(await (await downloadPromise).path()));
  expect(pdf).toContain("/Subtype /Link");
  expect(pdf).toContain("/URI (https://example.com/budget)");
  await page.getByTestId("print-preview-close").click();

  // AC-56: the lines closing the preview turned on follow the model — the
  // print area cleared (a name), then a value typed below the data.
  const rowLine = (index: number) => page.locator(`[data-testid="sheet-page-break-line"][data-axis="row"][data-index="${index}"]`);
  await expect(page.locator('[data-testid="sheet-page-break-line"][data-axis="row"]')).toHaveCount(2);
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("sheet-menu-print-area").click();
  await page.getByTestId("sheet-menu-clear-print-area").click();
  await expect(page.getByTestId("sheet-menu-print-area-flyout")).toBeHidden();
  await expect(rowLine(89)).toHaveCount(0);
  await page.getByTestId("name-box").fill("A120");
  await page.getByTestId("name-box").press("Enter");
  await page.keyboard.type("Below the data");
  await page.keyboard.press("Enter");
  await expect(page.locator('[data-testid="sheet-cell"][data-ref="A120"]')).toHaveText("Below the data");
  await expect(rowLine(89)).toHaveAttribute("data-kind", "auto");

  // AC-36: a new sheet is empty — nothing to print, and Print is off.
  await page.getByTestId("sheet-add").click();
  await openPreview(page);
  await expect(count).toHaveText("0 pages");
  await expect(page.getByTestId("print-no-pages")).toHaveText("This sheet has nothing to print.");
  await expect(page.getByTestId("print-confirm")).toBeDisabled();
  await expect(page.getByTestId("sheet-print-page")).toHaveCount(0);
});

test("spec136-print.ods: the read half, and a header edit on Budget leaves Notes' master page untouched (AC-83)", async ({ page }) => {
  const source = await unzip(new Uint8Array(readFileSync(join(fixtures, "spec136-print.ods"))));
  const sourceStyles = new TextDecoder().decode(source.get("styles.xml"));
  const defaultMaster = /<style:master-page style:name="Default"[\s\S]*?<\/style:master-page>/.exec(sourceStyles)![0];
  await open(page, join(fixtures, "spec136-print.ods"));
  await expectFixtureSetup(page, { left: "&A", center: "&P of &N" });
  await expectThreePages(page);

  // Edit Budget's header, download.
  await openPageSetup(page);
  await page.getByTestId("sheet-page-setup-tab-headerFooter").click();
  await page.getByTestId("sheet-header-center").fill("Budget &P of &N");
  await page.getByTestId("sheet-page-setup-ok").click();
  await expect(page.getByTestId("sheet-page-setup-dialog")).toBeHidden();
  const savedPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const saved = await savedPromise;
  const written = await unzip(new Uint8Array(readFileSync(await saved.path())));
  const styles = new TextDecoder().decode(written.get("styles.xml"));
  const content = new TextDecoder().decode(written.get("content.xml"));
  // The shared master page keeps its bytes; Budget's table names a minted one; Notes still the shared one.
  expect(styles).toContain(defaultMaster);
  expect(styles).toMatch(/<style:master-page style:name="de-mp\d+" style:page-layout-name="de-pm\d+">/);
  const budgetStyle = /<table:table table:name="Budget" table:style-name="([^"]+)"/.exec(content)![1]!;
  expect(budgetStyle).toMatch(/^de-ta\d+$/);
  expect(content).toMatch(new RegExp(`<style:style style:name="${budgetStyle}" style:family="table" style:master-page-name="de-mp\\d+"`));
  expect(content).toContain(`<table:table table:name="Notes" table:style-name="ta1">`);
});
