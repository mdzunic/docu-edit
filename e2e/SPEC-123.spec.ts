import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzip } from "../src/lib/core/zip";

// SPEC-123 — cell formatting (AC-73, AC-74), with the real pointer and
// keyboard: open spec123-styles.xlsx → select B2:C3 → Bold, fill Accent 1
// lighter 40%, All borders → Currency from the dropdown (B2 reads
// "$1,234.50") → select D2 (0.125) → Ctrl/Cmd+1 → Custom 0.0% previews
// "12.5%" → OK → apply "Good" from the gallery → format painter from B2 onto
// E2:E4 → select column F → Bold (no stored cell beyond the used rows: the
// column's style lands on its span, read from a download) → Clear formats on
// C3 → undo twice → download → reopen → the same formatting and named
// styles. Then the same core flow on spec123-styles.ods.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

// AC-19: at 1440px no band of the strip collapses.
test.use({ viewport: { width: 1440, height: 900 } });

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

/** Clicks the centre of a cell through its column and row headers. */
async function clickCell(page: Page, ref: string, modifiers: ("Shift" | "ControlOrMeta")[] = []): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  for (const key of modifiers) {
    await page.keyboard.down(key);
  }
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  for (const key of modifiers) {
    await page.keyboard.up(key);
  }
}

async function cellCenter(page: Page, ref: string): Promise<{ x: number; y: number }> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  return { x: column!.x + column!.width / 2, y: row!.y + row!.height / 2 };
}

async function open(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("sheet-format-toolbar")).toBeVisible();
  await expect(page.getByTestId("sheet-cell-styles")).toBeVisible();
  await expect(page.getByTestId("sheet-format-toolbar").getByTestId("toolbar-overflow-button")).toBeHidden();
}

async function download(page: Page, fileName: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const saved = await downloadPromise;
  const path = join(tmpdir(), `spec123-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  return path;
}

const fontWeight = (page: Page, ref: string): Promise<string> =>
  cell(page, ref).evaluate((element) => getComputedStyle(element.querySelector(".sheet-cell-text") ?? element).fontWeight);

/** The shared flow up to the named style and the painter (AC-73's first half). */
async function formatCoreFlow(page: Page): Promise<void> {
  await clickCell(page, "B2");
  await clickCell(page, "C3", ["Shift"]);
  await expect(page.getByTestId("name-box")).toHaveValue("B2");

  await page.getByTestId("sheet-bold").click();
  await expect(page.getByTestId("sheet-bold")).toHaveAttribute("aria-pressed", "true");
  expect(Number(await fontWeight(page, "C3"))).toBeGreaterThanOrEqual(700);

  await page.getByTestId("sheet-fill-menu-button").click();
  await page.getByTestId("sheet-fill-menu").getByLabel("Accent 1, lighter 40%").click();
  await page.getByTestId("sheet-borders-menu-button").click();
  await page.getByTestId("sheet-border-all").click();

  await page.getByTestId("sheet-number-format").click();
  await expect(page.getByTestId("sheet-number-format-currency")).toContainText("$1,234.50");
  await page.getByTestId("sheet-number-format-currency").click();
  await expect(cell(page, "B2")).toHaveText("$1,234.50");
  await expect(page.getByTestId("sheet-number-format")).toContainText("Currency");

  // D2 (0.125) through Format Cells' Custom category.
  await clickCell(page, "D2");
  await page.keyboard.press("ControlOrMeta+1");
  const dialog = page.getByTestId("format-cells-dialog");
  await expect(dialog).toBeVisible();
  await page.getByTestId("format-cells-category").selectOption("custom");
  await page.getByTestId("format-cells-custom-code").fill("0.0%");
  await expect(page.getByTestId("format-cells-sample")).toHaveText("12.5%");
  await expect(page.getByTestId("format-cells-preview")).toContainText("12.5%");
  await page.getByTestId("format-cells-ok").click();
  await expect(dialog).toHaveCount(0);
  await expect(cell(page, "D2")).toHaveText("12.5%");

  // The "Good" cell style from the gallery.
  await page.getByTestId("sheet-cell-styles").click();
  await page.locator('[data-testid="sheet-cell-style"][data-style="Good"]').click();
  await expect(cell(page, "D2")).toHaveText("12.5%");

  // The format painter from B2 onto E2:E4.
  await clickCell(page, "B2");
  await page.getByTestId("sheet-format-painter").click();
  await expect(page.getByTestId("spreadsheet-grid")).toHaveAttribute("data-format-painter", "once");
  const from = await cellCenter(page, "E2");
  const to = await cellCenter(page, "E4");
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 4 });
  await page.mouse.up();
  await expect(page.getByTestId("spreadsheet-grid")).not.toHaveAttribute("data-format-painter", /.+/);
  await expect(cell(page, "E2")).toHaveText("$10.00");
  await expect(cell(page, "E4")).toHaveText("$30.00");
}

async function expectFormatted(page: Page): Promise<void> {
  await expect(cell(page, "B2")).toHaveText("$1,234.50");
  await expect(cell(page, "C3")).toHaveText("$75.00");
  await expect(cell(page, "D2")).toHaveText("12.5%");
  await expect(cell(page, "E3")).toHaveText("$20.00");
  expect(Number(await fontWeight(page, "B2"))).toBeGreaterThanOrEqual(700);
  await clickCell(page, "C3");
  await expect(page.getByTestId("sheet-bold")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("sheet-number-format")).toContainText("Currency");
}

test("format spec123-styles.xlsx, undo, download and reopen (AC-73)", async ({ page }) => {
  await open(page, join(fixtures, "spec123-styles.xlsx"));
  await formatCoreFlow(page);

  // Column F bold: one span, no stored cell past the used rows.
  await page.getByTestId("sheet-col-header").filter({ hasText: /^F$/ }).click();
  await page.getByTestId("sheet-bold").click();
  expect(Number(await fontWeight(page, "F2"))).toBeGreaterThanOrEqual(700);
  const midway = await unzip(new Uint8Array(readFileSync(await download(page, "midway.xlsx"))));
  const budget = new TextDecoder().decode(midway.get("xl/worksheets/sheet1.xml")!);
  expect(budget).toMatch(/<col min="6" max="6"[^>]* style="\d+"/);
  const storedRows = [...budget.matchAll(/<c r="F(\d+)"/g)].map((match) => Number(match[1]));
  expect(Math.max(...storedRows)).toBeLessThanOrEqual(6);

  await clickCell(page, "C3");
  await page.getByTestId("sheet-clear").click();
  await page.getByTestId("sheet-clear-formats").click();
  await expect(cell(page, "C3")).toHaveText("75");

  await page.keyboard.press("ControlOrMeta+z");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "C3")).toHaveText("$75.00");
  expect(Number(await fontWeight(page, "F2"))).toBeLessThan(700);

  const path = await download(page, "styles.xlsx");
  const saved = await unzip(new Uint8Array(readFileSync(path)));
  const styles = new TextDecoder().decode(saved.get("xl/styles.xml")!);
  // The source's named styles stay; D2's new xf names Good (cellStyleXfs 1).
  expect(styles).toContain(`<cellStyle name="Good" xfId="1" builtinId="26"/>`);
  expect(styles).toMatch(/<xf numFmtId="\d+" fontId="\d+" fillId="\d+" borderId="\d+" xfId="1"/);
  expect(styles).toContain(`<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>`);

  await open(page, path);
  await expectFormatted(page);
  // Format Cells reopens D2's 0.0% as what it is: Percentage, one decimal.
  await clickCell(page, "D2");
  await page.keyboard.press("ControlOrMeta+1");
  await expect(page.getByTestId("format-cells-category")).toHaveValue("percentage");
  await expect(page.getByTestId("format-cells-decimals")).toHaveValue("1");
});

test("format spec123-styles.ods, download and reopen (AC-74)", async ({ page }) => {
  await open(page, join(fixtures, "spec123-styles.ods"));
  await formatCoreFlow(page);
  const path = await download(page, "styles.ods");
  const saved = await unzip(new Uint8Array(readFileSync(path)));
  const content = new TextDecoder().decode(saved.get("content.xml")!);
  expect(content).toMatch(/style:parent-style-name="Good"/);
  await open(page, path);
  await expectFormatted(page);
});

/** "#rrggbb" as getComputedStyle reports it. */
function rgb(hex: string | null): string {
  const [r, g, b] = [1, 3, 5].map((at) => parseInt(hex!.slice(at, at + 2), 16));
  return `rgb(${r}, ${g}, ${b})`;
}

// A theme colour picked inside Format Cells reached patchStyle as a Vue
// reactive Proxy (the dialog held it in a deep ref): structuredClone threw
// DataCloneError, the preview never showed the colour and OK silently
// applied nothing.
test("Format Cells applies theme colours to a mixed selection without flattening it (AC-42, AC-43)", async ({ page }) => {
  const errors: string[] = [];
  page.on("console", (message) => {
    if (message.type() === "error") {
      errors.push(message.text());
    }
  });
  page.on("pageerror", (error) => errors.push(error.message));
  await open(page, join(fixtures, "spec123-styles.xlsx"));
  // B2 bold, the rest of B2:C3 not — B2 active.
  await clickCell(page, "B2");
  await page.getByTestId("sheet-bold").click();
  await clickCell(page, "C3", ["Shift"]);
  await page.keyboard.press("ControlOrMeta+1");
  const dialog = page.getByTestId("format-cells-dialog");
  await expect(dialog).toBeVisible();

  await page.getByTestId("format-cells-tab-fill").click();
  const fillSwatch = page.getByTestId("format-cells-fill-theme-swatch").and(page.getByLabel("Accent 2", { exact: true }));
  const fill = rgb(await fillSwatch.getAttribute("data-color"));
  await fillSwatch.click();
  await expect(page.getByTestId("format-cells-preview")).toHaveCSS("background-color", fill);

  await page.getByTestId("format-cells-tab-font").click();
  const fontColor = page.getByTestId("format-cells-font-color");
  await fontColor.getByTestId("text-color-button").click();
  const textSwatch = fontColor.getByTestId("text-color-theme-swatch").and(page.getByLabel("Accent 6", { exact: true }));
  const text = rgb(await textSwatch.getAttribute("data-color"));
  await textSwatch.click();
  await expect(page.getByTestId("format-cells-preview").locator("span")).toHaveCSS("color", text);

  await page.getByTestId("format-cells-ok").click();
  await expect(dialog).toHaveCount(0);
  await clickCell(page, "A4");
  for (const ref of ["B2", "C2", "B3", "C3"]) {
    await expect(cell(page, ref)).toHaveCSS("background-color", fill);
    await expect(cell(page, ref).locator(".sheet-cell-text")).toHaveCSS("color", text);
  }
  // Only the changed fields applied: B2 stays bold, C3 stays regular.
  expect(Number(await fontWeight(page, "B2"))).toBeGreaterThanOrEqual(700);
  expect(Number(await fontWeight(page, "C3"))).toBeLessThan(700);
  await expect(page.getByTestId("spreadsheet-undo")).toHaveAttribute("title", /Format cells/);
  expect(errors).toEqual([]);
});
