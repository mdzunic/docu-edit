import { test, expect, type Locator, type Page } from "@playwright/test";
import { existsSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-125 — copy, paste, fill and move (AC-118, AC-119), with the real
// keyboard, pointer and clipboard (Chromium, clipboard-read/clipboard-write
// granted): open spec125-fill.xlsx → copy B2:C3 (the marching ants show) →
// paste at E5 → E6's formula bar shows the shifted formula → Escape clears
// the ants → cut A10:A12 → paste at H10 → Sheet2's SUM over A10:A12 now
// reads H10:H12 → [only when e2e/fixtures/spec125-clipboard/excel.html is
// captured: write it to the clipboard, Ctrl/Cmd+V, the currency cell is the
// number 1234.5 with its format, Ctrl/Cmd+Shift+V with Values + Transpose]
// → drag the fill handle from "Mon" down six rows → Tue…Sun → "Copy cells"
// in the chip → every cell reads "Mon" → double-click the fill handle beside
// a 20-row column → it fills 20 rows → Ctrl/Cmd+D → Alt/Option-drag a block
// to copy it → undo three times → download → reopen → the cells as left.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const EXCEL_CAPTURE = join(fixtures, "spec125-clipboard", "excel.html");

test.use({ viewport: { width: 1440, height: 900 }, permissions: ["clipboard-read", "clipboard-write"] });

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

function colHeader(page: Page, letters: string): Locator {
  return page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) });
}

function rowHeader(page: Page, row: number): Locator {
  return page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${row}$`) });
}

function sheetTab(page: Page, name: string): Locator {
  return page.getByTestId("sheet-tab").filter({ hasText: new RegExp(`^\\s*${name}\\s*$`) });
}

/** The centre of a cell, through its column and row headers. */
async function centre(page: Page, ref: string): Promise<{ x: number; y: number }> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = (await colHeader(page, letters!).boundingBox())!;
  const row = (await rowHeader(page, Number(digits)).boundingBox())!;
  return { x: column.x + column.width / 2, y: row.y + row.height / 2 };
}

async function clickCell(page: Page, ref: string, modifiers: "Shift"[] = []): Promise<void> {
  const at = await centre(page, ref);
  for (const key of modifiers) {
    await page.keyboard.down(key);
  }
  await page.mouse.click(at.x, at.y);
  for (const key of modifiers) {
    await page.keyboard.up(key);
  }
}

async function open(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
}

async function download(page: Page, fileName: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const saved = await downloadPromise;
  const path = join(tmpdir(), `spec125-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  return path;
}

const formulaBar = (page: Page): Locator => page.getByTestId("formula-bar-content");

test("copy, paste, cut, fill, move and undo on spec125-fill.xlsx, then download and reopen (AC-118, AC-119)", async ({ page }) => {
  await open(page, join(fixtures, "spec125-fill.xlsx"));
  const ants = page.getByTestId("sheet-copy-marquee");

  // Copy B2:C3: the marching ants outline it.
  await clickCell(page, "B2");
  await clickCell(page, "C3", ["Shift"]);
  await page.keyboard.press("ControlOrMeta+c");
  await expect(ants).toHaveAttribute("data-mode", "copy");

  // Paste at E5: E6 holds the shifted formula; the ants stay for another paste.
  await clickCell(page, "E5");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "E5")).toHaveText("10");
  await expect(cell(page, "F6")).toHaveText("30");
  await clickCell(page, "E6");
  await expect(formulaBar(page)).toHaveText("=E5*2");
  await expect(ants).toHaveCount(1);

  // Escape clears the ants.
  await page.keyboard.press("Escape");
  await expect(ants).toHaveCount(0);

  // Cut A10:A12 and paste at H10: the cells move, and Sheet2's SUM follows them.
  await clickCell(page, "A10");
  await clickCell(page, "A12", ["Shift"]);
  await page.keyboard.press("ControlOrMeta+x");
  await expect(ants).toHaveAttribute("data-mode", "cut");
  await expect(cell(page, "A10")).toHaveText("5");
  await clickCell(page, "H10");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "H10")).toHaveText("5");
  await expect(cell(page, "H12")).toHaveText("7");
  await expect(cell(page, "A10")).toHaveText("");
  await expect(ants).toHaveCount(0);
  await sheetTab(page, "Sheet2").click();
  await clickCell(page, "A1");
  await expect(formulaBar(page)).toHaveText("=SUM(Sheet1!H10:H12)");
  await expect(cell(page, "A1")).toHaveText("18");
  await sheetTab(page, "Sheet1").click();

  // The external paste runs only once excel.html is a real capture.
  const captured = existsSync(EXCEL_CAPTURE);
  if (captured) {
    await page.evaluate(async (html) => {
      await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }) })]);
    }, readFileSync(EXCEL_CAPTURE, "utf8"));
    await clickCell(page, "A40");
    await page.keyboard.press("ControlOrMeta+v");
    await clickCell(page, "C40");
    await expect(formulaBar(page)).toHaveText("1234.5");
    await expect(cell(page, "C40")).toHaveText(/\$1,234\.50/);
    await clickCell(page, "A45");
    await page.keyboard.press("ControlOrMeta+Shift+v");
    const dialog = page.getByTestId("sheet-paste-special-dialog");
    await expect(dialog).toBeVisible();
    await page.getByTestId("sheet-paste-special-values").check();
    await page.getByTestId("sheet-paste-special-transpose").check();
    await page.getByTestId("sheet-paste-special-ok").click();
    await expect(dialog).toHaveCount(0);
    await clickCell(page, "A47");
    await expect(formulaBar(page)).toHaveText("1234.5");
  } else {
    test.info().annotations.push({ type: "fixture", description: "fixture not yet captured: excel.html" });
  }

  // Drag the fill handle from J2 ("Mon") down six rows: Tue…Sun.
  await clickCell(page, "J2");
  const handle = page.getByTestId("sheet-fill-handle");
  await expect(handle).toHaveCSS("cursor", "crosshair");
  const grip = (await handle.boundingBox())!;
  expect([Math.round(grip.width), Math.round(grip.height)]).toEqual([7, 7]);
  await page.mouse.move(grip.x + grip.width / 2, grip.y + grip.height / 2);
  await page.mouse.down();
  const j8 = await centre(page, "J8");
  await page.mouse.move(j8.x, j8.y, { steps: 8 });
  await expect(page.getByTestId("sheet-fill-preview")).toBeVisible();
  await page.mouse.up();
  await expect(page.getByTestId("sheet-fill-preview")).toHaveCount(0);
  const days = ["Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  for (const [index, day] of days.entries()) {
    await expect(cell(page, `J${index + 3}`)).toHaveText(day);
  }

  // The chip: Fill series is pressed; Copy cells re-applies the fill as a copy.
  const chip = page.getByTestId("sheet-fill-options");
  await expect(chip).toHaveAttribute("role", "toolbar");
  await expect(chip.getByTestId("sheet-fill-series")).toHaveAttribute("aria-pressed", "true");
  await chip.getByTestId("sheet-fill-copy").click();
  for (let row = 3; row <= 8; row += 1) {
    await expect(cell(page, `J${row}`)).toHaveText("Mon");
  }
  await expect(chip.getByTestId("sheet-fill-copy")).toHaveAttribute("aria-pressed", "true");
  await page.keyboard.press("Escape");
  await expect(chip).toHaveCount(0);

  // Double-click the fill handle on M2, beside L2:L21: it fills M2:M21.
  await clickCell(page, "M2");
  const seed = (await page.getByTestId("sheet-fill-handle").boundingBox())!;
  await page.mouse.dblclick(seed.x + seed.width / 2, seed.y + seed.height / 2);
  await expect(cell(page, "M21")).toHaveText("40");
  await expect(cell(page, "M22")).toHaveText("");

  // Ctrl/Cmd+D fills N2's "Total" down N3:N6.
  await clickCell(page, "N2");
  await clickCell(page, "N6", ["Shift"]);
  await page.keyboard.press("ControlOrMeta+d");
  await expect(cell(page, "N6")).toHaveText("Total");

  // Alt/Option-drag B2:C3 by its outline to P2: a copy (formulas shifted).
  await clickCell(page, "B2");
  await clickCell(page, "C3", ["Shift"]);
  const outline = (await page.getByTestId("sheet-selection").boundingBox())!;
  const from = { x: outline.x + 1, y: outline.y + outline.height / 4 };
  await page.mouse.move(from.x, from.y);
  await expect(page.getByTestId("spreadsheet-grid")).toHaveCSS("cursor", "move");
  await page.mouse.down();
  const shift = (await centre(page, "P2")).x - (await centre(page, "B2")).x;
  await page.keyboard.down("Alt");
  await page.mouse.move(from.x + shift, from.y, { steps: 10 });
  await expect(page.getByTestId("sheet-move-preview")).toBeVisible();
  await expect(page.getByTestId("sheet-move-tooltip")).toHaveText("P2:Q3");
  await page.mouse.up();
  await page.keyboard.up("Alt");
  await expect(cell(page, "P2")).toHaveText("10");
  await expect(cell(page, "Q3")).toHaveText("30");
  await expect(cell(page, "B2")).toHaveText("10");
  await clickCell(page, "P3");
  await expect(formulaBar(page)).toHaveText("=P2*2");

  // Undo three times: the copy, the fill down and the double-click fill.
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "P2")).toHaveText("");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "N6")).toHaveText("");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "M21")).toHaveText("");
  await expect(cell(page, "J8")).toHaveText("Mon");

  // Download and reopen: the cells as left.
  const path = await download(page, "fill.xlsx");
  await open(page, path);
  await expect(cell(page, "E5")).toHaveText("10");
  await clickCell(page, "E6");
  await expect(formulaBar(page)).toHaveText("=E5*2");
  await expect(cell(page, "H10")).toHaveText("5");
  await expect(cell(page, "A10")).toHaveText("");
  for (let row = 3; row <= 8; row += 1) {
    await expect(cell(page, `J${row}`)).toHaveText("Mon");
  }
  await expect(cell(page, "M21")).toHaveText("");
  await expect(cell(page, "N6")).toHaveText("");
  await expect(cell(page, "P2")).toHaveText("");
  if (captured) {
    await clickCell(page, "C40");
    await expect(formulaBar(page)).toHaveText("1234.5");
    await clickCell(page, "A47");
    await expect(formulaBar(page)).toHaveText("1234.5");
  }
  await sheetTab(page, "Sheet2").click();
  await clickCell(page, "A1");
  await expect(formulaBar(page)).toHaveText("=SUM(Sheet1!H10:H12)");
});
