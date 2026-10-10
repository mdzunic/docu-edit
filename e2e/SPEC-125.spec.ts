import { test, expect, type Locator, type Page } from "@playwright/test";
import { existsSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzip } from "../src/lib/core/zip";

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
//
// The tests after the chain drive, the same way, what the chain does not
// reach, one concern each: the refusals — over 1,000,000 cells to copy
// (spec116-wide.ods, AC-7), past the grid's edge (AC-27), part of an array
// (spec119-formulas.xlsx, AC-29), a protected sheet (spec122-shapes.xlsx,
// AC-30); another workbook open in a second tab (spec116-1904.xlsx) —
// formulas naming a sheet it lacks paste as values and dates convert
// between the 1900 and 1904 systems both ways (AC-37, AC-38), styles intern
// there and keep their ids at home, and a cut pastes there as a copy, as it
// does at home once an edit has cancelled it (AC-36, AC-45); the menu's
// Paste recognising the view's own copy by its text (AC-21); LibreOffice's
// HTML, and plain text over a styled cell (AC-57, AC-64); date and counter
// series (AC-73, AC-76); Fill down/right of one row and one column (AC-88);
// the row refit after pastes and a fill (AC-96); and a cut moving once —
// #REF! for overwritten cells, moved formulas unshifted, merges and arrays
// moving (AC-40, AC-42, AC-43, AC-44).

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const EXCEL_CAPTURE = join(fixtures, "spec125-clipboard", "excel.html");
const FILL_FIXTURE = join(fixtures, "spec125-fill.xlsx");

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

/** spec125-fill.xlsx in `page` and, in a second tab of the same browser
 * (one clipboard), spec116-1904.xlsx — another workbook, on the 1904 date
 * system, whose one sheet is "Dates". */
async function openTwoWorkbooks(page: Page): Promise<Page> {
  await open(page, FILL_FIXTURE);
  const other = await page.context().newPage();
  await open(other, join(fixtures, "spec116-1904.xlsx"));
  await page.bringToFront();
  return other;
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

function toast(page: Page, message: string): Locator {
  return page.getByTestId("toast").filter({ hasText: message });
}

async function select(page: Page, from: string, to: string): Promise<void> {
  await clickCell(page, from);
  await clickCell(page, to, ["Shift"]);
}

/** Puts flavours on the clipboard, as another application's copy does. */
async function writeClipboard(page: Page, flavours: Record<string, string>): Promise<void> {
  await page.evaluate(async (items) => {
    const blobs: Record<string, Blob> = {};
    for (const [type, text] of Object.entries(items)) {
      blobs[type] = new Blob([text], { type });
    }
    await navigator.clipboard.write([new ClipboardItem(blobs)]);
  }, flavours);
}

/** Drags the selection's fill handle to `to`, then dismisses the chip. */
async function dragFill(page: Page, to: string): Promise<void> {
  const grip = (await page.getByTestId("sheet-fill-handle").boundingBox())!;
  await page.mouse.move(grip.x + grip.width / 2, grip.y + grip.height / 2);
  await page.mouse.down();
  const target = await centre(page, to);
  await page.mouse.move(target.x, target.y, { steps: 8 });
  await page.mouse.up();
  await expect(page.getByTestId("sheet-fill-options")).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("sheet-fill-options")).toHaveCount(0);
}

/** Sets columns `from`:`to` to `chars` characters through the header menu's
 * "Column width…". */
async function widenColumns(page: Page, from: string, to: string, chars: string): Promise<void> {
  await colHeader(page, from).click();
  await colHeader(page, to).click({ modifiers: ["Shift"] });
  await colHeader(page, to).click({ button: "right" });
  await page.getByTestId("sheet-size-band").click();
  await page.getByTestId("sheet-column-width-input").fill(chars);
  await page.getByTestId("sheet-size-ok").click();
  await expect(page.getByTestId("sheet-column-width-dialog")).toHaveCount(0);
}

/** A date cell's value, as the formula bar shows it (the fixture's columns
 * are too narrow for some dates, which then show as "########"). */
async function expectDate(page: Page, ref: string, date: string): Promise<void> {
  await clickCell(page, ref);
  await expect(formulaBar(page)).toHaveText(date);
}

async function rowHeight(page: Page, row: number): Promise<number> {
  return (await rowHeader(page, row).boundingBox())!.height;
}

/** A cell's `s` attribute in a sheet part. */
function styleAttribute(xml: string, ref: string): string | undefined {
  return new RegExp(`<c r="${ref}"(?: [^>]*?)? s="(\\d+)"`).exec(xml)?.[1];
}

test("copy, paste, cut, fill, move and undo on spec125-fill.xlsx, then download and reopen (AC-118, AC-119)", async ({ page }) => {
  // A journey with a download and a reopen: the allowance SPEC-077's have.
  test.setTimeout(120_000);
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
    // Rows the grid shows at this viewport (it ends at row 35); column C wide
    // enough for "$1,234.50", which its default width can show as "########".
    await widenColumns(page, "C", "C", "14");
    await clickCell(page, "A25");
    await page.keyboard.press("ControlOrMeta+v");
    await clickCell(page, "C25");
    await expect(formulaBar(page)).toHaveText("1234.5");
    await expect(cell(page, "C25")).toHaveText(/\$1,234\.50/);
    await clickCell(page, "A27");
    await page.keyboard.press("ControlOrMeta+Shift+v");
    const dialog = page.getByTestId("sheet-paste-special-dialog");
    await expect(dialog).toBeVisible();
    await page.getByTestId("sheet-paste-special-values").check();
    await page.getByTestId("sheet-paste-special-transpose").check();
    await page.getByTestId("sheet-paste-special-ok").click();
    await expect(dialog).toHaveCount(0);
    await clickCell(page, "A29");
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
    await clickCell(page, "C25");
    await expect(formulaBar(page)).toHaveText("1234.5");
    await clickCell(page, "A29");
    await expect(formulaBar(page)).toHaveText("1234.5");
  }
  await sheetTab(page, "Sheet2").click();
  await clickCell(page, "A1");
  await expect(formulaBar(page)).toHaveText("=SUM(Sheet1!H10:H12)");
});

test("Copy refuses more than 1,000,000 cells (AC-7)", async ({ page }) => {
  // spec116-wide.ods's one value sits at row 900,000, so columns A:B copy
  // 1,800,000 cells: refused, and no ants.
  await open(page, join(fixtures, "spec116-wide.ods"));
  await colHeader(page, "A").click();
  await colHeader(page, "B").click({ modifiers: ["Shift"] });
  await page.keyboard.press("ControlOrMeta+c");
  await expect(toast(page, "That's too many cells to copy — select fewer than 1,000,000 cells")).toBeVisible();
  await expect(page.getByTestId("sheet-copy-marquee")).toHaveCount(0);
});

for (const [edge, arrow, landing] of [
  ["last column", "ArrowRight", "XFD30"],
  ["last row", "ArrowDown", "A1048576"],
] as const) {
  test(`a paste reaching past the grid's ${edge} is refused (AC-27)`, async ({ page }) => {
    await open(page, FILL_FIXTURE);
    await select(page, "B2", "C3");
    await page.keyboard.press("ControlOrMeta+c");
    // From A30, an empty row and column: Ctrl/Cmd+arrow runs to the edge.
    await clickCell(page, "A30");
    await page.keyboard.press(`ControlOrMeta+${arrow}`);
    await expect(page.getByTestId("name-box")).toHaveValue(landing);
    await page.keyboard.press("ControlOrMeta+v");
    await expect(toast(page, "The paste area goes past the edge of the sheet")).toBeVisible();
    await expect(page.getByTestId("name-box")).toHaveValue(landing);
  });
}

test("a paste onto part of an array is refused; onto all of it, it lands (AC-29)", async ({ page }) => {
  // spec119-formulas.xlsx's "Calc" holds the array E2:E4 ({=A2:A4*2}).
  await open(page, join(fixtures, "spec119-formulas.xlsx"));
  await writeClipboard(page, { "text/plain": "7" });
  await clickCell(page, "E3");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(toast(page, "You can't change part of an array")).toBeVisible();
  await expect(cell(page, "E3")).toHaveText("2");
  await select(page, "E2", "E4");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "E2")).toHaveText("7");
  await expect(cell(page, "E4")).toHaveText("7");
});

test("a paste onto a protected sheet's locked cell is refused with SPEC-122's message (AC-30)", async ({ page }) => {
  // spec122-shapes.xlsx's "Locked" is protected: A1 is locked, B2 is its one
  // unlocked cell.
  await open(page, join(fixtures, "spec122-shapes.xlsx"));
  await sheetTab(page, "Locked").click();
  const before = await cell(page, "A1").textContent();
  await writeClipboard(page, { "text/plain": "7" });
  await clickCell(page, "A1");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(toast(page, "This cell is on a protected sheet")).toBeVisible();
  await expect(cell(page, "A1")).toHaveText(before ?? "");
  await clickCell(page, "B2");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "B2")).toHaveText("7");
});

test("another workbook: formulas naming a sheet it lacks paste as their values; dates convert between the 1900 and 1904 systems (AC-37, AC-38)", async ({ page }) => {
  // Two workbooks to open: the allowance SPEC-077's journeys have.
  test.setTimeout(120_000);
  const other = await openTwoWorkbooks(page);

  // Sheet2's two formulas both read Sheet1, which spec116-1904.xlsx lacks:
  // they paste there as their cached values, with the toast.
  await sheetTab(page, "Sheet2").click();
  await select(page, "A1", "A2");
  await page.keyboard.press("ControlOrMeta+c");
  await other.bringToFront();
  await clickCell(other, "D1");
  await other.keyboard.press("ControlOrMeta+v");
  await expect(toast(other, "2 formulas referring to other sheets were pasted as values")).toBeVisible();
  await expect(cell(other, "D1")).toHaveText("18");
  await expect(cell(other, "D2")).toHaveText("30");
  await expect(formulaBar(other)).toHaveText("18");

  // A 1904 date (B1, 2026-10-08) pasted into the 1900 workbook is the same
  // date — the formula bar shows it as typed (the columns are too narrow to
  // show every date in full)…
  await clickCell(other, "B1");
  await expect(formulaBar(other)).toHaveText("10/8/2026");
  await other.keyboard.press("ControlOrMeta+c");
  await page.bringToFront();
  await sheetTab(page, "Sheet1").click();
  await clickCell(page, "P20");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(formulaBar(page)).toHaveText("10/8/2026");
  // …and a 1900 date (Series!E2, 2026-01-31) pasted into the 1904 one too.
  await sheetTab(page, "Series").click();
  await clickCell(page, "E2");
  await expect(formulaBar(page)).toHaveText("1/31/2026");
  await page.keyboard.press("ControlOrMeta+c");
  await other.bringToFront();
  await clickCell(other, "D5");
  await other.keyboard.press("ControlOrMeta+v");
  await expect(formulaBar(other)).toHaveText("1/31/2026");
});

test("another workbook: styles intern there and keep their ids at home; a cut pastes there as a copy, as it does at home once an edit cancelled it (AC-36, AC-45)", async ({ page }) => {
  // Two workbooks to open and two downloads: the allowance SPEC-077's
  // journeys have.
  test.setTimeout(120_000);
  const other = await openTwoWorkbooks(page);

  // Sheet1's bold A1 pasted at home keeps its style id; pasted there, it is
  // interned as a bold style.
  await clickCell(page, "A1");
  await page.keyboard.press("ControlOrMeta+c");
  await clickCell(page, "A28");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "A28")).toHaveText("Amounts");
  await expect(page.getByTestId("sheet-bold")).toHaveAttribute("aria-pressed", "true");
  await other.bringToFront();
  await clickCell(other, "E1");
  await other.keyboard.press("ControlOrMeta+v");
  await expect(cell(other, "E1")).toHaveText("Amounts");
  await expect(other.getByTestId("sheet-bold")).toHaveAttribute("aria-pressed", "true");

  // A cut of A10:A12 pasted into the other workbook is a copy: the cells
  // land there, and nothing moves at home.
  await page.bringToFront();
  await select(page, "A10", "A12");
  await page.keyboard.press("ControlOrMeta+x");
  await expect(page.getByTestId("sheet-copy-marquee")).toHaveAttribute("data-mode", "cut");
  await other.bringToFront();
  await clickCell(other, "F1");
  await other.keyboard.press("ControlOrMeta+v");
  await expect(cell(other, "F1")).toHaveText("5");
  await expect(cell(other, "F3")).toHaveText("7");
  await page.bringToFront();
  await expect(cell(page, "A10")).toHaveText("5");
  await expect(cell(page, "A12")).toHaveText("7");

  // At home an edit cancels the pending cut (the ants go), so its paste is
  // a copy too: A10:A12 stay, and Sheet2's SUM still reads them.
  await clickCell(page, "P1");
  await page.keyboard.type("x");
  await page.keyboard.press("Enter");
  await expect(page.getByTestId("sheet-copy-marquee")).toHaveCount(0);
  await clickCell(page, "H10");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "H10")).toHaveText("5");
  await expect(cell(page, "H12")).toHaveText("7");
  await expect(cell(page, "A10")).toHaveText("5");
  await expect(cell(page, "A12")).toHaveText("7");
  await sheetTab(page, "Sheet2").click();
  await clickCell(page, "A1");
  await expect(formulaBar(page)).toHaveText("=SUM(Sheet1!A10:A12)");

  // Saved: A28 has A1's style index and no style was added at home; the
  // other workbook's E1 is still bold after a reopen.
  const home = await unzip(new Uint8Array(readFileSync(await download(page, "home.xlsx"))));
  const sheet1 = new TextDecoder().decode(home.get("xl/worksheets/sheet1.xml")!);
  expect(styleAttribute(sheet1, "A28")).toBeDefined();
  expect(styleAttribute(sheet1, "A28")).toBe(styleAttribute(sheet1, "A1"));
  const xfCount = (bytes: Uint8Array): string | undefined => /<cellXfs count="(\d+)"/.exec(new TextDecoder().decode(bytes))?.[1];
  const fixture = await unzip(new Uint8Array(readFileSync(FILL_FIXTURE)));
  expect(xfCount(home.get("xl/styles.xml")!)).toBe(xfCount(fixture.get("xl/styles.xml")!));
  await other.bringToFront();
  await open(other, await download(other, "other.xlsx"));
  await clickCell(other, "E1");
  await expect(other.getByTestId("sheet-bold")).toHaveAttribute("aria-pressed", "true");
  await expect(cell(other, "F3")).toHaveText("7");
});

test("the menu's Paste takes the view's own copy by its text when the HTML lacks the payload (AC-21)", async ({ page }) => {
  await open(page, FILL_FIXTURE);
  await select(page, "B2", "C3");
  await page.keyboard.press("ControlOrMeta+c");
  const tsv = await page.evaluate(() => navigator.clipboard.readText());
  expect(tsv).toBe("10\t20\r\n20\t30");

  // HTML without the attribute, and the copy's text: the copy's payload
  // pastes — E6 holds the shifted formula, not the number 20.
  await writeClipboard(page, {
    "text/html": "<table><tr><td>10</td><td>20</td></tr><tr><td>20</td><td>30</td></tr></table>",
    "text/plain": tsv,
  });
  await clickCell(page, "E5");
  const e5 = await centre(page, "E5");
  await page.mouse.click(e5.x, e5.y, { button: "right" });
  await page.getByTestId("sheet-cell-menu-paste").click();
  await expect(cell(page, "F6")).toHaveText("30");
  await clickCell(page, "E6");
  await expect(formulaBar(page)).toHaveText("=E5*2");

  // No HTML at all, the same text: the same rule.
  await writeClipboard(page, { "text/plain": tsv });
  await clickCell(page, "E20");
  const e20 = await centre(page, "E20");
  await page.mouse.click(e20.x, e20.y, { button: "right" });
  await page.getByTestId("sheet-cell-menu-paste").click();
  await expect(cell(page, "F21")).toHaveText("30");
  await clickCell(page, "F21");
  await expect(formulaBar(page)).toHaveText("=E20+F20");
});

test("LibreOffice's HTML imports through sdval and sdnum; plain text keeps the destination's style (AC-57, AC-64)", async ({ page }) => {
  await open(page, FILL_FIXTURE);
  // Columns A:I wide enough for every format to show in full.
  await widenColumns(page, "A", "I", "14");
  await writeClipboard(page, {
    "text/html":
      `<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><html><head>` +
      `<meta http-equiv="content-type" content="text/html; charset=utf-8"/><meta name="generator" content="LibreOffice 24.8.4.2 (MacOSX)"/>` +
      `<style type="text/css">body,div,table,thead,tbody,tfoot,tr,th,td,p { font-family:"Liberation Sans"; font-size:x-small }</style></head>` +
      `<body><table cellspacing="0" border="0"><colgroup span="9" width="85"></colgroup><tr>` +
      `<td height="17" align="left"><b>Widget</b></td>` +
      `<td align="right" sdval="1234567" sdnum="1033;0;#,##0">1,234,567</td>` +
      `<td align="right" sdval="1234.5" sdnum="1033;0;[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00">$1,234.50</td>` +
      `<td align="right" sdval="0.25" sdnum="1033;0;0%">25%</td>` +
      `<td align="right" sdval="46096" sdnum="1033;0;MM/DD/YY">03/15/26</td>` +
      `<td align="right" sdval="1" sdnum="1033;0;BOOLEAN">TRUE</td>` +
      `<td align="right" sdval="1235801.5" sdnum="1033;0;General">1235801.5</td>` +
      `<td colspan=2 align="left"><font color="#C9211E"><i><u>Merged</u></i></font></td>` +
      `</tr></table></body></html>`,
    "text/plain": "Widget\t1,234,567\t$1,234.50\t25%\t03/15/26\tTRUE\t1235801.5\tMerged",
  });
  await clickCell(page, "A24");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "A24")).toHaveText("Widget");
  await expect(page.getByTestId("sheet-bold")).toHaveAttribute("aria-pressed", "true");
  await expect(cell(page, "B24")).toHaveText("1,234,567");
  await clickCell(page, "B24");
  await expect(formulaBar(page)).toHaveText("1234567");
  await expect(cell(page, "C24")).toHaveText(/\$1,234\.50/);
  await clickCell(page, "C24");
  await expect(formulaBar(page)).toHaveText("1234.5");
  await expect(cell(page, "D24")).toHaveText("25%");
  await expect(cell(page, "E24")).toHaveText("03/15/26");
  await expect(cell(page, "F24")).toHaveText("TRUE");
  await expect(cell(page, "G24")).toHaveText("1235801.5");
  await expect(cell(page, "H24")).toHaveText("Merged");
  await expect(cell(page, "H24")).toHaveAttribute("aria-colspan", "2");
  await clickCell(page, "H24");
  await expect(page.getByTestId("sheet-italic")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("sheet-underline")).toHaveAttribute("aria-pressed", "true");

  // A bold A26 takes "1,234.50" as the number 1234.5 in the format the
  // text implies, and stays bold.
  await clickCell(page, "A26");
  await page.keyboard.press("ControlOrMeta+b");
  await expect(page.getByTestId("sheet-bold")).toHaveAttribute("aria-pressed", "true");
  await writeClipboard(page, { "text/plain": "1,234.50" });
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "A26")).toHaveText("1,234.50");
  await expect(formulaBar(page)).toHaveText("1234.5");
  await expect(page.getByTestId("sheet-bold")).toHaveAttribute("aria-pressed", "true");
});

test("the fill handle continues dates by day, month-end, month (EDATE), year and week (AC-73)", async ({ page }) => {
  await open(page, FILL_FIXTURE);
  await sheetTab(page, "Series").click();
  await clickCell(page, "E2");
  await dragFill(page, "E4");
  await expectDate(page, "E3", "2/1/2026");
  await expectDate(page, "E4", "2/2/2026");
  await select(page, "F2", "F3");
  await dragFill(page, "F5");
  await expectDate(page, "F4", "3/31/2026");
  await expectDate(page, "F5", "4/30/2026");
  await select(page, "G2", "G3");
  await dragFill(page, "G5");
  await expectDate(page, "G4", "3/15/2026");
  await expectDate(page, "G5", "4/15/2026");
  await select(page, "H2", "H3");
  await dragFill(page, "H4");
  await expectDate(page, "H4", "3/1/2026");
  await select(page, "I2", "I3");
  await dragFill(page, "I4");
  await expectDate(page, "I4", "1/15/2026");
});

test("the fill handle continues a trailing integer (its zero padding kept, a step of two) and ordinals (AC-76)", async ({ page }) => {
  await open(page, FILL_FIXTURE);
  await sheetTab(page, "Series").click();
  await clickCell(page, "O2");
  await dragFill(page, "O4");
  await expect(cell(page, "O3")).toHaveText("Item 10");
  await expect(cell(page, "O4")).toHaveText("Item 11");
  await clickCell(page, "P2");
  await dragFill(page, "P3");
  await expect(cell(page, "P3")).toHaveText("Item 010");
  await select(page, "Q2", "Q3");
  await dragFill(page, "Q5");
  await expect(cell(page, "Q4")).toHaveText("Item 5");
  await expect(cell(page, "Q5")).toHaveText("Item 7");
  await clickCell(page, "R2");
  await dragFill(page, "R5");
  await expect(cell(page, "R3")).toHaveText("2nd");
  await expect(cell(page, "R4")).toHaveText("3rd");
  await expect(cell(page, "R5")).toHaveText("4th");
});

test("Fill down of one row copies the row above; Fill right of one column copies the column to its left (AC-88)", async ({ page }) => {
  await open(page, FILL_FIXTURE);
  // One row, M3:N3: Fill down copies the row above (M2's formula shifted).
  await select(page, "M3", "N3");
  await page.keyboard.press("ControlOrMeta+d");
  await expect(cell(page, "M3")).toHaveText("4");
  await expect(cell(page, "N3")).toHaveText("Total");
  await clickCell(page, "M3");
  await expect(formulaBar(page)).toHaveText("=L3*2");
  // One column, K2: Fill right copies the column to its left.
  await clickCell(page, "K2");
  await page.keyboard.press("ControlOrMeta+r");
  await expect(cell(page, "K2")).toHaveText("Mon");
});

test("pastes and fills refit the rows they write (AC-96)", async ({ page }) => {
  await open(page, FILL_FIXTURE);
  const normal = await rowHeight(page, 30);
  // A 24 pt cell pasted from a web table makes its row taller.
  await writeClipboard(page, { "text/html": `<table><tr><td style="font-size:24pt">Big</td></tr></table>`, "text/plain": "Big" });
  await clickCell(page, "A30");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "A30")).toHaveText("Big");
  await expect.poll(() => rowHeight(page, 30)).toBeGreaterThan(normal + 10);
  const tall = await rowHeight(page, 30);
  // Fill down brings the style, and the refit, to row 31.
  await select(page, "A30", "A31");
  await page.keyboard.press("ControlOrMeta+d");
  await expect(cell(page, "A31")).toHaveText("Big");
  await expect.poll(() => rowHeight(page, 31)).toBe(tall);
  // The view's own copy pasted at A33: the same; undo puts row 33 back.
  await clickCell(page, "A30");
  await page.keyboard.press("ControlOrMeta+c");
  await clickCell(page, "A33");
  await page.keyboard.press("ControlOrMeta+v");
  await expect.poll(() => rowHeight(page, 33)).toBe(tall);
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "A33")).toHaveText("");
  await expect.poll(() => rowHeight(page, 33)).toBe(normal);
});

test("a cut moves once: references to overwritten cells become #REF!, moved formulas keep their references, merges and arrays move (AC-40, AC-42, AC-43, AC-44)", async ({ page }) => {
  await open(page, FILL_FIXTURE);
  const ants = page.getByTestId("sheet-copy-marquee");
  // B3:C3 (=B2*2, =B2+C2) moved to E20: their references, outside what
  // moved, stay as they were — not shifted as a copy would be.
  await select(page, "B3", "C3");
  await page.keyboard.press("ControlOrMeta+x");
  await clickCell(page, "E20");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "B3")).toHaveText("");
  await clickCell(page, "E20");
  await expect(formulaBar(page)).toHaveText("=B2*2");
  await clickCell(page, "F20");
  await expect(formulaBar(page)).toHaveText("=B2+C2");

  // A10:A12 moved onto B1:B3 overwrites B2, which Sheet2!A2 read: #REF!.
  await select(page, "A10", "A12");
  await page.keyboard.press("ControlOrMeta+x");
  await clickCell(page, "B1");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "B2")).toHaveText("6");
  await expect(cell(page, "A10")).toHaveText("");
  await expect(ants).toHaveCount(0);
  // The cut pasted once: pasting again lands a copy and moves nothing.
  await clickCell(page, "H20");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "H21")).toHaveText("6");
  await expect(cell(page, "B2")).toHaveText("6");
  await sheetTab(page, "Sheet2").click();
  await clickCell(page, "A2");
  await expect(formulaBar(page)).toHaveText("=Sheet1!#REF!*3");
  await clickCell(page, "A1");
  await expect(formulaBar(page)).toHaveText("=SUM(Sheet1!B1:B3)");
  await sheetTab(page, "Sheet1").click();

  // The merge E1:F1 moves to E30 whole.
  await clickCell(page, "E1");
  await page.keyboard.press("ControlOrMeta+x");
  await clickCell(page, "E30");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "E30")).toHaveText("Merged");
  await expect(cell(page, "E30")).toHaveAttribute("aria-colspan", "2");
  await expect(cell(page, "E1")).toHaveText("");
  await expect(cell(page, "E1")).not.toHaveAttribute("aria-colspan", "2");

  // spec119-formulas.xlsx's array E2:E4 ({=A2:A4*2}) moves to E20 whole.
  await open(page, join(fixtures, "spec119-formulas.xlsx"));
  await select(page, "E2", "E4");
  await page.keyboard.press("ControlOrMeta+x");
  await clickCell(page, "E20");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "E21")).toHaveText("2");
  await clickCell(page, "E21");
  await expect(formulaBar(page)).toHaveText("{=A2:A4*2}");
  await expect(cell(page, "E2")).toHaveText("");
});
