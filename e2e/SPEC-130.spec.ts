import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-130 — data validation and in-cell dropdowns (AC-50), with the real
// pointer and keyboard: open spec130-validation.xlsx → B2's input message
// reads "Priority" / "Pick one" → Alt+Down lists Low, Medium, High →
// ArrowDown, ArrowDown, Enter picks Medium → "Urgent" is stopped → Retry
// leaves "Urgent" selected in the editor → "High" is accepted → 42 in B3
// (whole 1–10, warning) asks "Continue?" → No → Escape leaves B3 as it was →
// B8 (showDropDown="1") has no arrow → D2 gets a List rule "Yes,No" through
// Data ▸ Data validation… and shows the arrow → download → reopen → D2 still
// offers Yes, No. Then spec130-validation.ods: C500, inside a rule carried by
// a repeated run of empty cells, shows its input message. And B10, whose
// Source is the defined name Colours, lists the named cells (AC-5) — in a
// window so short that B10 is cut by the grid's bottom edge too, where its
// list opens inside the grid (AC-23). And an input message on a cell at the
// grid's right edge flips to the cell's left at its natural width (AC-20).
//
// QA's additions: in a 900x600 window the message stays beside its cell — to
// the left of and below a right-edge cell, to the left of and above the
// corner cell — and re-places on a resize (AC-20); the warning and
// information alerts commit, Ctrl/Cmd+Enter's first failing target decides
// and a rule without its error alert commits silently (AC-9, AC-11, AC-14,
// AC-15); paste, fill and undo never validate, a copy leaves the rules where
// they were and a cut takes its cell's rule along (AC-16–AC-18, AC-47); an
// autofilter header keeps Alt+Down for its filter (AC-22); inserted rows and
// a renamed sheet carry the rules' ranges and sources (AC-45, AC-46); and an
// untouched workbook downloads byte for byte as it was opened (AC-40).

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 }, permissions: ["clipboard-read", "clipboard-write"] });

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

/** Clicks the centre of a cell through its column and row headers. */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
}

/** Moves the active cell down its column to `row` with the keyboard. */
async function keyDownTo(page: Page, column: string, row: number): Promise<void> {
  const nameBox = page.getByTestId("name-box");
  for (let step = 0; step < 200; step += 1) {
    const current = Number(/\d+$/.exec(await nameBox.inputValue())![0]);
    if (current === row) {
      break;
    }
    await page.keyboard.press(current < row ? (row - current > 40 ? "PageDown" : "ArrowDown") : "ArrowUp");
  }
  await expect(nameBox).toHaveValue(`${column}${row}`);
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
  const path = join(tmpdir(), `spec130-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  return path;
}

/** Whether an element is what a pointer at its centre would hit — shown,
 * not clipped away or under other chrome. */
const hitAtCentre = (target: Locator): Promise<boolean> =>
  target.evaluate((el) => {
    const box = el.getBoundingClientRect();
    const top = el.ownerDocument.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2);
    return top !== null && el.contains(top);
  });

/** The selected text of the in-cell editor. */
const editorSelection = (page: Page): Promise<string> =>
  page.getByTestId("cell-editor").evaluate((host) => {
    const selection = host.ownerDocument.getSelection();
    return selection && host.contains(selection.anchorNode) ? selection.toString() : "";
  });

test("the input message, the dropdown, the alerts, the dialog and a round trip (AC-50)", async ({ page }) => {
  await open(page, join(fixtures, "spec130-validation.xlsx"));

  // B2's input message, referenced by the active cell.
  await clickCell(page, "B2");
  const message = page.getByTestId("validation-input-message");
  await expect(message).toBeVisible();
  await expect(message).toHaveAttribute("role", "tooltip");
  await expect(message).toContainText("Priority");
  await expect(message).toContainText("Pick one");
  const messageId = await message.getAttribute("id");
  await expect(cell(page, "B2")).toHaveAttribute("aria-describedby", messageId!);

  // The arrow sits against the cell; Alt+Down opens the list.
  const arrow = page.getByTestId("validation-dropdown");
  await expect(arrow).toBeVisible();
  await expect(arrow).toHaveAttribute("aria-label", "Choose from list");
  await expect(arrow).toHaveAttribute("aria-expanded", "false");
  await page.keyboard.press("Alt+ArrowDown");
  const list = page.getByTestId("validation-list");
  await expect(list).toBeVisible();
  await expect(list).toHaveAttribute("role", "listbox");
  await expect(list.getByTestId("validation-option")).toHaveText(["Low", "Medium", "High"]);
  await expect(arrow).toHaveAttribute("aria-expanded", "true");
  await expect(message).toBeHidden();

  // B2 is empty: nothing is selected; the first ArrowDown selects Low, the second Medium.
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowDown");
  await expect(list.getByTestId("validation-option").nth(1)).toHaveAttribute("aria-selected", "true");
  await page.keyboard.press("Enter");
  await expect(list).toBeHidden();
  await expect(cell(page, "B2")).toHaveText("Medium");

  // "Urgent" is not in the list: the stop alert.
  await page.keyboard.type("Urgent");
  await page.keyboard.press("Enter");
  const alert = page.getByTestId("validation-alert");
  await expect(alert).toBeVisible();
  await expect(alert).toHaveAttribute("role", "alertdialog");
  await expect(alert.getByTestId("validation-alert-yes")).toHaveCount(0);
  await alert.getByTestId("validation-alert-retry").click();
  await expect(alert).toBeHidden();

  // Retry: back in the editor with the entered text selected.
  const editor = page.getByTestId("cell-editor");
  await expect(editor).toHaveText("Urgent");
  await expect.poll(() => editorSelection(page)).toBe("Urgent");
  await page.keyboard.type("High");
  await page.keyboard.press("Enter");
  await expect(alert).toBeHidden();
  await expect(cell(page, "B2")).toHaveText("High");

  // B3 (5) is a whole number between 1 and 10 with a warning: 42 asks to continue.
  await clickCell(page, "B3");
  await page.keyboard.type("42");
  await page.keyboard.press("Enter");
  await expect(alert).toBeVisible();
  await expect(alert).toContainText("Continue?");
  await alert.getByTestId("validation-alert-no").click();
  await expect(alert).toBeHidden();
  await expect(editor).toHaveText("42");
  await page.keyboard.press("Escape");
  await expect(cell(page, "B3")).toHaveText("5");

  // B8's file says showDropDown="1": no arrow.
  await clickCell(page, "B8");
  await expect(page.getByTestId("validation-dropdown")).toHaveCount(0);

  // D2: Data ▸ Data validation… → List, Source "Yes,No" → OK → the arrow shows.
  await clickCell(page, "D2");
  await expect(page.getByTestId("validation-dropdown")).toHaveCount(0);
  await page.getByTestId("menu-title-data").click();
  await page.getByTestId("sheet-menu-data-validation").click();
  const dialog = page.getByTestId("validation-dialog");
  await expect(dialog).toBeVisible();
  await dialog.getByTestId("validation-allow").selectOption({ label: "List" });
  await expect(dialog.getByLabel("Source")).toBeVisible();
  await dialog.getByTestId("validation-value-1").fill("Yes,No");
  await dialog.getByTestId("validation-ok").click();
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId("validation-dropdown")).toBeVisible();

  // Download → reopen: D2 still offers Yes, No.
  const saved = await download(page, "spec130-validation.xlsx");
  await open(page, saved);
  await clickCell(page, "D2");
  await page.getByTestId("validation-dropdown").click();
  await expect(page.getByTestId("validation-list").getByTestId("validation-option")).toHaveText(["Yes", "No"]);
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("validation-list")).toBeHidden();
  // The file's own rules came back too.
  await clickCell(page, "B2");
  await expect(page.getByTestId("validation-input-message")).toContainText("Priority");
});

test("spec130-validation.ods: a rule carried by a repeated run of empty cells shows its input message at C500 (AC-50)", async ({ page }) => {
  await open(page, join(fixtures, "spec130-validation.ods"));
  await clickCell(page, "C2");
  await keyDownTo(page, "C", 500);
  const message = page.getByTestId("validation-input-message");
  await expect(message).toBeVisible();
  // Its title, and its two-paragraph text on two lines.
  await expect(message).toContainText("Notes");
  expect(await message.innerText()).toContain("Free text\nKeep it short");
  await expect(cell(page, "C500")).toHaveAttribute("aria-describedby", (await message.getAttribute("id"))!);
});

test("B10's defined-name Source lists what Colours names, and its list opens inside a short grid (AC-5, AC-23)", async ({ page }) => {
  await open(page, join(fixtures, "spec130-validation.xlsx"));
  await clickCell(page, "B10");
  await page.keyboard.press("Alt+ArrowDown");
  const list = page.getByTestId("validation-list");
  const options = list.getByTestId("validation-option");
  await expect(options).toHaveText(["Red", "Green", "Blue"]);
  await page.keyboard.press("Escape");
  await expect(list).toBeHidden();

  // "Purple" is not one of them: the stop alert; Cancel leaves B10 empty.
  await page.keyboard.type("Purple");
  await page.keyboard.press("Enter");
  const alert = page.getByTestId("validation-alert");
  await expect(alert).toBeVisible();
  await alert.getByTestId("validation-alert-cancel").click();
  await expect(alert).toBeHidden();
  await expect(cell(page, "B10")).toHaveText("");

  // A window so short that B10, the grid's last row, is cut by its bottom edge.
  await page.setViewportSize({ width: 900, height: 260 });
  await page.keyboard.press("ArrowUp");
  await page.keyboard.press("ArrowDown");
  await expect(page.getByTestId("name-box")).toHaveValue("B10");
  await page.getByTestId("validation-dropdown").click();
  await expect(options).toHaveCount(3);
  await expect.poll(() => hitAtCentre(options.first())).toBe(true);
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowDown");
  await expect(options.nth(1)).toHaveAttribute("aria-selected", "true");
  await expect.poll(() => hitAtCentre(options.nth(1))).toBe(true);
  await page.keyboard.press("Enter");
  await expect(list).toBeHidden();
  await expect(cell(page, "B10")).toHaveText("Green");
});

test("an input message at the grid's right edge flips to the cell's left at its natural width (AC-20)", async ({ page }) => {
  await open(page, join(fixtures, "spec130-validation.xlsx"));

  // Row 3 of the last column whose cell still has less room on its right
  // than the message is wide.
  const grid = (await page.getByTestId("spreadsheet-grid").boundingBox())!;
  const headers = page.getByTestId("sheet-col-header");
  let letters = "";
  for (let index = (await headers.count()) - 1; index >= 0 && letters === ""; index -= 1) {
    const box = await headers.nth(index).boundingBox();
    if (box && box.x + box.width <= grid.x + grid.width - 20) {
      letters = (await headers.nth(index).innerText()).trim();
    }
  }
  expect(letters).toMatch(/^[A-Z]+$/);
  const ref = `${letters}3`;
  await clickCell(page, ref);

  // An input message on it through Data ▸ Data validation….
  await page.getByTestId("menu-title-data").click();
  await page.getByTestId("sheet-menu-data-validation").click();
  const dialog = page.getByTestId("validation-dialog");
  await expect(dialog).toBeVisible();
  await dialog.getByTestId("validation-tab-input").click();
  await dialog.getByTestId("validation-show-input").check();
  await dialog.getByTestId("validation-input-title").fill("Score");
  await dialog.getByTestId("validation-input-text").fill("Enter a whole number between 1 and 10");
  await dialog.getByTestId("validation-ok").click();
  await expect(dialog).toBeHidden();

  // It sits to the cell's left, still below it, whole inside the grid — not
  // squeezed into the strip left at the right edge, words broken mid-word.
  const message = page.getByTestId("validation-input-message");
  await expect(message).toBeVisible();
  await expect(message).toContainText("Score");
  const cellBox = (await cell(page, ref).boundingBox())!;
  const box = (await message.boundingBox())!;
  expect(box.width).toBeGreaterThan(120);
  expect(box.x + box.width).toBeLessThanOrEqual(cellBox.x + 1);
  expect(box.x).toBeGreaterThanOrEqual(grid.x);
  expect(box.y).toBeGreaterThanOrEqual(cellBox.y + cellBox.height);
  const titleBox = (await message.getByTestId("validation-input-title").boundingBox())!;
  expect(titleBox.height).toBeLessThan(cellBox.height * 1.5);
});

/** Opens Data ▸ Data validation… on the selection. */
async function openValidationDialog(page: Page): Promise<Locator> {
  await page.getByTestId("menu-title-data").click();
  await page.getByTestId("sheet-menu-data-validation").click();
  const dialog = page.getByTestId("validation-dialog");
  await expect(dialog).toBeVisible();
  return dialog;
}

/** Gives the selection a List rule with `source` through the dialog. */
async function setList(page: Page, source: string): Promise<void> {
  const dialog = await openValidationDialog(page);
  await dialog.getByTestId("validation-allow").selectOption({ label: "List" });
  await dialog.getByTestId("validation-value-1").fill(source);
  await dialog.getByTestId("validation-ok").click();
  await expect(dialog).toBeHidden();
}

/** Gives the active cell an input message through the dialog. */
async function setInputMessage(page: Page, title: string, text: string): Promise<void> {
  const dialog = await openValidationDialog(page);
  await dialog.getByTestId("validation-tab-input").click();
  await dialog.getByTestId("validation-show-input").check();
  await dialog.getByTestId("validation-input-title").fill(title);
  await dialog.getByTestId("validation-input-text").fill(text);
  await dialog.getByTestId("validation-ok").click();
  await expect(dialog).toBeHidden();
}

/** The first value field (Source, Minimum…) the dialog loads for the active cell. */
async function loadedSource(page: Page): Promise<string> {
  const dialog = await openValidationDialog(page);
  const value = await dialog.getByTestId("validation-value-1").inputValue();
  await dialog.getByTestId("validation-cancel").click();
  await expect(dialog).toBeHidden();
  return value;
}

/** The items the active cell's arrow lists (the list closed again). */
async function listedItems(page: Page): Promise<string[]> {
  await page.getByTestId("validation-dropdown").click();
  const list = page.getByTestId("validation-list");
  await expect(list).toBeVisible();
  const items = await list.getByTestId("validation-option").allInnerTexts();
  await page.keyboard.press("Escape");
  await expect(list).toBeHidden();
  return items;
}

/** The text of the last header (column or row) whose box passes `fits`. */
async function lastHeader(page: Page, testId: string, fits: (box: { x: number; y: number; width: number; height: number }) => boolean): Promise<string> {
  const headers = page.getByTestId(testId);
  for (let index = (await headers.count()) - 1; index >= 0; index -= 1) {
    const box = await headers.nth(index).boundingBox();
    if (box && fits(box)) {
      return (await headers.nth(index).innerText()).trim();
    }
  }
  return "";
}

test("in a 900x600 window an input message stays beside its cell — left of and below a right-edge cell, left of and above the corner cell — and re-places on a resize (AC-20)", async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 600 });
  await open(page, join(fixtures, "spec130-validation.xlsx"));
  const grid = (await page.getByTestId("spreadsheet-grid").boundingBox())!;
  const letters = await lastHeader(page, "sheet-col-header", (box) => box.x + box.width <= grid.x + grid.width - 20);
  const digits = await lastHeader(page, "sheet-row-header", (box) => box.y + box.height <= grid.y + grid.height - 5);
  expect(letters).toMatch(/^[A-Z]+$/);
  expect(Number(digits)).toBeGreaterThan(10);
  const message = page.getByTestId("validation-input-message");

  // Row 3 of the right-most column: left of the cell and below it — not
  // pinned to the top of the area over the cell.
  const edge = `${letters}3`;
  await clickCell(page, edge);
  await setInputMessage(page, "Score", "Enter a whole number between 1 and 10");
  await expect(message).toBeVisible();
  let cellBox = (await cell(page, edge).boundingBox())!;
  let box = (await message.boundingBox())!;
  expect(box.width).toBeGreaterThan(120);
  expect(box.x + box.width).toBeLessThanOrEqual(cellBox.x + 1);
  expect(box.x).toBeGreaterThanOrEqual(grid.x);
  expect(box.y).toBeGreaterThanOrEqual(cellBox.y + cellBox.height);

  // The bottom-right corner cell: left of it and above it, inside the grid.
  const corner = `${letters}${digits}`;
  await clickCell(page, corner);
  await setInputMessage(page, "Corner note", "A longer message that wraps over several lines at its natural width and must stay inside the grid.");
  await expect(message).toContainText("Corner note");
  cellBox = (await cell(page, corner).boundingBox())!;
  box = (await message.boundingBox())!;
  expect(box.x + box.width).toBeLessThanOrEqual(cellBox.x + 1);
  expect(box.y + box.height).toBeLessThanOrEqual(cellBox.y + 1);
  expect(box.y).toBeGreaterThanOrEqual(grid.y);

  // A wider, taller window leaves room below-right: the message moves there.
  await page.setViewportSize({ width: 1440, height: 900 });
  await expect
    .poll(async () => {
      const at = (await cell(page, corner).boundingBox())!;
      const placed = (await message.boundingBox())!;
      return placed.x >= at.x + at.width - 1 && placed.y >= at.y + at.height;
    })
    .toBe(true);
});

test("the warning and information alerts commit, Ctrl/Cmd+Enter's first failing target decides, and a rule without its error alert commits silently (AC-9, AC-11, AC-14, AC-15)", async ({ page }) => {
  await open(page, join(fixtures, "spec130-validation.xlsx"));
  const alert = page.getByTestId("validation-alert");
  const icon = alert.getByTestId("validation-alert-icon");

  // B3 (whole 1–10, warning): Yes commits 42.
  await clickCell(page, "B3");
  await page.keyboard.type("42");
  await page.keyboard.press("Enter");
  await expect(alert).toBeVisible();
  await expect(icon).toHaveAttribute("data-icon", "warning");
  await expect(alert.getByTestId("validation-alert-cancel")).toBeVisible();
  await alert.getByTestId("validation-alert-yes").click();
  await expect(alert).toBeHidden();
  await expect(cell(page, "B3")).toHaveText("42");

  // B4 (decimal > 0, information): OK and Cancel only; OK commits -5.
  await clickCell(page, "B4");
  await page.keyboard.type("-5");
  await page.keyboard.press("Enter");
  await expect(alert).toBeVisible();
  await expect(icon).toHaveAttribute("data-icon", "info");
  await expect(alert).toContainText("Invalid entry");
  await expect(alert).toContainText("This value doesn't match the data validation rules for this cell.");
  await expect(alert.getByTestId("validation-alert-retry")).toHaveCount(0);
  await expect(alert.getByTestId("validation-alert-cancel")).toBeVisible();
  await alert.getByTestId("validation-alert-ok").click();
  await expect(alert).toBeHidden();
  await expect(cell(page, "B4")).toHaveText("-5");

  // Ctrl/Cmd+Enter "abc" into B6:B7: B6 (length ≤ 5) takes it, B7 (custom:
  // contains "@") refuses — the stop alert; Cancel leaves both as they were.
  await clickCell(page, "B6");
  await page.keyboard.press("Shift+ArrowDown");
  await page.keyboard.type("abc");
  await page.keyboard.press("ControlOrMeta+Enter");
  await expect(alert).toBeVisible();
  await expect(icon).toHaveAttribute("data-icon", "error");
  await alert.getByTestId("validation-alert-cancel").click();
  await expect(alert).toBeHidden();
  await expect(cell(page, "B6")).toHaveText("");
  await expect(cell(page, "B7")).toHaveText("");
  // "a@b" passes both and fills both.
  await clickCell(page, "B6");
  await page.keyboard.press("Shift+ArrowDown");
  await page.keyboard.type("a@b");
  await page.keyboard.press("ControlOrMeta+Enter");
  await expect(cell(page, "B7")).toHaveText("a@b");
  await expect(cell(page, "B6")).toHaveText("a@b");
  await expect(alert).toBeHidden();

  // H3: whole ≤ 5 with "Show error alert" off — 999 commits silently.
  await clickCell(page, "H3");
  const dialog = await openValidationDialog(page);
  await dialog.getByTestId("validation-allow").selectOption({ label: "Whole number" });
  await dialog.getByTestId("validation-operator").selectOption({ label: "less than or equal to" });
  await dialog.getByTestId("validation-value-1").fill("5");
  await dialog.getByTestId("validation-tab-error").click();
  await dialog.getByTestId("validation-show-error").uncheck();
  await dialog.getByTestId("validation-ok").click();
  await expect(dialog).toBeHidden();
  await page.keyboard.type("999");
  await page.keyboard.press("Enter");
  await expect(cell(page, "H3")).toHaveText("999");
  await expect(alert).toBeHidden();
});

test("paste, fill and undo write without a check, a copy leaves the rules where they were, and a cut takes its cell's rule along (AC-16, AC-17, AC-18, AC-47)", async ({ page }) => {
  await open(page, join(fixtures, "spec130-validation.xlsx"));
  const alert = page.getByTestId("validation-alert");
  const arrow = page.getByTestId("validation-dropdown");

  // "Urgent" pasted into B2 (Low, Medium, High): written, B2's rule unchanged.
  await clickCell(page, "F2");
  await page.keyboard.type("Urgent");
  await page.keyboard.press("Enter");
  await clickCell(page, "F2");
  await page.keyboard.press("ControlOrMeta+c");
  await clickCell(page, "B2");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "B2")).toHaveText("Urgent");
  await expect(alert).toBeHidden();
  expect(await listedItems(page)).toEqual(["Low", "Medium", "High"]);

  // Filled down into B3 (whole 1–10), undone and redone: never checked.
  await page.keyboard.press("Shift+ArrowDown");
  await page.keyboard.press("ControlOrMeta+d");
  await expect(cell(page, "B3")).toHaveText("Urgent");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "B3")).toHaveText("5");
  await page.keyboard.press("ControlOrMeta+y");
  await expect(cell(page, "B3")).toHaveText("Urgent");
  await expect(alert).toBeHidden();

  // B2 copied to F4: the value comes, the rule does not.
  await clickCell(page, "B2");
  await page.keyboard.press("ControlOrMeta+c");
  await clickCell(page, "F4");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "F4")).toHaveText("Urgent");
  await expect(arrow).toHaveCount(0);
  await expect(page.getByTestId("validation-input-message")).toHaveCount(0);

  // J2 of a rule on J2:J10 cut and pasted on L2 (its own rule): L2 takes
  // J2's rule, J2 loses it, J3 keeps it.
  await clickCell(page, "J2");
  for (let step = 0; step < 8; step += 1) {
    await page.keyboard.press("Shift+ArrowDown");
  }
  await setList(page, "a,b");
  await clickCell(page, "L2");
  await setList(page, "x,y");
  await clickCell(page, "J2");
  await page.keyboard.type("a");
  await page.keyboard.press("Enter");
  await clickCell(page, "J2");
  await page.keyboard.press("ControlOrMeta+x");
  await clickCell(page, "L2");
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell(page, "L2")).toHaveText("a");
  await expect(cell(page, "J2")).toHaveText("");
  await clickCell(page, "L2");
  expect(await listedItems(page)).toEqual(["a", "b"]);
  await clickCell(page, "J2");
  await expect(arrow).toHaveCount(0);
  await clickCell(page, "J3");
  expect(await listedItems(page)).toEqual(["a", "b"]);
});

test("on an autofilter header Alt+Down stays the filter's, and the arrow opens the list (AC-22)", async ({ page }) => {
  await open(page, join(fixtures, "spec130-validation.xlsx"));
  await clickCell(page, "B1");
  await setList(page, "Entry,Value");
  await clickCell(page, "A1");
  for (let step = 0; step < 9; step += 1) {
    await page.keyboard.press("Shift+ArrowDown");
  }
  await page.keyboard.press("Shift+ArrowRight");
  await page.getByTestId("menu-title-data").click();
  await page.getByTestId("sheet-menu-filter").click();

  await clickCell(page, "B1");
  await page.keyboard.press("Alt+ArrowDown");
  const popover = page.getByTestId("filter-popover");
  await expect(popover).toBeVisible();
  await expect(page.getByTestId("validation-list")).toHaveCount(0);
  await page.keyboard.press("Escape");
  await expect(popover).toBeHidden();
  expect(await listedItems(page)).toEqual(["Entry", "Value"]);
});

test("inserted rows move a rule and its list source, and a renamed sheet renames a rule's source (AC-45, AC-46)", async ({ page }) => {
  await open(page, join(fixtures, "spec130-validation.xlsx"));

  // Rows 2–3 → "Insert 2 rows above": B8's rule (Source $E$2:$E$6, below the
  // insertion) is now B10's, its source $E$4:$E$8 — still listing the options.
  const row = (digits: string) => page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) });
  await row("2").click();
  await row("3").click({ modifiers: ["Shift"] });
  await row("3").click({ button: "right" });
  await page.getByRole("menuitem", { name: "Insert 2 rows above" }).click();
  await clickCell(page, "B10");
  expect(await loadedSource(page)).toBe("=$E$4:$E$8");
  await clickCell(page, "B4");
  expect(await listedItems(page)).toEqual(["Low", "Medium", "High"]);

  // Lists renamed Catalog: B9's cross-sheet source, now B11's, follows.
  await page.getByTestId("sheet-tab").filter({ hasText: /^\s*Lists\s*$/ }).dblclick();
  const input = page.getByTestId("sheet-tab-rename-input");
  await input.fill("Catalog");
  await input.press("Enter");
  await expect(input).toHaveCount(0);
  await page.getByTestId("sheet-tab").filter({ hasText: /^\s*Form\s*$/ }).click();
  await clickCell(page, "B11");
  expect(await loadedSource(page)).toBe("=Catalog!$A$1:$A$4");
  expect(await listedItems(page)).toEqual(["North", "South", "East", "West"]);
});

test("an untouched workbook downloads byte for byte as it was opened, in both formats (AC-40)", async ({ page }) => {
  for (const name of ["spec130-validation.xlsx", "spec130-validation.ods"]) {
    await open(page, join(fixtures, name));
    const saved = await download(page, name);
    expect(readFileSync(saved).equals(readFileSync(join(fixtures, name))), name).toBe(true);
  }
});
