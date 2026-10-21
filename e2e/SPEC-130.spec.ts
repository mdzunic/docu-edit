import { test, expect, type Locator, type Page } from "@playwright/test";
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
// a repeated run of empty cells, shows its input message.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 } });

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
  await expect(cell(page, "C500")).toHaveAttribute("aria-describedby", (await message.getAttribute("id"))!);
});
