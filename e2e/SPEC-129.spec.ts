import { test, expect, type Locator, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { SPEC129_CF_XLSX } from "../src/lib/spreadsheet/fixtures/build";

// SPEC-129 — conditional formatting (AC-52), with the real pointer and
// keyboard: open spec129-cf.xlsx → B3 (140, "> 100") shows Light red fill
// with dark red text → the C column's data bars widen with their values and
// the negative cell's bar sits left of the axis → D2 shows the 3Arrows green
// up arrow → Format ▸ Conditional formatting… → select B2:B13 → Add ▸ Greater
// than… → 120 → "Green fill with dark green text" → Save → B3 turns green
// (the new rule ranks first) → Move down → B3 is red again → type 130 into
// B4 → it turns red → Ctrl/Cmd+Z → plain again → download → reopen → the
// manager lists the same rules in the same order. Then spec129-cf.ods: its
// calcext:condition cell shows its named style's ("Bad") fill.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const SCORES = SPEC129_CF_XLSX.scores;

/** Excel's presets as getComputedStyle reports them. */
const RED_FILL = "rgb(255, 199, 206)";
const RED_TEXT = "rgb(156, 0, 6)";
const GREEN_FILL = "rgb(198, 239, 206)";
const GREEN_TEXT = "rgb(0, 97, 0)";

test.use({ viewport: { width: 1440, height: 900 } });

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

/** Clicks the centre of a cell through its column and row headers. */
async function clickCell(page: Page, ref: string, modifiers: "Shift"[] = []): Promise<void> {
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
  const path = join(tmpdir(), `spec129-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  return path;
}

async function openManager(page: Page): Promise<Locator> {
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("sheet-menu-conditional-formatting").click();
  const manager = page.getByTestId("cf-manager");
  await expect(manager).toBeVisible();
  return manager;
}

const backgroundOf = (page: Page, ref: string): Promise<string> => cell(page, ref).evaluate((element) => getComputedStyle(element).backgroundColor);

const textColorOf = (page: Page, ref: string): Promise<string> =>
  cell(page, ref).evaluate((element) => getComputedStyle(element.querySelector(".sheet-cell-text") ?? element).color);

const barWidth = async (page: Page, ref: string): Promise<number> => Number(await cell(page, ref).getByTestId("sheet-cf-databar").getAttribute("data-width-pct"));

test("conditional formats render from the file, and the manager adds, reorders, undoes and round-trips rules (AC-52)", async ({ page }) => {
  await open(page, join(fixtures, "spec129-cf.xlsx"));

  // B3 = 140 meets the file's "> 100" rule: Light red fill with dark red text.
  await expect(cell(page, "B3")).toHaveAttribute("data-cf", "true");
  expect(await backgroundOf(page, "B3")).toBe(RED_FILL);
  expect(await textColorOf(page, "B3")).toBe(RED_TEXT);

  // The data bars widen with their values; the negative cell's bar is drawn
  // left of the axis, a positive one right of it.
  const changes = SCORES.changes.map((value, index) => ({ value, ref: `C${index + 2}` }));
  const positives = changes.filter(({ value }) => value > 0).toSorted((a, b) => a.value - b.value);
  const widths: number[] = [];
  for (const { ref } of positives) {
    await expect(cell(page, ref).getByTestId("sheet-cf-databar")).toHaveAttribute("aria-hidden", "true");
    widths.push(await barWidth(page, ref));
  }
  for (let at = 1; at < widths.length; at += 1) {
    expect(widths[at]!, `${positives[at]!.ref} (${positives[at]!.value})`).toBeGreaterThan(widths[at - 1]!);
  }
  const negative = SCORES.negativeCells[0]!;
  const negativeBar = (await cell(page, negative).getByTestId("sheet-cf-databar").boundingBox())!;
  const negativeAxis = (await cell(page, negative).locator(".sheet-cf-axis").boundingBox())!;
  expect(negativeBar.x + negativeBar.width).toBeLessThanOrEqual(negativeAxis.x + 1);
  const positiveBar = (await cell(page, positives[0]!.ref).getByTestId("sheet-cf-databar").boundingBox())!;
  const positiveAxis = (await cell(page, positives[0]!.ref).locator(".sheet-cf-axis").boundingBox())!;
  expect(positiveBar.x).toBeGreaterThanOrEqual(positiveAxis.x - 1);

  // D2, the Trend column's largest value, shows the top arrow.
  const icon = cell(page, SCORES.arrows.topCell).getByTestId("sheet-cf-icon");
  await expect(icon).toHaveAttribute("data-icon", SCORES.arrows.topIcon);
  await expect(icon).toHaveAttribute("role", "img");
  await expect(icon).toHaveAttribute("aria-label", "Green up arrow");

  // Format ▸ Conditional formatting…, then select B2:B13 and add "> 120" in green.
  const manager = await openManager(page);
  await clickCell(page, "B2");
  await clickCell(page, "B13", ["Shift"]);
  await manager.getByTestId("cf-add").click();
  await page.getByTestId("cf-preset-greater-than").click();
  const editor = page.getByTestId("cf-editor");
  await expect(editor).toBeVisible();
  await editor.getByTestId("cf-value-1").fill("120");
  await editor.getByTestId("cf-format").selectOption({ label: "Green fill with dark green text" });
  await editor.getByTestId("cf-save").click();
  await expect(editor).toBeHidden();

  // The new rule ranks first: B3 turns green.
  await expect(manager.getByTestId("cf-rule-description").first()).toHaveText("Cell value > 120");
  await expect.poll(() => backgroundOf(page, "B3")).toBe(GREEN_FILL);
  expect(await textColorOf(page, "B3")).toBe(GREEN_TEXT);

  // Moving it down puts the file's "> 100" first again: B3 is red.
  await manager.getByTestId("cf-move-down").click();
  await expect(manager.getByTestId("cf-rule-description").nth(1)).toHaveText("Cell value > 120");
  await expect.poll(() => backgroundOf(page, "B3")).toBe(RED_FILL);

  // 130 in B4 meets both rules; "> 100" ranks first.
  await clickCell(page, "B4");
  await page.keyboard.type("130");
  await page.keyboard.press("Enter");
  await expect(cell(page, "B4")).toHaveText("130");
  await expect.poll(() => backgroundOf(page, "B4")).toBe(RED_FILL);
  expect(await textColorOf(page, "B4")).toBe(RED_TEXT);

  // Undo restores 90: no rule holds, the cell is plain.
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "B4")).toHaveText(String(SCORES.cellIs.coolValue));
  await expect.poll(() => backgroundOf(page, "B4")).not.toBe(RED_FILL);
  await expect(cell(page, "B4")).not.toHaveAttribute("data-cf", "true");

  // Download → reopen: the manager lists the same rules in the same order.
  const listed = await manager.getByTestId("cf-rule-description").allInnerTexts();
  expect(listed.slice(0, 2)).toEqual(["Cell value > 100", "Cell value > 120"]);
  const saved = await download(page, "spec129-cf.xlsx");
  await open(page, saved);
  const reopened = await openManager(page);
  await expect(reopened.getByTestId("cf-rule-description")).toHaveText(listed);
  await expect.poll(() => backgroundOf(page, "B3")).toBe(RED_FILL);
});

test("spec129-cf.ods: a calcext:condition cell shows its named style's fill (AC-52)", async ({ page }) => {
  await open(page, join(fixtures, "spec129-cf.ods"));
  // B3 = 140 meets "> 100", which applies the named style "Bad".
  await expect(cell(page, "B3")).toHaveAttribute("data-cf", "true");
  await expect.poll(() => backgroundOf(page, "B3")).toBe(RED_FILL);
  expect(await textColorOf(page, "B3")).toBe(RED_TEXT);
  // B4 = 90 does not.
  expect(await backgroundOf(page, "B4")).not.toBe(RED_FILL);
});
