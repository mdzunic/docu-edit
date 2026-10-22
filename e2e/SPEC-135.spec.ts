import { test, expect, type Locator, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-135 — notes, comments and hyperlinks on cells (AC-65–AC-71), with the
// real pointer and keyboard.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 } });

// The links (AC-68, AC-69 and the links of AC-71's .ods read): select C7,
// Ctrl/Cmd+K, link it to the place Data!A10 → the chip's Open activates
// "Data" with A10 selected; Ctrl/Cmd+click on the javascript: link's text
// shows LINK_NOT_OPENABLE_MESSAGE and opens nothing; in the .ods twin A7 and
// B7 carry data-link and following B7 activates "Data" with A1 selected.
test.describe("SPEC-135 links", () => {
  const NOT_OPENABLE = "This link can't be opened from the browser";

  function cellAt(page: Page, ref: string): Locator {
    return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
  }

  function sheetTab(page: Page, name: string): Locator {
    return page.getByTestId("sheet-tab").filter({ hasText: new RegExp(`^\\s*${name}\\s*$`) });
  }

  /** Clicks the centre of a cell through its column and row headers. */
  async function pressCell(page: Page, ref: string): Promise<void> {
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
    const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
    const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
    await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
    await expect(page.getByTestId("name-box")).toHaveValue(ref);
  }

  /** Opens a fixture, recording every window.open the page makes. */
  async function openFixture(page: Page, name: string): Promise<void> {
    await page.addInitScript(() => {
      const calls: unknown[] = [];
      Object.defineProperty(window, "__linkOpens", { value: calls });
      const original = window.open.bind(window);
      window.open = (...args: Parameters<typeof window.open>) => {
        calls.push(args);
        return original(...args);
      };
    });
    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(join(fixtures, name));
    await expect(page).toHaveURL(/\/sheets\/.+/);
    await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
  }

  const opens = (page: Page): Promise<unknown[]> => page.evaluate(() => (window as unknown as { __linkOpens: unknown[] }).__linkOpens);

  test("Ctrl/Cmd+K links C7 to the place Data!A10, and the chip's Open activates Data with A10 selected", async ({ page }) => {
    await openFixture(page, "spec135-annotations.xlsx");
    await expect(cellAt(page, "C7")).not.toHaveAttribute("data-link", "true");
    await pressCell(page, "C7");
    await page.keyboard.press("ControlOrMeta+k");
    const dialog = page.getByTestId("sheet-link-dialog");
    await expect(dialog).toBeVisible();
    await expect(page.getByTestId("sheet-link-text")).toHaveValue("Details");
    await page.getByTestId("sheet-link-kind-place").check();
    await page.getByTestId("sheet-link-sheet").selectOption("Data");
    await page.getByTestId("sheet-link-ref").fill("A10");
    await expect(page.getByTestId("sheet-link-error")).toHaveText("");
    await page.getByTestId("sheet-link-save").click();
    await expect(dialog).toBeHidden();
    await expect(cellAt(page, "C7")).toHaveAttribute("data-link", "true");
    await expect(cellAt(page, "C7")).toHaveText("Details");
    await expect(page.getByTestId("link-chip")).toBeVisible();
    await expect(page.getByTestId("link-chip-target")).toHaveText("Data!A10");
    await page.getByTestId("link-chip-open").click();
    await expect(sheetTab(page, "Data")).toHaveAttribute("aria-selected", "true");
    await expect(page.getByTestId("name-box")).toHaveValue("A10");
    expect(await opens(page)).toEqual([]);
  });

  test("Ctrl/Cmd+click on the javascript: link's text shows the refusal and opens nothing", async ({ page }) => {
    await openFixture(page, "spec135-annotations.xlsx");
    const popups: Page[] = [];
    page.context().on("page", (opened) => popups.push(opened));
    await expect(cellAt(page, "B8")).toHaveAttribute("data-link", "true");
    await pressCell(page, "A1");
    await cellAt(page, "B8").locator(".sheet-cell-link").click({ modifiers: ["ControlOrMeta"] });
    await expect(page.getByTestId("toast").filter({ hasText: NOT_OPENABLE })).toBeVisible();
    // Followed, not added: the selection is still A1 alone.
    await expect(page.getByTestId("name-box")).toHaveValue("A1");
    await expect(page.getByTestId("sheet-selection")).toHaveCount(0);
    expect(await opens(page)).toEqual([]);
    expect(popups).toEqual([]);
    expect(page.context().pages()).toHaveLength(1);
  });

  test("an .ods workbook's links read: A7 and B7 carry data-link, and following B7 activates Data with A1 selected", async ({ page }) => {
    await openFixture(page, "spec135-annotations.ods");
    await expect(cellAt(page, "A7")).toHaveAttribute("data-link", "true");
    await expect(cellAt(page, "B7")).toHaveAttribute("data-link", "true");
    await expect(cellAt(page, "A1")).not.toHaveAttribute("data-link", "true");
    // Leave Data's selection elsewhere, so the follow is what selects A1.
    await sheetTab(page, "Data").click();
    await pressCell(page, "A10");
    await sheetTab(page, "Summary").click();
    await expect(sheetTab(page, "Summary")).toHaveAttribute("aria-selected", "true");
    await cellAt(page, "B7").locator(".sheet-cell-link").click({ modifiers: ["ControlOrMeta"] });
    await expect(sheetTab(page, "Data")).toHaveAttribute("aria-selected", "true");
    await expect(page.getByTestId("name-box")).toHaveValue("A1");
    expect(await opens(page)).toEqual([]);
  });
});
