import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-076 — table header repetition, the style gallery, sorting and both
// text↔table conversions (AC-77). The unit suites cover every rule in
// isolation — the four-level resolution, the band arithmetic, the stable
// multi-key sort, both conversions and both round trips; this file exercises
// the seams only a real browser can confirm: that a header run actually
// redraws at the top of a continuation page, that a gallery click repaints
// the live table, that toggling banded rows off and back leaves the cells
// exactly as they were, and that the two conversions land in the canvas.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

const content = (page: Page) => page.getByTestId("document-content");

async function openFixture(page: Page, name: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(join(fixturesDir, name));
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page).locator("p, h1, h2").first()).toBeVisible();
}

async function waitForPagination(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-pages")).toHaveAttribute(
    "data-pagination-state",
    "clean",
    { timeout },
  );
}

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page).locator("p").first()).toBeVisible();
}

/** Opens the Table menu and picks one of its items. */
async function tableMenu(page: Page, itemTestId: string): Promise<void> {
  await page.getByTestId("menu-title-table").click();
  await page.getByTestId(itemTestId).click();
}

/** Puts the caret in a table cell by clicking its first paragraph. */
async function clickCell(page: Page, row: number, col: number): Promise<void> {
  await content(page).getByTestId(`table-cell-${row}-${col}`).locator("p").first().click();
}

test.describe("SPEC-076 repeating header rows", () => {
  test("a table's header run redraws at the top of every continuation (AC-1, AC-7)", async ({ page }) => {
    await openFixture(page, "tables-repeating-header.docx");
    await waitForPagination(page);

    // The fixture is long enough to run over several pages.
    const pages = page.getByTestId("document-pages").locator(".page");
    expect(await pages.count()).toBeGreaterThan(1);

    // Every continuation slice of the table draws the header run again…
    const repeats = content(page).locator("[data-repeated-header]");
    expect(await repeats.count()).toBeGreaterThan(0);
    await expect(repeats.first()).toContainText("Item");

    // …as a layout artefact: uneditable, hidden from assistive tech, and
    // never a second copy of the model's own header row.
    await expect(repeats.first()).toHaveAttribute("aria-hidden", "true");
    await expect(repeats.first()).toHaveAttribute("contenteditable", "false");
    // The real header cell keeps its testid; the repeat gets its own, so the
    // model still holds exactly one row 0 (AC-3).
    await expect(content(page).getByTestId("table-cell-0-0")).toHaveCount(1);
  });

  test("clicking a repeated header cell puts the caret in the real header cell (AC-8)", async ({ page }) => {
    await openFixture(page, "tables-repeating-header.docx");
    await waitForPagination(page);
    const repeatCell = content(page).getByTestId("repeated-header-cell-0-0").first();
    await expect(repeatCell).toBeVisible();
    // Click near the end of "Item", then type: the character must land in the
    // REAL header cell, at the offset nearest the click — the repeat itself
    // is uneditable and holds no model text of its own.
    // The repeat lives on a later page: scroll it into view so the synthetic
    // click lands on it rather than on whatever the viewport currently holds.
    await repeatCell.scrollIntoViewIfNeeded();
    await waitForPagination(page);
    const box = await repeatCell.boundingBox();
    await page.mouse.click(box!.x + box!.width - 4, box!.y + box!.height / 2);
    await page.keyboard.type("!");
    await waitForPagination(page);
    await expect(content(page).getByTestId("table-cell-0-0")).toContainText("Item!");
    // …and the repeat, being a copy of that row, shows the edit too.
    await expect(content(page).getByTestId("repeated-header-cell-0-0").first()).toContainText("Item!");
  });

  test("the word count is unaffected by the repeat (AC-10)", async ({ page }) => {
    await openFixture(page, "tables-repeating-header.docx");
    await waitForPagination(page);
    const before = await page.getByTestId("word-count-button").textContent();
    // Scrolling realizes more pages, which draws more repeats — the count
    // walks the model, so it cannot move.
    await page.mouse.wheel(0, 4000);
    await waitForPagination(page);
    await expect(page.getByTestId("word-count-button")).toHaveText(before ?? "");
  });
});

test.describe("SPEC-076 the table style gallery", () => {
  test("applying a look repaints the table, and toggling banded rows off and back restores it (AC-31..AC-35)", async ({ page }) => {
    await openFixture(page, "tables-sorted.docx");
    await waitForPagination(page);
    await clickCell(page, 1, 0);

    const bandedCell = content(page).getByTestId("table-cell-1-0");
    const plain = await bandedCell.evaluate((el) => globalThis.getComputedStyle(el).backgroundColor);

    await tableMenu(page, "table-menu-styles");
    await expect(page.getByTestId("table-style-gallery")).toBeVisible();
    // Five built-ins plus "None" (AC-31, AC-34).
    await expect(page.getByTestId("table-style-none")).toBeVisible();
    await expect(page.getByTestId("table-style-TableStyleAccent")).toBeVisible();

    await page.getByTestId("table-style-TableStyleBandedRows").click();
    await waitForPagination(page);
    const banded = await bandedCell.evaluate((el) => globalThis.getComputedStyle(el).backgroundColor);
    expect(banded).not.toBe(plain);

    // The gallery stays open; the six look checkboxes act on the live table.
    await page.getByTestId("table-look-bandedRows").click();
    await waitForPagination(page);
    await expect
      .poll(async () => bandedCell.evaluate((el) => globalThis.getComputedStyle(el).backgroundColor))
      .toBe(plain);

    await page.getByTestId("table-look-bandedRows").click();
    await waitForPagination(page);
    await expect
      .poll(async () => bandedCell.evaluate((el) => globalThis.getComputedStyle(el).backgroundColor))
      .toBe(banded);
  });

  test("a banded table's shading and borders reach the exported PDF (AC-29)", async ({ page }) => {
    await openFixture(page, "tables-sorted.docx");
    await waitForPagination(page);
    await clickCell(page, 1, 0);

    await tableMenu(page, "table-menu-styles");
    // The Accent tile sits in the gallery's third column, whose live table
    // preview can push it past the e2e viewport's right edge; tile
    // clickability is the gallery test's concern, this test's is the export,
    // so the click is dispatched rather than positional.
    await page.getByTestId("table-style-TableStyleAccent").dispatchEvent("click");
    await page.getByTestId("table-style-close").click();
    await waitForPagination(page);
    // The look landed before we export: the header row carries Accent's
    // first-row shading.
    await expect
      .poll(async () =>
        content(page)
          .getByTestId("table-cell-0-0")
          .evaluate((el) => globalThis.getComputedStyle(el).backgroundColor),
      )
      .toBe("rgb(47, 84, 150)");

    // Export the PDF and inspect its bytes: the builder writes content
    // streams uncompressed, so the resolved cell paint must show up as
    // filled paths in the exact colours the screen resolves.
    await page.getByTestId("menu-title-file").click();
    await page.getByTestId("file-menu-export").click();
    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-pdf").click(),
    ]);
    const path = await download.path();
    expect(path).toBeTruthy();
    const pdf = readFileSync(path!).toString("latin1");

    // Filled closed paths exist at all (before this spec, a table exported
    // as bare text — not one fill or stroke operator in the stream)…
    expect(pdf).toMatch(/h\nf Q/);
    // …in Accent's band shading (#d9e2f3) and first-row (#2f5496) colours,
    expect(pdf).toContain("0.85 0.89 0.95 rg");
    expect(pdf).toContain("0.18 0.33 0.59 rg");
    // …plus Accent's black grid borders as black-filled strips — `rg`
    // followed by a moveto is a path fill state, not a text colour.
    expect(pdf).toMatch(/q 0 0 0 rg [\d.]+ [\d.]+ m\n/);
  });
});

test.describe("SPEC-076 sorting", () => {
  test("sorts by two columns (AC-38, AC-41, AC-46, AC-51)", async ({ page }) => {
    await openFixture(page, "tables-tab-separated.docx");
    await waitForPagination(page);

    // Turn the list into a table first, so this test owns its own data.
    await content(page).locator("p").filter({ hasText: "Name\tRole" }).first().click();
    await page.keyboard.press(`Shift+${MOD}+End`);
    await tableMenu(page, "table-menu-text-to-table");
    await expect(page.getByTestId("table-separator-dialog")).toBeVisible();
    await page.getByTestId("convert-confirm").click();
    await waitForPagination(page);
    await expect(content(page).getByTestId("doc-table")).toHaveCount(1);

    await clickCell(page, 1, 0);
    await tableMenu(page, "table-menu-sort");
    await expect(page.getByTestId("table-sort-dialog")).toBeVisible();
    // Key one: column 1 (Name), descending. Key two: column 2 (Role).
    await page.getByTestId("sort-column-0").selectOption("0");
    await page.getByTestId("sort-direction-0").selectOption("desc");
    await page.getByTestId("sort-column-1").selectOption("1");
    await page.getByTestId("sort-confirm").click();
    await waitForPagination(page);

    // The header row stayed put and the body reversed.
    await expect(content(page).getByTestId("table-cell-0-0")).toContainText("Name");
    await expect(content(page).getByTestId("table-cell-1-0")).toContainText("Grace");

    // One undo entry puts the original order back (AC-51).
    await page.keyboard.press(`${MOD}+z`);
    await waitForPagination(page);
    await expect(content(page).getByTestId("table-cell-1-0")).toContainText("Ada");
  });
});

test.describe("SPEC-076 conversions", () => {
  test("converts a tab-separated list to a table and back (AC-52..AC-64)", async ({ page }) => {
    await openFixture(page, "tables-tab-separated.docx");
    await waitForPagination(page);

    await content(page).locator("p").filter({ hasText: "Name\tRole" }).first().click();
    await page.keyboard.press(`Shift+${MOD}+End`);
    await tableMenu(page, "table-menu-text-to-table");
    const dialog = page.getByTestId("table-separator-dialog");
    await expect(dialog).toBeVisible();
    // The preview counts rows and columns live (AC-54).
    await expect(page.getByTestId("convert-preview")).toContainText("columns");
    await page.getByTestId("convert-confirm").click();
    await waitForPagination(page);

    const table = content(page).getByTestId("doc-table");
    await expect(table).toHaveCount(1);
    await expect(content(page).getByTestId("table-cell-0-0")).toContainText("Name");
    await expect(content(page).getByTestId("table-cell-0-2")).toContainText("City");

    // …and back again.
    await clickCell(page, 0, 0);
    await tableMenu(page, "table-menu-table-to-text");
    await expect(dialog).toBeVisible();
    await page.getByTestId("convert-confirm").click();
    await waitForPagination(page);
    await expect(content(page).getByTestId("doc-table")).toHaveCount(0);
    await expect(content(page).locator("p").filter({ hasText: "Name" }).first()).toBeVisible();
  });

  test("an invalid custom separator disables OK with an inline error (AC-55)", async ({ page }) => {
    await createDocument(page);
    await content(page).locator("p").first().click();
    await page.keyboard.type("a,b");
    await page.keyboard.press(`Shift+Home`);
    await tableMenu(page, "table-menu-text-to-table");
    await expect(page.getByTestId("table-separator-dialog")).toBeVisible();
    await page.getByTestId("separator-custom").click();
    await expect(page.getByTestId("separator-error")).toBeVisible();
    await expect(page.getByTestId("convert-confirm")).toBeDisabled();
    await page.getByTestId("separator-custom-char").fill(",");
    await expect(page.getByTestId("separator-error")).toHaveCount(0);
    await expect(page.getByTestId("convert-confirm")).toBeEnabled();
  });
});

test.describe("SPEC-076 the Table menu", () => {
  test("sits between Format and Tools and disables its items outside a table (AC-16, AC-18)", async ({ page }) => {
    await createDocument(page);
    const titles = page.getByTestId("menu-bar").locator("[data-menu-title]");
    await expect(titles.nth(4)).toHaveText("Format");
    await expect(titles.nth(5)).toHaveText("Table");
    await expect(titles.nth(6)).toHaveText("Tools");

    await page.getByTestId("menu-title-table").click();
    await expect(page.getByTestId("table-menu-styles")).toBeDisabled();
    await expect(page.getByTestId("table-menu-sort")).toBeDisabled();
    await expect(page.getByTestId("table-menu-repeat-header")).toBeDisabled();
    // Disabled, never hidden — with the reason stated (SPEC-049).
    await expect(page.getByTestId("table-menu-styles")).toHaveAttribute("title", /not in a table/);
  });

  test("Repeat header row reflects and toggles the caret row (AC-17)", async ({ page }) => {
    await openFixture(page, "tables-sorted.docx");
    await waitForPagination(page);
    await clickCell(page, 0, 0);
    await page.getByTestId("menu-title-table").click();
    await expect(page.getByTestId("table-menu-repeat-header")).toHaveAttribute("aria-checked", "true");
    await page.getByTestId("table-menu-repeat-header").click();
    await waitForPagination(page);
    await clickCell(page, 0, 0);
    await page.getByTestId("menu-title-table").click();
    await expect(page.getByTestId("table-menu-repeat-header")).toHaveAttribute("aria-checked", "false");
  });
});
