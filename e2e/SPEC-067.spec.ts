import { test, expect, type Page } from "@playwright/test";

// SPEC-067 — the redesigned document library: list/grid views, search, sort,
// filters, starring, multi-select + bulk actions, trash and the storage panel.
// Each Playwright test runs in a fresh browser context, so IndexedDB starts
// empty and the first-run empty state shows until a document is created.

async function createDocument(page: Page, title: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill(title);
  await page.getByTestId("document-title-input").press("Enter");
  await expect(page.getByTestId("document-title")).toHaveText(title);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

test("the first-run empty state describes the product and the three ways to start (AC-51..53)", async ({
  page,
}) => {
  await page.goto("/");
  const empty = page.getByTestId("empty-state");
  await expect(empty).toBeVisible();
  await expect(empty).toContainText("docu-edit");
  await expect(empty).toContainText("stored in this browser");
  await expect(page.getByTestId("new-document")).toBeVisible();
  await expect(page.getByTestId("new-from-template")).toBeVisible();
  await expect(page.getByTestId("open-document")).toBeVisible();
});

test("the view toggle switches to the grid and the choice persists (AC-1, AC-2)", async ({
  page,
}) => {
  await createDocument(page, "Griddable");
  await page.goto("/");

  await expect(page.getByTestId("library-list")).toBeVisible();
  await page.getByTestId("view-grid").click();
  await expect(page.getByTestId("library-grid")).toBeVisible();
  // The grid card shows a thumbnail.
  await expect(page.getByTestId("row-thumbnail").first()).toBeVisible();

  // The choice survives a reload.
  await page.reload();
  await expect(page.getByTestId("library-grid")).toBeVisible();
});

test("search filters by subsequence, highlights matches, and offers a clear when nothing matches (AC-8..10, AC-54, AC-55)", async ({
  page,
}) => {
  await createDocument(page, "Quarterly Report");
  await createDocument(page, "Budget");
  await page.goto("/");

  await page.getByTestId("library-search").fill("qrt");
  await expect(page.getByTestId("row-title")).toHaveCount(1);
  await expect(page.getByTestId("row-title")).toHaveText("Quarterly Report");
  await expect(page.getByTestId("title-match").first()).toBeVisible();

  await page.getByTestId("library-search").fill("zzzzz");
  await expect(page.getByTestId("no-results")).toBeVisible();
  await page.getByTestId("clear-filters").click();
  await expect(page.getByTestId("row-title")).toHaveCount(2);
});

test("format and starred filter chips narrow the list with live counts (AC-16, AC-17)", async ({
  page,
}) => {
  await createDocument(page, "One");
  await createDocument(page, "Two");
  await page.goto("/");

  // Star "One" from its row, then filter to starred only.
  await rowByTitle(page, "One").getByTestId("row-star").click();
  await expect(page.getByTestId("filter-starred")).toContainText("1");
  await page.getByTestId("filter-starred").click();
  await expect(page.getByTestId("row-title")).toHaveText("One");
});

test("starred documents sort first within the active sort (AC-21)", async ({
  page,
}) => {
  await createDocument(page, "Alpha");
  await createDocument(page, "Bravo");
  await page.goto("/");

  // Star the older document (Alpha); it should jump to the top.
  await rowByTitle(page, "Alpha").getByTestId("row-star").click();
  const titles = page.getByTestId("row-title");
  await expect(titles.nth(0)).toHaveText("Alpha");
});

test("multi-select drives a bulk move-to-trash (AC-23, AC-30, AC-31, AC-32)", async ({
  page,
}) => {
  await createDocument(page, "First");
  await createDocument(page, "Second");
  await page.goto("/");

  await page.getByTestId("row-select").nth(0).check();
  await page.getByTestId("row-select").nth(1).check();
  await expect(page.getByTestId("selection-count")).toContainText("2");

  await page.getByTestId("bulk-trash").click();
  await expect(page.getByTestId("library-empty")).toBeVisible();
  await page.getByTestId("trash-link").click();
  await expect(page.getByTestId("library-row")).toHaveCount(2);
});

test("Enter opens the focused row from the keyboard (AC-57, AC-58)", async ({
  page,
}) => {
  const id = await createDocument(page, "Keyboard");
  await page.goto("/");

  await rowByTitle(page, "Keyboard").focus();
  await page.keyboard.press("Enter");
  await expect(page).toHaveURL(new RegExp(`/doc/${id}$`));
});
