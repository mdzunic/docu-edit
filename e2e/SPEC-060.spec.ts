import { test, expect, type Page } from "@playwright/test";

// SPEC-060 — close the document and return to the library.
//
// Closing is a commit boundary: the Close control (and File → Close, the
// product-name home link and the browser Back button) flushes any pending
// autosave and only then returns to `/`, so the library reads the just-saved
// row rather than a stale one. Each Playwright test runs in a fresh context,
// so IndexedDB starts empty.

/** Creates a document via the UI and returns its id from the URL. */
async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

async function renameTo(page: Page, title: string): Promise<void> {
  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill(title);
  await page.getByTestId("document-title-input").press("Enter");
  await expect(page.getByTestId("document-title")).toHaveText(title);
}

async function typeBody(page: Page, text: string): Promise<void> {
  await page.getByTestId("document-content").click();
  await page.keyboard.type(text);
}

/** The library row whose title cell shows exactly `title`. */
function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

// AC-29: edit, close via the app-bar Close control, and read the edited row
// straight off the library — no waitForTimeout, no reload, no retry loop. The
// flush is awaited before navigation, so the row is already current on arrival.
test("closing flushes so the library shows the edited title and content (AC-1, AC-3, AC-7, AC-8, AC-29)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await renameTo(page, "Meeting notes");
  await typeBody(page, "Agenda for Monday");

  await page.getByTestId("close-document").click();

  await expect(page).toHaveURL(/\/$/);
  await expect(rowByTitle(page, "Meeting notes").getByTestId("row-title")).toHaveText(
    "Meeting notes",
  );

  // The committed content is there too: reopening the same id shows the body.
  await rowByTitle(page, "Meeting notes").click();
  await expect(page).toHaveURL(new RegExp(`/doc/${id}$`));
  await expect(page.getByTestId("document-content")).toContainText(
    "Agenda for Monday",
  );
});

test("the Close control is labelled and icon-only, and is absent in the library (AC-2, AC-24, AC-25)", async ({
  page,
}) => {
  await createDocument(page);
  const close = page.getByTestId("close-document");
  await expect(close).toHaveAttribute("aria-label", "Close document");
  await expect(close).toHaveAttribute("title", "Close document");

  await close.click();
  await expect(page.getByTestId("library")).toBeVisible();
  // The library's app bar carries no Close control, and shows no real document
  // title — the placeholder is unchanged from today (AC-24, AC-25).
  await expect(page.getByTestId("close-document")).toHaveCount(0);
  await expect(page.getByTestId("document-title")).toHaveText("No document");
});

test("File → Close returns to the library (AC-4, AC-5)", async ({ page }) => {
  await createDocument(page);
  await renameTo(page, "Via the menu");

  await page.getByTestId("menu-title-file").click();
  await expect(page.getByTestId("menu-file")).toBeVisible();
  await page.getByTestId("file-menu-close").click();

  await expect(page.getByTestId("library")).toBeVisible();
  await expect(rowByTitle(page, "Via the menu")).toBeVisible();
});

test("the product name is a home link that flushes on the way out (AC-6, AC-6b)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await typeBody(page, "Draft body text");

  const home = page.getByTestId("product-name");
  // It reads as a link to assistive technology.
  await expect(home).toHaveJSProperty("tagName", "A");
  await home.click();

  await expect(page.getByTestId("library")).toBeVisible();
  // The pending edit was flushed before leaving: reopening shows it.
  await page.goto(`/doc/${id}`);
  await expect(page.getByTestId("document-content")).toContainText(
    "Draft body text",
  );
});

test("browser Back out of the editor takes the same flush path (AC-18)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await typeBody(page, "Back button content");

  await page.goBack();
  await expect(page.getByTestId("library")).toBeVisible();

  await page.goto(`/doc/${id}`);
  await expect(page.getByTestId("document-content")).toContainText(
    "Back button content",
  );
});

test("Escape does not close the document (AC-23)", async ({ page }) => {
  const id = await createDocument(page);
  await page.getByTestId("document-content").click();
  await page.keyboard.press("Escape");
  // Still in the editor, on the same document.
  await expect(page).toHaveURL(new RegExp(`/doc/${id}$`));
  await expect(page.getByTestId("close-document")).toBeVisible();
});

test("the just-closed row is the library's scroll target on return (AC-21)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await renameTo(page, "Return here");
  await page.getByTestId("close-document").click();

  await expect(page.getByTestId("library")).toBeVisible();
  // The closed row is addressable by its id — the anchor HomeView scrolls to.
  await expect(page.locator(`[data-doc-id="${id}"]`)).toBeVisible();
  // The id rode in router history state, not the URL (a privacy leak the spec
  // avoids) — this is what HomeView reads to pick the row to flash (AC-22).
  expect(page.url()).not.toContain(id);
  await expect
    .poll(() => page.evaluate(() => window.history.state?.closedDocId))
    .toBe(id);
});
