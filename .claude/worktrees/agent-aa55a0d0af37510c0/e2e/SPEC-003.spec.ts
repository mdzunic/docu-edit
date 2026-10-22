import { test, expect, type Page } from "@playwright/test";

// SPEC-003 — local document library.
//
// Each Playwright test runs in a fresh browser context, so IndexedDB starts
// empty and `/` shows the SPEC-001 empty state until a document is created.

/** Creates a document via the UI, optionally renames it in the app bar, and
 * returns its id from the URL. */
async function createDocument(page: Page, title?: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  if (title) {
    await page.getByTestId("document-title").click();
    await page.getByTestId("document-title-input").fill(title);
    await page.getByTestId("document-title-input").press("Enter");
    await expect(page.getByTestId("document-title")).toHaveText(title);
  }
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

/** The library row whose title cell shows exactly `title`. */
function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

// Fire-and-forget IndexedDB writes need a beat to commit before a reload.
async function settleWrites(page: Page) {
  await page.waitForTimeout(300);
}

test("documents survive a page reload with content intact (AC-1, AC-2, AC-9)", async ({
  page,
}) => {
  const id = await createDocument(page, "Persist me");
  await settleWrites(page);

  // Reloading the editor restores the persisted document.
  await page.reload();
  await expect(page.getByTestId("document-title")).toHaveText("Persist me");

  // And the library at / lists it; clicking the row reopens the same id.
  await page.goto("/");
  await expect(page.getByTestId("library-list")).toBeVisible();
  await rowByTitle(page, "Persist me").click();
  await expect(page).toHaveURL(new RegExp(`/doc/${id}$`));
  await expect(page.getByTestId("document-title")).toHaveText("Persist me");
});

test("the library lists title, relative time and format badge, newest first (AC-3..AC-7)", async ({
  page,
}) => {
  await createDocument(page, "First doc");
  await createDocument(page, "Second doc");
  await page.goto("/");

  // Columns are labelled.
  const table = page.getByTestId("library-list");
  await expect(table.getByRole("columnheader", { name: "Title" })).toBeVisible();
  await expect(
    table.getByRole("columnheader", { name: "Last modified" }),
  ).toBeVisible();
  await expect(
    table.getByRole("columnheader", { name: "Format" }),
  ).toBeVisible();

  // Both rows are present, newest (Second doc) first (AC-7).
  const titles = page.getByTestId("row-title");
  await expect(titles).toHaveCount(2);
  await expect(titles.nth(0)).toHaveText("Second doc");
  await expect(titles.nth(1)).toHaveText("First doc");

  // Relative last-modified (AC-5) and the "New" badge (AC-6) — nothing has
  // been saved as DOCX/ODT yet.
  await expect(page.getByTestId("row-modified").nth(0)).toHaveText("just now");
  await expect(page.getByTestId("row-format").nth(0)).toHaveText("New");
});

test("clicking a row opens the document at /doc/:id (AC-8)", async ({
  page,
}) => {
  const id = await createDocument(page, "Openable");
  await page.goto("/");
  await rowByTitle(page, "Openable").click();
  await expect(page).toHaveURL(new RegExp(`/doc/${id}$`));
  await expect(page.getByTestId("document-title")).toHaveText("Openable");
});

test("each row has an overflow menu with Rename, Duplicate, Delete (AC-10..AC-13)", async ({
  page,
}) => {
  await createDocument(page, "With menu");
  await page.goto("/");

  await expect(page.getByTestId("row-menu-list")).toBeHidden();
  await rowByTitle(page, "With menu").getByTestId("row-menu").click();
  await expect(page.getByTestId("menu-rename")).toHaveText("Rename");
  await expect(page.getByTestId("menu-duplicate")).toHaveText("Duplicate");
  await expect(page.getByTestId("menu-delete")).toHaveText("Delete");
});

test("Rename opens a prefilled modal; saving applies and persists the new title", async ({
  page,
}) => {
  await createDocument(page, "Old name");
  await page.goto("/");

  await rowByTitle(page, "Old name").getByTestId("row-menu").click();
  await page.getByTestId("menu-rename").click();
  await expect(page.getByTestId("rename-input")).toHaveValue("Old name");

  await page.getByTestId("rename-input").fill("New name");
  await page.getByTestId("rename-save").click();
  await expect(page.getByTestId("rename-dialog")).toBeHidden();
  await expect(page.getByTestId("row-title")).toHaveText("New name");

  // The rename survives a reload.
  await settleWrites(page);
  await page.reload();
  await expect(page.getByTestId("row-title")).toHaveText("New name");
});

test("renaming to a taken title appends a count; empty titles are rejected", async ({
  page,
}) => {
  await createDocument(page, "Report");
  await createDocument(page, "Draft");
  await page.goto("/");

  // Empty title: error shown, save disabled.
  await rowByTitle(page, "Draft").getByTestId("row-menu").click();
  await page.getByTestId("menu-rename").click();
  await page.getByTestId("rename-input").fill("   ");
  await expect(page.getByTestId("rename-error")).toHaveText(
    "Title cannot be empty",
  );
  await expect(page.getByTestId("rename-save")).toBeDisabled();

  // Duplicate title: gets a count appended.
  await page.getByTestId("rename-input").fill("Report");
  await page.getByTestId("rename-save").click();
  await expect(rowByTitle(page, "Report (2)")).toBeVisible();
  await expect(rowByTitle(page, "Report")).toBeVisible();
});

test("Duplicate creates '<title> (copy)' that opens as a separate document (AC-16..AC-18)", async ({
  page,
}) => {
  const originalId = await createDocument(page, "Original");
  await page.goto("/");

  await rowByTitle(page, "Original").getByTestId("row-menu").click();
  await page.getByTestId("menu-duplicate").click();

  // The copy exists alongside the original, newest first.
  const titles = page.getByTestId("row-title");
  await expect(titles).toHaveCount(2);
  await expect(titles.nth(0)).toHaveText("Original (copy)");
  await expect(titles.nth(1)).toHaveText("Original");

  // It opens as its own document under a different id.
  await rowByTitle(page, "Original (copy)").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  expect(page.url().endsWith(originalId)).toBe(false);
  await expect(page.getByTestId("document-title")).toHaveText(
    "Original (copy)",
  );
});

// SPEC-067 amends SPEC-003's delete AC: deleting from the library moves the
// document to the trash (with an Undo toast), and permanent deletion is a
// separate, explicit step from the Trash view that confirms once.
test("Delete moves to trash; permanent deletion is a separate confirmed step (AC-33, AC-36..39, AC-44)", async ({
  page,
}) => {
  await createDocument(page, "Doomed");
  await createDocument(page, "Survivor");
  await page.goto("/");

  // Delete moves to trash — no confirm dialog, an Undo toast instead.
  await rowByTitle(page, "Doomed").getByTestId("row-menu").click();
  await page.getByTestId("menu-delete").click();
  await expect(rowByTitle(page, "Doomed")).toBeHidden();
  await expect(page.getByTestId("toast")).toContainText("Doomed");

  // Trash lists it with Restore and Delete permanently.
  await page.getByTestId("trash-link").click();
  await expect(page.getByTestId("trash-view")).toContainText("Doomed");
  await page.getByTestId("row-menu").click();
  await page.getByTestId("menu-delete-permanently").click();
  await expect(page.getByTestId("delete-message")).toContainText("Doomed");
  await page.getByTestId("delete-confirm").click();
  await expect(page.getByTestId("trash-view")).not.toContainText("Doomed");

  // Survivor is untouched and stays after a reload.
  await page.getByTestId("back-to-library").click();
  await settleWrites(page);
  await page.reload();
  await expect(page.getByTestId("row-title")).toHaveText("Survivor");
});

test("the delete toast's Undo restores the document (AC-34, AC-35)", async ({
  page,
}) => {
  await createDocument(page, "Bring me back");
  await page.goto("/");

  await rowByTitle(page, "Bring me back").getByTestId("row-menu").click();
  await page.getByTestId("menu-delete").click();
  await expect(rowByTitle(page, "Bring me back")).toBeHidden();
  await page.getByTestId("toast-action").click();
  await expect(rowByTitle(page, "Bring me back")).toBeVisible();
});

test("deleting the open document moves it to trash, returns to / with a toast (AC-33, AC-19, AC-20)", async ({
  page,
}) => {
  await createDocument(page, "Open and doomed");

  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("delete-document").click();
  await expect(page.getByTestId("delete-message")).toContainText(
    "Open and doomed",
  );
  await page.getByTestId("delete-confirm").click();

  await expect(page).toHaveURL(/\/$/);
  await expect(page.getByTestId("toast")).toHaveText(
    "Deleted Open and doomed",
  );
  // It was the only document; it is now in the trash, not erased.
  await expect(page.getByTestId("library-empty")).toBeVisible();
  await page.getByTestId("trash-link").click();
  await expect(page.getByTestId("trash-view")).toContainText("Open and doomed");
});

test("without IndexedDB the app works in-memory and shows the banner (AC-21, AC-23)", async ({
  page,
}) => {
  await page.addInitScript(() => {
    Object.defineProperty(window, "indexedDB", {
      get: () => undefined,
      configurable: true,
    });
  });
  await page.goto("/");

  await expect(page.getByTestId("storage-warning")).toHaveText(
    "Changes will not be saved in this browser",
  );

  // Everything still works in-memory: create, rename, list. Navigate back
  // within the SPA (goBack = popstate, no reload) — a full reload would wipe
  // the in-memory store, which is exactly what "not saved" means here.
  await createDocument(page, "Ephemeral");
  await expect(page.getByTestId("storage-warning")).toBeVisible();
  // Browser Back out of a document that can't be saved (storage unavailable)
  // is a guarded exit (SPEC-060 AC-40/41): it warns before losing the
  // in-memory changes, so confirm "Close anyway" to return to the library.
  await page.goBack();
  await expect(page.getByTestId("close-without-saving-dialog")).toBeVisible();
  await page.getByTestId("close-without-saving-confirm").click();
  await expect(rowByTitle(page, "Ephemeral")).toBeVisible();
  await expect(page.getByTestId("storage-warning")).toBeVisible();
});

test("when writes fail (quota exceeded) the app works in-memory and shows the banner (AC-22, AC-23)", async ({
  page,
}) => {
  await page.addInitScript(() => {
    IDBObjectStore.prototype.put = () => {
      throw new DOMException("Quota exceeded", "QuotaExceededError");
    };
  });
  await page.goto("/");
  // Reads still work, so no banner until a write fails.
  await expect(page.getByTestId("storage-warning")).toBeHidden();

  const id = await createDocument(page, "Over quota");
  await expect(page.getByTestId("storage-warning")).toHaveText(
    "Changes will not be saved in this browser",
  );

  // The document remains fully usable in-memory (goBack stays in the SPA;
  // a full reload would wipe the unpersisted store). Because it can't be
  // saved, browser Back warns first (SPEC-060 AC-40/41) — "Close anyway"
  // returns to the library, where the in-memory record is still listed.
  await page.goBack();
  await expect(page.getByTestId("close-without-saving-dialog")).toBeVisible();
  await page.getByTestId("close-without-saving-confirm").click();
  await rowByTitle(page, "Over quota").click();
  await expect(page).toHaveURL(new RegExp(`/doc/${id}$`));
  await expect(page.getByTestId("document-title")).toHaveText("Over quota");
});
