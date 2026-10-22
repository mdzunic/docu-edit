import { test, expect, type Page } from "@playwright/test";
import { copyFileSync, writeFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import { MOD } from "./keys";

// SPEC-046 — template gallery and mail merge. Unit/component tests cover the
// engine (src/lib/merge.ts, mergeChips.ts) and every dialog in isolation
// (src/components/*.spec.ts); this file exercises the seams only a real
// browser can confirm: a merge chip renders as one caret atom in the actual
// canvas (the exact defect a prior review round flagged), the template
// gallery/import/save flow through real navigation, and a real merge run
// producing a real download.

const here = dirname(fileURLToPath(import.meta.url));
const docxFixture = join(here, "..", "src", "lib", "docx", "fixtures", "word-authored.docx");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const chips = (page: Page) => page.locator(".doc-merge-chip");

async function openMenu(page: Page, menuId: string): Promise<void> {
  await page.getByTestId(`menu-title-${menuId}`).click();
}

test("New from template offers the six shipped templates alongside Blank document, and opening one produces an editable document (AC-1..4)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("new-from-template").click();
  await expect(page.getByTestId("template-gallery-dialog")).toBeVisible();
  await expect(page.getByTestId("template-card-blank")).toBeVisible();
  const cards = page.getByTestId("template-card");
  await expect(cards).toHaveCount(6);
  const ids = await cards.evaluateAll((els) => els.map((el) => el.getAttribute("data-template-id")));
  expect(ids.sort()).toEqual(["cv", "invoice", "letter", "meeting-minutes", "memo", "report"].sort());

  await page.locator('[data-template-id="invoice"]').click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page)).toContainText("Invoice");
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.type(" — edited");
  await expect(content(page)).toContainText("edited");
});

test("a merge field chip renders as {{name}}, is a single caret atom, and survives a .docx round trip (AC-12..17)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Dear ");

  await openMenu(page, "insert");
  await page.getByTestId("insert-menu-merge-field").click();
  await expect(page.getByTestId("merge-field-picker-dialog")).toBeVisible();
  await page.getByTestId("merge-field-picker-custom-input").fill("client_name");
  await page.getByTestId("merge-field-picker-custom-insert").click();
  await expect(page.getByTestId("merge-field-picker-dialog")).not.toBeVisible();

  await expect(chips(page)).toHaveCount(1);
  await expect(chips(page).first()).toHaveText("{{client_name}}");

  // One Backspace removes the whole chip (AC-14: a field run and a single
  // caret atom) — place the caret right after it deterministically, the
  // same technique SPEC-031's note-reference test uses.
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press("Backspace");
  await expect(chips(page)).toHaveCount(0);
  await expect(content(page)).toContainText("Dear");

  // Re-insert and round-trip through .docx.
  await openMenu(page, "insert");
  await page.getByTestId("insert-menu-merge-field").click();
  await page.getByTestId("merge-field-picker-custom-input").fill("client_name");
  await page.getByTestId("merge-field-picker-custom-insert").click();
  await expect(chips(page)).toHaveCount(1);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const download = await downloadPromise;
  const path = await download.path();
  expect(path).toBeTruthy();

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path!);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(chips(page)).toHaveCount(1);
  await expect(chips(page).first()).toHaveText("{{client_name}}");
});

test("a merge chip inside a table cell renders as the same {{name}} pill as in the body, never the raw placeholder (AC-13, AC-29)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Body: ");
  await openMenu(page, "insert");
  await page.getByTestId("insert-menu-merge-field").click();
  await page.getByTestId("merge-field-picker-custom-input").fill("client_name");
  await page.getByTestId("merge-field-picker-custom-insert").click();
  await expect(chips(page)).toHaveCount(1);

  // A table, then the same insertion inside a cell. The cell renderer paints
  // runs itself (DocTableBlock, not usePaginationView's precomputed entries),
  // which is exactly where a QA round caught chips surfacing as the raw
  // one-character field placeholder "1" instead of their {{...}} label.
  await openMenu(page, "insert");
  await page.getByTestId("table-grid-cell-2-2").click();
  await expect(content(page).locator("table.doc-table")).toHaveCount(1);

  const firstCell = page.getByTestId("table-cell-0-0");
  await firstCell.locator(".doc-paragraph").first().click();
  await openMenu(page, "insert");
  await page.getByTestId("insert-menu-merge-field").click();
  await page.getByTestId("merge-field-picker-custom-input").fill("sku");
  await page.getByTestId("merge-field-picker-custom-insert").click();

  await expect(firstCell.locator(".doc-merge-chip")).toHaveText(["{{sku}}"]);
  // The paragraph's whole rendered text is the label — nothing of the model
  // placeholder character leaks into the cell.
  await expect(firstCell.locator(".doc-paragraph").first()).toHaveText("{{sku}}");

  // "Repeat row for each…" (AC-29) plants an {{#each …}} chip at the start
  // of the row's first cell — the repeating-row authoring surface lives in
  // cells, so it must render as a chip there too.
  await firstCell.locator(".doc-paragraph").first().click({ button: "right" });
  await page.getByTestId("context-menu-repeat-row").click();
  await page.getByTestId("merge-field-picker-custom-input").fill("line_items");
  await page.getByTestId("merge-field-picker-custom-insert").click();

  await expect(firstCell.locator(".doc-merge-chip")).toHaveText(["{{#each line_items}}", "{{sku}}"]);
});

test("Save as template stores the document, it appears in the gallery, and can be renamed and deleted (AC-5..8)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("A template body");

  await openMenu(page, "file");
  await page.getByTestId("file-menu-save-as-template").click();
  await expect(page.getByTestId("rename-dialog")).toBeVisible();
  await page.getByTestId("rename-input").fill("My E2E Template");
  await page.getByTestId("rename-save").click();
  await expect(page.getByTestId("toast")).toContainText("My E2E Template");

  await page.goto("/");
  await page.getByTestId("new-from-template").click();
  const userCard = page.getByTestId("user-template-card").filter({ hasText: "My E2E Template" });
  await expect(userCard).toBeVisible();

  await userCard.getByTestId("user-template-rename").click();
  await page.getByTestId("rename-input").fill("Renamed E2E Template");
  await page.getByTestId("rename-save").click();
  await expect(page.getByTestId("template-gallery-dialog")).toContainText("Renamed E2E Template");

  const renamedCard = page.getByTestId("user-template-card").filter({ hasText: "Renamed E2E Template" });
  await renamedCard.getByTestId("user-template-delete").click();
  await expect(page.getByTestId("delete-dialog")).toBeVisible();
  await page.getByTestId("delete-confirm").click();
  await expect(page.getByTestId("user-template-card")).toHaveCount(0);
});

test("opening a .dotx creates a new untitled document, not the same document reopened (AC-9, AC-11)", async ({
  page,
}) => {
  const dotxPath = join(tmpdir(), `${randomUUID()}-letter.dotx`);
  copyFileSync(docxFixture, dotxPath);

  await page.goto("/");
  await page.getByTestId("new-from-template").click();
  await page.getByTestId("template-gallery-import-input").setInputFiles(dotxPath);
  await expect(page).toHaveURL(/\/doc\/.+/);
  const firstId = page.url().split("/doc/")[1];

  await page.goto("/");
  await page.getByTestId("new-from-template").click();
  await page.getByTestId("template-gallery-import-input").setInputFiles(dotxPath);
  await expect(page).toHaveURL(/\/doc\/.+/);
  const secondId = page.url().split("/doc/")[1];

  expect(firstId).not.toBe(secondId);
});

test("running a mail merge downloads a .zip built from an attached CSV (AC-21, AC-40, AC-41)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Dear ");
  await openMenu(page, "insert");
  await page.getByTestId("insert-menu-merge-field").click();
  await page.getByTestId("merge-field-picker-custom-input").fill("client_name");
  await page.getByTestId("merge-field-picker-custom-insert").click();

  const csvPath = join(tmpdir(), `${randomUUID()}-records.csv`);
  writeFileSync(csvPath, "client_name\nAcme\nBeta\n");

  await openMenu(page, "file");
  await page.getByTestId("file-menu-mail-merge").click();
  await expect(page.getByTestId("merge-run-dialog")).toBeVisible();

  await page.getByTestId("merge-run-attach-input").setInputFiles(csvPath);
  await expect(page.getByTestId("merge-run-record-count")).toContainText("2 records");

  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("merge-run-run").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toMatch(/-merge\.zip$/);
  await expect(page.getByTestId("merge-run-report-summary")).toContainText("2 records merged");
});
