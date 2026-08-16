import { test, expect, type Page } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { writeFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { MOD } from "./keys";
import { createDocument as buildDocumentModel, createFieldSpecRun } from "../src/lib/documentModel";
import type { Block, Paragraph } from "../src/lib/documentModel";
import { writeDocx } from "../src/lib/docx/writer";

// SPEC-053 — the unified field system. Unit/component tests cover the evaluator,
// the readers/writers, the dialogs and the field commands in isolation; this
// file exercises the seams only a real browser can confirm and that a prior QA
// round found broken: inserting a field through the Date and time dialog (the
// selection the handler reads was always null), a field being a clickable,
// shadeable one-character atom in the actual canvas, fields showing their
// SAVED cached values on open (never a live recompute), F9 / Update-all
// recalculation, cross-reference insertion, and a page field in a header
// resolving per page.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function openFixture(page: Page, name: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(join(fixturesDir, name));
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const fieldAtoms = (page: Page) => page.locator('[data-testid="field"]');

async function openMenu(page: Page, menuId: string): Promise<void> {
  await page.getByTestId(`menu-title-${menuId}`).click();
}

async function insertDateField(page: Page): Promise<void> {
  await openMenu(page, "insert");
  await page.getByTestId("insert-date-time-button").click();
  await expect(page.getByTestId("date-time-dialog")).toBeVisible();
  await page.getByTestId("date-time-format-d MMMM yyyy").check();
  await page.getByTestId("date-time-update-automatically").check();
  await page.getByTestId("date-time-insert").click();
  await expect(page.getByTestId("date-time-dialog")).not.toBeVisible();
}

test("Insert → Date and time inserts plain text with the toggle off and a date field with it on (AC-4, AC-5)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Signed ");

  // Toggle off → plain text, never a field.
  await openMenu(page, "insert");
  await page.getByTestId("insert-date-time-button").click();
  await expect(page.getByTestId("date-time-dialog")).toBeVisible();
  await page.getByTestId("date-time-format-d MMMM yyyy").check();
  await page.getByTestId("date-time-insert").click();
  await expect(page.getByTestId("date-time-dialog")).not.toBeVisible();
  await expect(fieldAtoms(page)).toHaveCount(0);
  await expect(content(page)).toContainText("Signed");

  // Toggle on → exactly one date field atom is inserted.
  await insertDateField(page);
  await expect(fieldAtoms(page)).toHaveCount(1);
});

test("clicking a field selects its whole atom — it shades and shows the toolbar chip — and typing over it replaces it (AC-7, AC-8, AC-10, AC-14)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await insertDateField(page);

  const field = fieldAtoms(page).first();
  await expect(field).toHaveCount(1);

  // A plain click lands the selection on the field: it shades (AC-10) and the
  // field toolbar chip appears (AC-14) — both were broken when the caret never
  // reached the field.
  await field.click();
  await expect(field).toHaveClass(/doc-field-shaded/);
  await expect(page.getByTestId("field-toolbar")).toBeVisible();

  // Typing over the selected field replaces the whole field (AC-8).
  await page.keyboard.type("Z");
  await expect(fieldAtoms(page)).toHaveCount(0);
  await expect(content(page)).toContainText("Z");
});

test("View → Field shading shades a field regardless of the caret (AC-11)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Draft ");
  await insertDateField(page);

  // Move the caret well away from the field so any shading is the View toggle's
  // doing, not caret proximity.
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.type(" copy");
  const field = fieldAtoms(page).first();
  await expect(field).not.toHaveClass(/doc-field-shaded/);

  await openMenu(page, "view");
  await page.getByTestId("view-menu-field-shading").click();
  await expect(field).toHaveClass(/doc-field-shaded/);
});

// The word-fields.docx fixture ships one field of every catalogue type at a
// SAVED value that differs from any value the clock could produce today. On
// open, that saved value must show verbatim — a document open must never
// recalculate a field.
test("opening a document shows saved cached values, never a live recompute (AC-19, AC-20, AC-51)", async ({
  page,
}) => {
  await openFixture(page, "word-fields.docx");
  await expect(content(page)).toBeVisible();
  await expect(content(page)).toContainText("4 March 2026"); // saved date field
  await expect(content(page)).toContainText("14:30"); // saved time field
  const body = await content(page).innerText();
  expect(body).not.toContain("16 August 2026"); // today's date must not appear
});

// F9 recalculates the field at the caret; Tools → Update all recalculates every
// unlocked field, leaves locked ones alone, and reports deleted-target refs.
test("F9 updates the field at the caret and Update-all updates every unlocked field (AC-15, AC-16, AC-28)", async ({
  page,
}) => {
  await openFixture(page, "word-fields.docx");
  const dateField = fieldAtoms(page).filter({ hasText: "4 March 2026" }).first();
  await expect(dateField).toHaveCount(1);
  await dateField.click();
  await expect(page.getByTestId("field-toolbar")).toBeVisible();
  await page.keyboard.press("F9");
  await expect(content(page)).toContainText("16 August 2026"); // AC-15 recomputed to today

  // Update-all: the deleted-target cross-reference is reported, not fatal.
  await openMenu(page, "tools");
  await page.getByTestId("tools-menu-update-all-fields").click();
  await expect(page.getByText(/point to a deleted target/i)).toBeVisible(); // AC-28
  await expect(content(page)).toContainText("1 March 2025"); // AC-16 locked field untouched
});

// Insert → Cross-reference inserts a reference field for each "show" mode. The
// handler read a null selection before the fix, so nothing was inserted.
test("Insert → Cross-reference inserts text / page / position references (AC-23, AC-24, AC-25)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Introduction");
  await page.keyboard.press(`${MOD}+Alt+1`);
  await page.keyboard.press("Enter");
  await page.keyboard.press(`${MOD}+Alt+0`);
  await page.keyboard.type("See ");

  async function insertRef(show: string): Promise<void> {
    await openMenu(page, "insert");
    await page.getByTestId("insert-cross-reference-button").click();
    await expect(page.getByTestId("cross-reference-dialog")).toBeVisible();
    await page.getByTestId("cross-reference-show").selectOption(show);
    await page.getByTestId("cross-reference-heading-0").check();
    await page.getByTestId("cross-reference-insert").click();
    await expect(page.getByTestId("cross-reference-dialog")).not.toBeVisible();
  }

  await insertRef("text");
  await expect(fieldAtoms(page)).toHaveCount(1);
  await expect(content(page)).toContainText("Introduction");
  await page.keyboard.press(`${MOD}+End`);
  await insertRef("page");
  await expect(fieldAtoms(page)).toHaveCount(2);
  await page.keyboard.press(`${MOD}+End`);
  await insertRef("position");
  await expect(fieldAtoms(page)).toHaveCount(3);
});

// A page field in a header resolves to each page's own number. No UI inserts a
// field into a header, so build a .docx with a page-field header and a body
// long enough to span several pages, then read each page's header band.
test("a page field in a header shows the correct value on every page (AC-35)", async ({ page }) => {
  const model = buildDocumentModel("spec053-page-header");
  const body: Block[] = [];
  for (let i = 0; i < 140; i++) {
    body.push({
      type: "paragraph",
      runs: [{ text: `Body line ${i} lorem ipsum dolor sit amet consectetur.`, formatting: {} }],
    });
  }
  const headerPara: Paragraph = { type: "paragraph", runs: [createFieldSpecRun({ kind: "page", format: "decimal" })] };
  model.sections = [{ blocks: body, headers: { default: { blocks: [headerPara] } } }];
  model.title = "spec053-page-header.docx";
  const bytes = (await writeDocx(model)).bytes;
  const file = join(mkdtempSync(join(tmpdir(), "spec053-")), "spec053-page-header.docx");
  writeFileSync(file, bytes);

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page)).toBeVisible();

  const indicator = await page.getByTestId("page-indicator").textContent();
  const total = Number((indicator ?? "").match(/of\s+(\d+)/)?.[1] ?? "1");
  expect(total).toBeGreaterThanOrEqual(2);

  const bands = page.getByTestId("page-header-band");
  const values: string[] = [];
  for (let i = 0; i < (await bands.count()); i++) {
    const atom = bands.nth(i).locator('[data-testid="field"]').first();
    if (await atom.count()) values.push(((await atom.textContent()) ?? "").trim());
  }
  expect(values).toEqual(Array.from({ length: total }, (_, i) => String(i + 1)));
});
