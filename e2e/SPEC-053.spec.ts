import { test, expect, type Page } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { writeFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { MOD } from "./keys";
import { createDocument as buildDocumentModel, createFieldSpecRun } from "../src/lib/documentModel";
import type { Block, Paragraph } from "../src/lib/documentModel";
import { writeDocx } from "../src/lib/docx/writer";
import { renderDateTime } from "../src/lib/fields";
import { DEFAULT_LANGUAGE } from "../src/lib/core/language";

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

/** Today under the fixture's date-field pattern, rendered by the same catalogue
 * function the editor recomputes through (src/lib/fields.ts). Both assertions
 * about "today" below used to spell out the date this spec was written on
 * (16 August 2026), so at the next midnight main went red on AC-15 while the
 * AC-19 check that today's date must NOT appear quietly became vacuous — it
 * was naming a date no clock would ever produce again. */
function todayLong(): string {
  const rendered = renderDateTime("d MMMM yyyy", new Date(), DEFAULT_LANGUAGE);
  // A null here would make the assertions below trivially true.
  expect(rendered).not.toBeNull();
  return rendered ?? "";
}

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
  expect(body).not.toContain(todayLong()); // today's date must not appear
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
  await expect(content(page)).toContainText(todayLong()); // AC-15 recomputed to today
  // The saved value it replaced is gone, so an F9 that did nothing fails here
  // even on a day whose date happens to appear elsewhere in the fixture.
  await expect(content(page)).not.toContainText("4 March 2026");

  // Update-all: the deleted-target cross-reference is reported, not fatal.
  await openMenu(page, "tools");
  await page.getByTestId("tools-menu-update-all-fields").click();
  await expect(page.getByText(/point to a deleted target/i)).toBeVisible(); // AC-28
  await expect(content(page)).toContainText("1 March 2025"); // AC-16 locked field untouched
});

// F9 with the caret in a header band. The band routes its keydown to the
// header/footer handler and stops it before the canvas shortcut registry, so
// F9 used to be inert there and Tools → Update all fields was a header field's
// only refresh — AC-15 puts no such limit on where the caret is, and AC-33 puts
// fields in headers. No UI inserts a field into a header, so the fixture is
// built here, the same way the AC-35 test builds its page-field header.
test("F9 updates the field at the caret inside a header (AC-15, AC-33)", async ({ page }) => {
  const model = buildDocumentModel("spec053-header-f9");
  const headerPara: Paragraph = {
    type: "paragraph",
    runs: [createFieldSpecRun({ kind: "date", format: "d MMMM yyyy", cached: "1 January 2020" })],
  };
  // A body date field saved at the same stale value: it must still be stale
  // afterwards, which is what proves F9 acted on the header caret rather than
  // on the body caret the registry path would have read.
  const bodyPara: Paragraph = {
    type: "paragraph",
    runs: [
      { text: "Signed ", formatting: {} },
      createFieldSpecRun({ kind: "date", format: "d MMMM yyyy", cached: "1 January 2020" }),
    ],
  };
  model.sections = [{ blocks: [bodyPara], headers: { default: { blocks: [headerPara] } } }];
  model.title = "spec053-header-f9.docx";
  const bytes = (await writeDocx(model)).bytes;
  const file = join(mkdtempSync(join(tmpdir(), "spec053-")), "spec053-header-f9.docx");
  writeFileSync(file, bytes);

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page)).toBeVisible();

  const band = page.getByTestId("page-header-band").first();
  await expect(band).toContainText("1 January 2020");

  // Enter header editing: the caret lands at the start of the header's first
  // block, which is the field atom's leading edge — "the field at the caret".
  const box = await band.boundingBox();
  if (!box) {
    throw new Error("header band has no bounding box");
  }
  await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
  await expect(band).toHaveAttribute("data-hf-editable", "1");

  await page.keyboard.press("F9");
  await expect(band).toContainText(todayLong());
  await expect(band).not.toContainText("1 January 2020");
  // The body field was not at the caret, so it keeps its saved value.
  await expect(content(page)).toContainText("Signed 1 January 2020");
});

/** A stale date field, saved at a value no clock will produce again — the
 * "before" every F9 check below asserts against. */
function staleDateRun() {
  return createFieldSpecRun({ kind: "date", format: "d MMMM yyyy", cached: "1 January 2020" });
}

/** Writes `model` as a .docx in a fresh temp dir and returns its path — the
 * surface fixtures below all need a saved file, since no UI inserts a field
 * into a note or a shape's text frame. */
async function writeFixture(model: Parameters<typeof writeDocx>[0], name: string): Promise<string> {
  model.title = `${name}.docx`;
  const bytes = (await writeDocx(model)).bytes;
  const file = join(mkdtempSync(join(tmpdir(), "spec053-")), `${name}.docx`);
  writeFileSync(file, bytes);
  return file;
}

// F9 with the caret in a footnote body — the header band's defect exactly:
// DocPage routes the note's keydown to the note handler and stops it before the
// canvas shortcut registry, so F9 was inert there too. AC-15 puts no limit on
// where the caret is and AC-33 lists footnotes among the surfaces fields work
// in.
test("F9 updates the field at the caret inside a footnote (AC-15, AC-33)", async ({ page }) => {
  const model = buildDocumentModel("spec053-note-f9");
  // The note's content hangs off the reference mark in the body: the note area
  // is derived from the marks, so without this run the note never renders.
  const refPara: Paragraph = {
    type: "paragraph",
    runs: [
      { text: "Body text", formatting: {} },
      { text: "", formatting: {}, note: { noteId: "n1", kind: "footnote" } },
    ],
  };
  // A body date field saved at the same stale value: it must still be stale
  // afterwards, which is what proves F9 acted on the note caret rather than on
  // the body caret the registry path would have read.
  const bodyPara: Paragraph = {
    type: "paragraph",
    runs: [{ text: "Signed ", formatting: {} }, staleDateRun()],
  };
  model.sections = [{ blocks: [refPara, bodyPara] }];
  model.notes = [
    { noteId: "n1", kind: "footnote", blocks: [{ type: "paragraph", runs: [staleDateRun()] }] },
  ];
  const file = await writeFixture(model, "spec053-note-f9");

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page)).toBeVisible();

  const noteBody = page.getByTestId("note-body").first();
  await expect(noteBody).toContainText("1 January 2020");

  // Clicking a read-only note body enters editing with the caret at the end of
  // its last block — for a lone field atom that is the atom's trailing edge,
  // which is still "the field at the caret".
  await noteBody.click();
  await expect(noteBody).toHaveAttribute("data-note-editable", "1");

  await page.keyboard.press("F9");
  await expect(noteBody).toContainText(todayLong());
  await expect(noteBody).not.toContainText("1 January 2020");
  // The body field was not at the caret, so it keeps its saved value.
  await expect(content(page)).toContainText("Signed 1 January 2020");
});

// F9 with the caret in a shape's text frame — the third surface with the same
// defect (DocShapeBlock stops the frame's keydown before the canvas registry).
// AC-33 lists shape text boxes.
test("F9 updates the field at the caret inside a shape text box (AC-15, AC-33)", async ({ page }) => {
  const model = buildDocumentModel("spec053-shape-f9");
  const bodyPara: Paragraph = {
    type: "paragraph",
    runs: [{ text: "Signed ", formatting: {} }, staleDateRun()],
  };
  const textBox: Block = {
    type: "shape",
    geometry: "rect",
    widthCm: 8,
    heightCm: 3,
    wrap: "square",
    text: { blocks: [{ type: "paragraph", runs: [staleDateRun()] }], vAlign: "top" },
  };
  model.sections = [{ blocks: [bodyPara, textBox] }];
  const file = await writeFixture(model, "spec053-shape-f9");

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page)).toBeVisible();

  const frame = content(page).locator('[data-testid="drawing-0"] [data-testid="shape-text"]');
  await expect(frame).toContainText("1 January 2020");

  // Double-clicking the shape enters text editing with the caret at the end of
  // the frame's last block — the field atom's trailing edge (AC-17).
  await content(page).locator('[data-testid="drawing-0"]').dblclick();
  await expect(frame).toHaveAttribute("contenteditable", "true");

  await page.keyboard.press("F9");
  await expect(frame).toContainText(todayLong());
  await expect(frame).not.toContainText("1 January 2020");
  // The body field was not at the caret, so it keeps its saved value.
  await expect(content(page)).toContainText("Signed 1 January 2020");
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

  // Pagination is deferred off the idle callback, so this used to sample
  // "of 1" before the second page existed — a one-shot textContent() read has
  // no retry, and the test failed roughly one run in three. Wait for the
  // indicator to settle on a multi-page count, then read the settled text.
  const indicator = page.getByTestId("page-indicator");
  await expect(indicator).toContainText(/of\s+([2-9]|\d\d+)/);
  const total = Number(((await indicator.textContent()) ?? "").match(/of\s+(\d+)/)?.[1] ?? "1");
  expect(total).toBeGreaterThanOrEqual(2);

  const bands = page.getByTestId("page-header-band");
  const expected = Array.from({ length: total }, (_, i) => String(i + 1));
  // The bands paint page by page as pagination lands, so the indicator can
  // already read "of 4" while only three bands exist. Collect them under a
  // retrying assertion; the expectation itself is unchanged.
  await expect(async () => {
    const values: string[] = [];
    for (let i = 0; i < (await bands.count()); i++) {
      const atom = bands.nth(i).locator('[data-testid="field"]').first();
      if (await atom.count()) values.push(((await atom.textContent()) ?? "").trim());
    }
    expect(values).toEqual(expected);
  }).toPass({ timeout: 10_000 });
});
