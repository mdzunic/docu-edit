import { test, expect } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-004 — open .docx: text and paragraphs.
const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "src", "lib", "docx", "fixtures");
const wordFixture = join(fixturesDir, "word-authored.docx");

test("'Open document' opens a file picker limited to supported formats (AC-1)", async ({ page }) => {
  await page.goto("/");
  const input = page.getByTestId("open-document-input");
  // SPEC-004 shipped this as ".docx"; SPEC-005's AC-1 ("the file picker
  // accepts .odt alongside .docx"), SPEC-050's AC-1 ("opening a .md
  // or .markdown file — from the library, the file picker, or a drop") and
  // SPEC-107's AC-1 (pdf and the five image formats join OPENABLE_EXTENSIONS)
  // supersede that exact value. The picker must still be limited to formats
  // the app can open.
  await expect(input).toHaveAttribute(
    "accept",
    ".docx,.odt,.md,.markdown,.pptx,.odp,.pdf,.png,.jpg,.jpeg,.gif,.svg,.webp",
  );
});

test("choosing a .docx file parses it and opens it at /doc/:id (AC-2)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(wordFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByRole("heading", { name: "Quarterly Report" })).toBeVisible();
});

test("the app bar title comes from the file name without extension (AC-29)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(wordFixture);
  await expect(page.getByTestId("document-title")).toHaveText("word-authored");
});

test("the format badge in the app bar reads DOCX (AC-30)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(wordFixture);
  await expect(page.getByTestId("format-badge")).toHaveText("DOCX");
});

test("opened content renders headings and direct formatting on the canvas", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(wordFixture);
  await expect(page.getByRole("heading", { name: "Quarterly Report", level: 1 })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Section Two", level: 2 })).toBeVisible();
  await expect(page.getByText("bold", { exact: true })).toBeVisible();
});

test("a file that is not a ZIP shows the specific error and stays on the home page (AC-31, AC-33)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({
    name: "broken.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("this is not a zip file"),
  });
  await expect(page.getByTestId("open-document-error")).toHaveText(
    "This file isn't a Word document we can open",
  );
  await expect(page).toHaveURL(/\/$/);
  // The library redesign dropped the "No document" placeholder: the library
  // is not a document, so no title region renders there at all (see
  // SPEC-001). No title region still proves no document was opened.
  await expect(page.getByTestId("document-title")).toHaveCount(0);
});

test("a .doc (binary Word 97) file is rejected with the legacy-format error (AC-34)", async ({ page }) => {
  await page.goto("/");
  const docMagic = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);
  await page.getByTestId("open-document-input").setInputFiles({
    name: "legacy.docx",
    mimeType: "application/msword",
    buffer: docMagic,
  });
  await expect(page.getByTestId("open-document-error")).toHaveText(
    "Old .doc files aren't supported — save as .docx first",
  );
  await expect(page).toHaveURL(/\/$/);
});

test("rejecting an invalid file leaves a previously open document untouched (AC-33)", async ({ page }) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill("Keep me");
  await page.getByTestId("document-title-input").press("Enter");
  await expect(page.getByTestId("document-title")).toHaveText("Keep me");

  // Client-side navigation (History API, no full reload) back to the home
  // page, so the in-memory document from the "New document" click above
  // survives — a full page reload would trivially "reset" it either way.
  // Home shows the library rather than the empty state here, because the
  // document created above is already in it.
  await page.goBack();
  await expect(page.getByTestId("library")).toBeVisible();

  await page.getByTestId("open-document-input").setInputFiles({
    name: "broken.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: Buffer.from("not a zip"),
  });
  await expect(page.getByTestId("open-document-error")).toBeVisible();

  await page.goForward();
  await expect(page.getByTestId("document-title")).toHaveText("Keep me");
});
