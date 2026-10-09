import { test, expect } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-005 — open .odt: text and paragraphs.
const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "src", "lib", "odf", "fixtures");
const writerFixture = join(fixturesDir, "libreoffice-writer.odt");

test("'Open document' accepts .odt alongside .docx (AC-1)", async ({ page }) => {
  await page.goto("/");
  const input = page.getByTestId("open-document-input");
  // SPEC-050's AC-1 adds .md/.markdown, SPEC-107's AC-1 adds .pdf and the
  // image extensions and SPEC-116's AC-1 adds the workbook extensions to the
  // same picker; .odt must still be there alongside .docx, which is what this
  // AC is about.
  await expect(input).toHaveAttribute(
    "accept",
    ".docx,.odt,.md,.markdown,.pptx,.odp,.key,.xlsx,.xlsm,.xltx,.xltm,.ods,.ots,.pdf,.png,.jpg,.jpeg,.gif,.svg,.webp",
  );
});

test("choosing a .odt file parses it and opens it at /doc/:id (AC-3, AC-5, AC-6)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(writerFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByRole("heading", { name: "Project Plan", level: 1 })).toBeVisible();
  await expect(page.getByRole("heading", { name: "Schedule", level: 2 })).toBeVisible();
  // "underlined"/"struck" appear exactly once; "bold" also occurs (trimmed)
  // in the nested-spans paragraph and would violate strict mode.
  await expect(page.getByText("underlined", { exact: true })).toBeVisible();
  await expect(page.getByText("struck", { exact: true })).toBeVisible();
});

test("the format badge in the app bar reads ODT (AC-28)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(writerFixture);
  await expect(page.getByTestId("format-badge")).toHaveText("ODT");
});

test("format comes from the package mimetype, not the extension (AC-2)", async ({ page }) => {
  // The same ODF package bytes under a .docx name still open as ODT.
  const { readFileSync } = await import("node:fs");
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({
    name: "renamed.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: readFileSync(writerFixture),
  });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("format-badge")).toHaveText("ODT");
});

test("a flat XML ODF (.fodt) file is rejected with the specific message and stays on the home page (AC-29)", async ({ page }) => {
  const flatOdf = `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"/>`;
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({
    name: "flat.fodt",
    mimeType: "application/vnd.oasis.opendocument.text",
    buffer: Buffer.from(flatOdf),
  });
  await expect(page.getByTestId("open-document-error")).toHaveText(
    "Flat ODF files aren't supported yet",
  );
  await expect(page).toHaveURL(/\/$/);
  // The library redesign dropped the "No document" placeholder: the library
  // is not a document, so no title region renders there at all (see
  // SPEC-001). No title region still proves no document was opened.
  await expect(page.getByTestId("document-title")).toHaveCount(0);
});
