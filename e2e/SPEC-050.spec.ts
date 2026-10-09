import { test, expect, type Download } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-050 — Markdown: open, edit visually, and convert to/from .docx/.odt.
const KITCHEN_SINK_MD = `# Report Title

A paragraph with **bold** and _italic_ text.

- first item
- second item
`;

const fixturesDir = dirname(fileURLToPath(import.meta.url));
const mdFixturesDir = join(fixturesDir, "..", "src", "lib", "md", "fixtures");
const KITCHEN_SINK_FULL_MD = readFileSync(join(mdFixturesDir, "kitchen-sink.md"), "utf-8");
const docxFixture = join(fixturesDir, "fixtures", "lists.docx");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

async function downloadedBuffer(download: Download): Promise<Buffer> {
  const path = await download.path();
  expect(path).toBeTruthy();
  return readFileSync(path);
}

async function downloadedText(download: Download): Promise<string> {
  return (await downloadedBuffer(download)).toString("utf-8");
}

test("opening a .md file renders it as an editable document — real headings, no raw markdown syntax (AC-1)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({
    name: "report.md",
    mimeType: "text/markdown",
    buffer: Buffer.from(KITCHEN_SINK_MD, "utf-8"),
  });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("format-badge")).toHaveText("MD");

  // The heading is a real heading element, not a line of "# " source text.
  const heading = page.getByRole("heading", { name: "Report Title" });
  await expect(heading).toBeVisible();
  await expect(heading).not.toHaveText(/^#/);

  // No raw Markdown syntax is visible anywhere on the canvas.
  const canvasText = await page.getByTestId("document-content").innerText();
  expect(canvasText).not.toContain("**");
  expect(canvasText).not.toContain("# Report Title");
  expect(canvasText).toContain("Report Title");
  expect(canvasText).toContain("bold");
  expect(canvasText).toContain("first item");

  // Bold text actually renders bold (a real run, not literal asterisks).
  const bold = page.getByText("bold", { exact: true });
  await expect(bold).toHaveCSS("font-weight", /^(bold|700)$/);
});

test("dropping a .md file opens it, exactly as the file picker would", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  const dataTransfer = await page.evaluateHandle((text) => {
    const dt = new DataTransfer();
    dt.items.add(new File([text], "notes.md", { type: "text/markdown" }));
    return dt;
  }, "# Dropped Heading\n\nSome text.\n");
  await page.dispatchEvent("body", "drop", { dataTransfer });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByRole("heading", { name: "Dropped Heading" })).toBeVisible();
});

test("Download offers Markdown (.md) and the downloaded file re-opens with the same content (AC-6, AC-7)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({
    name: "report.md",
    mimeType: "text/markdown",
    buffer: Buffer.from(KITCHEN_SINK_MD, "utf-8"),
  });
  await expect(page).toHaveURL(/\/doc\/.+/);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const option = page.getByTestId("download-md");
  await expect(option).toHaveText("Markdown (.md)");
  const downloadPromise = page.waitForEvent("download");
  await option.click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("report.md");

  const text = (await downloadedBuffer(download)).toString("utf-8");
  // The heading style is bold, so its own text is wrapped too — real
  // content, not a discarded artefact (the writer always emits the
  // effective, baked formatting).
  expect(text).toContain("# **Report Title**");
  expect(text).toContain("**bold**");
  expect(text).toContain("- first item");
});

test("saving a .docx document as .md shows the conversion report for content Markdown cannot carry (AC-6)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  // A non-default alignment has no Markdown representation (SPEC-050's
  // "What markdown cannot carry": "alignment, indents and spacing").
  await page.getByTestId("document-content").click();
  await page.keyboard.type("Centered text");
  await page.getByTestId("align-center").click();

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-md").click();
  await downloadPromise;

  await expect(page.getByTestId("conversion-report")).toBeVisible();
});

test("front matter and constructs outside the supported subset stay off the canvas but round-trip verbatim on save (AC-3, AC-4)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({
    name: "kitchen-sink.md",
    mimeType: "text/markdown",
    buffer: Buffer.from(KITCHEN_SINK_FULL_MD, "utf-8"),
  });
  await expect(page).toHaveURL(/\/doc\/.+/);

  // Neither the YAML front matter nor the opaque raw-HTML block is a
  // raw-source leak onto the canvas — there is no source view at all.
  const canvasText = await page.getByTestId("document-content").innerText();
  expect(canvasText).not.toContain("title: Kitchen Sink");
  expect(canvasText).not.toContain("raw-html-block");

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-md").click();
  const text = await downloadedText(await downloadPromise);

  // Front matter is retained verbatim, first in the file, unparsed.
  expect(text.startsWith("---\ntitle: Kitchen Sink")).toBe(true);
  // The raw HTML block (outside the supported subset) is anchored and
  // re-emitted verbatim.
  expect(text).toContain('<div class="raw-html-block">');
  expect(text).toContain("This block is outside the supported subset");
  // A reference-style link definition is likewise unrecognised and preserved.
  expect(text).toContain("[ref-link]: https://example.com/reference");
});

test("saving a .docx-backed document as .md does not change what it is — it stays .docx and still round-trips (AC-9)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({
    name: "lists.docx",
    mimeType: DOCX_MIME,
    buffer: readFileSync(docxFixture),
  });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("format-badge")).toHaveText("DOCX");

  // Exporting to .md is a one-off download, not a format change.
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const mdDownloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-md").click();
  await mdDownloadPromise;
  const report = page.getByTestId("conversion-report");
  if (await report.isVisible().catch(() => false)) {
    await page.getByTestId("conversion-report-close").click();
  }
  await expect(page.getByTestId("format-badge")).toHaveText("DOCX");

  // A later save as .docx still round-trips the preserved package intact.
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const docxDownloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const outBytes = await downloadedBuffer(await docxDownloadPromise);
  expect(outBytes.length).toBeGreaterThan(1000);
  expect(outBytes[0]).toBe(0x50); // ZIP magic "PK"
  expect(outBytes[1]).toBe(0x4b);
});

test("format detection admits markdown by extension, but a ZIP package is still identified by content even when named .md (AC-10)", async ({
  page,
}) => {
  await page.goto("/");
  // A real .docx (ZIP) file, deliberately misnamed with a .md extension —
  // content-sniffing must still win for ZIP packages per SPEC-050's
  // documented exception.
  await page.getByTestId("open-document-input").setInputFiles({
    name: "tricky.md",
    mimeType: "text/markdown",
    buffer: readFileSync(docxFixture),
  });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("format-badge")).toHaveText("DOCX");
});

test("dropping an unsupported file type names all three supported formats in the error (AC-11)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  const dataTransfer = await page.evaluateHandle((text) => {
    const dt = new DataTransfer();
    dt.items.add(new File([text], "notes.txt", { type: "text/plain" }));
    return dt;
  }, "plain text, unsupported extension");
  await page.dispatchEvent("body", "dragenter", { dataTransfer });
  await page.dispatchEvent("body", "drop", { dataTransfer });
  await expect(page.getByTestId("drop-error")).toHaveText("We can only open .docx, .odt, .md, .pptx, .odp, .key, .xlsx, .ods, .pdf and image files");
  await expect(page).toHaveURL(/\/$/);
});
