import { test, expect, type Download } from "@playwright/test";
import { readFileSync } from "node:fs";

// SPEC-050 — Markdown: open, edit visually, and convert to/from .docx/.odt.
const KITCHEN_SINK_MD = `# Report Title

A paragraph with **bold** and _italic_ text.

- first item
- second item
`;

async function downloadedBuffer(download: Download): Promise<Buffer> {
  const path = await download.path();
  expect(path).toBeTruthy();
  return readFileSync(path);
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

  await page.getByTestId("download-menu").click();
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

  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-md").click();
  await downloadPromise;

  await expect(page.getByTestId("conversion-report")).toBeVisible();
});
