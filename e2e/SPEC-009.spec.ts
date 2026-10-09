import { test, expect } from "@playwright/test";
import { readFileSync, writeFileSync, mkdtempSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";

// SPEC-009 — drag-and-drop and open-from-URL.
const docxFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "src", "lib", "docx", "fixtures");
const odfFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "src", "lib", "odf", "fixtures");
const wordFixture = join(docxFixturesDir, "word-authored.docx");
const odtFixture = join(odfFixturesDir, "libreoffice-writer.odt");
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const ODT_MIME = "application/vnd.oasis.opendocument.text";

/** Navigates to /open?url=… and confirms the fetch. The confirmation panel
 * is the whole point of the route (a link must not be able to fetch and
 * store a document on its own), so every URL-open test goes through it. */
async function openFromUrl(page: import("@playwright/test").Page, url: string) {
  await page.goto("/open?url=" + encodeURIComponent(url));
  await expect(page.getByTestId("open-from-url-confirm")).toBeVisible();
  await page.getByTestId("open-from-url-confirm-button").click();
}

/** Builds a real DataTransfer holding real File objects inside the page —
 * jsdom can't do this (no DragEvent/DataTransfer), but a live browser can,
 * and it's the standard way Playwright recommends testing drag-and-drop
 * file uploads (https://playwright.dev/docs/input#drag-and-drop). */
async function dataTransferWithFiles(page: import("@playwright/test").Page, files: { name: string; mime: string; base64: string }[]) {
  return page.evaluateHandle((items) => {
    const dt = new DataTransfer();
    for (const item of items) {
      const bytes = Uint8Array.from(atob(item.base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], item.name, { type: item.mime }));
    }
    return dt;
  }, files);
}

function toBase64(path: string): string {
  return readFileSync(path).toString("base64");
}

// `page.goto` resolves on navigation, not once Vue has mounted and
// registered its window-level drag listeners (mounting awaits
// initLibrary()); dispatching a drag event before that point is dropped
// like any DOM event with no listener yet. Waiting for a bit of chrome to
// render is the same signal Playwright's own auto-waiting locators use.
async function waitUntilAppReady(page: import("@playwright/test").Page) {
  await expect(page.getByTestId("app-bar")).toBeVisible();
}

test("dragging a file over the window shows a full-window 'Drop to open' overlay (AC-1)", async ({ page }) => {
  await page.goto("/");
  await waitUntilAppReady(page);
  const dataTransfer = await dataTransferWithFiles(page, [{ name: "word-authored.docx", mime: DOCX_MIME, base64: toBase64(wordFixture) }]);
  await page.dispatchEvent("body", "dragenter", { dataTransfer });
  await expect(page.getByTestId("drop-overlay")).toHaveText("Drop to open");
});

test("dropping a .docx file opens it, exactly as the file picker would (AC-2)", async ({ page }) => {
  await page.goto("/");
  await waitUntilAppReady(page);
  const dataTransfer = await dataTransferWithFiles(page, [{ name: "word-authored.docx", mime: DOCX_MIME, base64: toBase64(wordFixture) }]);
  await page.dispatchEvent("body", "dragenter", { dataTransfer });
  await page.dispatchEvent("body", "drop", { dataTransfer });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByRole("heading", { name: "Quarterly Report" })).toBeVisible();
  await expect(page.getByTestId("drop-overlay")).toBeHidden();
});

test("dropping a .odt file opens it, exactly as the file picker would (AC-3)", async ({ page }) => {
  await page.goto("/");
  await waitUntilAppReady(page);
  const dataTransfer = await dataTransferWithFiles(page, [{ name: "libreoffice-writer.odt", mime: ODT_MIME, base64: toBase64(odtFixture) }]);
  await page.dispatchEvent("body", "drop", { dataTransfer });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("format-badge")).toHaveText("ODT");
});

test("dropping an unsupported file type shows the specific error and does not navigate away (AC-4, AC-5)", async ({ page }) => {
  await page.goto("/");
  await waitUntilAppReady(page);
  const dataTransfer = await dataTransferWithFiles(page, [{ name: "notes.txt", mime: "text/plain", base64: Buffer.from("hello").toString("base64") }]);
  await page.dispatchEvent("body", "drop", { dataTransfer });
  await expect(page.getByTestId("drop-error")).toHaveText("We can only open .docx, .odt, .md, .pptx, .odp, .key, .xlsx, .ods, .pdf and image files");
  await expect(page).toHaveURL(/\/$/);
});

test("dropping multiple files opens the first and shows a toast naming it (AC-6, AC-7)", async ({ page }) => {
  await page.goto("/");
  await waitUntilAppReady(page);
  const dataTransfer = await dataTransferWithFiles(page, [
    { name: "First.docx", mime: DOCX_MIME, base64: toBase64(wordFixture) },
    { name: "Second.docx", mime: DOCX_MIME, base64: toBase64(wordFixture) },
  ]);
  await page.dispatchEvent("body", "drop", { dataTransfer });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("toast")).toHaveText("Opened First — only one file at a time");
});

test("dragging over the window while a document is open does not modify it until the drop completes (AC-8)", async ({ page }) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();
  await expect(page.getByTestId("document-title")).toHaveAttribute("role", "button");
  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill("Keep me");
  await page.getByTestId("document-title-input").press("Enter");
  await expect(page.getByTestId("document-title")).toHaveText("Keep me");
  const openUrl = page.url();

  const dataTransfer = await dataTransferWithFiles(page, [{ name: "word-authored.docx", mime: DOCX_MIME, base64: toBase64(wordFixture) }]);
  await page.dispatchEvent("body", "dragenter", { dataTransfer });
  await expect(page.getByTestId("drop-overlay")).toBeVisible();
  expect(page.url()).toBe(openUrl);
  await expect(page.getByTestId("document-title")).toHaveText("Keep me");

  await page.dispatchEvent("body", "dragleave", { dataTransfer });
  await expect(page.getByTestId("drop-overlay")).toBeHidden();
  expect(page.url()).toBe(openUrl);
  await expect(page.getByTestId("document-title")).toHaveText("Keep me");
});

test("files larger than 50 MB are refused before parsing (AC-14)", async ({ page }) => {
  const dir = mkdtempSync(join(tmpdir(), "spec-009-"));
  const bigFile = join(dir, "big.docx");
  writeFileSync(bigFile, Buffer.alloc(51 * 1024 * 1024));

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(bigFile);
  await expect(page.getByTestId("open-document-error")).toHaveText(
    "That file is too large to open (limit 50 MB)",
  );
  await expect(page).toHaveURL(/\/$/);
});

test("/open?url= shows the address and fetches nothing until the reader confirms (AC-9)", async ({ page }) => {
  const bytes = readFileSync(wordFixture);
  let fetchCount = 0;
  await page.route("**/remote/report.docx", async (route) => {
    fetchCount += 1;
    await route.fulfill({ status: 200, contentType: DOCX_MIME, body: bytes });
  });

  await page.goto("/open?url=" + encodeURIComponent("https://example.com/remote/report.docx"));
  await expect(page.getByTestId("open-from-url-confirm")).toBeVisible();
  await expect(page.getByTestId("open-from-url-host")).toHaveText("example.com");
  await expect(page.getByTestId("open-from-url-target")).toHaveText(
    "https://example.com/remote/report.docx",
  );
  // The point of the panel: sitting on it sends no request.
  expect(fetchCount).toBe(0);

  await page.getByTestId("open-from-url-confirm-button").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  expect(fetchCount).toBe(1);
});

test("/open?url= refuses a non-http(s) scheme rather than fetching it (AC-9)", async ({ page }) => {
  // A data: URL carries its own content type, so without the scheme check it
  // would walk past the content-type gate and land in the library.
  const dataUrl = "data:text/html,<script>alert(1)</script>";
  await page.goto("/open?url=" + encodeURIComponent(dataUrl));
  await page.getByTestId("open-from-url-confirm-button").click();
  await expect(page.getByTestId("open-from-url-error")).toContainText("Couldn't open that link");
  await expect(page.getByTestId("open-from-url-reason")).toContainText("http://");
  await expect(page).toHaveURL(/\/open/);
});

test("/open?url= fetches the document at that URL, shows a loading state, then opens it (AC-9, AC-10)", async ({ page }) => {
  const bytes = readFileSync(wordFixture);
  let resolveResponse: () => void = () => {};
  const responseGate = new Promise<void>((resolve) => {
    resolveResponse = resolve;
  });
  await page.route("**/remote/report.docx", async (route) => {
    await responseGate;
    await route.fulfill({ status: 200, contentType: DOCX_MIME, body: bytes });
  });

  await openFromUrl(page, "https://example.com/remote/report.docx");
  await expect(page.getByTestId("open-from-url-loading")).toBeVisible();
  resolveResponse();

  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByRole("heading", { name: "Quarterly Report" })).toBeVisible();
});

test("a URL fetch that fails shows 'Couldn't open that link' with a reason and a fallback button (AC-11)", async ({ page }) => {
  await page.route("**/remote/missing.docx", (route) => route.fulfill({ status: 404, body: "not found" }));
  await openFromUrl(page, "https://example.com/remote/missing.docx");
  await expect(page.getByTestId("open-from-url-error")).toContainText("Couldn't open that link");
  await expect(page.getByTestId("open-from-url-reason")).toContainText("404");
  await expect(page.getByTestId("open-from-url-pick-file")).toHaveText("Choose a file instead");
});

test("a URL fetch blocked by CORS shows 'Couldn't open that link' with a reason and a fallback button (AC-12)", async ({ page }) => {
  await page.route("**/remote/blocked.docx", (route) => route.abort("failed"));
  await openFromUrl(page, "https://example.com/remote/blocked.docx");
  await expect(page.getByTestId("open-from-url-error")).toContainText("Couldn't open that link");
  await expect(page.getByTestId("open-from-url-reason")).not.toBeEmpty();
  await expect(page.getByTestId("open-from-url-pick-file")).toBeVisible();
});

test("a URL fetch returning a non-document content type shows 'Couldn't open that link' with the reason (AC-13)", async ({ page }) => {
  await page.route("**/remote/page", (route) => route.fulfill({ status: 200, contentType: "text/html", body: "<html></html>" }));
  await openFromUrl(page, "https://example.com/remote/page");
  await expect(page.getByTestId("open-from-url-error")).toContainText("Couldn't open that link");
  await expect(page.getByTestId("open-from-url-reason")).toContainText("text/html");
});

test("the fallback button on a failed URL fetch opens a local file (AC-11)", async ({ page }) => {
  await page.route("**/remote/missing.docx", (route) => route.fulfill({ status: 404, body: "not found" }));
  await openFromUrl(page, "https://example.com/remote/missing.docx");
  await expect(page.getByTestId("open-from-url-error")).toBeVisible();

  await page.getByTestId("open-from-url-pick-file").click();
  await page.getByTestId("open-from-url-input").setInputFiles(wordFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByRole("heading", { name: "Quarterly Report" })).toBeVisible();
});

test("a document opened from a URL is stored in the local library and is not re-fetched on reload (AC-15, AC-16)", async ({ page }) => {
  const bytes = readFileSync(wordFixture);
  let fetchCount = 0;
  await page.route("**/remote/library.docx", async (route) => {
    fetchCount += 1;
    await route.fulfill({ status: 200, contentType: DOCX_MIME, body: bytes });
  });

  await openFromUrl(page, "https://example.com/remote/library.docx");
  await expect(page).toHaveURL(/\/doc\/.+/);
  expect(fetchCount).toBe(1);
  const docUrl = page.url();

  await page.reload();
  await expect(page.getByRole("heading", { name: "Quarterly Report" })).toBeVisible();
  expect(page.url()).toBe(docUrl);
  expect(fetchCount).toBe(1);

  await page.goto("/");
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(page.getByTestId("row-title").first()).toHaveText("library");
});
