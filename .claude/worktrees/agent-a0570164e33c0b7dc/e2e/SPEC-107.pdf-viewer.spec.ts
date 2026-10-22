import { test, expect, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-107 AC-12 (QA round 1's AC-33) — a PDF at /media/:id is shown by the
// browser's own viewer, under SEC-6's Content-Security-Policy, rather than by
// the fallback panel.
//
// Two faults kept it from doing so, and AC-16's journey in SPEC-107.spec.ts
// only asserts the object is ATTACHED, which held throughout:
//  - Chromium hosts that viewer in a frame inside the <object>, so `frame-src`
//    governs it even though the element is an <object>: with `frame-src
//    'none'` every PDF was refused ("Framing 'blob:…' violates … frame-src
//    'none'") and the viewer area showed only Chromium's blocked-frame icon.
//  - useMediaDocument minted a second blob: URL on its post-initLibrary
//    re-sync, revoking the one the <object> had just begun loading. Chromium
//    then dropped the object to its fallback content most of the time: the
//    nested <embed> carried the viewer, and the "cannot display PDFs" panel
//    sat laid out (clipped) behind it, in the accessibility tree.
//
// Full Chromium, not the headless shell: Playwright's default headless browser
// (chromium-headless-shell) ships no PDF viewer at all —
// navigator.pdfViewerEnabled is false and every PDF <object> reads "Couldn't
// load plugin." — so it cannot tell a working viewer from a blocked one.
// channel "chromium" runs the full build in new headless mode, which carries
// the viewer a desktop Chrome does.
test.use({ channel: "chromium" });

const pdfFixture = join(dirname(fileURLToPath(import.meta.url)), "fixtures", "spec107-pages.pdf");

async function dropPdf(page: Page): Promise<void> {
  const transfer = await page.evaluateHandle((base64) => {
    const dt = new DataTransfer();
    const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
    dt.items.add(new File([bytes], "spec107-pages.pdf", { type: "application/pdf" }));
    return dt;
  }, readFileSync(pdfFixture).toString("base64"));
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
}

test("the browser's PDF viewer renders the document, not the fallback panel (AC-12)", async ({ page }) => {
  const cspViolations: string[] = [];
  page.on("console", (message) => {
    if (message.type() === "error" && message.text().includes("Content Security Policy")) {
      cspViolations.push(message.text());
    }
  });

  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  expect(
    await page.evaluate(() => navigator.pdfViewerEnabled),
    "this spec needs a Chromium that carries its PDF viewer (channel: chromium)",
  ).toBe(true);

  await dropPdf(page);
  await expect(page).toHaveURL(/\/media\/.+/);
  await expect(page.getByTestId("detail-pages")).toHaveText("3 pages");

  // The viewer's frame navigates to the object's blob: URL; a refused frame
  // lands on chrome-error:// instead. Waiting for it first also keeps the
  // fallback check below from passing in the instant before the object has
  // resolved either way.
  const object = page.locator('object[type="application/pdf"]');
  await expect(object).toBeVisible();
  const src = await object.getAttribute("data");
  expect(src).toMatch(/^blob:/);
  await expect.poll(() => page.frames().some((frame) => frame.url() === src)).toBe(true);

  // An <object> that renders lays out none of its fallback content — neither
  // the nested <embed> nor the "cannot display PDFs" panel.
  await expect(page.getByTestId("media-pdf-fallback")).toBeHidden();
  await expect(page.locator("embed")).toBeHidden();

  expect(cspViolations).toEqual([]);
});
