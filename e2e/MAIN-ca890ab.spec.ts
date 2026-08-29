import { test, expect, type Page } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";

// MAIN-ca890ab — main-repair for the SPEC-047 firstPageMs perf gate. The fix
// defers the .docx image base64 encode off the first-paint path: parseDocx
// gains an opt-in `deferMedia` (set by the interactive open funnel), and image
// `imageSrc` resolution moved out of the all-blocks pass into renderEntryFor,
// which runs only for realised pages. A still-deferred media part is therefore
// force-encoded on demand as its page renders.
//
// The perf budget itself is guarded by e2e/SPEC-047.perf.spec.ts (firstPageMs
// median-of-5). This spec guards the *behaviour* the deferral must not
// regress — the two runtime facts a read-only review could not observe and QA
// had to confirm in a browser:
//   1. a page-1 image (inline AND floating) actually paints on first render
//      through the deferred → encode-on-realise path, not a blank/broken <img>.
//   2. a save/export performed through that same path still writes every
//      image, so an open → edit → export → reopen round trip loses nothing.
// Both use the images.docx fixture (one inline 300x200 picture, one floating
// 120x120 picture) opened through the real interactive picker that sets
// deferMedia.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxImagesFixture = join(fixturesDir, "images.docx");

const content = (page: Page) => page.getByTestId("document-content");
const images = (page: Page) => content(page).locator('[data-testid="doc-image"]');

/** Open a file through the interactive funnel (the code path that sets
 * deferMedia) by driving the hidden open-document input directly. */
async function openViaPicker(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** Assert every doc-image on the page has actually decoded and painted —
 * `complete` with a non-zero natural size and a real data: URI — not merely
 * that an <img> element exists. Polled: the on-realise encode sets `src`
 * synchronously but image decode settles a tick later. */
async function expectAllImagesPainted(page: Page, expectedCount: number): Promise<void> {
  await expect(images(page)).toHaveCount(expectedCount);
  await expect
    .poll(async () =>
      images(page).evaluateAll((els) =>
        (els as HTMLImageElement[]).map((img) => ({
          complete: img.complete,
          naturalWidth: img.naturalWidth,
          dataUri: (img.currentSrc || "").startsWith("data:image/"),
        })),
      ),
    )
    .toEqual(
      Array.from({ length: expectedCount }, () => ({ complete: true, naturalWidth: expect.any(Number), dataUri: true })),
    );
  const widths = await images(page).evaluateAll((els) => (els as HTMLImageElement[]).map((img) => img.naturalWidth));
  for (const w of widths) {
    expect(w).toBeGreaterThan(0);
  }
}

test("a .docx opened through the picker paints its page-1 inline and floating images on first render (AC-1, AC-2)", async ({
  page,
}) => {
  await openViaPicker(page, docxImagesFixture);

  // Two images, both decoded and painted from real data: URIs.
  await expectAllImagesPainted(page, 2);

  // The inline picture is 300x200, the floating one 120x120 — the deferred
  // part was force-encoded at its source size, not dropped to a placeholder.
  const inline = (await images(page).nth(0).boundingBox())!;
  const floating = (await images(page).nth(1).boundingBox())!;
  expect(inline.width).toBeGreaterThan(300 * 0.9);
  expect(inline.width).toBeLessThan(300 * 1.1);
  expect(floating.width).toBeGreaterThan(120 * 0.9);
  expect(floating.width).toBeLessThan(120 * 1.1);
});

test("an edit + export + reopen through the deferred-media path loses no image and keeps the edit (AC-2 round trip)", async ({
  page,
}) => {
  await openViaPicker(page, docxImagesFixture);
  await expectAllImagesPainted(page, 2);

  // Edit: append a unique marker at the very end of the document.
  const marker = "QAROUNDTRIP1837";
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.type(` ${marker}`);
  await expect(content(page)).toContainText(marker);

  // Export to .docx through the File menu, capturing the download.
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]);
  const path = await download.path();
  expect(path).toBeTruthy();

  // Reopen the exported file: both images must repaint and the edit survive.
  await openViaPicker(page, path!);
  await expectAllImagesPainted(page, 2);
  await expect(content(page)).toContainText(marker);
});
