import { test, expect, type Download, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { inflateSync } from "node:zlib";

// SPEC-104 — hidden slides and deck print/export. Open the hidden-slide
// fixture (three slides, the middle one hidden): its thumbnail is marked
// hidden. Export a PDF as "Handout 4 per page" (one sheet for the two
// visible slides), then as "Slides" (two pages), then with hidden slides
// included (three) — that 2-vs-3 delta pins the option. Export slide 1 as a
// 2× PNG and read its IHDR. Present from slide 1: one advance steps over the
// hidden slide 2 onto slide 3. Back in the editor, hide slide 1 from the
// filmstrip menu, undo it, then download and reopen: slide 2 is still
// hidden and slide 1 is not.
const here = dirname(fileURLToPath(import.meta.url));
const HIDDEN_DECK = join(here, "fixtures", "spec104-hidden.pptx");

async function openDeck(page: Page, file: string): Promise<void> {
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function capture(page: Page, act: () => Promise<void>): Promise<{ download: Download; bytes: Buffer }> {
  const downloadPromise = page.waitForEvent("download");
  await act();
  const download = await downloadPromise;
  return { download, bytes: readFileSync(await download.path()) };
}

/** The PDF's page-tree count (the /Pages dictionary is never compressed). */
function pageTreeCount(pdf: Buffer): number {
  const match = /\/Type \/Pages \/Kids \[[^\]]*\] \/Count (\d+)/.exec(pdf.toString("latin1"));
  return match ? Number(match[1]) : -1;
}

/** Every FlateDecode stream of the PDF, inflated and concatenated. */
function inflatedStreams(pdf: Buffer): string {
  const text = pdf.toString("latin1");
  let out = "";
  for (const match of text.matchAll(/\/Filter \/FlateDecode[^>]*>>\nstream\n/g)) {
    const start = (match.index ?? 0) + match[0].length;
    const end = text.indexOf("\nendstream", start);
    try {
      out += inflateSync(pdf.subarray(start, end)).toString("latin1");
    } catch {
      // A stream that does not inflate cleanly is not page content.
    }
  }
  return out;
}

test("hidden slides: filmstrip mark, PDF/handout and PNG export, playback skip, hide + undo round trip", async ({
  page,
}) => {
  await page.goto("/");
  await openDeck(page, HIDDEN_DECK);

  // The middle thumbnail is marked hidden (AC-43).
  const thumbs = page.getByTestId("filmstrip-slide");
  await expect(thumbs).toHaveCount(3);
  await expect(thumbs.nth(1)).toHaveClass(/filmstrip-hidden/);
  await expect(thumbs.nth(0)).not.toHaveClass(/filmstrip-hidden/);
  await expect(thumbs.nth(2)).not.toHaveClass(/filmstrip-hidden/);

  // Download menu → "PDF (.pdf)…" opens the dialog on its PDF tab.
  await page.getByTestId("presentation-download-menu").click();
  await page.getByTestId("presentation-export-pdf").click();
  const dialog = page.getByTestId("deck-print-dialog");
  await expect(dialog).toBeVisible();
  await expect(page.getByTestId("deck-print-tab-pdf")).toHaveAttribute("aria-selected", "true");

  // Handout 4 per page, range All, hidden off → one sheet (AC-44).
  await page.getByTestId("deck-print-layout").selectOption({ label: "Handout 4 per page" });
  await page.getByTestId("deck-print-range-all").check();
  await expect(page.getByTestId("deck-print-hidden")).not.toBeChecked();
  await expect(page.getByTestId("deck-print-frame")).toBeChecked();
  const handout = await capture(page, () => page.getByTestId("deck-print-pdf").click());
  expect(handout.download.suggestedFilename()).toBe("spec104-hidden.pdf");
  expect(handout.bytes.subarray(0, 8).toString("latin1")).toBe("%PDF-1.7");
  expect(pageTreeCount(handout.bytes)).toBe(1);
  // Real pages, not screenshots: bookmarks per slide, the deck title, and
  // text as Tj runs against an embedded face with a /ToUnicode map.
  const handoutText = handout.bytes.toString("latin1");
  expect(handoutText).toContain("/Title (spec104-hidden)");
  expect(handoutText).toContain("/Title (Welcome)");
  expect(handoutText).toContain("/Title (Thank you)");
  expect(handoutText).not.toContain("/Title (Backup data)");
  expect(handoutText).toContain("/ToUnicode");
  expect(inflatedStreams(handout.bytes)).toContain(" Tj ");
  // The sheet is unmounted once the export is done.
  await expect(page.getByTestId("deck-print-sheet")).toHaveCount(0);

  // Switch to the Slides layout → two pages; include hidden slides → three.
  await page.getByTestId("deck-print-layout").selectOption({ label: "Slides" });
  await expect(page.getByTestId("deck-print-frame")).not.toBeChecked();
  const slidesPdf = await capture(page, () => page.getByTestId("deck-print-pdf").click());
  expect(pageTreeCount(slidesPdf.bytes)).toBe(2);
  await page.getByTestId("deck-print-hidden").check();
  const withHidden = await capture(page, () => page.getByTestId("deck-print-pdf").click());
  expect(pageTreeCount(withHidden.bytes)).toBe(3);
  expect(withHidden.bytes.toString("latin1")).toContain("/Title (Backup data)");

  // Images, scale 2×, range "1" → a 2560×1440 PNG (AC-45).
  await page.getByTestId("deck-print-tab-images").click();
  await page.getByTestId("deck-print-scale").selectOption("2");
  await page.getByTestId("deck-print-range-custom").check();
  await page.getByTestId("deck-print-range-input").fill("1");
  await expect(page.getByTestId("deck-print-estimate")).toContainText("2560 × 1440");
  const png = await capture(page, () => page.getByTestId("deck-print-images").click());
  expect(png.download.suggestedFilename()).toBe("spec104-hidden-slide-01.png");
  expect(png.bytes.subarray(1, 4).toString("latin1")).toBe("PNG");
  expect(png.bytes.subarray(12, 16).toString("latin1")).toBe("IHDR");
  expect([png.bytes.readUInt32BE(16), png.bytes.readUInt32BE(20)]).toEqual([2560, 1440]);
  await page.getByTestId("deck-print-cancel").click();
  await expect(dialog).toHaveCount(0);

  // Cmd/Ctrl+P opens the dialog on the PDF tab (AC-18).
  await page.getByTestId("slide-filmstrip").focus();
  await page.keyboard.press("ControlOrMeta+p");
  await expect(dialog).toBeVisible();
  await expect(page.getByTestId("deck-print-tab-pdf")).toHaveAttribute("aria-selected", "true");
  await page.getByTestId("deck-print-cancel").click();
  await expect(dialog).toHaveCount(0);

  // With slide 1 selected, Present opens on slide 1; one advance steps over
  // the hidden slide 2 onto slide 3 (AC-46).
  const deckId = page.url().match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  await thumbs.nth(0).click();
  await expect(thumbs.nth(0)).toHaveAttribute("aria-selected", "true");
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  await expect(page.getByTestId("presentation-play")).toBeVisible();
  await page.keyboard.press("ArrowRight");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/3$`));
  await expect(page.getByTestId("presentation-play").getByTestId("slide-stage").getByText("Thank you")).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("presentation-view")).toBeVisible();

  // Hide slide 1 from the filmstrip menu, then undo it (AC-47).
  await thumbs.nth(0).click({ button: "right" });
  await expect(page.getByTestId("slide-menu-hide")).toHaveText("Hide slide");
  await page.getByTestId("slide-menu-hide").click();
  await expect(thumbs.nth(0)).toHaveClass(/filmstrip-hidden/);
  await page.getByTestId("presentation-undo").click();
  await expect(thumbs.nth(0)).not.toHaveClass(/filmstrip-hidden/);
  // Still fully reachable in the editor: selecting the hidden slide works.
  await thumbs.nth(1).click();
  await expect(thumbs.nth(1)).toHaveAttribute("aria-selected", "true");

  // Download, reopen: slide 2 is still hidden and slide 1 is not.
  const saved = await capture(page, () => page.getByTestId("presentation-download").click());
  const savedPath = join(tmpdir(), `spec104-${Date.now()}.pptx`);
  await saved.download.saveAs(savedPath);
  await page.goto("/");
  await openDeck(page, savedPath);
  const reopened = page.getByTestId("filmstrip-slide");
  await expect(reopened).toHaveCount(3);
  await expect(reopened.nth(1)).toHaveClass(/filmstrip-hidden/);
  await expect(reopened.nth(0)).not.toHaveClass(/filmstrip-hidden/);
});
