import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PPTX = join(here, "fixtures", "mini-deck.pptx");

// SPEC-083 — speaker notes: open the pptx fixture, toggle the notes panel, see
// slide 1's fixture notes, edit them, add notes to slide 2, switch slides (the
// notes are kept per slide), then download, re-open and confirm the edited
// notes survived the .pptx round trip.

function editor(page: Page) {
  return page.getByTestId("slides-notes-editor");
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("read, edit and round-trip speaker notes across slides", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE_PPTX);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();

  // The panel is collapsed until toggled (default per SPEC-066 preference).
  await expect(page.getByTestId("slides-notes-panel")).toHaveCount(0);
  await page.getByTestId("slides-notes-toggle").click();
  await expect(page.getByTestId("slides-notes-panel")).toBeVisible();

  // Slide 1's fixture notes are shown, formatting flattened, trailing empty gone.
  await expect(editor(page)).toHaveValue("Open with the Q3 highlight\nPause before the revenue chart");

  // Edit slide 1's notes.
  await editor(page).fill("Rehearsed opener\nWatch the clock");
  await expect(editor(page)).toHaveValue("Rehearsed opener\nWatch the clock");

  // Slide 2 has no notes yet — add some, then switch away and back.
  await page.getByTestId("filmstrip-slide").nth(1).click();
  await expect(editor(page)).toHaveValue("");
  await editor(page).fill("First note for slide two");
  await page.getByTestId("filmstrip-slide").nth(0).click();
  await expect(editor(page)).toHaveValue("Rehearsed opener\nWatch the clock");
  await page.getByTestId("filmstrip-slide").nth(1).click();
  await expect(editor(page)).toHaveValue("First note for slide two");

  // Download and re-open: the edited notes survive the round trip.
  const file = await downloadTo(page, "spec083-deck.pptx");
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();

  // The open/collapsed choice persists per device, so the reopened deck shows
  // the panel already expanded (no re-toggle needed).
  await expect(page.getByTestId("slides-notes-panel")).toBeVisible();
  await expect(editor(page)).toHaveValue("Rehearsed opener\nWatch the clock");
  await page.getByTestId("filmstrip-slide").nth(1).click();
  await expect(editor(page)).toHaveValue("First note for slide two");
});
