import { test, expect, type Page } from "@playwright/test";
import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-087 — importing an Apple Keynote deck: drag mini-deck.key onto the
// app, see the non-blocking import notice, land in the editor with slide 1's
// title on the stage, find the deck in the library under a PPTX badge, and
// download a .pptx that re-opens with the same content.
//
// The fixture is a REAL Keynote-written file (CI cannot run Keynote), so
// these tests fail loudly while it is absent rather than passing vacuously
// (AC-51's rule, applied to the e2e half of the contract).
const here = dirname(fileURLToPath(import.meta.url));
const keyFixture = join(here, "fixtures", "mini-deck.key");

const FIXTURE_MISSING =
  "e2e/fixtures/mini-deck.key is missing — run scripts/make-keynote-fixture.applescript on a Mac with Keynote and commit the result (SPEC-087's prerequisite).";

test.beforeEach(() => {
  if (!existsSync(keyFixture)) {
    throw new Error(FIXTURE_MISSING);
  }
});

/** Drags mini-deck.key onto the app shell — the drop path, not the picker
 * (AC-52 specifies dragging). The bytes travel as base64 because a
 * DataTransfer must be built inside the page. */
async function dropFixture(page: Page): Promise<void> {
  const base64 = readFileSync(keyFixture).toString("base64");
  const dataTransfer = await page.evaluateHandle((payload) => {
    const binary = atob(payload);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    const dt = new DataTransfer();
    dt.items.add(new File([bytes], "mini-deck.key", { type: "application/x-iwork-keynote-sffkey" }));
    return dt;
  }, base64);
  await page.dispatchEvent("body", "dragenter", { dataTransfer });
  await page.dispatchEvent("body", "drop", { dataTransfer });
}

async function openFixtureByDrop(page: Page): Promise<void> {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFixture(page);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

test("dropping mini-deck.key shows the import notice (AC-52)", async ({ page }) => {
  await openFixtureByDrop(page);
  await expect(page.getByTestId("toast").filter({ hasText: "Imported from Keynote." })).toBeVisible();
});

test("the deck opens at /slides/:id with slide 1's title on the stage (AC-39, AC-53)", async ({ page }) => {
  await openFixtureByDrop(page);
  const stage = page.getByTestId("stage-workspace");
  await expect(stage.getByText("Keynote fixture deck")).toBeVisible();
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(2);
  // Fully editable (AC-39): the header title is the rename control and the
  // deck carries the normal editing chrome, not a read-only note.
  await expect(page.getByTestId("presentation-note")).toHaveCount(0);
});

test("the library lists the deck as \"mini-deck\" with a PPTX badge (AC-54, AC-56)", async ({ page }) => {
  await openFixtureByDrop(page);
  await expect(page.getByTestId("document-title")).toContainText("mini-deck");

  await page.getByTestId("presentation-close").click();
  await expect(page.getByTestId("library")).toBeVisible();
  const row = page.getByTestId("library-row").filter({ hasText: "mini-deck" });
  await expect(row.getByTestId("row-format")).toHaveText("PPTX");
});

test("downloading produces a .pptx that re-opens with the same content contract (AC-55)", async ({ page }, testInfo) => {
  await openFixtureByDrop(page);

  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("mini-deck.pptx");
  const filePath = testInfo.outputPath("mini-deck.pptx");
  await download.saveAs(filePath);

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(filePath);
  await expect(page).toHaveURL(/\/slides\/.+/);

  const stage = page.getByTestId("stage-workspace");
  await expect(stage.getByText("Keynote fixture deck")).toBeVisible();
  await expect(stage.getByText("Imported by docu-edit")).toBeVisible();
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(2);

  await page.getByTestId("filmstrip-slide").nth(1).click();
  await expect(stage.getByText("Second slide")).toBeVisible();
  await expect(stage.getByText("First bullet point")).toBeVisible();
  await expect(stage.getByText("Second bullet point")).toBeVisible();
  // The image survived the round trip onto the rebuilt slide.
  await expect(stage.locator("img").first()).toBeVisible();

  // The speaker note re-imported into SPEC-083's notes editor (slide 2 is
  // still selected from the assertions above).
  await page.getByTestId("slides-notes-toggle").click();
  await expect(page.getByTestId("slides-notes-editor")).toHaveValue("A speaker note for the importer");
});
