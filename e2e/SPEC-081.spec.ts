import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));

// SPEC-081 — slide text editing: open the pptx fixture, retype the slide-1
// title (bolded via the toolbar), add a bulleted body line on slide 2,
// download and re-open, and confirm both edits survive; a second download with
// no further edits re-downloads the edited deck stably.

const FIXTURE = join(here, "fixtures", "mini-deck.pptx");

/** The text frames on the MAIN stage (not the filmstrip thumbnails). */
function stageFrames(page: Page) {
  return page.getByTestId("stage-workspace").getByTestId("slide-text-frame");
}

async function openFixture(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

test("edit a title + add a bulleted body line, round-trip the download, re-download stably", async ({ page }) => {
  await openFixture(page);
  await expect(stageFrames(page).first()).toContainText("Quarterly Review");

  // Slide 1: double-click the title, select all, type a new title.
  await stageFrames(page).first().dblclick();
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await page.keyboard.press("Control+a");
  await page.keyboard.type("New Title");
  // Select all again and bold it via the toolbar.
  await page.keyboard.press("Control+a");
  await page.getByTestId("slides-bold").click();
  await expect(page.getByTestId("slides-bold")).toHaveAttribute("aria-pressed", "true");
  await page.keyboard.press("Escape");
  await expect(stageFrames(page).first()).toContainText("New Title");

  // Slide 2: add a bulleted body line.
  await page.getByTestId("filmstrip-slide").nth(1).click();
  const body = stageFrames(page).nth(1);
  await expect(body).toContainText("Ship presentations MVP");
  await body.dblclick();
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await page.keyboard.press("Enter");
  await page.keyboard.type("A brand new line");
  await page.getByTestId("slides-bullet").click();
  await page.keyboard.press("Escape");
  await expect(stageFrames(page).nth(1)).toContainText("A brand new line");

  // Download and re-open — both edits are visible on the stage.
  const download = await downloadTo(page, "spec081-edited.pptx");
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(download);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(stageFrames(page).first()).toContainText("New Title");
  await page.getByTestId("filmstrip-slide").nth(1).click();
  await expect(stageFrames(page).nth(1)).toContainText("A brand new line");

  // AC-47: a second download with zero further edits re-downloads the edited
  // deck stably — re-opening it still shows both edits.
  const again = await downloadTo(page, "spec081-again.pptx");
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(again);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(stageFrames(page).first()).toContainText("New Title");
});

test("double-click drops the caret at the click point, not the end (AC-1)", async ({ page }) => {
  await openFixture(page);
  const title = stageFrames(page).first();
  await expect(title).toContainText("Quarterly Review");

  // Double-click just inside the left edge, over the very start of the text, so
  // the caret lands before 'Q' — not at the end of the frame's content.
  const box = await title.boundingBox();
  if (!box) throw new Error("no title box");
  await page.mouse.dblclick(box.x + 3, box.y + box.height / 2);
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await page.keyboard.type("XXX");
  await page.keyboard.press("Escape");

  // Caret at the click point → the typed text is at the START, not appended.
  await expect(title).toContainText("XXXQuarterly Review");
});

test("the presentation font-size control clamps typed entry to 6-96 (AC-18)", async ({ page }) => {
  await openFixture(page);
  // Select the title frame so the toolbar targets it (no editing session needed).
  await stageFrames(page).first().click();
  const size = page.getByTestId("font-size-input");
  await expect(size).toBeEnabled();

  await size.fill("200");
  await size.press("Enter");
  await expect(size).toHaveValue("96"); // clamped to the presentation max

  await size.fill("2");
  await size.press("Enter");
  await expect(size).toHaveValue("6"); // clamped to the presentation min
});

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}
