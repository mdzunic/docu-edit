import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PPTX = join(here, "fixtures", "spec095-media.pptx");

// SPEC-095 — slide media: open the media fixture and confirm the embedded audio
// frame shows its POSTER PLATE (not a grey "Media" box); open the Media panel,
// set Start automatically + Loop; Present and confirm the audio frame either
// reports playing or degrades to the media-blocked affordance; then advance,
// exit, download → reopen and confirm the panel still reads Start automatically
// + Loop and the linked clip still shows its link placeholder.

function stage(page: Page) {
  return page.getByTestId("stage-workspace");
}

async function openFixture(page: Page, file: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("media opens to a poster plate, plays in Present, and round-trips (AC-57/58/59)", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);
  const deckId = page.url().match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  expect(deckId).not.toBe("");

  // AC-57: the audio frame shows its poster plate, not a grey "Media" box.
  const audio = stage(page).getByRole("group", { name: "Audio" });
  await expect(audio).toBeVisible();
  await expect(audio.getByTestId("media-poster")).toBeVisible();
  // Only the deliberately out-of-subset element is a foreign "Media" box; the
  // audio clip is a modelled media frame, not a placeholder.
  await expect(stage(page).getByRole("group", { name: "Media" })).toHaveCount(1);
  await expect(stage(page).getByTestId("media-link-placeholder")).toBeVisible();

  // Open the Media panel on the audio frame; set Start automatically + Loop.
  await audio.click();
  await page.getByTestId("slides-media-toggle").click();
  await expect(page.getByTestId("slides-media-panel")).toBeVisible();
  await page.getByTestId("media-start-mode").selectOption("auto");
  await page.getByTestId("media-loop").check();

  // Present from slide 1.
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  await expect(page.getByTestId("presentation-play")).toBeVisible();

  // AC-58: the audio frame reports playing, or shows media-blocked — one of two.
  const play = page.getByTestId("presentation-play");
  await expect(async () => {
    const blocked = (await play.getByTestId("media-blocked").count()) > 0;
    const controls = play.getByTestId("media-controls");
    const playing =
      (await controls.count()) > 0 && (await controls.evaluate((el) => !(el as HTMLMediaElement).paused));
    expect(blocked || playing).toBe(true);
  }).toPass({ timeout: 5000 });

  // Advance past it and exit playback.
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("presentation-view")).toBeVisible();

  // AC-59: download → reopen; the panel still reads Start automatically + Loop,
  // and the linked clip still shows its link placeholder. (The video frame's
  // media-unsupported degrade is a deliberate non-assertion — no real codec
  // runs in CI, so chromium reports the synthetic MP4 as playable until load.)
  const file = await downloadTo(page, "spec095-roundtrip.pptx");
  await openFixture(page, file);
  // The Media panel's open state is a persisted preference, so it is already
  // open on reopen; selecting the audio frame shows its (round-tripped) flags.
  await stage(page).getByRole("group", { name: "Audio" }).click();
  await expect(page.getByTestId("slides-media-panel")).toBeVisible();
  await expect(page.getByTestId("media-start-mode")).toHaveValue("auto");
  await expect(page.getByTestId("media-loop")).toBeChecked();
  await expect(stage(page).getByTestId("media-link-placeholder")).toBeVisible();
});
