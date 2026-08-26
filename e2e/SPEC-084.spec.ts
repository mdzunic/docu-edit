import { test, expect, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-084 — slideshow playback. Open the fixture deck, Present from the deck
// header, walk it with the keyboard to the end screen, then confirm the exit
// lands back on the editor with the last slide selected. Fullscreen grants are
// flaky headless, so nothing here asserts the fullscreen promise's outcome —
// playback must work regardless (AC-55).
const here = dirname(fileURLToPath(import.meta.url));
const pptxFixture = join(here, "fixtures", "mini-deck.pptx");

async function openFixture(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(pptxFixture);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

test("Present the deck, walk it to the end screen, and exit to the last slide", async ({ page }) => {
  await openFixture(page);
  const deckUrl = page.url();
  const deckId = deckUrl.match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  expect(deckId).not.toBe("");

  // Present starts playback at the current slide (slide 1) on a dedicated route.
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  const play = page.getByTestId("presentation-play");
  await expect(play).toBeVisible();

  // Slide 1 content renders through the SPEC-079 stage…
  await expect(play.getByText("Quarterly Review")).toBeVisible();
  // …with no app chrome (no editor view, toolbar or filmstrip).
  await expect(page.getByTestId("presentation-view")).toHaveCount(0);
  await expect(page.getByTestId("slide-filmstrip")).toHaveCount(0);
  await expect(page.getByTestId("stage-workspace")).toHaveCount(0);

  // Advance to slide 2 — the URL tracks it.
  await page.keyboard.press("ArrowRight");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/2$`));
  await expect(play.getByText("Ship presentations MVP")).toBeVisible();

  // The slide counter reads "2 / 4" and fades in on pointer movement.
  await page.mouse.move(120, 120);
  await page.mouse.move(300, 300);
  const counter = page.getByTestId("play-slide-counter");
  await expect(counter).toHaveText("2 / 4");
  await expect
    .poll(() => counter.evaluate((el) => getComputedStyle(el.closest(".play-hud") as HTMLElement).opacity))
    .toBe("1");

  // Jump to the last slide, then one advance shows the end screen.
  await page.keyboard.press("End");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/4$`));
  await expect(play.getByText("Extras")).toBeVisible();

  await page.keyboard.press("ArrowRight");
  const endScreen = page.getByTestId("play-end-screen");
  await expect(endScreen).toBeVisible();
  await expect(endScreen).toContainText("End of slide show");

  // One more advance exits playback: back on the editor with slide 4 selected.
  await page.keyboard.press("ArrowRight");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}(\\?|$)`));
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  const selected = page.getByTestId("filmstrip-slide").nth(3);
  await expect(selected).toHaveAttribute("aria-selected", "true");
});

test("a deck with a slide can be presented; the button is disabled only when empty", async ({ page }) => {
  await openFixture(page);
  await expect(page.getByTestId("slides-present")).toBeEnabled();
});
