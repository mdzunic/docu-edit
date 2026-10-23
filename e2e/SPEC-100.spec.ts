import { test, expect, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-100 — presenter view. The editor tab becomes the presenter window and
// opens the audience window (the existing /present/:n route with ?follow=) as
// a popup; the two stay in step over the same-origin channel whichever window
// takes the input. Fullscreen grants are flaky headless, so nothing here
// asserts the fullscreen promise's outcome (as SPEC-084's e2e).
const here = dirname(fileURLToPath(import.meta.url));
const miniDeck = join(here, "fixtures", "mini-deck.pptx");
const animationsDeck = join(here, "fixtures", "spec086-animations.pptx");

async function openFixture(page: Page, file: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  const deckId = page.url().match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  expect(deckId).not.toBe("");
  return deckId;
}

/** Present menu → "Presenter view": the presenter route here, the audience
 * window as the popup. */
async function startPresenterView(page: Page, deckId: string): Promise<Page> {
  await page.getByTestId("presentation-present-menu").click();
  await expect(page.getByTestId("presentation-present-menu-popup")).toBeVisible();
  const popup = page.waitForEvent("popup");
  await page.getByRole("menuitem", { name: "Presenter view" }).click();
  const audience = await popup;
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/presenter/1$`));
  await expect(page.getByTestId("presenter-view")).toBeVisible();
  await expect(audience).toHaveURL(new RegExp(`/slides/${deckId}/present/1\\?follow=[^&]+`));
  await expect(audience.getByTestId("presentation-play")).toBeVisible();
  return audience;
}

test("presenter view drives the audience window from either window, blanks it, jumps by grid and exits", async ({ page }) => {
  const deckId = await openFixture(page, miniDeck);
  const audience = await startPresenterView(page, deckId);
  const audienceStage = audience.getByTestId("presentation-play").getByTestId("slide-stage");

  // Slide 1 on both; the next-slide pane shows slide 2, the notes pane slide
  // 1's first note.
  const current = page.getByTestId("presenter-current-slide");
  const counter = page.getByTestId("presenter-slide-counter");
  await expect(counter).toHaveText("Slide 1 of 4");
  await expect(current.getByText("Quarterly Review")).toBeVisible();
  await expect(page.getByTestId("presenter-next-slide").getByText("Ship presentations MVP")).toBeVisible();
  await expect(page.getByTestId("presenter-notes")).toContainText("Open with the Q3 highlight");
  await expect(audienceStage.getByText("Quarterly Review")).toBeVisible();
  // The timer runs from the start of the show (M:SS; the exact second
  // depends on how loaded the machine is).
  await expect(page.getByTestId("presenter-timer")).toHaveText(/^\d+:\d\d$/);

  // ArrowRight in the presenter moves both windows to slide 2.
  await page.keyboard.press("ArrowRight");
  await expect(counter).toHaveText("Slide 2 of 4");
  await expect(audience).toHaveURL(new RegExp(`/slides/${deckId}/present/2\\?follow=`));
  await expect(audienceStage.getByText("Ship presentations MVP")).toBeVisible();

  // ArrowRight in the AUDIENCE window moves both to slide 3: it forwards the
  // input to the presenter and renders what comes back.
  await audience.keyboard.press("ArrowRight");
  await expect(counter).toHaveText("Slide 3 of 4");
  await expect(audience).toHaveURL(new RegExp(`/slides/${deckId}/present/3\\?follow=`));
  await expect(audienceStage.getByText("Thank you")).toBeVisible();
  await expect(page.getByTestId("presenter-notes")).toContainText("No notes for this slide");

  // b blanks the audience and badges the presenter; ArrowRight clears it.
  await page.keyboard.press("b");
  const blank = audience.getByTestId("play-blank");
  await expect(blank).toBeVisible();
  await expect(blank).toHaveCSS("background-color", "rgb(0, 0, 0)");
  await expect(audience.getByTestId("play-live")).toHaveText("Screen blanked");
  await expect(page.getByTestId("presenter-blank-badge")).toBeVisible();
  await expect(page.getByTestId("presenter-next-button")).toBeEnabled();
  await page.keyboard.press("ArrowRight");
  await expect(blank).toHaveCount(0);
  await expect(page.getByTestId("presenter-blank-badge")).toHaveCount(0);
  await expect(counter).toHaveText("Slide 3 of 4");

  // The grid jumps both windows to slide 4.
  await page.getByTestId("presenter-grid-toggle").click();
  const grid = page.getByTestId("presenter-grid");
  await expect(grid).toBeVisible();
  await expect(grid).toHaveAttribute("role", "listbox");
  await expect(page.getByTestId("presenter-grid-slide")).toHaveCount(4);
  await page.getByTestId("presenter-grid-slide").nth(3).click();
  await expect(grid).toHaveCount(0);
  await expect(counter).toHaveText("Slide 4 of 4");
  await expect(audience).toHaveURL(new RegExp(`/slides/${deckId}/present/4\\?follow=`));
  await expect(audienceStage.getByText("Extras")).toBeVisible();
  await expect(page.getByTestId("presenter-next-end")).toHaveText("End of show");

  // Escape in the presenter exits to the editor with slide 4 selected and
  // closes the audience window.
  await page.keyboard.press("Escape");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}\\?return=4$`));
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  await expect(page.getByTestId("filmstrip-slide").nth(3)).toHaveAttribute("aria-selected", "true");
  await expect.poll(() => audience.isClosed()).toBe(true);
});

test("the presenter's build indicator counts builds as the show advances", async ({ page }) => {
  const deckId = await openFixture(page, animationsDeck);
  const audience = await startPresenterView(page, deckId);
  const indicator = page.getByTestId("presenter-build-indicator");
  const counter = page.getByTestId("presenter-slide-counter");
  const current = page.getByTestId("presenter-current-slide");

  // Slide 1 has no builds: no indicator.
  await expect(counter).toHaveText("Slide 1 of 3");
  await expect(indicator).toHaveCount(0);

  // Slide 2 carries one click build (a fade-in with a fly-in after it).
  await page.keyboard.press("ArrowRight");
  await expect(counter).toHaveText("Slide 2 of 3");
  await expect(indicator).toHaveText("Build 0 of 1");
  await expect(audience).toHaveURL(new RegExp(`/slides/${deckId}/present/2\\?follow=`));

  await page.keyboard.press("ArrowRight");
  await expect(indicator).toHaveText("Build 1 of 1");
  // Both windows play the build; an advance mid-build would complete it and
  // hold, so let the presenter's mirror finish first.
  await expect(current).toHaveAttribute("data-building", "false");
  await expect(audience.getByTestId("presentation-play")).toHaveAttribute("data-building", "false");

  // On to slide 3, whose timing the model only preserves: no builds again.
  await page.keyboard.press("ArrowRight");
  await expect(counter).toHaveText("Slide 3 of 3");
  await expect(indicator).toHaveCount(0);
  await expect(audience).toHaveURL(new RegExp(`/slides/${deckId}/present/3\\?follow=`));
});
