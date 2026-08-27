import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { pathToFileURL } from "node:url";
import { fileURLToPath } from "node:url";

// SPEC-088 — export a presentation to a single self-contained HTML file. Open
// the animation fixture, add an entrance build to a slide-1 frame and a Push
// transition to slide 2, attach a speaker note, then export via the Download
// menu (notes off, then notes on). Drive each exported file straight from
// `file://`: slide 1 is visible, one advance builds the hidden frame in before
// the slide changes, the next advance plays a transition to slide 2, a
// `#2` deep-link reload lands on slide 2, and the whole run makes no network
// request. With notes on, pressing V reveals the note text.
const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE = join(here, "fixtures", "spec086-animations.pptx");

async function openFixture(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function exportHtml(page: Page, includeNotes: boolean, savePath: string): Promise<void> {
  await page.getByTestId("presentation-download-menu").click();
  await page.getByTestId("presentation-export-html").click();
  await expect(page.getByTestId("export-html-dialog")).toBeVisible();
  const notes = page.getByTestId("export-html-notes");
  if (includeNotes) {
    await notes.check();
  }
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("export-html-confirm").click();
  const download = await downloadPromise;
  await download.saveAs(savePath);
}

test("export a deck to self-contained HTML and drive it offline", async ({ page }) => {
  await openFixture(page);
  const workspace = page.getByTestId("stage-workspace");

  // Slide 1: add "Fade in, on click" to the body frame (a build).
  await page.getByTestId("slides-animate-toggle").click();
  await expect(page.getByTestId("slides-animation-panel")).toBeVisible();
  await workspace.getByText("Body line").click();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(1);

  // Slide 1: attach a speaker note.
  await page.getByTestId("slides-notes-toggle").click();
  await page.getByTestId("slides-notes-editor").fill("Remember the punchline");

  // Slide 2: a Push/Left transition, 1s (slow enough to observe mid-flight).
  await page.getByTestId("filmstrip-slide").nth(1).click();
  await page.getByTestId("slides-transition-toggle").click();
  await expect(page.getByTestId("slides-transition-panel")).toBeVisible();
  await page.getByTestId("transition-effect-push").check();
  await page.getByTestId("transition-direction").selectOption("left");
  await page.getByTestId("transition-duration").fill("1");
  await page.getByTestId("transition-duration").blur();

  // Export twice: notes off, then notes on.
  const fileOff = join(tmpdir(), "spec088-notes-off.html");
  const fileOn = join(tmpdir(), "spec088-notes-on.html");
  await exportHtml(page, false, fileOff);
  await exportHtml(page, true, fileOn);

  // --- drive the notes-off file straight from file:// ----------------------
  // The notes-off export omits the #notes aside entirely, so the viewer must
  // never touch it during navigation; capture any uncaught page error so a
  // regression there fails the run.
  const network: string[] = [];
  const pageErrors: string[] = [];
  page.on("request", (req) => {
    if (!req.url().startsWith("file:")) {
      network.push(req.url());
    }
  });
  page.on("pageerror", (err) => pageErrors.push(err.message));
  const urlOff = pathToFileURL(fileOff).href;
  await page.goto(urlOff);

  // Slide 1 is visible; the built frame starts hidden.
  await expect(page.locator("#s1")).toBeVisible();
  await expect(page.locator("#s2")).toBeHidden();
  await expect(page.locator("#s1").getByText("Body line")).toBeHidden();

  // One advance builds the frame in — the slide has not changed yet.
  await page.locator("body").click();
  await expect(page.locator("#s1").getByText("Body line")).toBeVisible();
  await expect(page.locator("#s1")).toBeVisible();
  await expect(page.locator("#s2")).toBeHidden();
  // Let the build fully settle so the next input advances rather than just
  // completing the in-flight build (a mid-build input completes it — AC-37).
  await page.waitForTimeout(700);

  // The next advance plays the transition to slide 2: the incoming section
  // carries an inline transform mid-flight (an intermediate state), then settles.
  await page.keyboard.press("ArrowRight");
  await expect
    .poll(() => page.$eval("#s2", (el) => (el as HTMLElement).style.transform), { timeout: 2000 })
    .not.toBe("");
  await expect(page.locator("#s2")).toBeVisible();

  // Navigation on the notes-off file must sync the hash (a null #notes deref
  // used to throw before syncHash ran, leaving the hash empty).
  await expect.poll(() => page.evaluate(() => location.hash)).toBe("#2");

  // Not a single network request — nor an uncaught error — while driving it.
  expect(network).toEqual([]);
  expect(pageErrors).toEqual([]);

  // A #2 deep-link reload lands on slide 2.
  await page.goto(`${urlOff}#2`);
  await expect(page.locator("#s2")).toBeVisible();
  await expect(page.locator("#s1")).toBeHidden();

  // --- the notes-on file exposes the note via V ----------------------------
  await page.goto(pathToFileURL(fileOn).href);
  await expect(page.getByText("Remember the punchline")).toBeHidden();
  await page.keyboard.press("v");
  await expect(page.getByText("Remember the punchline")).toBeVisible();
});
