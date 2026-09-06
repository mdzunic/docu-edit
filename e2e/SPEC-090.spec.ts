import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-090 — deck setup. Open a deck, change its slide size to 4:3 with "Scale
// content" (the stage aspect ratio and slide 4's frames change proportionally),
// undo/redo the size change, then edit deck metadata and confirm both the size
// preset and the metadata survive a download → reopen. The size half repeats
// against the .odp twin (odp does not serialise the first slide number).
const here = dirname(fileURLToPath(import.meta.url));
const PPTX = join(here, "fixtures", "mini-deck.pptx");
const ODP = join(here, "fixtures", "mini-deck.odp");

async function openFixture(page: Page, file: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function stageAspect(page: Page): Promise<number> {
  const box = await page.getByTestId("stage-workspace").getByTestId("slide-stage").boundingBox();
  if (!box) throw new Error("no stage box");
  return box.width / box.height;
}

async function openDeckSetup(page: Page): Promise<void> {
  // Deck setup… left the deck's action cluster when it folded into the app bar
  // (the eight controls overflowed a 1024px bar). It is a File-menu command
  // now — where it had always also been — so this reaches it the way a user
  // does. The dialog and everything below are unchanged.
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("slides-menu-deck-setup").click();
  await expect(page.getByTestId("deck-setup-dialog")).toBeVisible();
}

async function applyPreset43Scale(page: Page): Promise<void> {
  await openDeckSetup(page);
  await page.getByTestId("deck-size-rescale-scale").check();
  await page.getByTestId("deck-size-preset").selectOption("4:3");
  await page.getByTestId("deck-setup-apply").click();
  await expect(page.getByTestId("deck-setup-dialog")).toBeHidden();
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("pptx: resize to 4:3, undo/redo, edit metadata, and round-trip", async ({ page }) => {
  await openFixture(page, PPTX);

  // Select slide 4 and record a frame's box before resizing.
  await page.getByTestId("filmstrip-slide").nth(3).click();
  const frame = page.getByTestId("stage-workspace").getByTestId("slide-frame").first();
  await expect(frame).toBeVisible();
  const wideAspect = await stageAspect(page);
  const frameWide = await frame.boundingBox();
  if (!frameWide) throw new Error("no frame box");

  // Deck setup → 4:3 standard + Scale content → Apply.
  await applyPreset43Scale(page);
  await expect.poll(() => stageAspect(page)).toBeLessThan(wideAspect - 0.1);
  const squareAspect = await stageAspect(page);
  expect(squareAspect).toBeGreaterThan(1.2);
  expect(squareAspect).toBeLessThan(1.45); // ~4:3 = 1.333
  // Slide 4's frame geometry changed with the deck size.
  await expect
    .poll(async () => (await frame.boundingBox())?.width ?? 0)
    .not.toBe(frameWide.width);

  // Undo restores the widescreen aspect; redo re-applies 4:3.
  await page.getByTestId("presentation-undo").click();
  await expect.poll(() => stageAspect(page)).toBeGreaterThan(wideAspect - 0.05);
  await page.getByTestId("presentation-redo").click();
  await expect.poll(() => stageAspect(page)).toBeLessThan(wideAspect - 0.1);

  // Properties tab → set author + subject → Apply.
  await openDeckSetup(page);
  await page.getByTestId("deck-setup-tab-properties").click();
  await page.getByTestId("deck-properties-author").fill("Grace Hopper");
  await page.getByTestId("deck-properties-subject").fill("Quarterly review");
  await page.getByTestId("deck-setup-apply").click();
  await expect(page.getByTestId("deck-setup-dialog")).toBeHidden();

  // Download, reopen: the preset reads 4:3 and the metadata persisted.
  const file = await downloadTo(page, "spec090-mini-deck.pptx");
  await openFixture(page, file);
  await openDeckSetup(page);
  await expect(page.getByTestId("deck-size-preset")).toHaveValue("4:3");
  await page.getByTestId("deck-setup-tab-properties").click();
  await expect(page.getByTestId("deck-properties-author")).toHaveValue("Grace Hopper");
  await expect(page.getByTestId("deck-properties-subject")).toHaveValue("Quarterly review");
});

test("odp: resize to 4:3 with scale, undo and redo", async ({ page }) => {
  await openFixture(page, ODP);
  await page.getByTestId("filmstrip-slide").nth(3).click();
  const wideAspect = await stageAspect(page);

  await applyPreset43Scale(page);
  await expect.poll(() => stageAspect(page)).toBeLessThan(wideAspect - 0.1);

  await page.getByTestId("presentation-undo").click();
  await expect.poll(() => stageAspect(page)).toBeGreaterThan(wideAspect - 0.05);
  await page.getByTestId("presentation-redo").click();
  await expect.poll(() => stageAspect(page)).toBeLessThan(wideAspect - 0.1);

  // The resized 4:3 preset survives a download → reopen.
  const file = await downloadTo(page, "spec090-mini-deck.odp");
  await openFixture(page, file);
  await openDeckSetup(page);
  await expect(page.getByTestId("deck-size-preset")).toHaveValue("4:3");
});
