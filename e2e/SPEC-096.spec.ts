import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PPTX = join(here, "fixtures", "spec096-objects.pptx");

// SPEC-096 — slide object depth: open the objects fixture, marquee the three
// unaligned shapes (selection-bounds, frame count unchanged), Align left,
// Distribute vertically (gaps equalise), Group to one frame and Ungroup back
// to three, rotate the image with Shift (readout a multiple of 15°), set an
// exact 90° through Position & Size, then download → reopen and confirm the
// rotation, the alignment and the tuned roundRect adjustment all survived.

// Fixture geometry (cm) — see fixtures/build.ts SPEC096_* constants.
const SLIDE_W = 33.867;
const SLIDE_H = 19.05;

function stageFrames(page: Page) {
  return page.getByTestId("stage-workspace").getByTestId("slide-frame");
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

/** A slide-space cm point as viewport px, via the stage's on-screen box. */
async function cmPoint(page: Page, xCm: number, yCm: number): Promise<{ x: number; y: number }> {
  const box = await page.getByTestId("stage-workspace").getByTestId("slide-stage").boundingBox();
  if (!box) throw new Error("no stage box");
  return { x: box.x + (box.width * xCm) / SLIDE_W, y: box.y + (box.height * yCm) / SLIDE_H };
}

function rectangles(page: Page) {
  return page.getByTestId("stage-workspace").getByRole("group", { name: "Rectangle" });
}

test("marquee, align, distribute, group/ungroup, rotate, position dialog, round trip", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);
  const frameCount = await stageFrames(page).count();
  expect(frameCount).toBe(7); // 3 rects + image + roundRect + 2 verbatim groups

  // --- marquee-drag over the three shapes -----------------------------------
  const from = await cmPoint(page, 1, 1);
  const to = await cmPoint(page, 19.5, 14.5);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 10 });
  await expect(page.getByTestId("marquee-box")).toBeVisible();
  await page.mouse.up();
  await expect(page.getByTestId("selection-bounds")).toBeVisible();
  await expect(stageFrames(page)).toHaveCount(frameCount); // selection only

  // --- Align → "Align left": their left edges match -------------------------
  await page.getByTestId("align-menu").click();
  await expect(page.getByTestId("align-popover")).toBeVisible();
  await page.getByTestId("align-left").click();
  const xs: number[] = [];
  for (let i = 0; i < 3; i++) {
    const box = await rectangles(page).nth(i).boundingBox();
    xs.push(box?.x ?? -1);
  }
  expect(Math.max(...xs) - Math.min(...xs)).toBeLessThan(1.5);

  // --- "Distribute vertically": gaps equalise -------------------------------
  await page.getByTestId("align-menu").click();
  await page.getByTestId("distribute-v").click();
  const boxes = [];
  for (let i = 0; i < 3; i++) {
    const box = await rectangles(page).nth(i).boundingBox();
    if (!box) throw new Error("missing rect box");
    boxes.push(box);
  }
  boxes.sort((a, b) => a.y - b.y);
  const gap1 = boxes[1].y - (boxes[0].y + boxes[0].height);
  const gap2 = boxes[2].y - (boxes[1].y + boxes[1].height);
  expect(Math.abs(gap1 - gap2)).toBeLessThan(2);
  expect(gap1).toBeGreaterThan(10); // visibly spaced (2.5cm ≈ 30+ px)

  // --- Group → one frame remains; Ungroup restores three --------------------
  await page.getByTestId("group-button").click();
  await expect(stageFrames(page)).toHaveCount(frameCount - 2);
  await page.getByTestId("ungroup-button").click();
  await expect(stageFrames(page)).toHaveCount(frameCount);

  // --- rotate the image with Shift: readout reads a multiple of 15 ----------
  const image = page.getByTestId("stage-workspace").getByRole("group", { name: "Image" });
  await image.click();
  const handle = page.getByTestId("rotate-handle");
  await expect(handle).toBeVisible();
  const handleBox = await handle.boundingBox();
  if (!handleBox) throw new Error("no rotate handle box");
  await page.keyboard.down("Shift");
  await page.mouse.move(handleBox.x + handleBox.width / 2, handleBox.y + handleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(handleBox.x + 90, handleBox.y + 60, { steps: 6 });
  const readout = (await page.getByTestId("rotation-readout").textContent()) ?? "";
  const angle = Number(readout.replace("°", "").trim());
  expect(Number.isFinite(angle)).toBe(true);
  expect(angle % 15).toBe(0);
  await page.mouse.up();
  await page.keyboard.up("Shift");

  // --- Position & Size…: set rotation exactly 90 ----------------------------
  await page.getByTestId("slides-position-toggle").click();
  await expect(page.getByTestId("slides-position-dialog")).toBeVisible();
  await page.getByTestId("position-rotation").fill("90");
  await page.getByTestId("position-rotation").press("Enter");
  await page.getByTestId("position-close").click();

  // --- download → reopen: rotation, alignment, adjustment all survive -------
  const file = await downloadTo(page, "spec096-objects-roundtrip.pptx");
  await openFixture(page, file);
  await expect(stageFrames(page)).toHaveCount(7);

  // Rotation survived: the image's Position & Size dialog reads 90.
  await page.getByTestId("stage-workspace").getByRole("group", { name: "Image" }).click();
  await page.getByTestId("slides-position-toggle").click();
  await expect(page.getByTestId("position-rotation")).toHaveValue("90");
  await page.getByTestId("position-close").click();

  // Alignment survived: the three rectangles share a left edge.
  const xsBack: number[] = [];
  for (let i = 0; i < 3; i++) {
    const box = await rectangles(page).nth(i).boundingBox();
    xsBack.push(box?.x ?? -1);
  }
  expect(Math.max(...xsBack) - Math.min(...xsBack)).toBeLessThan(1.5);

  // The tuned adjustment survived: the roundRect's diamond reads 40 (0.4).
  await page.getByTestId("stage-workspace").getByRole("group", { name: "Rounded rectangle" }).click();
  const adjust = page.getByTestId("adjust-handle");
  await expect(adjust).toBeVisible();
  await expect(adjust).toHaveAttribute("aria-valuenow", "40");
});
