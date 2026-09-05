import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";
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

/** The on-screen boxes of a group wrapper's child frames, keyed off the
 * group wrapper with the smallest x (the authored test group sits far left
 * of the fixture's two verbatim groups). */
async function leftmostGroupChildBoxes(page: Page): Promise<{ x: number; y: number; width: number; height: number }[]> {
  const wrappers = page.getByTestId("stage-workspace").locator(".group-frame");
  const count = await wrappers.count();
  let best = -1;
  let bestX = Number.POSITIVE_INFINITY;
  for (let i = 0; i < count; i++) {
    const box = await wrappers.nth(i).boundingBox();
    if (box && box.x < bestX) {
      bestX = box.x;
      best = i;
    }
  }
  if (best < 0) throw new Error("no group wrapper found");
  const children = wrappers.nth(best).locator(".frame");
  const boxes = [];
  for (let i = 0; i < (await children.count()); i++) {
    const box = await children.nth(i).boundingBox();
    if (!box) throw new Error("missing group child box");
    boxes.push(box);
  }
  return boxes;
}

// AC-27 regression (QA round 1): Escape mid-rotation-drag must restore the
// starting angle even after SEVERAL pointermoves — the live preview's frame
// object is re-identified by every refresh, so the restore has to reach the
// current one, not the pointer-down capture.
test("Escape mid-rotation-drag restores the starting angle", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);
  await rectangles(page).nth(0).click(); // the blue rect at (2,2) 4×2 — centre (4,3)
  const handle = page.getByTestId("rotate-handle");
  await expect(handle).toBeVisible();
  const handleBox = await handle.boundingBox();
  if (!handleBox) throw new Error("no rotate handle box");

  await page.mouse.move(handleBox.x + handleBox.width / 2, handleBox.y + handleBox.height / 2);
  await page.mouse.down();
  const east = await cmPoint(page, 8, 3); // due east of the centre → ~90°
  await page.mouse.move(east.x, east.y, { steps: 5 });
  const south = await cmPoint(page, 4, 6); // then due south → ~180°
  await page.mouse.move(south.x, south.y, { steps: 5 });
  await expect(page.getByTestId("rotation-readout")).toHaveText(/\d+°/);
  await page.keyboard.press("Escape");
  await page.mouse.up();

  // The gesture recorded nothing and the angle is back at 0.
  await expect(page.getByTestId("rotation-readout")).toHaveText("");
  await page.getByTestId("slides-position-toggle").click();
  await expect(page.getByTestId("position-rotation")).toHaveValue("0");
  await page.getByTestId("position-close").click();
});

// AC-45/AC-69 regression (QA round 1): rotating an authored group must move
// its children (the rotation composes onto them about the GROUP's centre, not
// the stage origin), and one undo restores geometry and children together.
test("rotating an authored group moves its children about the group centre; undo restores them", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);
  await rectangles(page).nth(0).click();
  await rectangles(page).nth(1).click({ modifiers: ["Shift"] });
  await page.getByTestId("group-button").click();
  await expect(stageFrames(page)).toHaveCount(6); // 7 - 2 rects + 1 group

  const before = await leftmostGroupChildBoxes(page);
  expect(before).toHaveLength(2);

  // Drag the rotate handle to a point east of the group centre (7, 4.5) → ~90°.
  const handle = page.getByTestId("rotate-handle");
  await expect(handle).toBeVisible();
  const handleBox = await handle.boundingBox();
  if (!handleBox) throw new Error("no rotate handle box");
  await page.mouse.move(handleBox.x + handleBox.width / 2, handleBox.y + handleBox.height / 2);
  await page.mouse.down();
  const target = await cmPoint(page, 12, 4.5);
  await page.mouse.move(target.x, target.y, { steps: 5 });
  await page.mouse.up();

  // Both children moved (rigid rotation about the group centre).
  const after = await leftmostGroupChildBoxes(page);
  expect(after).toHaveLength(2);
  const moved = (i: number) =>
    Math.hypot(
      after[i].x + after[i].width / 2 - (before[i].x + before[i].width / 2),
      after[i].y + after[i].height / 2 - (before[i].y + before[i].height / 2),
    );
  expect(moved(0)).toBeGreaterThan(20);
  expect(moved(1)).toBeGreaterThan(20);

  // One undo restores the children's geometry together (AC-69).
  await page.keyboard.press(`${MOD}+z`);
  const restored = await leftmostGroupChildBoxes(page);
  for (let i = 0; i < 2; i++) {
    expect(Math.abs(restored[i].x - before[i].x)).toBeLessThan(3);
    expect(Math.abs(restored[i].y - before[i].y)).toBeLessThan(3);
  }
});

// AC-42/43/44 (blocked in QA round 1 by the 0×0 group-wrapper regression):
// both verbatim groups render at their real geometry and can be clicked; a
// 0°-rotation group releases silently while the rotated one parks in the
// confirm, whose cancel changes nothing.
test("verbatim groups select by click and release through the rotation confirm", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);
  const groups = page.getByTestId("stage-workspace").getByRole("group", { name: "Group", exact: true });
  await expect(groups).toHaveCount(2);

  // The non-identity group's child maps to (20,12)→(22,15) in slide space —
  // clicking it selects the group, whose wrapper carries the real box.
  const inChild = await cmPoint(page, 21, 13.5);
  await page.mouse.click(inChild.x, inChild.y);
  await expect(page.getByTestId("frame-selection")).toBeVisible();

  // rotation 0 → releases silently, no confirm (AC-44).
  await page.getByTestId("ungroup-button").click();
  await expect(page.getByTestId("group-release-confirm")).toHaveCount(0);
  await expect(groups).toHaveCount(1);
  await expect(stageFrames(page)).toHaveCount(7); // one child promoted

  // The rotated group's child (28,12 2×3, spun 45° about the group centre) —
  // click its transformed centre.
  const rotatedChild = page.getByTestId("stage-workspace").locator(".group-frame .frame").first();
  await rotatedChild.click();
  await expect(page.getByTestId("frame-selection")).toBeVisible();

  // Releasing the rotated VERBATIM group parks in the confirm (AC-42/43).
  await page.getByTestId("ungroup-button").click();
  const confirm = page.getByRole("alertdialog", { name: "Replace the original grouping" });
  await expect(confirm).toBeVisible();
  await confirm.getByTestId("group-release-cancel").click();
  await expect(groups).toHaveCount(1); // cancel changes nothing

  await page.getByTestId("ungroup-button").click();
  await confirm.getByTestId("group-release-confirm").click();
  await expect(groups).toHaveCount(0);
  await expect(stageFrames(page)).toHaveCount(7);
});
