import { test, expect, type Page } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-057 — native drawing-group containers. The SPEC-051 fixtures carry a
// legacy `de:group` pair that migrates on open to one native DrawingGroup
// container (AC-65), which renders as a selectable, enterable group (AC-77).
// The rest is exercised through the create/ungroup gestures on a fresh
// document, so it needs no new binary fixtures.
//
// Fixture drawing-index map after migration (see e2e/SPEC-051.spec.ts):
//   0..13 catalogue shapes, 14 migrated group container, 15 text box,
//   16 unsupported geometry.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxShapesFixture = join(fixturesDir, "word-shapes.docx");
const odtShapesFixture = join(fixturesDir, "lo-shapes.odt");

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]:not([data-testid*="-member-"])');
const drawingAt = (page: Page, index: number) => content(page).locator(`[data-testid="drawing-${index}"]`);

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function insertShapeAt(page: Page, geometry: string, x: number, y: number): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`shape-gallery-${geometry}`).click();
  await content(page).click({ position: { x, y } });
}

for (const [label, fixture] of [
  ["word-shapes.docx", docxShapesFixture],
  ["lo-shapes.odt", odtShapesFixture],
] as const) {
  test.describe(`SPEC-057 — migrated ${label} opens as a native container`, () => {
    test("the legacy de:group pair renders as one container, selectable and enterable (AC-65, AC-77)", async ({ page }) => {
      await openFixture(page, fixture);
      const container = drawingAt(page, 14);
      await expect(container).toHaveAttribute("role", "group");
      const memberA = page.locator('[data-testid="drawing-14-member-0"]');
      await expect(memberA).toHaveCount(1);
      await expect(page.locator('[data-testid="drawing-14-member-1"]')).toHaveCount(1);

      // First click selects the outermost container (AC-22): its toolbar shows.
      await memberA.click();
      await expect(page.getByTestId("group-toolbar")).toBeVisible();
      await expect(page.getByTestId("group-handles")).toBeVisible();

      // A second click descends to the member under the pointer (AC-23) — the
      // member proxy becomes focusable (tabindex 0) once its container is
      // entered.
      await memberA.click();
      await expect(memberA).toHaveAttribute("tabindex", "0");

      // Escape steps back out one level (AC-24): the container is selected
      // again, its member no longer focusable.
      await page.keyboard.press("Escape");
      await expect(page.getByTestId("group-toolbar")).toBeVisible();
      await expect(memberA).toHaveAttribute("tabindex", "-1");
    });
  });
}

test.describe("SPEC-057 — Group and Ungroup gestures", () => {
  test("Group replaces two shapes with a container; Ctrl+Shift+G ungroups them back (AC-3, AC-5)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    await expect(drawings(page)).toHaveCount(2);

    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
    const [leftBox, rightBox] = firstBox.x <= secondBox.x ? [firstBox, secondBox] : [secondBox, firstBox];
    await leftShape.click({ position: { x: leftBox.width / 2, y: leftBox.height * 0.72 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await rightShape.click({ modifiers: ["Shift"], position: { x: rightBox.width / 2, y: rightBox.height * 0.72 } });

    // Ctrl/Cmd+G groups the multi-selection into one container.
    await page.keyboard.press("ControlOrMeta+g");
    await expect(drawings(page)).toHaveCount(1);
    await expect(drawingAt(page, 0)).toHaveAttribute("role", "group");
    await expect(page.getByTestId("group-toolbar")).toBeVisible();

    // Ctrl/Cmd+Shift+G ungroups one level: the two shapes re-emerge.
    await page.keyboard.press("ControlOrMeta+Shift+g");
    await expect(drawings(page)).toHaveCount(2);
  });

  test("a second click selects a member, whose toolbar offers the reorder verbs; Delete dissolves a two-member container (AC-23, AC-31/32, AC-33..36)", async ({ page }) => {
    await openFixture(page, docxShapesFixture);
    const member0 = page.locator('[data-testid="drawing-14-member-0"]');
    // First click selects the container, second descends to the member.
    await member0.click();
    await member0.click();
    await expect(page.getByTestId("group-member-reorder")).toBeVisible();
    // Send it to the back — reorders it among its siblings (member order is
    // paint order).
    await page.getByTestId("group-member-reorder").selectOption("back");
    // Deleting a member of a two-member container dissolves it into the
    // remaining member: the container proxy at index 14 is gone, replaced by a
    // plain shape.
    await page.locator('[data-testid="drawing-14-member-0"]').click();
    await page.locator('[data-testid="drawing-14-member-0"]').click();
    await page.keyboard.press("Delete");
    await expect(drawingAt(page, 14)).toHaveAttribute("role", "img");
  });

  test("the container carries editable alt text from its toolbar (AC-49)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
    await leftShape.click({ position: { x: firstBox.width / 2, y: firstBox.height * 0.72 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await rightShape.click({ modifiers: ["Shift"], position: { x: secondBox.width / 2, y: secondBox.height * 0.72 } });
    await page.keyboard.press("ControlOrMeta+g");

    const altInput = page.getByTestId("group-alt-text");
    await expect(altInput).toBeVisible();
    await altInput.fill("My diagram");
    // The container proxy reflects the alt text as its aria-label.
    await expect(drawingAt(page, 0)).toHaveAttribute("aria-label", "My diagram");
  });
});
