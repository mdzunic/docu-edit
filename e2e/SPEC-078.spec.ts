import { test, expect, type Page, type Locator } from "@playwright/test";

// SPEC-078 — explicit anchors for floating drawings: two shapes anchored to
// one paragraph both mark it as their anchor, re-anchoring one by drag
// leaves the other's rendered position pixel-identical, and both pairings
// (with their offsets) survive a save → reopen in each format.
//
// The setup deliberately reproduces the ambiguity SPEC-078 exists to kill:
// consecutive shape blocks beside one paragraph, where the old
// adjacency-only inference flipped the second drawing's anchor to the
// following paragraph.

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]');
const drawingAt = (page: Page, index: number) => content(page).locator(`[data-testid="drawing-${index}"]`);
const anchorMarker = (page: Page) => page.getByTestId("shape-anchor-marker");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** Three real paragraphs (typed with real Enter keys, so the model holds
 * three blocks, not one block with line breaks). */
async function typeThreeParagraphs(page: Page): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.keyboard.type("First anchor paragraph with some words in it.");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Second paragraph sits in the middle of the page.");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Third paragraph takes the drop near the bottom.");
}

/** "Insert -> Shape" with the caret first placed at the FIRST paragraph's
 * start, then a single click placing `geometry` at its default size — both
 * shapes in these tests anchor to that same first paragraph. */
async function insertShapeOnFirstParagraph(page: Page, geometry: string, x: number, y: number): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`shape-gallery-${geometry}`).click();
  await content(page).click({ position: { x, y } });
}

async function boxOf(locator: Locator): Promise<{ x: number; y: number; width: number; height: number }> {
  const box = await locator.boundingBox();
  if (!box) {
    throw new Error("locator has no bounding box");
  }
  return box;
}

/** Asserts the selected shape's anchor marker sits at `paragraph`'s
 * top-left corner (the marker is a small flag drawn there). */
async function expectMarkerOnParagraph(page: Page, paragraph: Locator): Promise<void> {
  await expect(anchorMarker(page)).toBeVisible();
  const marker = await boxOf(anchorMarker(page));
  const target = await boxOf(paragraph);
  expect(Math.abs(marker.y - target.y)).toBeLessThan(14);
  expect(Math.abs(marker.x - target.x)).toBeLessThan(24);
}

async function dragBy(page: Page, locator: Locator, dx: number, dy: number): Promise<void> {
  const box = await boxOf(locator);
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + dx, box.y + box.height / 2 + dy, { steps: 6 });
  await page.mouse.up();
}

test.describe("SPEC-078 — two drawings, one anchor paragraph", () => {
  test("a drawn shape lands exactly where the pointer released — x AND y (the SPEC-051 known bug, fixed)", async ({
    page,
  }) => {
    await createDocument(page);
    await typeThreeParagraphs(page);
    const contentBox = await boxOf(content(page));
    await insertShapeOnFirstParagraph(page, "rect", 320, 260);
    const drawn = await boxOf(drawingAt(page, 0));
    expect(Math.abs(drawn.x - (contentBox.x + 320))).toBeLessThanOrEqual(1.5);
    expect(Math.abs(drawn.y - (contentBox.y + 260))).toBeLessThanOrEqual(1.5);
  });

  test("both shapes' anchor markers point at the shared first paragraph", async ({ page }) => {
    await createDocument(page);
    await typeThreeParagraphs(page);
    await insertShapeOnFirstParagraph(page, "rect", 320, 260);
    await insertShapeOnFirstParagraph(page, "ellipse", 460, 260);
    await expect(drawings(page)).toHaveCount(2);
    const firstParagraph = content(page).locator("p", { hasText: "First anchor paragraph" });

    await drawingAt(page, 0).click();
    await expectMarkerOnParagraph(page, firstParagraph);
    await drawingAt(page, 1).click();
    await expectMarkerOnParagraph(page, firstParagraph);
  });

  test("dragging one of the two re-anchors only it — the other's box is pixel-identical", async ({ page }) => {
    await createDocument(page);
    await typeThreeParagraphs(page);
    await insertShapeOnFirstParagraph(page, "rect", 320, 260);
    await insertShapeOnFirstParagraph(page, "ellipse", 460, 260);
    const rect = drawingAt(page, 0);
    const ellipse = drawingAt(page, 1);
    const rectBefore = await boxOf(rect);
    const ellipseBefore = await boxOf(ellipse);

    // Down far enough to cross into the third paragraph's territory.
    await dragBy(page, ellipse, 0, 140);

    const ellipseAfter = await boxOf(ellipse);
    expect(Math.abs(ellipseAfter.y - (ellipseBefore.y + 140))).toBeLessThanOrEqual(2);
    expect(Math.abs(ellipseAfter.x - ellipseBefore.x)).toBeLessThanOrEqual(2);
    // The sibling never moves — the regression the explicit pairing kills:
    // before SPEC-078 this drag spliced blocks in a way that flipped the
    // OTHER drawing's inferred anchor and jumped its rendered position.
    const rectAfter = await boxOf(rect);
    expect(Math.abs(rectAfter.x - rectBefore.x)).toBeLessThanOrEqual(1);
    expect(Math.abs(rectAfter.y - rectBefore.y)).toBeLessThanOrEqual(1);

    // The dragged shape's marker moved off the first paragraph — it now
    // marks a paragraph below it; the sibling still marks the first.
    const firstParagraph = content(page).locator("p", { hasText: "First anchor paragraph" });
    const firstBox = await boxOf(firstParagraph);
    const movedMarker = await boxOf(anchorMarker(page));
    expect(movedMarker.y).toBeGreaterThan(firstBox.y + 10);
    await rect.click();
    await expectMarkerOnParagraph(page, firstParagraph);
  });
});

/** Downloads the current document in `format` and reopens the file from
 * the library screen, returning once the editor is back. */
async function saveAndReopen(page: Page, format: "docx" | "odt"): Promise<void> {
  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId(`download-${format}`).click();
  const download = await downloadPromise;
  const path = await download.path();
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

for (const format of ["docx", "odt"] as const) {
  test(`SPEC-078 — both pairings and positions survive save and reopen (.${format})`, async ({ page }) => {
    await createDocument(page);
    await typeThreeParagraphs(page);
    await insertShapeOnFirstParagraph(page, "rect", 320, 260);
    await insertShapeOnFirstParagraph(page, "ellipse", 460, 300);
    await expect(drawings(page)).toHaveCount(2);
    const contentBefore = await boxOf(content(page));
    const rectBefore = await boxOf(drawingAt(page, 0));
    const ellipseBefore = await boxOf(drawingAt(page, 1));

    await saveAndReopen(page, format);
    await expect(drawings(page)).toHaveCount(2);

    // Positions are compared content-relative: the reopened page may sit at
    // a different scroll/viewport offset, but each shape's place ON the
    // page must be identical.
    const contentAfter = await boxOf(content(page));
    await expect
      .poll(async () => {
        const box = await boxOf(drawingAt(page, 0));
        return Math.abs(box.y - contentAfter.y - (rectBefore.y - contentBefore.y));
      })
      .toBeLessThanOrEqual(2);
    const rectAfter = await boxOf(drawingAt(page, 0));
    const ellipseAfter = await boxOf(drawingAt(page, 1));
    expect(Math.abs(rectAfter.x - contentAfter.x - (rectBefore.x - contentBefore.x))).toBeLessThanOrEqual(2);
    expect(Math.abs(ellipseAfter.x - contentAfter.x - (ellipseBefore.x - contentBefore.x))).toBeLessThanOrEqual(2);
    expect(Math.abs(ellipseAfter.y - contentAfter.y - (ellipseBefore.y - contentBefore.y))).toBeLessThanOrEqual(2);

    // The PAIRING survived, not just the offsets: dragging one shape after
    // the reopen still leaves the other pixel-identical.
    await dragBy(page, drawingAt(page, 1), 0, 140);
    const rectFinal = await boxOf(drawingAt(page, 0));
    expect(Math.abs(rectFinal.x - rectAfter.x)).toBeLessThanOrEqual(1);
    expect(Math.abs(rectFinal.y - rectAfter.y)).toBeLessThanOrEqual(1);
  });
}
