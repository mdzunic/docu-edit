import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzip } from "../src/lib/core/zip";

// SPEC-131 — images and shapes on sheets (AC-52), with the real pointer and
// keyboard: open spec131-drawings.xlsx → the picture, the rounded
// rectangle's text and the group paint; the slicer shows Excel's fallback
// text box and the form control its "Form control" placeholder; the hidden
// shape does not paint → select the picture (eight handles) → arrow-nudge it
// → insert two rows above it through the row header menu (SPEC-124) → it
// moved down by exactly the inserted height → Insert ▸ Shapes ▸ Rectangle,
// drag a box → Send to Back → Alt Text… → Insert ▸ Image with
// spec107-shot.png → download, reopen → every drawing is where it was left,
// and the untouched slicer's anchor XML is unchanged apart from its rows.
// The same flow on spec131-drawings.ods (its own drawings: the image, the
// "Assumptions" box, the group and the form control's placeholder).

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 } });

const decoder = new TextDecoder();

function drawing(page: Page, name: string): Locator {
  return page.locator(`[data-testid="sheet-drawing"][data-primary="true"][data-drawing-name="${name}"]`);
}

function rowHeader(page: Page, row: number): Locator {
  return page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${row}$`) });
}

async function open(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
}

async function download(page: Page, fileName: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const saved = await downloadPromise;
  const path = join(tmpdir(), `spec131-${Date.now()}-${fileName}`);
  await saved.saveAs(path);
  return path;
}

type Box = { x: number; y: number; width: number; height: number };

async function box(locator: Locator): Promise<Box> {
  const found = await locator.boundingBox();
  expect(found).not.toBeNull();
  return found!;
}

function expectNear(actual: Box, expected: Box, tolerance = 1.5): void {
  expect(Math.abs(actual.x - expected.x), `x ${actual.x} vs ${expected.x}`).toBeLessThanOrEqual(tolerance);
  expect(Math.abs(actual.y - expected.y), `y ${actual.y} vs ${expected.y}`).toBeLessThanOrEqual(tolerance);
  expect(Math.abs(actual.width - expected.width), `width ${actual.width} vs ${expected.width}`).toBeLessThanOrEqual(tolerance);
  expect(Math.abs(actual.height - expected.height), `height ${actual.height} vs ${expected.height}`).toBeLessThanOrEqual(tolerance);
}

/** Selects rows `first`..`last` through their headers and inserts as many above. */
async function insertRowsAbove(page: Page, first: number, last: number): Promise<void> {
  await rowHeader(page, first).click();
  await rowHeader(page, last).click({ modifiers: ["Shift"] });
  await rowHeader(page, first).click({ button: "right" });
  const menu = page.getByTestId("sheet-header-menu");
  await expect(menu).toBeVisible();
  const insert = menu.getByTestId("sheet-insert-before");
  await expect(insert).toHaveText(`Insert ${last - first + 1} rows above`);
  await insert.click();
}

/** Insert ▸ Shapes ▸ Rectangle, then a drag on the grid from `from` to `to` (page px). */
async function drawRectangle(page: Page, from: { x: number; y: number }, to: { x: number; y: number }): Promise<Locator> {
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("sheet-menu-insert-shape").click();
  const popover = page.getByTestId("sheet-shape-gallery-popover");
  await expect(popover).toBeVisible();
  // The catalogue only: no connectors or freeform tools on a sheet.
  await expect(popover.getByTestId("connector-gallery-group")).toHaveCount(0);
  await expect(popover.getByTestId("freeform-gallery-group")).toHaveCount(0);
  await popover.getByTestId("shape-gallery-rect").click();
  await expect(popover).toBeHidden();
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move((from.x + to.x) / 2, (from.y + to.y) / 2, { steps: 4 });
  await page.mouse.move(to.x, to.y, { steps: 4 });
  await page.mouse.up();
  const created = page.locator(`[data-testid="sheet-drawing"][data-primary="true"][data-kind="shape"][aria-selected="true"]`);
  await expect(created).toBeVisible();
  return created;
}

async function contextMenu(page: Page, target: Locator, entry: string): Promise<void> {
  await target.click({ button: "right" });
  const item = page.getByTestId(entry);
  await expect(item).toBeVisible();
  await item.click();
}

async function setAltText(page: Page, target: Locator, text: string): Promise<void> {
  await contextMenu(page, target, "sheet-drawing-alt-text");
  const input = page.getByTestId("sheet-drawing-alt-text-input");
  await expect(input).toBeVisible();
  await input.fill(text);
  await page.getByTestId("sheet-drawing-alt-text-ok").click();
  await expect(input).toBeHidden();
}

async function insertImage(page: Page): Promise<Locator> {
  await page.getByTestId("menu-title-insert").click();
  const chooser = page.waitForEvent("filechooser");
  await page.getByTestId("sheet-menu-insert-image").click();
  await (await chooser).setFiles(join(fixtures, "spec107-shot.png"));
  const inserted = page.locator(`[data-testid="sheet-drawing"][data-primary="true"][data-kind="image"][aria-selected="true"]`);
  await expect(inserted).toBeVisible();
  return inserted;
}

/** The drawing names in z-order (the layer paints back to front). */
async function zOrder(page: Page): Promise<string[]> {
  const names = await page.locator(`[data-testid="sheet-drawing-layer"] [data-testid="sheet-drawing"][data-primary="true"]`).evaluateAll((els) => els.map((el) => el.getAttribute("data-drawing-name") ?? ""));
  return [...new Set(names)];
}

/** Every painted drawing's box by name. */
async function boxes(page: Page, names: readonly string[]): Promise<Map<string, Box>> {
  const out = new Map<string, Box>();
  for (const name of names) {
    out.set(name, await box(drawing(page, name)));
  }
  return out;
}

/** An anchor element's text with its rows blanked: the xdr:row numbers, and
 * the vertical offset of the content's own a:xfrm that follows them (the
 * writer gives the transform the anchor's box). */
const withoutRows = (xml: string): string =>
  xml.replace(/<xdr:row>\d+<\/xdr:row>/g, "<xdr:row>#</xdr:row>").replace(/(<a:off x="\d+" y=")\d+(")/g, "$1#$2");

/** The mc:AlternateContent wrapping the slicer, from a drawing part. */
function slicerWrapper(xml: string): string {
  const at = xml.indexOf(`Requires="sle15"`);
  expect(at).toBeGreaterThan(0);
  const start = xml.lastIndexOf("<mc:AlternateContent", at);
  const end = xml.indexOf("</mc:AlternateContent>", at) + "</mc:AlternateContent>".length;
  return xml.slice(start, end);
}

test("the xlsx drawing layer: paints, nudge, row insert, draw a rectangle, z-order, alt text, insert an image, round trip (AC-52)", async ({ page }) => {
  await open(page, join(fixtures, "spec131-drawings.xlsx"));
  const layer = page.getByTestId("sheet-drawing-layer").first();
  await expect(layer).toBeVisible();

  // The paintings.
  const picture = drawing(page, "Picture 1");
  await expect(picture).toBeVisible();
  await expect(picture).toHaveAttribute("role", "img");
  await expect(picture).toHaveAttribute("data-kind", "image");
  await expect(picture).toHaveAttribute("aria-label", "Company logo");
  await expect(picture.locator("img")).toBeVisible();
  await expect(drawing(page, "Rounded Rectangle 2")).toContainText("Assumptions");
  await expect(drawing(page, "Rounded Rectangle 2")).toHaveAttribute("role", "group");
  await expect(drawing(page, "Group 8")).toBeVisible();
  await expect(drawing(page, "Group 8")).toHaveAttribute("data-kind", "group");
  await expect(drawing(page, "Region")).toContainText("This shape represents a table slicer.");
  await expect(drawing(page, "Check Box 1")).toContainText("Form control");
  await expect(page.locator(`[data-testid="sheet-drawing"][data-drawing-name="Rectangle 9"]`)).toHaveCount(0);
  await expect(page.locator(`[data-testid="sheet-drawing"][data-drawing-name="Picture 12"] [data-testid="sheet-drawing-placeholder"]`).first()).toContainText("Picture");

  // Select the picture: aria-selected and eight handles.
  await picture.click();
  await expect(picture).toHaveAttribute("aria-selected", "true");
  for (const handle of ["nw", "n", "ne", "e", "se", "s", "sw", "w"]) {
    await expect(page.getByTestId(`sheet-drawing-handle-${handle}`)).toBeVisible();
  }
  const start = await box(picture);

  // Arrow-nudge: 1 px right, then Shift+Down 10 px.
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("Shift+ArrowDown");
  const nudged = await box(picture);
  expectNear(nudged, { ...start, x: start.x + 1, y: start.y + 10 });

  // Escape returns focus to the grid with the cell selection unchanged.
  const nameBox = page.getByTestId("name-box");
  const before = await nameBox.inputValue();
  await page.keyboard.press("Escape");
  await expect(picture).not.toHaveAttribute("aria-selected", "true");
  // Focus is back in the grid (its in-cell editor holds it).
  expect(await page.getByTestId("spreadsheet-grid").evaluate((grid) => grid.contains(grid.ownerDocument.activeElement))).toBe(true);
  await expect(nameBox).toHaveValue(before);

  // Two rows inserted above it: it moves down by exactly their height.
  const rowHeight = (await box(rowHeader(page, 1))).height;
  await insertRowsAbove(page, 1, 2);
  const moved = await box(picture);
  expectNear(moved, { ...nudged, y: nudged.y + 2 * rowHeight });

  // Draw a rectangle by drag over empty cells (rows 26–29, columns D–F).
  const d26 = await box(page.getByTestId("sheet-col-header").filter({ hasText: /^D$/ }));
  const r26 = await box(rowHeader(page, 26));
  const from = { x: d26.x + 10, y: r26.y + 5 };
  const rectangle = await drawRectangle(page, from, { x: from.x + 150, y: from.y + 60 });
  const drawn = await box(rectangle);
  expectNear(drawn, { x: from.x, y: from.y, width: 150, height: 60 }, 3);
  const rectangleName = (await rectangle.getAttribute("data-drawing-name"))!;
  expect(rectangleName).toMatch(/^Rectangle \d+$/);

  // Send to Back: it is first in the z-order.
  await contextMenu(page, rectangle, "sheet-drawing-back");
  await expect.poll(async () => (await zOrder(page))[0]).toBe(rectangleName);

  // Alt text.
  await setAltText(page, drawing(page, rectangleName), "Highlight box");
  await expect(drawing(page, rectangleName)).toHaveAttribute("aria-label", "Highlight box");

  // Insert ▸ Image at the active cell.
  await page.keyboard.press("Escape");
  await page.getByTestId("sheet-col-header").filter({ hasText: /^H$/ }).click();
  const image = await insertImage(page);
  const imageName = (await image.getAttribute("data-drawing-name"))!;
  expect(imageName).toMatch(/^Picture \d+$/);

  // Download, reopen: every drawing where it was left.
  const names = ["Picture 1", "Rounded Rectangle 2", "Group 8", "Region", "Check Box 1", rectangleName, imageName];
  const left = await boxes(page, names);
  const path = await download(page, "spec131-drawings.xlsx");
  const saved = await unzip(new Uint8Array(readFileSync(path)));
  const source = await unzip(new Uint8Array(readFileSync(join(fixtures, "spec131-drawings.xlsx"))));
  const savedDrawing = decoder.decode(saved.get("xl/drawings/drawing1.xml"));
  const sourceDrawing = decoder.decode(source.get("xl/drawings/drawing1.xml"));
  // The untouched slicer's anchor XML is unchanged apart from its rows.
  expect(withoutRows(slicerWrapper(savedDrawing))).toBe(withoutRows(slicerWrapper(sourceDrawing)));
  expect(slicerWrapper(savedDrawing)).not.toBe(slicerWrapper(sourceDrawing));

  await open(page, path);
  await expect(drawing(page, "Picture 1")).toBeVisible();
  const reopened = await boxes(page, names);
  for (const name of names) {
    expectNear(reopened.get(name)!, left.get(name)!);
  }
  await expect(drawing(page, rectangleName)).toHaveAttribute("aria-label", "Highlight box");
  expect((await zOrder(page))[0]).toBe(rectangleName);
});

test("the ods drawing layer: the same flow on spec131-drawings.ods (AC-52)", async ({ page }) => {
  await open(page, join(fixtures, "spec131-drawings.ods"));
  const picture = drawing(page, "Image 1");
  await expect(picture).toBeVisible();
  await expect(picture).toHaveAttribute("role", "img");
  await expect(picture).toHaveAttribute("aria-label", "Company logo");
  await expect(drawing(page, "Assumptions box")).toContainText("Assumptions");
  await expect(drawing(page, "Group 5")).toHaveAttribute("data-kind", "group");
  await expect(drawing(page, "Check Box 1")).toContainText("Form control");

  await picture.click();
  await expect(picture).toHaveAttribute("aria-selected", "true");
  const start = await box(picture);
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("Shift+ArrowDown");
  const nudged = await box(picture);
  expectNear(nudged, { ...start, x: start.x + 1, y: start.y + 10 });
  await page.keyboard.press("Escape");

  const rowHeight = (await box(rowHeader(page, 1))).height;
  await insertRowsAbove(page, 1, 2);
  expectNear(await box(picture), { ...nudged, y: nudged.y + 2 * rowHeight });

  const d = await box(page.getByTestId("sheet-col-header").filter({ hasText: /^D$/ }));
  const r = await box(rowHeader(page, 28));
  const from = { x: d.x + 10, y: r.y + 4 };
  const rectangle = await drawRectangle(page, from, { x: from.x + 140, y: from.y + 50 });
  const rectangleName = (await rectangle.getAttribute("data-drawing-name"))!;
  await contextMenu(page, rectangle, "sheet-drawing-back");
  await expect.poll(async () => (await zOrder(page))[0]).toBe(rectangleName);
  await setAltText(page, drawing(page, rectangleName), "Highlight box");

  await page.keyboard.press("Escape");
  await page.getByTestId("sheet-col-header").filter({ hasText: /^J$/ }).click();
  const image = await insertImage(page);
  const imageName = (await image.getAttribute("data-drawing-name"))!;

  const names = ["Image 1", "Assumptions box", "Group 5", "Check Box 1", rectangleName, imageName];
  const left = await boxes(page, names);
  const path = await download(page, "spec131-drawings.ods");
  await open(page, path);
  await expect(drawing(page, "Image 1")).toBeVisible();
  const reopened = await boxes(page, names);
  for (const name of names) {
    expectNear(reopened.get(name)!, left.get(name)!, 2);
  }
  await expect(drawing(page, rectangleName)).toHaveAttribute("aria-label", "Highlight box");
  expect((await zOrder(page))[0]).toBe(rectangleName);
});
