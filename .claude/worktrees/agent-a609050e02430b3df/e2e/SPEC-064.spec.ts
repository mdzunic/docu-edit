import { test, expect, type Locator, type Page } from "@playwright/test";

// SPEC-064 — the vertical ruler and app-wide measurement units. AC-16 drives
// the top-margin marker end to end (the page's first line moves and the
// section margin changes), then switches the unit to inches and asserts the
// ruler's read-out and the Page setup dialog agree. Helpers mirror
// e2e/SPEC-027.spec.ts (page setup) and e2e/SPEC-014.spec.ts (model reads).

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

const SAVED = "All changes saved";
async function waitForSaved(page: Page): Promise<void> {
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);
}

const content = (page: Page) => page.getByTestId("document-content");
const firstParagraph = (page: Page) => content(page).locator(".doc-paragraph").first();

interface ModelShape {
  sections: Array<{ blocks: unknown[]; margins?: { topCm: number; bottomCm: number } }>;
  defaultStyles: { margins: { topCm: number; bottomCm: number } };
  measurementUnit?: string;
}

async function readModel(page: Page, id: string): Promise<ModelShape> {
  await expect.poll(async () => (await readPersisted(page, id)) !== null).toBe(true);
  const model = await readPersisted(page, id);
  if (!model) {
    throw new Error(`no persisted model for ${id}`);
  }
  return model;
}

async function readPersisted(page: Page, id: string): Promise<ModelShape | null> {
  return page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const get = (key: string) =>
      new Promise<{ lastModified?: number; savedAt?: number; model: unknown } | undefined>((resolve, reject) => {
        const tx = db.transaction("documents", "readonly");
        const rq = tx.objectStore("documents").get(key);
        rq.onsuccess = () => resolve(rq.result);
        rq.onerror = () => reject(rq.error);
      });
    const committed = await get(docId);
    const draft = await get("draft:" + docId);
    db.close();
    if (draft && (!committed || (draft.savedAt ?? 0) > (committed.lastModified ?? 0))) {
      return draft.model;
    }
    return committed?.model ?? null;
  }, id) as Promise<ModelShape | null>;
}

/** The document's effective top margin in cm — the section's own, else the
 * document default (what setPageMargins edits on a blank document). */
function topMarginCm(model: ModelShape): number {
  return model.sections[0].margins?.topCm ?? model.defaultStyles.margins.topCm;
}

/** The document's effective bottom margin in cm, same fall-through as above. */
function bottomMarginCm(model: ModelShape): number {
  return model.sections[0].margins?.bottomCm ?? model.defaultStyles.margins.bottomCm;
}

async function openPageSetup(page: Page): Promise<void> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("page-setup-button").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeVisible();
}

async function applyPageSetup(page: Page): Promise<void> {
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).toHaveCount(0);
}

const marker = (page: Page) => page.getByTestId("vertical-ruler-marker-margin-top");

async function markerCentre(locator: Locator): Promise<{ x: number; y: number }> {
  const box = (await locator.boundingBox())!;
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}

/** A slider marker's current value in the display unit (its `aria-valuenow`). */
async function markerValue(locator: Locator): Promise<number> {
  return Number(await locator.getAttribute("aria-valuenow"));
}

/** The print-canvas scroll offset — the axis a mid-drag scroll-jump moves. */
const canvasScrollTop = (page: Page): Promise<number> =>
  page.getByTestId("canvas").evaluate((el) => el.scrollTop);

test("dragging the top-margin marker moves the first line and the section margin, and the unit switches to inches (AC-16, AC-16a)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Vertically ruled");
  await waitForSaved(page);

  // The vertical ruler is present in print layout, aligned to the page (AC-1).
  await expect(page.getByTestId("document-vertical-ruler")).toBeVisible();

  const before = await readModel(page, id);
  const beforeTopCm = topMarginCm(before);
  const beforeLineTop = (await firstParagraph(page).boundingBox())!.y;

  // Drag the top-margin marker down by ~2cm (96px/inch → ~76px for 2cm).
  const grab = await markerCentre(marker(page));
  await page.mouse.move(grab.x, grab.y);
  await page.mouse.down();
  await page.mouse.move(grab.x, grab.y + 76, { steps: 8 });
  await page.mouse.up();
  await waitForSaved(page);

  // The section's top margin grew, and the first line moved down with it.
  const after = await readModel(page, id);
  expect(topMarginCm(after)).toBeGreaterThan(beforeTopCm + 1);
  const afterLineTop = (await firstParagraph(page).boundingBox())!.y;
  expect(afterLineTop).toBeGreaterThan(beforeLineTop);

  // Switch the unit to inches through the Page setup dialog (AC-16a).
  await openPageSetup(page);
  await page.getByTestId("page-setup-unit").selectOption("in");
  await applyPageSetup(page);

  const inches = (topMarginCm(after) / 2.54).toFixed(2);
  const inchesTrimmed = String(Number(inches));

  // The ruler's read-out reports the top margin in inches...
  await expect(marker(page)).toHaveAttribute("aria-valuetext", new RegExp(`${inchesTrimmed}\\s+inches`));
  // ...and the ruler now draws unit labels (the origin "0" among them).
  await expect(
    page.getByTestId("document-vertical-ruler").locator(".v-tick-label", { hasText: /^0$/ }).first(),
  ).toBeVisible();

  // ...and the Page setup dialog agrees: unit "in", margin field in inches.
  await openPageSetup(page);
  await expect(page.getByTestId("page-setup-unit")).toHaveValue("in");
  const fieldValue = await page.getByTestId("page-setup-margin-top").inputValue();
  expect(Number(fieldValue)).toBeCloseTo(Number(inchesTrimmed), 2);
});

test("dragging the bottom-margin marker after scrolling it into view moves the margin by the drag distance, not a scroll-jump (AC-2a)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Bottom-edge marker");
  await waitForSaved(page);

  // The bottom-margin marker sits near the page's bottom edge — below the fold
  // on a normal viewport — so it must be scrolled into view to be grabbed. This
  // is the case the shipped top-margin test never exercised: an edit fired per
  // pointermove must not snap the viewport back to the caret (still at the top)
  // and desync the pointer→content mapping (SPEC-064 AC-2a scroll-jump fix).
  const bottomMarker = page.getByTestId("vertical-ruler-marker-margin-bottom");
  await bottomMarker.scrollIntoViewIfNeeded();
  await expect(bottomMarker).toBeInViewport();

  const beforeBottomCm = bottomMarginCm(await readModel(page, id));

  // Drag the marker up by ~2cm (76px). Up moves the bottom-margin boundary
  // toward the page bottom, growing the margin.
  const grab = await markerCentre(bottomMarker);
  await page.mouse.move(grab.x, grab.y);
  await page.mouse.down();
  await page.mouse.move(grab.x, grab.y - 76, { steps: 8 });
  await page.mouse.up();
  await waitForSaved(page);

  // The margin grew by roughly the 2cm the pointer travelled — a small, sane
  // change. Before the fix a few-pixel move jumped it past 13cm because the
  // viewport snapped back to the caret mid-drag; assert it stayed in-page.
  const afterBottomCm = bottomMarginCm(await readModel(page, id));
  expect(afterBottomCm).toBeGreaterThan(beforeBottomCm + 1);
  expect(afterBottomCm).toBeLessThan(beforeBottomCm + 4);
});

test("dragging the footer marker after scrolling it into view moves it by the drag distance, not a scroll-jump (AC-3b)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Footer band marker");
  await waitForSaved(page);

  // The footer boundary sits near the page's bottom edge — below the fold, like
  // the bottom margin — so reaching it means scrolling the caret off-screen. A
  // per-pointermove edit must not snap the viewport back to the caret and
  // desync the pointer→content mapping (SPEC-064 AC-3b scroll-jump fix).
  const footer = page.getByTestId("vertical-ruler-marker-footer");
  await footer.scrollIntoViewIfNeeded();
  await expect(footer).toBeInViewport();

  const beforeScroll = await canvasScrollTop(page);
  const beforeCm = await markerValue(footer);

  const grab = await markerCentre(footer);
  await page.mouse.move(grab.x, grab.y);
  await page.mouse.down();
  await page.mouse.move(grab.x, grab.y - 30, { steps: 8 }); // ~0.8cm up
  await page.mouse.up();
  await waitForSaved(page);

  // The viewport did not snap back to the caret mid-drag...
  expect(Math.abs((await canvasScrollTop(page)) - beforeScroll)).toBeLessThan(50);
  // ...and the value changed by the small amount the pointer travelled. Before
  // the fix a few-pixel move jumped the footer distance past 13cm.
  const afterCm = await markerValue(footer);
  expect(afterCm).toBeGreaterThan(beforeCm);
  expect(afterCm).toBeLessThan(beforeCm + 2);
});

test("dragging a below-the-fold table-row boundary sets the row height without a scroll-jump (AC-5, AC-5a)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();

  // Insert a tall table so its lower row boundaries fall below the fold, then
  // place the caret inside it — that is when the ruler shows the row markers.
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-10-2").click();
  await content(page).locator("table td, table th").first().click();
  await waitForSaved(page);

  // A lower row boundary requires scrolling into view (AC-5 shows all rows).
  const rowMarker = page.getByTestId("vertical-ruler-row-8");
  await rowMarker.scrollIntoViewIfNeeded();
  await expect(rowMarker).toBeInViewport();

  const beforeScroll = await canvasScrollTop(page);
  const beforeCm = await markerValue(rowMarker);

  const grab = await markerCentre(rowMarker);
  await page.mouse.move(grab.x, grab.y);
  await page.mouse.down();
  await page.mouse.move(grab.x, grab.y + 40, { steps: 8 }); // grow the row ~1cm
  await page.mouse.up();
  await waitForSaved(page);

  // No mid-drag snap-back, and the row height grew by the sane drag distance
  // (a scroll-jump would have thrown the height wildly off).
  expect(Math.abs((await canvasScrollTop(page)) - beforeScroll)).toBeLessThan(50);
  const afterCm = await markerValue(rowMarker);
  expect(afterCm).toBeGreaterThan(beforeCm);
  expect(afterCm).toBeLessThan(beforeCm + 3);
});

test("a length field accepts an explicit unit suffix and converts it (AC-11)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Suffixed margins");
  await waitForSaved(page);

  // The display unit is centimetres, but the user types inches with a suffix.
  await openPageSetup(page);
  const topField = page.getByTestId("page-setup-margin-top");
  await topField.fill("0.5in");
  await applyPageSetup(page);
  await waitForSaved(page);

  // 0.5in = 1.27cm reached the model despite the cm display unit (AC-11).
  const after = await readModel(page, id);
  expect(topMarginCm(after)).toBeCloseTo(1.27, 2);
});

test("View → Ruler toggles both rulers together with one control (AC-7)", async ({ page }) => {
  await createDocument(page);
  await expect(page.getByTestId("document-ruler")).toBeVisible();
  await expect(page.getByTestId("document-vertical-ruler")).toBeVisible();

  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("toggle-ruler-button").click();
  await expect(page.getByTestId("document-ruler")).toHaveCount(0);
  await expect(page.getByTestId("document-vertical-ruler")).toHaveCount(0);

  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("toggle-ruler-button").click();
  await expect(page.getByTestId("document-ruler")).toBeVisible();
  await expect(page.getByTestId("document-vertical-ruler")).toBeVisible();
});

test("View → Horizontal ruler and View → Vertical ruler toggle each ruler on its own", async ({ page }) => {
  await createDocument(page);
  await expect(page.getByTestId("document-ruler")).toBeVisible();
  await expect(page.getByTestId("document-vertical-ruler")).toBeVisible();

  // Hiding the vertical ruler alone leaves the horizontal one, and the group
  // control reports the partial state.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("toggle-vertical-ruler-button").click();
  await expect(page.getByTestId("document-vertical-ruler")).toHaveCount(0);
  await expect(page.getByTestId("document-ruler")).toBeVisible();
  await page.getByTestId("menu-title-view").click();
  await expect(page.getByTestId("toggle-ruler-button")).toHaveAttribute("aria-checked", "mixed");
  await expect(page.getByTestId("toggle-vertical-ruler-button")).toHaveAttribute("aria-checked", "false");
  await expect(page.getByTestId("toggle-horizontal-ruler-button")).toHaveAttribute("aria-checked", "true");

  // Hiding the horizontal ruler too empties the group; showing it again brings
  // back only the horizontal ruler.
  await page.getByTestId("toggle-horizontal-ruler-button").click();
  await expect(page.getByTestId("document-ruler")).toHaveCount(0);
  await page.getByTestId("menu-title-view").click();
  await expect(page.getByTestId("toggle-ruler-button")).toHaveAttribute("aria-checked", "false");
  await page.getByTestId("toggle-horizontal-ruler-button").click();
  await expect(page.getByTestId("document-ruler")).toBeVisible();
  await expect(page.getByTestId("document-vertical-ruler")).toHaveCount(0);

  // The group control from a partial state shows both.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("toggle-ruler-button").click();
  await expect(page.getByTestId("document-ruler")).toBeVisible();
  await expect(page.getByTestId("document-vertical-ruler")).toBeVisible();
});

test("the vertical ruler sits one gutter left of the page, as thick as the horizontal ruler is tall (AC-1)", async ({
  page,
}) => {
  await createDocument(page);
  const horizontal = (await page.getByTestId("document-ruler").boundingBox())!;
  const vertical = (await page.getByTestId("document-vertical-ruler").boundingBox())!;
  const pageBox = (await page.locator(".page").first().boundingBox())!;
  // Same thickness, and the same gap to the page the horizontal ruler keeps.
  expect(vertical.width).toBeCloseTo(horizontal.height, 0);
  const horizontalGap = pageBox.y - (horizontal.y + horizontal.height);
  const verticalGap = pageBox.x - (vertical.x + vertical.width);
  expect(verticalGap).toBeGreaterThan(0);
  expect(verticalGap).toBeCloseTo(horizontalGap, 0);
  expect(vertical.y).toBeCloseTo(pageBox.y, 0);
});

test("the vertical ruler is hidden in web layout, the horizontal ruler stays (AC-8)", async ({ page }) => {
  await createDocument(page);
  await expect(page.getByTestId("document-vertical-ruler")).toBeVisible();

  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-layout").click();
  await page.getByTestId("layout-web").click();
  await expect(page.locator(".is-web-layout")).toHaveCount(1);

  await expect(page.getByTestId("document-vertical-ruler")).toHaveCount(0);
  await expect(page.getByTestId("document-ruler")).toBeVisible();
});

test("right-clicking the vertical ruler sets the unit, the same value Page setup shows (AC-10a, AC-10b)", async ({
  page,
}) => {
  await createDocument(page);
  const ruler = page.getByTestId("document-vertical-ruler");
  await ruler.click({ button: "right" });
  await expect(page.getByTestId("ruler-unit-menu")).toBeVisible();
  await page.getByTestId("ruler-unit-mm").click();
  await expect(page.getByTestId("ruler-unit-menu")).toHaveCount(0);

  await openPageSetup(page);
  await expect(page.getByTestId("page-setup-unit")).toHaveValue("mm");
});

test("the top-margin marker is keyboard-operable and announces its value (AC-13, AC-14)", async ({ page }) => {
  const id = await createDocument(page);
  const m = marker(page);
  await expect(m).toHaveAttribute("role", "slider");
  await m.focus();

  const before = topMarginCm(await readModel(page, id));
  await page.keyboard.press("ArrowDown"); // one minor division toward the page bottom
  await waitForSaved(page);
  const afterMinor = topMarginCm(await readModel(page, id));
  expect(afterMinor).toBeGreaterThan(before);

  // aria-valuetext reads the marker name and its value in the display unit.
  await expect(m).toHaveAttribute("aria-valuetext", /Top margin, .* centimetres/);
});

test("the ruler anchors to page 1 — not the container centre — when zoomed-out pages tile side by side (AC-1 regression)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Intro paragraph up top here.");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Second line of body text.");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-page").click();
  await page.keyboard.type("Page two text lives here.");

  // Fit-page tiles the two pages side by side: page 1 sits in the LEFT grid
  // cell, so the old "50% of .pages minus half a page width" anchor painted
  // the ruler strip over page 1's content — and its draggable strip then
  // swallowed every click underneath.
  await page.getByTestId("zoom-percent").click();
  await page.getByTestId("zoom-fit-page").click();

  const ruler = page.getByTestId("document-vertical-ruler");
  await expect(ruler).toBeVisible();
  await expect(async () => {
    const rulerBox = (await ruler.boundingBox())!;
    const pageBox = (await page.getByTestId("document-page").first().boundingBox())!;
    // The strip sits fully in the gutter, hugging page 1's left edge.
    const gap = pageBox.x - (rulerBox.x + rulerBox.width);
    expect(gap).toBeGreaterThanOrEqual(0);
    expect(gap).toBeLessThanOrEqual(16);
  }).toPass({ timeout: 5000 });

  // The click the displaced strip used to intercept must land: Playwright's
  // actionability check fails on an intercepting overlay, so a plain click
  // succeeding IS the regression assertion.
  await content(page).locator(".doc-paragraph", { hasText: "Intro paragraph" }).first().click();
});
