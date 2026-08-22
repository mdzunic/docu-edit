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
  sections: Array<{ blocks: unknown[]; margins?: { topCm: number } }>;
  defaultStyles: { margins: { topCm: number } };
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
