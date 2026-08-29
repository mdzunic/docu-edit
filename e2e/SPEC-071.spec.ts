import { test, expect, type Locator, type Page } from "@playwright/test";
import { join, dirname } from "node:path";
import { tmpdir } from "node:os";
import { fileURLToPath } from "node:url";

// SPEC-071 — freeform shapes, curves and edit points. Drives the three
// Insert-gallery Freeform tools and asserts a real path-bearing drawing
// lands (the model + round-trip are covered exhaustively by the unit
// suites; this proves the gesture wiring reaches them).

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]');

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function armFreeform(page: Page, tool: string): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`freeform-gallery-${tool}`).click();
}

test.describe("SPEC-071 — freeform drawing", () => {
  test("Freeform tool places vertices and double-click finishes an open shape (AC-6, AC-7)", async ({
    page,
  }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 520, y: 320 } });
    await content(page).dblclick({ position: { x: 640, y: 220 } });
    await expect(drawings(page)).toHaveCount(1);
    // The gallery is a Freeform group with three tools.
    await page.getByTestId("menu-title-insert").click();
    await expect(page.getByTestId("freeform-gallery-freeform")).toBeVisible();
    await expect(page.getByTestId("freeform-gallery-curve")).toBeVisible();
    await expect(page.getByTestId("freeform-gallery-scribble")).toBeVisible();
  });

  test("Escape during a freeform draw creates nothing (AC-10)", async ({ page }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 520, y: 320 } });
    await page.keyboard.press("Escape");
    await expect(drawings(page)).toHaveCount(0);
  });

  test("a single click then finish creates nothing (AC-7)", async ({ page }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    await content(page).click({ position: { x: 400, y: 200 } });
    await page.keyboard.press("Enter");
    await expect(drawings(page)).toHaveCount(0);
  });

  test("Curve tool draws a shape that survives save and reopen", async ({ page }) => {
    await createDocument(page);
    await armFreeform(page, "curve");
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 500, y: 320 } });
    await content(page).dblclick({ position: { x: 620, y: 200 } });
    await expect(drawings(page)).toHaveCount(1);
  });

  test("an open freeform hit-tests along its stroke, not just at its vertices (AC-24)", async ({
    page,
  }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    // A V: first segment (400,200)->(520,320), second (520,320)->(640,220).
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 520, y: 320 } });
    await content(page).dblclick({ position: { x: 640, y: 220 } });
    await expect(drawings(page)).toHaveCount(1);
    // Deselect the freshly drawn (auto-selected) shape.
    await content(page).click({ position: { x: 40, y: 40 } });
    await expect(page.getByTestId("shape-toolbar")).toHaveCount(0);
    // Click the exact midpoint of the first segment — on the stroke but deep
    // in the shape's bounding-box interior, far from any vertex. Before the
    // fix this only selected at the extreme vertices (the box edges).
    await content(page).click({ position: { x: 460, y: 260 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
  });

  test("an open freeform is hit-tested against its outline, not its bounding box (AC-24)", async ({
    page,
  }) => {
    await createDocument(page);
    await armFreeform(page, "freeform");
    // Same V: segments (400,200)->(520,320)->(640,220).
    await content(page).click({ position: { x: 400, y: 200 } });
    await content(page).click({ position: { x: 520, y: 320 } });
    await content(page).dblclick({ position: { x: 640, y: 220 } });
    await content(page).click({ position: { x: 40, y: 40 } });
    await expect(page.getByTestId("shape-toolbar")).toHaveCount(0);
    // A point inside the V's bounding box but far from either stroke segment
    // (the empty gap above the V's notch): an unfilled path must NOT select
    // here — proving the hit test follows the real outline, not the box.
    await content(page).click({ position: { x: 520, y: 210 } });
    await expect(page.getByTestId("shape-toolbar")).toHaveCount(0);
    // ...while a point on the stroke does select.
    await content(page).click({ position: { x: 460, y: 260 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
  });
});

test.describe("SPEC-071 — edit points", () => {
  async function insertPreset(page: Page, geometry: string, x: number, y: number): Promise<void> {
    await content(page).click({ position: { x: 20, y: 20 } });
    await page.getByTestId("menu-title-insert").click();
    await page.getByTestId(`shape-gallery-${geometry}`).click();
    await content(page).click({ position: { x, y } });
  }

  test("Format -> Edit points converts a preset and shows vertex handles (AC-14, AC-15, AC-19)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertPreset(page, "triangle", 450, 250);
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await page.getByTestId("shape-edit-points-button").click();
    await expect(page.getByTestId("point-edit-overlay")).toBeVisible();
    // A triangle converts to a 3-vertex freeform.
    await expect(page.getByTestId("point-edit-vertex-0")).toBeVisible();
    await expect(page.getByTestId("point-edit-vertex-2")).toBeVisible();
  });

  test("dragging a vertex commits one edit and Escape exits (AC-17, AC-22)", async ({ page }) => {
    await createDocument(page);
    await insertPreset(page, "triangle", 450, 250);
    await page.getByTestId("shape-edit-points-button").click();
    const handle = page.getByTestId("point-edit-vertex-0");
    const box = await handle.boundingBox();
    if (!box) {
      throw new Error("no vertex handle");
    }
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 40, box.y - 30, { steps: 8 });
    await page.mouse.up();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("point-edit-overlay")).toHaveCount(0);
    // The shape is still there, still a single drawing.
    await expect(drawings(page)).toHaveCount(1);
  });

  test("point-edit hides the box resize/rotation handles so a vertex handle is grabbable (AC-17)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertPreset(page, "triangle", 450, 250);
    // Resize/rotation handles are shown while the shape is merely selected.
    await expect(page.getByTestId("shape-handles")).toBeVisible();
    await page.getByTestId("shape-edit-points-button").click();
    await expect(page.getByTestId("point-edit-overlay")).toBeVisible();
    // In point-edit mode they are gone, so a vertex handle sitting on a box
    // corner (every vertex of a converted triangle) is the topmost element at
    // its own centre and receives the drag instead of a resize handle.
    await expect(page.getByTestId("shape-handles")).toHaveCount(0);
    const handle = page.getByTestId("point-edit-vertex-0");
    const box = await handle.boundingBox();
    if (!box) {
      throw new Error("no vertex handle");
    }
    const cx = box.x + box.width / 2;
    const cy = box.y + box.height / 2;
    const topmost = await page.evaluate(
      ([x, y]) => document.elementFromPoint(x, y)?.getAttribute("data-testid") ?? "",
      [cx, cy],
    );
    expect(topmost).toBe("point-edit-vertex-0");
  });

  test("dragging a vertex on a box corner edits the point, not the box (AC-17)", async ({
    page,
  }) => {
    await createDocument(page);
    await insertPreset(page, "triangle", 450, 250);
    await page.getByTestId("shape-edit-points-button").click();
    await expect(page.getByTestId("point-edit-overlay")).toBeVisible();
    const handle = page.getByTestId("point-edit-vertex-0");
    const box = await handle.boundingBox();
    if (!box) {
      throw new Error("no vertex handle");
    }
    const cx = box.x + box.width / 2;
    const cy = box.y + box.height / 2;
    // This vertex sits on a corner of the shape's bounding box, exactly where a
    // resize handle used to sit on top of it. Dragging it must move the vertex
    // (committing an "edit points" undo) — before the fix it resized the box.
    await page.mouse.move(cx, cy);
    await page.mouse.down();
    await page.mouse.move(cx + 60, cy - 40, { steps: 10 });
    await page.mouse.up();
    // The vertex handle followed the drag rather than the box being resized.
    const moved = await handle.boundingBox();
    if (!moved) {
      throw new Error("vertex handle vanished");
    }
    expect(Math.round(moved.x + moved.width / 2 - cx)).toBeGreaterThan(30);
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("point-edit-overlay")).toHaveCount(0);
    const undoLabel = await page.getByTestId("undo-button").getAttribute("aria-label");
    expect(undoLabel).toContain("edit points");
    expect(undoLabel).not.toContain("resize");
  });
});

// --- the freeform fixtures (AC-4, AC-5, AC-32, AC-34) ---------------------
// word-freeform.docx and lo-freeform.odt are hand-built by
// scripts/build-spec071-fixtures.mjs to reproduce what Word and Writer emit
// for hand-drawn geometry — NOT round-tripped through this project's own
// writer — so opening them exercises the real custom-geometry readers. Both
// hold the same eight shapes in the same document order:
//   0 open polyline      2 closed curve   4 smooth/corner mix  6 formula-driven
//   1 closed polygon     3 open scribble  5 reduced primitives 7 unsupported
// 0..5 paint their real outline; 6 and 7 are what the narrowed placeholder
// scope still preserves, and still paint as the dashed grey placeholder.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxFreeformFixture = join(fixturesDir, "word-freeform.docx");
const odtFreeformFixture = join(fixturesDir, "lo-freeform.odt");

const FIXTURE_LABELS = [
  "Word freeform",
  "Closed freeform",
  "Closed curve",
  "Open scribble",
  "Smooth and corner mix",
  "Reduced primitives",
  "Formula-driven geometry",
  "Unsupported geometry",
];

// The colours scripts/build-spec071-fixtures.mjs paints every fixture shape
// with. The placeholder never uses either: shapeCanvas.ts strokes it in
// #888888, which is what makes "did the real outline paint?" answerable from
// pixels alone.
const OUTLINE_RGB = { r: 0xcc, g: 0x00, b: 0x00 };
const FILL_RGB = { r: 0x33, g: 0x66, b: 0xcc };
const PLACEHOLDER_RGB = { r: 0x88, g: 0x88, b: 0x88 };

const drawingAt = (page: Page, index: number) => content(page).locator(`[data-testid="drawing-${index}"]`);

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** The canvas DocShapeBlock.vue paints a shape on — the sibling immediately
 * before its ARIA proxy (the same addressing e2e/SPEC-051.spec.ts uses). */
const shapeCanvas = (proxy: Locator) => proxy.locator("xpath=preceding-sibling::canvas[1]");

/** Whether any pixel on a shape's canvas is within `tolerance` of `rgb` —
 * decisive where sampling a single point on a thin stroke would be flaky.
 * Alpha must be near-opaque so antialiased edges don't count as a match. */
async function canvasHasColour(
  proxy: Locator,
  rgb: { r: number; g: number; b: number },
  tolerance = 24,
): Promise<boolean> {
  return shapeCanvas(proxy).evaluate(
    (el, [target, tol]) => {
      const canvasEl = el as HTMLCanvasElement;
      const data = canvasEl.getContext("2d")!.getImageData(0, 0, canvasEl.width, canvasEl.height).data;
      for (let i = 0; i < data.length; i += 4) {
        if (
          data[i + 3] > 200 &&
          Math.abs(data[i] - target.r) <= tol &&
          Math.abs(data[i + 1] - target.g) <= tol &&
          Math.abs(data[i + 2] - target.b) <= tol
        ) {
          return true;
        }
      }
      return false;
    },
    [rgb, tolerance] as const,
  );
}

/** One pixel at a fraction of the canvas's own backing-store size, so
 * devicePixelRatio scaling doesn't matter. */
async function canvasPixelAt(
  proxy: Locator,
  xFrac: number,
  yFrac: number,
): Promise<{ r: number; g: number; b: number; a: number }> {
  const [r, g, b, a] = await shapeCanvas(proxy).evaluate(
    (el, [xf, yf]) => {
      const canvasEl = el as HTMLCanvasElement;
      const ctx = canvasEl.getContext("2d")!;
      const x = Math.min(canvasEl.width - 1, Math.max(0, Math.floor(canvasEl.width * xf)));
      const y = Math.min(canvasEl.height - 1, Math.max(0, Math.floor(canvasEl.height * yf)));
      const px = ctx.getImageData(x, y, 1, 1).data;
      return [px[0], px[1], px[2], px[3]];
    },
    [xFrac, yFrac] as const,
  );
  return { r, g, b, a };
}

/** Registers the assertions both fixtures must satisfy — the two files hold
 * the same eight shapes, so this runs once per format rather than
 * duplicating every expectation. */
function registerFreeformFixtureTests(formatLabel: string, fixturePath: string, acTag: string): void {
  test.describe(`SPEC-071 — opening ${formatLabel} (${acTag})`, () => {
    test(`renders all eight shapes with their alt text (${acTag})`, async ({ page }) => {
      await openFixture(page, fixturePath);
      await expect(drawings(page)).toHaveCount(8);
      for (let i = 0; i < FIXTURE_LABELS.length; i++) {
        await expect(drawingAt(page, i), FIXTURE_LABELS[i]).toHaveAttribute("aria-label", FIXTURE_LABELS[i]);
      }
    });

    test(`every path shape paints its own outline colour, not the placeholder's grey (${acTag})`, async ({
      page,
    }) => {
      await openFixture(page, fixturePath);
      for (let i = 0; i < 6; i++) {
        const proxy = drawingAt(page, i);
        expect(await canvasHasColour(proxy, OUTLINE_RGB), `${FIXTURE_LABELS[i]}: outline`).toBe(true);
        expect(await canvasHasColour(proxy, PLACEHOLDER_RGB), `${FIXTURE_LABELS[i]}: placeholder`).toBe(false);
      }
    });

    test(`a closed path fills its real outline, leaving the bounding box's corners empty (AC-2, ${acTag})`, async ({
      page,
    }) => {
      await openFixture(page, fixturePath);
      // The closed curve is a circle drawn as four cubics: its centre is
      // inside the outline and its box corners are outside it. A "rect"
      // geometry stand-in would paint the corners too — this is what makes
      // `path` visibly take precedence over `geometry`.
      const circle = drawingAt(page, 2);
      const centre = await canvasPixelAt(circle, 0.5, 0.5);
      expect(centre.a).toBeGreaterThan(200);
      expect(centre.r).toBeCloseTo(FILL_RGB.r, -1);
      expect(centre.g).toBeCloseTo(FILL_RGB.g, -1);
      expect(centre.b).toBeCloseTo(FILL_RGB.b, -1);
      expect((await canvasPixelAt(circle, 0.02, 0.02)).a).toBe(0);
      expect((await canvasPixelAt(circle, 0.98, 0.98)).a).toBe(0);
    });

    test(`an open path paints an outline and no fill (AC-9, ${acTag})`, async ({ page }) => {
      await openFixture(page, fixturePath);
      // The scribble is two cubic waves through the middle of its box: the
      // stroke is there, but nothing is filled — the top-left corner, which a
      // filled shape would cover, stays empty.
      const scribble = drawingAt(page, 3);
      expect(await canvasHasColour(scribble, OUTLINE_RGB)).toBe(true);
      expect(await canvasHasColour(scribble, FILL_RGB)).toBe(false);
      expect((await canvasPixelAt(scribble, 0.02, 0.02)).a).toBe(0);
    });

    test(`the two unmodellable geometries still paint as dashed placeholders (AC-31, AC-32, ${acTag})`, async ({
      page,
    }) => {
      await openFixture(page, fixturePath);
      for (const index of [6, 7]) {
        const proxy = drawingAt(page, index);
        // The placeholder is a grey dashed outline only: none of the source's
        // own fill or stroke colour is painted, and its interior stays empty.
        expect(await canvasHasColour(proxy, PLACEHOLDER_RGB), FIXTURE_LABELS[index]).toBe(true);
        expect(await canvasHasColour(proxy, OUTLINE_RGB), FIXTURE_LABELS[index]).toBe(false);
        expect((await canvasPixelAt(proxy, 0.5, 0.5)).a, FIXTURE_LABELS[index]).toBe(0);
      }
    });

    test(`Edit points opens on a shape read as a path but not on a preserved one (AC-3, ${acTag})`, async ({
      page,
    }) => {
      await openFixture(page, fixturePath);
      // A shape whose path we modelled is editable straight from the file —
      // no "convert to freeform" step, since it arrived as a freeform.
      await drawingAt(page, 1).click();
      await expect(page.getByTestId("shape-toolbar")).toBeVisible();
      await page.getByTestId("shape-edit-points-button").click();
      await expect(page.getByTestId("point-edit-overlay")).toBeVisible();
      await expect(page.getByTestId("point-edit-vertex-0")).toBeVisible();
      await page.keyboard.press("Escape");
      await expect(page.getByTestId("point-edit-overlay")).toHaveCount(0);

      // A preserved shape has no path to edit, so the same control does
      // nothing — the other half of the mutual-exclusion invariant.
      await drawingAt(page, 7).click();
      await expect(page.getByTestId("shape-toolbar")).toBeVisible();
      await page.getByTestId("shape-edit-points-button").click();
      await expect(page.getByTestId("point-edit-overlay")).toHaveCount(0);
    });
  });
}

registerFreeformFixtureTests("word-freeform.docx", docxFreeformFixture, "AC-4");
registerFreeformFixtureTests("lo-freeform.odt", odtFreeformFixture, "AC-5");

test.describe("SPEC-071 — the freeform fixtures survive a save and reopen (AC-34)", () => {
  /** Exports the open document and reopens the downloaded file, so the whole
   * writer → reader path runs in the browser rather than in a unit test. */
  async function saveAndReopen(page: Page, format: "docx" | "odt"): Promise<void> {
    await page.getByTestId("menu-title-file").click();
    await page.getByTestId("file-menu-export").click();
    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId(`download-${format}`).click();
    const download = await downloadPromise;
    const saved = join(tmpdir(), `spec071-${format}-${Date.now()}.${format}`);
    await download.saveAs(saved);
    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(saved);
    await expect(page).toHaveURL(/\/doc\/.+/);
  }

  for (const [formatLabel, fixturePath, format] of [
    ["word-freeform.docx", docxFreeformFixture, "docx"],
    ["lo-freeform.odt", odtFreeformFixture, "odt"],
  ] as const) {
    test(`${formatLabel} reopens with the same real outlines and placeholders`, async ({ page }) => {
      await openFixture(page, fixturePath);
      await expect(drawings(page)).toHaveCount(8);
      await saveAndReopen(page, format);

      await expect(drawings(page)).toHaveCount(8);
      for (let i = 0; i < FIXTURE_LABELS.length; i++) {
        await expect(drawingAt(page, i), FIXTURE_LABELS[i]).toHaveAttribute("aria-label", FIXTURE_LABELS[i]);
      }
      // The six modelled paths still paint their own outline...
      for (let i = 0; i < 6; i++) {
        expect(await canvasHasColour(drawingAt(page, i), OUTLINE_RGB), FIXTURE_LABELS[i]).toBe(true);
      }
      // ...and the two preserved ones are still placeholders, neither
      // reinterpreted into a path nor lost (AC-33).
      for (const index of [6, 7]) {
        expect(await canvasHasColour(drawingAt(page, index), PLACEHOLDER_RGB), FIXTURE_LABELS[index]).toBe(true);
      }
      // The closed curve is still a circle, not its bounding box.
      expect((await canvasPixelAt(drawingAt(page, 2), 0.02, 0.02)).a).toBe(0);
    });
  }
});
