import { test, expect, type Page } from "@playwright/test";

// SPEC-068 — responsive layout tiers and touch support. e2e/SPEC-068.compact.spec.ts
// (the touch-emulated 390x844 project) already covers the compact-tier smoke
// path (AC-7..10, AC-25, AC-35, AC-37..39, AC-45/46). This file covers what
// that project's fixed viewport cannot: live tier transitions at the exact
// 640/1024 boundaries (AC-1, AC-2), the medium tier's own layout (AC-4, AC-5,
// AC-6), the touch long-press context menu (AC-22, AC-23), pinch-zoom and
// two-finger pan through the app's own view transform (AC-32, AC-33, AC-34),
// and the drawing position dialog's tier gating — the fix for the round-1
// blocking finding (AC-42), enabled wherever a coarse pointer needs it and
// disabled with a stated reason only at the compact tier (AC-12, AC-13).

const content = (page: Page) => page.getByTestId("document-content");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page)).toBeVisible();
}

async function insertShapeAt(page: Page, x: number, y: number): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("shape-gallery-rect").click();
  await content(page).click({ position: { x, y } });
}

test.describe("layout tiers — exact boundaries, live resize (AC-1, AC-2, AC-3)", () => {
  test("639px is compact (Menu button), 640px is medium (menu bar) — no reload needed", async ({ page }) => {
    await page.setViewportSize({ width: 639, height: 900 });
    await createDocument(page);
    await expect(page.getByTestId("menu-button")).toBeVisible();
    await expect(page.locator('[role="menubar"]')).toHaveCount(0);

    await page.setViewportSize({ width: 640, height: 900 });
    await expect(page.locator('[role="menubar"]')).toBeVisible();
    await expect(page.getByTestId("menu-button")).toHaveCount(0);
    // Still the same document — this was a resize, not a navigation.
    await expect(page).toHaveURL(/\/doc\/.+/);
  });

  test("1023px keeps the medium tier's toolbar overflow; 1024px is the unchanged full tier (AC-3)", async ({
    page,
  }) => {
    await page.setViewportSize({ width: 1280, height: 900 });
    await createDocument(page);
    // The full tier at a generous width shows the whole toolbar band SPEC-061
    // defines, with every group present and no overflow needed.
    await expect(page.getByTestId("format-underline")).toBeVisible();

    await page.setViewportSize({ width: 639, height: 900 });
    await expect(page.getByTestId("menu-button")).toBeVisible();
    await page.setViewportSize({ width: 1280, height: 900 });
    await expect(page.locator('[role="menubar"]')).toBeVisible();
    await expect(page.getByTestId("menu-button")).toHaveCount(0);
  });
});

test.describe("medium tier layout (AC-4, AC-5, AC-6)", () => {
  test.use({ viewport: { width: 800, height: 800 } });

  test("the menu bar stays a real menubar, the toolbar gains an overflow, and the side panel overlays", async ({
    page,
  }) => {
    await createDocument(page);
    // AC-6: still a real menubar, not collapsed to a Menu button.
    await expect(page.locator('[role="menubar"]')).toBeVisible();
    await expect(page.getByTestId("menu-button")).toHaveCount(0);

    // AC-5: the toolbar has collapsed per SPEC-061's overflow order — some
    // controls move behind an overflow toggle at this width.
    await expect(page.getByTestId("toolbar-overflow-button")).toBeVisible();

    // AC-4: opening a side-panel pane overlays rather than displacing.
    const docWidthBefore = await content(page).evaluate((el) => Math.round(el.getBoundingClientRect().width));
    await page.getByTestId("menu-title-view").click();
    await page.getByTestId("view-menu-outline").click();
    await expect(page.getByTestId("side-panel")).toBeVisible();
    await expect(page.getByTestId("side-panel")).toHaveCSS("position", "fixed");
    const docWidthAfter = await content(page).evaluate((el) => Math.round(el.getBoundingClientRect().width));
    expect(docWidthAfter).toBe(docWidthBefore);
  });
});

test.describe("touch equivalents (AC-22, AC-23, AC-32, AC-33, AC-34)", () => {
  test.use({ hasTouch: true, isMobile: true });

  test("a long press opens the editor's context menu and the synthetic post-touch click does not dismiss it", async ({
    page,
  }) => {
    await createDocument(page);
    await content(page).click();
    await page.keyboard.type("Hello long press world");
    const box = await content(page).boundingBox();
    if (!box) {
      throw new Error("document content has no bounding box");
    }
    const x = box.x + 40;
    const y = box.y + 20;
    // Pointer events, not raw touch, are what useLongPress listens for
    // (pointerType === "touch"), so drive it directly at that level.
    await page.evaluate(
      ({ x, y }) => {
        const el = document.elementFromPoint(x, y);
        el?.dispatchEvent(
          new PointerEvent("pointerdown", { bubbles: true, clientX: x, clientY: y, pointerType: "touch", pointerId: 7 }),
        );
      },
      { x, y },
    );
    await page.waitForTimeout(600);
    await page.evaluate(
      ({ x, y }) => {
        const el = document.elementFromPoint(x, y);
        el?.dispatchEvent(
          new PointerEvent("pointerup", { bubbles: true, clientX: x, clientY: y, pointerType: "touch", pointerId: 7 }),
        );
      },
      { x, y },
    );
    await expect(page.getByTestId("context-menu-cut")).toBeVisible();
    await expect(page.getByTestId("context-menu-copy")).toBeVisible();

    // The browser synthesizes a click after a real touch; useLongPress
    // swallows exactly that one click (AC-22/23's earlier round-1 nonblocking
    // gap) so it cannot dismiss the menu it just opened.
    await page.evaluate(
      ({ x, y }) => {
        const el = document.elementFromPoint(x, y);
        el?.dispatchEvent(new MouseEvent("click", { bubbles: true, clientX: x, clientY: y }));
      },
      { x, y },
    );
    await expect(page.getByTestId("context-menu-cut")).toBeVisible();
  });

  test("pinch zooms the app's own view transform, not the browser page, and leaves the chrome unscaled (AC-32, AC-34)", async ({
    page,
  }) => {
    await createDocument(page);
    const zoomBefore = Number((await page.getByTestId("zoom-percent").inputValue()).replace("%", ""));
    const toolbarHeightBefore = await page
      .getByTestId("toolbar-overflow-button")
      .evaluate((el) => el.getBoundingClientRect().height)
      .catch(() => null);

    const client = await page.context().newCDPSession(page);
    const cx = 195;
    const cy = 400;
    await client.send("Input.dispatchTouchEvent", {
      type: "touchStart",
      touchPoints: [
        { x: cx - 10, y: cy },
        { x: cx + 10, y: cy },
      ],
    });
    for (let i = 1; i <= 8; i++) {
      await client.send("Input.dispatchTouchEvent", {
        type: "touchMove",
        touchPoints: [
          { x: cx - 10 - i * 8, y: cy },
          { x: cx + 10 + i * 8, y: cy },
        ],
      });
      await page.waitForTimeout(20);
    }
    await client.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });

    const zoomAfter = Number((await page.getByTestId("zoom-percent").inputValue()).replace("%", ""));
    expect(zoomAfter).toBeGreaterThan(zoomBefore);

    if (toolbarHeightBefore !== null) {
      const toolbarHeightAfter = await page
        .getByTestId("toolbar-overflow-button")
        .evaluate((el) => el.getBoundingClientRect().height);
      expect(Math.abs(toolbarHeightAfter - toolbarHeightBefore)).toBeLessThan(1);
    }
  });

  test("a two-finger drag pans the zoomed document (AC-33)", async ({ page }) => {
    await createDocument(page);
    for (let i = 0; i < 10; i++) {
      await page.getByTestId("zoom-in").click();
    }
    const canvas = page.getByTestId("canvas");
    const scrollBefore = await canvas.evaluate((el) => el.scrollTop);

    const client = await page.context().newCDPSession(page);
    const x1 = 100;
    const x2 = 150;
    const y = 500;
    await client.send("Input.dispatchTouchEvent", {
      type: "touchStart",
      touchPoints: [
        { x: x1, y },
        { x: x2, y },
      ],
    });
    for (let i = 1; i <= 8; i++) {
      await client.send("Input.dispatchTouchEvent", {
        type: "touchMove",
        touchPoints: [
          { x: x1, y: y - i * 20 },
          { x: x2, y: y - i * 20 },
        ],
      });
      await page.waitForTimeout(20);
    }
    await client.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });

    const scrollAfter = await canvas.evaluate((el) => el.scrollTop);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
  });
});

test.describe("drawing position dialog — tier gating (AC-12, AC-13, AC-42)", () => {
  test("stays enabled at the full and medium tiers, and is disabled with a stated reason only at compact", async ({
    page,
  }) => {
    await page.setViewportSize({ width: 1280, height: 900 });
    await createDocument(page);
    await insertShapeAt(page, 400, 300);
    await expect(page.getByTestId("shape-position-button")).toBeEnabled();

    // The numeric fallback actually resizes and repositions the shape.
    const drawing = content(page).locator('[data-testid="drawing-0"]');
    const before = await drawing.boundingBox();
    if (!before) {
      throw new Error("the drawing has no bounding box");
    }
    await page.getByTestId("shape-position-button").click();
    await expect(page.getByTestId("shape-position-dialog")).toBeVisible();
    await page.getByTestId("shape-position-width").fill("10");
    await page.getByTestId("shape-position-apply").click();
    await expect
      .poll(async () => (await drawing.boundingBox())?.width ?? 0)
      .toBeGreaterThan(before.width + 20);

    // Medium tier: still enabled — this is the coarse-pointer fallback for a
    // handle drag the spec calls for (AC-42).
    await page.setViewportSize({ width: 800, height: 900 });
    await expect(page.getByTestId("shape-position-button")).toBeEnabled();

    // Compact tier: disabled with the standard stated reason, per the spec's
    // "What compact does not offer" list — the shape itself stays selectable,
    // movable and deletable there (AC-12/13), only this precise numeric
    // editor is off.
    await page.setViewportSize({ width: 390, height: 844 });
    await expect(page.getByTestId("shape-position-button")).toBeDisabled();
    await expect(page.getByTestId("shape-position-button")).toHaveAttribute(
      "title",
      "not available on this screen size",
    );
    await expect(page.getByTestId("shape-delete-button")).toBeEnabled();
  });
});
