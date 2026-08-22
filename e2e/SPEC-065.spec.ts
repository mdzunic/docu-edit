import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-065 — the consolidated side panel: the seven per-pane docks are replaced
// by one docked, tabbed, resizable, persisted region (SidePanel.vue +
// useSidePanel.ts) hosting every auxiliary pane as a tab. This spec drives the
// real browser the way the QA pass did: it leans hardest on the two behaviours
// that regressed in an earlier round — dragging the panel's inner edge to
// resize (AC-8) and double-clicking that same handle to collapse (AC-10), both
// of which were unreachable while .side-panel had no positioning context and
// the absolutely-positioned handle rendered half a viewport away.

const SIDE_KEY = "docu-edit:side-panel";

const panel = (page: Page) => page.getByTestId("side-panel");
const handle = (page: Page) => page.getByTestId("side-panel-resize");
const tab = (page: Page, id: string) => page.getByTestId(`side-panel-tab-${id}`);
const badge = (page: Page, id: string) => page.getByTestId(`side-panel-badge-${id}`);
const content = (page: Page) => page.getByTestId("document-content");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function openView(page: Page, testid: string): Promise<void> {
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId(testid).click();
}
async function openFile(page: Page, testid: string): Promise<void> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId(testid).click();
}
async function openTools(page: Page, testid: string): Promise<void> {
  await page.getByTestId("menu-title-tools").click();
  await page.getByTestId(testid).click();
}

function readState(page: Page) {
  return page.evaluate((k) => JSON.parse(localStorage.getItem(k) ?? "null"), SIDE_KEY);
}

test("consolidates the seven docks into one SidePanel that hosts each pane (AC-1, AC-4, AC-5, AC-22)", async ({
  page,
}) => {
  await createDocument(page);
  // Closed by default — no panel and no legacy dock wrappers.
  await expect(panel(page)).toHaveCount(0);
  const dockCount = () => page.evaluate(() => document.querySelectorAll("[class*='-dock']").length);
  expect(await dockCount()).toBe(0);

  await openView(page, "view-menu-outline");
  await expect(panel(page)).toHaveCount(1);
  await openView(page, "view-menu-styles-pane");

  // One rail, two tabs, exactly one visible pane body.
  await expect(page.getByRole("tablist")).toHaveCount(1);
  await expect(page.getByRole("tab")).toHaveCount(2);
  await expect(page.getByRole("tabpanel")).toHaveCount(1);
  expect(await dockCount()).toBe(0);

  // Each tab carries an icon element and its title (AC-5).
  await expect(tab(page, "outline").locator(".side-panel-tab-icon")).toBeVisible();
  await expect(tab(page, "outline")).toContainText("Outline");
  await expect(tab(page, "styles")).toContainText("Styles");

  // Panes keep their own root class + test-id, minus the dock wrapper (AC-22).
  await expect(page.locator(".styles-pane")).toHaveCount(1);
  await expect(page.getByTestId("styles-pane")).toBeVisible();
});

test("docks right by default and a header toggle flips + persists the side (AC-2, AC-3)", async ({
  page,
}) => {
  await createDocument(page);
  await openView(page, "view-menu-styles-pane");

  const panelBox = await panel(page).boundingBox();
  const docBox = await content(page).boundingBox();
  expect(panelBox!.x).toBeGreaterThan(docBox!.x); // to the right
  expect((await readState(page)).side).toBe("right");

  // Toggle to the left — styles stays the active tab, no outline tab appears.
  await page.getByTestId("side-panel-side-toggle").click();
  await expect(panel(page)).toHaveClass(/dock-left/);
  expect((await readState(page)).side).toBe("left");
  await expect(tab(page, "styles")).toHaveAttribute("aria-selected", "true");
  await expect(tab(page, "outline")).toHaveCount(0); // AC-3: outline not auto-activated

  await page.reload();
  expect((await readState(page)).side).toBe("left"); // persists
});

test("the header carries the active title, a close button, and closing empties the panel (AC-7, AC-7a, AC-17)", async ({
  page,
}) => {
  await createDocument(page);
  await openView(page, "view-menu-styles-pane");

  await expect(panel(page)).toHaveAttribute("role", "complementary");
  await expect(panel(page)).toHaveAttribute("aria-label", "Styles");
  await expect(page.locator(".side-panel-title")).toHaveText("Styles");
  await expect(page.getByTestId("side-panel-close")).toBeVisible();

  await page.getByTestId("side-panel-close").click();
  await expect(panel(page)).toHaveCount(0);
  const state = await readState(page);
  expect(state.openPanes).toEqual([]);
  expect(state.activePane).toBeNull();
});

test.describe("resize", () => {
  test.use({ viewport: { width: 1280, height: 800 } });

  test("dragging the panel's inner edge resizes it, with clamps, commit-on-release and document reflow (AC-8, AC-8a, AC-8b, AC-8c, AC-8d)", async ({
    page,
  }) => {
    await createDocument(page);
    await openView(page, "view-menu-styles-pane");

    // The handle is pinned to the panel's inner edge (right dock -> left edge),
    // not floating over the document: elementFromPoint at the inner edge is it.
    const pinned = await page.evaluate(() => {
      const p = document.querySelector('[data-testid="side-panel"]')!.getBoundingClientRect();
      const h = document.querySelector('[data-testid="side-panel-resize"]') as HTMLElement;
      const hit = document.elementFromPoint(p.left, p.top + p.height / 2);
      return {
        atInnerEdge: hit?.getAttribute("data-testid"),
        offsetParent: h.offsetParent?.className ?? null,
      };
    });
    expect(pinned.atInnerEdge).toBe("side-panel-resize");
    expect(pinned.offsetParent).toContain("side-panel");

    const widthOf = () => panel(page).evaluate((el) => Math.round(el.getBoundingClientRect().width));
    const docWidth = () => content(page).evaluate((el) => Math.round(el.getBoundingClientRect().width));

    expect(await widthOf()).toBe(320);
    const docNarrow = await docWidth();

    // Drag the inner edge left to widen the panel (right dock).
    async function dragHandleTo(clientX: number): Promise<void> {
      const h = (await handle(page).boundingBox())!;
      await page.mouse.move(h.x + h.width / 2, h.y + h.height / 2);
      await page.mouse.down();
      await page.mouse.move(clientX, h.y + h.height / 2, { steps: 10 });
      await page.mouse.up();
    }

    await dragHandleTo(1280 - 480); // width -> ~480
    expect(await widthOf()).toBeGreaterThan(440);
    // AC-8c: the persisted width equals the settled width at pointer-up.
    expect((await readState(page)).widthPx).toBe(await widthOf());
    // AC-8d: the document reflows narrower as the panel widens.
    expect(await docWidth()).toBeLessThan(docNarrow);

    // AC-8a: over-shrink clamps to the 240px minimum.
    await dragHandleTo(1279);
    expect(await widthOf()).toBe(240);

    // AC-8b: over-widen clamps to half the window (640px at 1280).
    await dragHandleTo(5);
    expect(await widthOf()).toBe(640);
    expect((await readState(page)).widthPx).toBe(640);
  });

  test("double-clicking the handle collapses to the tab rail and expands back to the persisted width; collapse persists (AC-10, AC-10a, AC-10b)", async ({
    page,
  }) => {
    await createDocument(page);
    // Outline restores on reload, so the collapsed panel is visible afterwards.
    await openView(page, "view-menu-outline");
    await openView(page, "view-menu-styles-pane");

    const widthOf = () => panel(page).evaluate((el) => Math.round(el.getBoundingClientRect().width));
    const expandedWidth = await widthOf();

    await handle(page).dblclick();
    await expect(panel(page)).toHaveClass(/collapsed/);
    await expect(page.getByRole("tabpanel")).toBeHidden(); // body hidden
    await expect(page.getByRole("tablist")).toBeVisible(); // rail still shown
    expect((await readState(page)).collapsed).toBe(true);

    // AC-10a: tabs remain reachable + clickable while collapsed.
    await tab(page, "styles").click();
    await expect(tab(page, "styles")).toHaveAttribute("aria-selected", "true");
    await expect(panel(page)).toHaveClass(/collapsed/);

    // Expand restores the persisted width.
    await handle(page).dblclick();
    await expect(panel(page)).not.toHaveClass(/collapsed/);
    expect(await widthOf()).toBe(expandedWidth);

    // AC-10b: collapse persists across a reload.
    await handle(page).dblclick();
    await expect(panel(page)).toHaveClass(/collapsed/);
    await page.reload();
    await expect(panel(page)).toBeVisible();
    await expect(panel(page)).toHaveClass(/collapsed/);
    await expect(page.getByRole("tabpanel")).toBeHidden();
  });

  test("the handle is a vertical separator supporting 8px arrow-key resize (AC-20, AC-20a)", async ({
    page,
  }) => {
    await createDocument(page);
    await openView(page, "view-menu-styles-pane");

    await expect(handle(page)).toHaveAttribute("role", "separator");
    await expect(handle(page)).toHaveAttribute("aria-orientation", "vertical");

    const width = () => readState(page).then((s) => s.widthPx);
    const start = await width();
    await handle(page).focus();
    await page.keyboard.press("ArrowLeft"); // right dock: widen
    expect(await width()).toBe(start + 8);
    await page.keyboard.press("ArrowRight"); // narrow
    expect(await width()).toBe(start);
  });
});

test("persists the whole panel state under one key with the documented shape (AC-11)", async ({
  page,
}) => {
  await createDocument(page);
  await openView(page, "view-menu-styles-pane");
  await openView(page, "view-menu-outline");

  const state = await readState(page);
  expect(Object.keys(state).sort()).toEqual([
    "activePane",
    "collapsed",
    "openPanes",
    "side",
    "widthPx",
  ]);
  expect(typeof state.side).toBe("string");
  expect(typeof state.widthPx).toBe("number");
  expect(typeof state.collapsed).toBe("boolean");
  expect(Array.isArray(state.openPanes)).toBe(true);
  expect(state.openPanes).toContain("styles");
});

test("migrates the three legacy keys once, then deletes them (AC-12, AC-12a, AC-12b, AC-12c, AC-12d)", async ({
  page,
}) => {
  await page.goto("/");
  await page.evaluate(() => {
    localStorage.removeItem("docu-edit:side-panel");
    localStorage.setItem("docu-edit:outline-pane", JSON.stringify({ open: true }));
    localStorage.setItem("docu-edit:comments-pane", JSON.stringify({ open: true }));
    localStorage.setItem("docu-edit:version-history-pane", JSON.stringify({ open: true }));
  });
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(panel(page)).toBeVisible();

  const state = await readState(page);
  expect(state.openPanes).toEqual(["outline", "comments", "versionHistory"]); // AC-12a order
  expect(state.activePane).toBe("outline"); // AC-12b first seeded
  // AC-12c: the legacy keys are gone.
  const legacy = await page.evaluate(() => [
    localStorage.getItem("docu-edit:outline-pane"),
    localStorage.getItem("docu-edit:comments-pane"),
    localStorage.getItem("docu-edit:version-history-pane"),
  ]);
  expect(legacy).toEqual([null, null, null]);

  // AC-12d: re-adding a legacy key and reloading does not re-migrate.
  await page.evaluate(() =>
    localStorage.setItem("docu-edit:outline-pane", JSON.stringify({ open: true })),
  );
  await page.reload();
  await expect(panel(page)).toBeVisible();
  expect((await readState(page)).openPanes).toEqual(["outline", "comments", "versionHistory"]);
  expect(await page.evaluate(() => localStorage.getItem("docu-edit:outline-pane"))).not.toBeNull();
});

test("drops an unknown pane id from openPanes and falls back for an unknown activePane (AC-13)", async ({
  page,
}) => {
  await page.goto("/");
  await page.evaluate(() =>
    localStorage.setItem(
      "docu-edit:side-panel",
      JSON.stringify({
        side: "right",
        widthPx: 320,
        collapsed: false,
        openPanes: ["frobnicate", "outline"],
        activePane: "frobnicate",
      }),
    ),
  );
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  const state = await readState(page);
  expect(state.openPanes).toEqual(["outline"]); // unknown dropped
  expect(state.activePane).toBe("outline"); // unknown activePane coerced to first open
  await expect(tab(page, "frobnicate")).toHaveCount(0);
  await expect(tab(page, "outline")).toBeVisible();
});

test("every menu command keeps its id and follows open/activate/close toggle semantics (AC-15, AC-15a, AC-15b, AC-15c, AC-16)", async ({
  page,
}) => {
  await createDocument(page);

  // AC-15: all seven command ids open their pane.
  await openView(page, "view-menu-outline");
  await openView(page, "view-menu-styles-pane");
  await openView(page, "view-menu-comments-pane");
  await openView(page, "view-menu-track-changes-pane");
  await openFile(page, "file-menu-version-history");
  await openTools(page, "tools-menu-accessibility");
  await openTools(page, "tools-menu-fonts");
  await expect(page.getByRole("tab")).toHaveCount(7);

  // Start clean for the toggle-semantics walk.
  await page.evaluate(() => localStorage.clear());
  await createDocument(page);

  // AC-15a: closed -> open, active, panel open.
  await openView(page, "view-menu-outline");
  await expect(panel(page)).toBeVisible();
  await expect(tab(page, "outline")).toHaveAttribute("aria-selected", "true");

  // AC-15b: open but not active -> make active. Open styles (now active),
  // outline becomes a background tab that stays open.
  await openView(page, "view-menu-styles-pane");
  await expect(tab(page, "outline")).toHaveAttribute("aria-selected", "false");
  // AC-16: the background pane's menu item is still checked ("open", not "active").
  await page.getByTestId("menu-title-view").click();
  await expect(page.getByTestId("view-menu-outline")).toHaveAttribute("aria-checked", "true");
  await page.keyboard.press("Escape");
  await openView(page, "view-menu-outline"); // toggle the background tab
  await expect(tab(page, "outline")).toHaveAttribute("aria-selected", "true");
  await expect(tab(page, "styles")).toBeVisible();

  // AC-15c: open + active -> close that tab; closing the last one closes the panel.
  await openView(page, "view-menu-outline");
  await expect(tab(page, "outline")).toHaveCount(0);
  await expect(tab(page, "styles")).toBeVisible();
  await openView(page, "view-menu-styles-pane");
  await expect(panel(page)).toHaveCount(0);
  expect((await readState(page)).openPanes).toEqual([]);
});

test("exposes tablist/tab/tabpanel ARIA and roving arrow-key navigation (AC-18, AC-18a, AC-18b, AC-19, AC-19a, AC-19b)", async ({
  page,
}) => {
  await createDocument(page);
  await openView(page, "view-menu-outline");
  await openView(page, "view-menu-styles-pane");
  await openView(page, "view-menu-comments-pane"); // comments active, last

  await expect(page.getByRole("tablist")).toHaveCount(1);
  await expect(tab(page, "comments")).toHaveAttribute("aria-selected", "true");
  await expect(tab(page, "comments")).toHaveAttribute("aria-controls", "side-panel-tabpanel");
  await expect(page.getByRole("tabpanel")).toHaveAttribute("id", "side-panel-tabpanel");
  await expect(page.getByRole("tabpanel")).toHaveAttribute(
    "aria-labelledby",
    "side-panel-tab-comments",
  );

  const activeTestId = () => page.evaluate(() => document.activeElement?.getAttribute("data-testid"));
  await tab(page, "comments").focus();
  await page.keyboard.press("ArrowRight"); // wraps to first
  expect(await activeTestId()).toBe("side-panel-tab-outline");
  await page.keyboard.press("ArrowLeft"); // wraps back to last
  expect(await activeTestId()).toBe("side-panel-tab-comments");
  await page.keyboard.press("Home");
  expect(await activeTestId()).toBe("side-panel-tab-outline");
  await page.keyboard.press("End");
  expect(await activeTestId()).toBe("side-panel-tab-comments");

  // AC-19b: Enter/Space activate the focused tab.
  await tab(page, "outline").focus();
  await page.keyboard.press("Enter");
  await expect(tab(page, "outline")).toHaveAttribute("aria-selected", "true");
  await tab(page, "styles").focus();
  await page.keyboard.press(" ");
  await expect(tab(page, "styles")).toHaveAttribute("aria-selected", "true");
});

test("returns focus to the document at the prior caret when the panel closes (AC-21, AC-21a)", async ({
  page,
}) => {
  await createDocument(page);
  await openView(page, "view-menu-styles-pane");

  // AC-21: the resize handle is last in DOM order (document -> rail -> body -> handle).
  const order = await page.evaluate(() => {
    const doc = document.querySelector('[data-testid="document-content"]')!;
    const rail = document.querySelector('[role="tablist"]')!;
    const body = document.querySelector('[role="tabpanel"]')!;
    const h = document.querySelector('[data-testid="side-panel-resize"]')!;
    const FOLLOWS = Node.DOCUMENT_POSITION_FOLLOWING;
    return (
      !!(doc.compareDocumentPosition(rail) & FOLLOWS) &&
      !!(rail.compareDocumentPosition(body) & FOLLOWS) &&
      !!(body.compareDocumentPosition(h) & FOLLOWS)
    );
  });
  expect(order).toBe(true);

  await content(page).click();
  await page.keyboard.type("Hello world here.");
  await page.getByTestId("side-panel-close").click();
  await expect(panel(page)).toHaveCount(0);

  // Focus is back in the document; typing lands at the prior caret.
  const inDoc = await page.evaluate(() => {
    const dc = document.querySelector('[data-testid="document-content"]');
    return !!dc && (dc === document.activeElement || dc.contains(document.activeElement));
  });
  expect(inDoc).toBe(true);
  await page.keyboard.type("XYZ");
  await expect(content(page)).toContainText("Hello world here.XYZ");
});

test("keeps switched-away panes mounted so transient state survives (AC-23)", async ({ page }) => {
  await createDocument(page);
  await openView(page, "view-menu-comments-pane");
  await openView(page, "view-menu-outline");

  await tab(page, "comments").click();
  await page.getByTestId("comment-author-name-input").fill("ZZUNIQUE");
  await tab(page, "outline").click();
  // still mounted while backgrounded (v-show), not destroyed
  await expect(page.getByTestId("comment-author-name-input")).toHaveCount(1);
  await tab(page, "comments").click();
  await expect(page.getByTestId("comment-author-name-input")).toHaveValue("ZZUNIQUE");
});

test("a pane that emits close is removed, the next pane activates, and the panel stays open (AC-24, AC-24a, AC-24b)", async ({
  page,
}) => {
  await createDocument(page);
  await openView(page, "view-menu-outline");
  await openView(page, "view-menu-styles-pane");
  await openView(page, "view-menu-comments-pane");
  expect((await readState(page)).openPanes).toEqual(["outline", "styles", "comments"]);

  await tab(page, "styles").click();
  await page.getByTestId("styles-pane-close").click(); // pane emits "close"

  expect((await readState(page)).openPanes).toEqual(["outline", "comments"]);
  await expect(tab(page, "styles")).toHaveCount(0);
  await expect(panel(page)).toBeVisible(); // AC-24b: stays open
  expect((await readState(page)).collapsed).toBe(false);
  // AC-24a: the next open pane becomes active.
  await expect(tab(page, "comments")).toHaveAttribute("aria-selected", "true");
});

test("shows a live count badge only for panes that report one (AC-6, AC-6a, AC-6c)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("The quick brown fox.");

  await openView(page, "view-menu-comments-pane");
  await expect(badge(page, "comments")).toHaveCount(0); // AC-6a: zero -> no badge

  // Add a comment on "The" -> badge shows 1 (AC-6).
  await content(page).click();
  await page.keyboard.press("Home");
  await page.keyboard.down("Shift");
  for (let i = 0; i < 3; i++) await page.keyboard.press("ArrowRight");
  await page.keyboard.up("Shift");
  await page.keyboard.press(`${MOD}+Alt+m`);
  await page.getByTestId("comment-composer-input").fill("a note");
  await page.getByTestId("comment-composer-submit").click();
  await expect(badge(page, "comments")).toHaveText("1");

  // AC-6c: a pane without a badge property never shows one.
  await openView(page, "view-menu-outline");
  await expect(badge(page, "outline")).toHaveCount(0);
});

test.describe("responsive", () => {
  test("overlays with a scrim below 900px, closing on scrim-click or Escape, without reflowing the document (AC-25, AC-25a, AC-25b, AC-25c)", async ({
    page,
  }) => {
    await page.setViewportSize({ width: 800, height: 800 });
    await createDocument(page);

    const docWidth = () => content(page).evaluate((el) => Math.round(el.getBoundingClientRect().width));
    const closedWidth = await docWidth();

    await openView(page, "view-menu-styles-pane");
    await expect(panel(page)).toBeVisible();
    // Out-of-flow overlay + a semi-transparent scrim behind it.
    await expect(panel(page)).toHaveCSS("position", "fixed");
    await expect(page.getByTestId("side-panel-scrim")).toBeVisible();
    // AC-25c: the document does not reflow when the overlay opens.
    expect(await docWidth()).toBe(closedWidth);

    // AC-25b: Escape closes it.
    await page.keyboard.press("Escape");
    await expect(panel(page)).toHaveCount(0);

    // AC-25a: clicking the scrim closes it.
    await openView(page, "view-menu-styles-pane");
    await page.getByTestId("side-panel-scrim").click();
    await expect(panel(page)).toHaveCount(0);
  });

  test("at 900px and above the panel displaces the document and shows no scrim (AC-26)", async ({
    page,
  }) => {
    await page.setViewportSize({ width: 1200, height: 800 });
    await createDocument(page);

    const docWidth = () => content(page).evaluate((el) => Math.round(el.getBoundingClientRect().width));
    const closedWidth = await docWidth();

    await openView(page, "view-menu-styles-pane");
    await expect(panel(page)).toBeVisible();
    await expect(panel(page)).not.toHaveCSS("position", "fixed"); // in flex flow
    await expect(page.getByTestId("side-panel-scrim")).toHaveCount(0);
    expect(await docWidth()).toBeLessThan(closedWidth); // displaced
  });
});
