import { test, expect, type Page } from "@playwright/test";
import { openHomeTab, openInsertTab } from "./slides";

// SPEC-049 — menu bar and toolbar consolidation. A menu bar (File, Edit, View,
// Insert, Format, Tools) sits between the app bar and the toolbar strip and
// absorbs every non-continuous-use control and the seven ad-hoc dropdowns, so
// the toolbar strip keeps only the continuous-use controls and stops
// overflowing at 1280px. The bar is a keyboard-navigable ARIA menubar whose
// items show SPEC-021 shortcut labels, disable (never hide) inapplicable
// commands, and render toggles as checked items reflecting live state.

// SPEC-076 AC-16 adds a seventh top-level menu, Table, between Format and
// Tools — so the bar this spec pins is now seven menus wide.
const MENU_ORDER = ["file", "edit", "view", "insert", "format", "table", "tools"] as const;
const MENU_LABELS = ["File", "Edit", "View", "Insert", "Format", "Table", "Tools"];

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("document-content")).toBeVisible();
}

function title(page: Page, id: string) {
  return page.getByTestId(`menu-title-${id}`);
}
function popup(page: Page, id: string) {
  return page.getByTestId(`menu-${id}`);
}
async function openMenu(page: Page, id: string): Promise<void> {
  await title(page, id).click();
  await expect(popup(page, id)).toBeVisible();
}
async function closeMenus(page: Page): Promise<void> {
  // Titles use @mousedown.prevent, so a click-opened menu leaves focus off the
  // title and Escape would not reach it. Close the way a user's pointer does —
  // a click outside the bar — which the menubar's document-mousedown handler
  // treats as dismiss. The neutral point is the app bar, above the menu bar.
  await page.mouse.click(640, 18);
  await expect(page.locator("[data-menu-popup-open]")).toHaveCount(0);
}

test.describe("SPEC-049 menu bar and toolbar consolidation", () => {
  test.beforeEach(async ({ page }) => {
    await createDocument(page);
  });

  // AC-1: menu bar between the app bar and the toolbar strip, menus in order.
  test("AC-1 menu bar sits between app bar and toolbar with menus in order", async ({ page }) => {
    const bar = page.getByTestId("menu-bar");
    await expect(bar).toBeVisible();

    // The Toolbar Redesign sets the bar's titles in the caption face, so what
    // the user reads is the uppercase form; the DOM text is unchanged (this is
    // text-transform, and allInnerTexts reports the rendered casing).
    const labels = await bar.locator("[data-menu-title]").allInnerTexts();
    expect(labels.map((t) => t.trim())).toEqual(MENU_LABELS.map((l) => l.toUpperCase()));
    const domLabels = await bar.locator("[data-menu-title]").allTextContents();
    expect(domLabels.map((t) => t.trim())).toEqual(MENU_LABELS);

    const barBox = await bar.boundingBox();
    const toolbarBox = await page.getByTestId("toolbar").boundingBox();
    expect(barBox).not.toBeNull();
    expect(toolbarBox).not.toBeNull();
    // Menu bar renders above (before) the toolbar strip.
    expect(barBox!.y).toBeLessThan(toolbarBox!.y);
  });

  // AC-2: no command appears in two menus (item test-ids are unique across all
  // menus). The toolbar-duplicated controls (undo/redo, colour, highlight) are
  // the deliberate toolbar+menu exception and do not violate menu-to-menu.
  test("AC-2 no command appears in two menus", async ({ page }) => {
    const seen = new Map<string, number>();
    for (const id of MENU_ORDER) {
      await openMenu(page, id);
      const ids = await popup(page, id)
        .locator("[role=menuitem][data-testid], [role=menuitemcheckbox][data-testid]")
        .evaluateAll((els) => els.map((e) => e.getAttribute("data-testid")!));
      for (const tid of ids) {
        seen.set(tid, (seen.get(tid) ?? 0) + 1);
      }
      await closeMenus(page);
    }
    const duplicates = [...seen.entries()].filter(([, n]) => n > 1);
    expect(duplicates).toEqual([]);
    // Sanity: menus actually carried a meaningful number of commands.
    expect(seen.size).toBeGreaterThan(20);
  });

  // AC-3: the toolbar strip retains only the continuous-use controls, and the
  // non-continuous controls no longer live in it.
  test("AC-3 toolbar retains continuous-use controls only", async ({ page }) => {
    const toolbar = page.getByTestId("toolbar");
    // Continuous-use controls remain in the strip.
    await expect(toolbar.getByTestId("style-picker-button")).toBeVisible();
    await expect(toolbar.getByTestId("format-bold")).toBeVisible();
    await expect(toolbar.getByTestId("format-italic")).toBeVisible();
    await expect(toolbar.getByTestId("format-underline")).toBeVisible();
    await expect(toolbar.getByTestId("format-strikethrough")).toBeVisible();

    // Non-continuous controls are gone from the toolbar strip (they now live
    // only inside menu popups).
    for (const tid of [
      "format-painter-button",
      "clear-formatting-button",
      "columns-button",
      "print-button",
      "page-setup-button",
    ]) {
      await expect(toolbar.getByTestId(tid)).toHaveCount(0);
    }
  });

  // AC-4: the toolbar strip no longer overflows at 1280px.
  test("AC-4 toolbar strip does not overflow at 1280px", async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 900 });
    const overflow = await page.getByTestId("toolbar").evaluate((el) => el.scrollWidth > el.clientWidth + 1);
    expect(overflow).toBe(false);
  });

  // AC-5: the seven ad-hoc dropdowns are absorbed as menu sections and no
  // longer exist as standalone toolbar dropdown buttons.
  test("AC-5 the seven ad-hoc dropdowns are absorbed into the menus", async ({ page }) => {
    // Representative command from each absorbed dropdown, in its new menu.
    const absorbed: Array<[string, string]> = [
      ["insert", "break-page"], // Insert break
      ["insert", "notes-menu-footnote"], // Notes
      ["insert", "hf-menu-header"], // Header/footer
      ["insert", "table-grid-cell-1-1"], // table insert control
      ["format", "multilevel-define-list"], // Multilevel list
      ["edit", "autocorrect-smart-quotes"], // Autocorrect
      ["view", "toggle-ruler-button"], // View
    ];
    for (const [menu, tid] of absorbed) {
      await openMenu(page, menu);
      await expect(popup(page, menu).getByTestId(tid)).toHaveCount(1);
      await closeMenus(page);
    }

    // The old standalone dropdown toolbar buttons are gone from the strip.
    const toolbar = page.getByTestId("toolbar");
    for (const tid of [
      "insert-break-menu",
      "notes-menu",
      "header-footer-menu",
      "multilevel-list-menu",
      "autocorrect-menu",
      "view-menu",
      "table-insert-control",
    ]) {
      await expect(toolbar.getByTestId(tid)).toHaveCount(0);
    }
  });

  // AC-6: a menu item with a keyboard shortcut shows that shortcut's label
  // (read from the SPEC-021 registry, not restated in the menu).
  test("AC-6 menu items show shortcut labels from the registry", async ({ page }) => {
    await openMenu(page, "file");
    // Print binds Ctrl/Cmd+P in shortcuts.ts; the label is rendered, not typed.
    await expect(page.getByTestId("file-menu-print")).toContainText(/Ctrl\+P|⌘P|Cmd\+P/);
    await closeMenus(page);

    await openMenu(page, "insert");
    await expect(page.getByTestId("break-page")).toContainText(/Ctrl\+Enter|⌘.*Enter|Cmd\+Enter/);
    await closeMenus(page);
  });

  // AC-7: a command that does not apply is present but disabled, never hidden.
  test("AC-7 inapplicable commands are disabled, not hidden", async ({ page }) => {
    // On a fresh document there is nothing to undo/redo.
    await openMenu(page, "edit");
    const undo = page.getByTestId("edit-menu-undo");
    await expect(undo).toBeVisible();
    await expect(undo).toBeDisabled();
    const redo = page.getByTestId("edit-menu-redo");
    await expect(redo).toBeVisible();
    await expect(redo).toBeDisabled();
    await closeMenus(page);
  });

  // AC-8: one tab stop; Left/Right between menus, Up/Down within an open menu,
  // Enter opens/activates, Escape closes and returns focus to the menu title.
  test("AC-8 keyboard model: arrows, enter, escape", async ({ page }) => {
    await title(page, "file").focus();
    await expect(title(page, "file")).toBeFocused();

    await page.keyboard.press("ArrowRight");
    await expect(title(page, "edit")).toBeFocused();
    await page.keyboard.press("ArrowLeft");
    await expect(title(page, "file")).toBeFocused();

    // Enter opens the menu and focus moves into it.
    await page.keyboard.press("Enter");
    await expect(title(page, "file")).toHaveAttribute("aria-expanded", "true");
    const first = await page.evaluate(() => document.activeElement?.getAttribute("data-testid"));
    await page.keyboard.press("ArrowDown");
    const second = await page.evaluate(() => document.activeElement?.getAttribute("data-testid"));
    expect(first).toBeTruthy();
    expect(second).toBeTruthy();
    expect(second).not.toBe(first);

    // Escape closes and returns focus to the menu title.
    await page.keyboard.press("Escape");
    await expect(title(page, "file")).toHaveAttribute("aria-expanded", "false");
    await expect(title(page, "file")).toBeFocused();
  });

  // AC-9: with a menu open, moving the pointer to a neighbour opens it without
  // a second click.
  test("AC-9 pointer cross-open between menus", async ({ page }) => {
    await openMenu(page, "file");
    await title(page, "edit").hover();
    await expect(popup(page, "edit")).toBeVisible();
    await expect(popup(page, "file")).toBeHidden();
  });

  // AC-10: ARIA menubar/menu/menuitem structure and expanded state.
  test("AC-10 exposes menubar/menu/menuitem roles and expanded state", async ({ page }) => {
    await expect(page.getByTestId("menu-bar")).toHaveAttribute("role", "menubar");

    await expect(title(page, "file")).toHaveAttribute("aria-expanded", "false");
    await openMenu(page, "file");
    await expect(title(page, "file")).toHaveAttribute("aria-expanded", "true");
    await expect(popup(page, "file")).toHaveAttribute("role", "menu");
    await expect(popup(page, "file").getByRole("menuitem").first()).toBeVisible();
    await closeMenus(page);

    // Toggle commands carry the menuitemcheckbox role.
    await openMenu(page, "view");
    await expect(page.getByTestId("toggle-ruler-button")).toHaveAttribute("role", "menuitemcheckbox");
    await closeMenus(page);
  });

  // AC-11: toggle commands render as checked items reflecting live state.
  test("AC-11 toggle commands render as checked items reflecting live state", async ({ page }) => {
    await openMenu(page, "view");
    // Ruler is on by default → checked.
    await expect(page.getByTestId("toggle-ruler-button")).toHaveAttribute("aria-checked", "true");

    // Comments-only view starts off; toggling flips the live check state.
    const commentsOnly = page.getByTestId("view-menu-comments-only");
    await expect(commentsOnly).toHaveAttribute("aria-checked", "false");
    await commentsOnly.click();
    await openMenu(page, "view");
    await expect(page.getByTestId("view-menu-comments-only")).toHaveAttribute("aria-checked", "true");
    // Restore.
    await page.getByTestId("view-menu-comments-only").click();
    await closeMenus(page);
  });

  // AC-12: at 1024px every menu shows without scrolling; toolbar still usable.
  test("AC-12 all seven menus visible at 1024px, toolbar usable", async ({ page }) => {
    await page.setViewportSize({ width: 1024, height: 900 });
    const bar = page.getByTestId("menu-bar");
    const titles = bar.locator("[data-menu-title]");
    await expect(titles).toHaveCount(MENU_ORDER.length);
    for (const id of MENU_ORDER) {
      await expect(title(page, id)).toBeInViewport();
    }
    const barScrolls = await bar.evaluate((el) => el.scrollWidth > el.clientWidth + 1);
    expect(barScrolls).toBe(false);

    // Toolbar remains usable per SPEC-001 — a core control is present.
    await expect(page.getByTestId("toolbar").getByTestId("format-bold")).toBeVisible();
  });
});

// The slides half of SPEC-049's model. The deck editor had only strips of
// buttons, so the same application offered two different ways to find a
// command — and named the same action twice ("Download" against "Export").
// The bar is a second route to commands the strips already carry, so these
// assert the routing rather than the commands, which their own specs cover.
//
// Four menus, where the word processor has seven: the deck's bar shares its
// row with the ribbon's Home and Insert tabs, so an "Insert" menu sat a
// thumb's width from the "Insert" tab in the same condensed uppercase face,
// and a "Format" menu beside a Home tab holding the same commands. Both were
// dropped INTO those tabs rather than deleted — the assertion below pins that
// every command they carried is still reachable from the strip that replaced
// them, so the shorter bar cannot quietly become a shorter product.
const SLIDE_MENU_LABELS = ["File", "Edit", "View", "Slide"];

async function newDeck(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-presentation").click();
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

function slideMenu(page: Page) {
  return page.getByRole("menubar", { name: "Presentation menu" });
}

test("the presentation editor has the same menu bar model as the document editor", async ({ page }) => {
  await newDeck(page);
  const bar = slideMenu(page);
  await expect(bar).toBeVisible();
  await expect(bar.getByRole("menuitem")).toHaveText(SLIDE_MENU_LABELS);
  // Insert and Format are the ribbon's two tabs, in the same row as the bar.
  await expect(page.getByTestId("slides-tab-insert")).toBeVisible();
  await expect(page.getByTestId("slides-tab-home")).toBeVisible();
});

// What the two dropped menus carried, item by item, now reachable from the tab
// that replaced them. Without this the menus could be removed and the commands
// quietly go with them.
test("every command the Insert and Format menus carried is on the strip that replaced them", async ({ page }) => {
  await newDeck(page);

  // Format ▸ Bold / Italic / Underline / Bulleted list / both indents.
  await openHomeTab(page);
  for (const testid of ["slides-bold", "slides-italic", "slides-underline", "slides-bullet", "slides-indent", "slides-outdent"]) {
    await expect(page.getByTestId(testid), `Format ▸ ${testid} lost`).toBeVisible();
  }
  // Format ▸ Align text ▸ Left / Centre / Right / Justified.
  const align = page.getByTestId("slides-align");
  for (const value of ["left", "center", "right", "justify"]) {
    await expect(align.locator(`[data-align="${value}"]`), `Format ▸ align ${value} lost`).toBeVisible();
  }

  // Insert ▸ Text box / Image… / Audio or video… / Shape ▸ / Chart ▸.
  await openInsertTab(page);
  for (const testid of ["insert-text-box", "insert-image", "insert-media", "insert-shape", "insert-chart"]) {
    await expect(page.getByTestId(testid), `Insert ▸ ${testid} lost`).toBeVisible();
  }
});

test("the slides menu runs commands and reflects their state", async ({ page }) => {
  await newDeck(page);
  const bar = slideMenu(page);
  const open = async (label: string) => bar.getByRole("menuitem", { name: label, exact: true }).click();

  // Undo starts unavailable on a fresh deck — a command is disabled, never
  // hidden (SPEC-049), so the menu stays a stable map of the product.
  await open("Edit");
  await expect(page.getByTestId("slides-menu-undo")).toBeDisabled();
  await page.keyboard.press("Escape");

  // Slide ▸ New slide adds one, which in turn makes Undo available.
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(1);
  await open("Slide");
  await page.getByTestId("slides-menu-new-slide").click();
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(2);

  await open("Edit");
  await expect(page.getByTestId("slides-menu-undo")).toBeEnabled();
  await page.getByTestId("slides-menu-undo").click();
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(1);
});

test("the slides View menu checks the panel it opened", async ({ page }) => {
  await newDeck(page);
  const bar = slideMenu(page);
  const openView = async () => bar.getByRole("menuitem", { name: "View", exact: true }).click();

  await openView();
  const design = page.getByTestId("slides-menu-design");
  await expect(design).toHaveAttribute("aria-checked", "false");
  await design.click();

  // Reopening reads the checkmark back from the panel's own state, so the
  // menu and the right-hand strip cannot disagree about what is open.
  await openView();
  await expect(page.getByTestId("slides-menu-design")).toHaveAttribute("aria-checked", "true");
});
