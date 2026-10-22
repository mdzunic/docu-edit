import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-063 — presentation-layer zoom (a CSS scale transform on the pages,
// never repaginating) plus three view modes (Print / Web / Focus). This spec
// guards the criteria that gate the merge, with particular attention to the
// three that regressed in QA and were fixed: Web-layout page chrome hidden
// (AC-39), the ruler staying registered with the page under zoom (AC-54), and
// the fixed-position selection toolbar staying registered with its selection
// under zoom (AC-57).

async function newDoc(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("document-content")).toBeVisible();
}

/** Push the document across page boundaries so the multi-page grid (AC-36) and
 * the Web-layout continuous column (AC-38/39) have several pages to work with. */
async function seedMultiPage(page: Page): Promise<void> {
  const content = page.getByTestId("document-content");
  await content.click();
  for (let i = 0; i < 8; i++) {
    await page.keyboard.type("The quick brown fox jumps over the lazy dog. ");
  }
  for (let i = 0; i < 90; i++) {
    await page.keyboard.press("Enter");
  }
  await page.keyboard.type("Final line.");
  await expect
    .poll(async () => page.locator('[data-testid="document-page"]').count())
    .toBeGreaterThan(1);
}

/** Set the zoom via the status-bar percentage field (commits on blur). Does not
 * depend on the editor being focused, unlike the keyboard ladder. */
async function setZoomInput(page: Page, value: string): Promise<void> {
  const input = page.getByTestId("zoom-percent");
  await input.click();
  await input.fill(value);
  await input.blur();
}

/** Pick a preset from the zoom menu (does not depend on editor focus). */
async function pickZoomPreset(page: Page, preset: number): Promise<void> {
  await page.getByTestId("zoom-percent").click();
  await page.getByTestId(`zoom-preset-${preset}`).click();
}

async function openLayoutMenu(page: Page): Promise<void> {
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-layout").click();
}

test.describe("SPEC-063 zoom and view modes", () => {
  test("zoom control sits in the status bar with −/percent/+ and the preset menu (AC-1..AC-5)", async ({
    page,
  }) => {
    await newDoc(page);
    const zc = page.getByTestId("zoom-control");
    await expect(zc).toBeVisible();
    // Exactly [zoom-out, zoom-percent, zoom-in] as direct children (AC-2, AC-3).
    const childIds = await zc.evaluate((el) =>
      [...el.children].map((c) => c.getAttribute("data-testid")),
    );
    expect(childIds).toEqual(["zoom-out", "zoom-percent", "zoom-in"]);

    // Menu entries in order, then Fit width / Fit page (AC-4).
    await page.getByTestId("zoom-percent").click();
    await expect(page.getByTestId("zoom-menu")).toBeVisible();
    const labels = await page
      .getByTestId("zoom-menu")
      .locator('[role="menuitemradio"]')
      .allInnerTexts();
    expect(labels.map((s) => s.trim())).toEqual([
      "50%",
      "75%",
      "90%",
      "100%",
      "125%",
      "150%",
      "200%",
      "Fit width",
      "Fit page",
    ]);
    // 100% is checked at default zoom (AC-5).
    await expect(page.getByTestId("zoom-preset-100")).toHaveAttribute("aria-checked", "true");
  });

  test("the −/+ ladder clamps, off-ladder snaps, and the field edits/clamps/reverts (AC-6..AC-10)", async ({
    page,
  }) => {
    await newDoc(page);
    const input = page.getByTestId("zoom-percent");
    const content = page.getByTestId("document-content");

    // + clamps at 200, − clamps at 50 (AC-6). Keyboard ladder needs editor focus.
    await content.click();
    for (let i = 0; i < 5; i++) await page.keyboard.press(`${MOD}+=`);
    await expect(input).toHaveValue("200%");
    for (let i = 0; i < 8; i++) await page.keyboard.press(`${MOD}+-`);
    await expect(input).toHaveValue("50%");

    // Off-ladder 110%: + → 125, − → 100 (AC-7).
    await setZoomInput(page, "110");
    await content.click();
    await page.keyboard.press(`${MOD}+=`);
    await expect(input).toHaveValue("125%");
    await setZoomInput(page, "110");
    await content.click();
    await page.keyboard.press(`${MOD}+-`);
    await expect(input).toHaveValue("100%");

    // Editable field: whitespace / trailing % trimmed (AC-8).
    await setZoomInput(page, "  240  ");
    await expect(input).toHaveValue("240%");
    await setZoomInput(page, "88%");
    await expect(input).toHaveValue("88%");
    // Clamp to 25..400 (AC-9).
    await setZoomInput(page, "10");
    await expect(input).toHaveValue("25%");
    await setZoomInput(page, "500");
    await expect(input).toHaveValue("400%");
    // Empty / non-numeric revert to last valid on blur (AC-10).
    await setZoomInput(page, "120");
    await setZoomInput(page, "abc");
    await expect(input).toHaveValue("120%");
    await setZoomInput(page, "");
    await expect(input).toHaveValue("120%");
  });

  test("zoom is a single scale transform with a top-center origin; the ruler scales in step without stretching (AC-16, AC-53)", async ({
    page,
  }) => {
    await newDoc(page);
    await setZoomInput(page, "150");
    const geom = await page.evaluate(() => {
      const pages = document.querySelector<HTMLElement>(".pages--scaled") ??
        document.querySelector<HTMLElement>('[data-testid="document-pages"]')!;
      const ruler = document.querySelector<HTMLElement>('[data-testid="document-ruler"]')!;
      const rulerBox = ruler.getBoundingClientRect();
      // The first labelled tick after the origin ("1") and the selector glyph.
      const labels = Array.from(ruler.querySelectorAll<HTMLElement>(".tick-label"));
      const one = labels.find((el) => el.textContent?.trim() === "1")!;
      const tickOne = one.parentElement!.getBoundingClientRect();
      const selector = ruler.querySelector<HTMLElement>('[data-testid="ruler-tab-selector"]')!.getBoundingClientRect();
      return {
        transform: getComputedStyle(pages).transform,
        origin: getComputedStyle(pages).transformOrigin,
        rulerTransform: getComputedStyle(ruler).transform,
        rulerWidth: rulerBox.width,
        rulerHeight: rulerBox.height,
        // Ticks are positioned inside the ruler's 1px border: measure from the
        // padding box (clientLeft is that border width).
        tickOneOffset: tickOne.left - (rulerBox.left + ruler.clientLeft),
        selectorAspect: selector.width / selector.height,
      };
    });
    // scale(1.5) on the pages…
    expect(geom.transform).toBe("matrix(1.5, 0, 0, 1.5, 0, 0)");
    // transform-origin's x is the horizontal centre (its y is the top: 0px).
    expect(geom.origin.endsWith(" 0px")).toBeTruthy();
    // …while the ruler scales by layout, not by a stretching transform: it is
    // laid out at the zoomed page width (21cm × 1.5 at 96dpi), its 1cm tick
    // sits 1.5cm from its left edge, and its chrome keeps its proportions —
    // the 1.5rem height and the square tab selector are not widened.
    expect(geom.rulerTransform).toBe("none");
    expect(geom.rulerWidth).toBeCloseTo(21 * 1.5 * (96 / 2.54), 0);
    expect(geom.tickOneOffset).toBeCloseTo(1.5 * (96 / 2.54), 0);
    expect(geom.rulerHeight).toBe(24);
    expect(geom.selectorAspect).toBeCloseTo(1, 1);
  });

  test("the ruler stays registered with the page at 50% and 200% (AC-53, AC-54)", async ({
    page,
  }) => {
    await newDoc(page);
    await page.getByTestId("document-content").click();
    await page.keyboard.type("Hello world — a single line for ruler alignment.");

    async function offsets() {
      return page.evaluate(() => {
        const ruler = document.querySelector('[data-testid="document-ruler"]')!.getBoundingClientRect();
        const pg = document.querySelector('[data-testid="document-page"]')!.getBoundingClientRect();
        return {
          dxCenter: ruler.left + ruler.width / 2 - (pg.left + pg.width / 2),
          rulerW: ruler.width,
          pageW: pg.width,
        };
      });
    }

    for (const pct of ["50", "200", "100"]) {
      await setZoomInput(page, pct);
      const o = await offsets();
      // A ruler centimetre stays over the document centimetre: the ruler and the
      // page share a centre (AC-54) and the ruler's laid-out width scales with
      // the page (AC-53).
      expect(Math.abs(o.dxCenter)).toBeLessThanOrEqual(2);
      expect(Math.abs(o.rulerW - o.pageW)).toBeLessThanOrEqual(2);
    }
  });

  test("spell underlines stay on their words straight after a zoom change, without a scroll (AC-25)", async ({
    page,
  }) => {
    await newDoc(page);
    await page.getByTestId("document-content").click();
    await page.keyboard.type("This is a teh mispeled wrod on the page.");
    const underline = page.locator('[data-testid="spell-underline"]').first();
    await expect(underline).toBeVisible();

    // The underline's offset from its word's glyph box, in client px.
    async function offset() {
      return page.evaluate(() => {
        const el = document.querySelector('[data-testid="spell-underline"]')!;
        const word = el.getAttribute("data-word") ?? "";
        const box = el.getBoundingClientRect();
        const block = document.querySelector('[data-testid="document-content"] [data-section][data-block]')!;
        const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
        let node: Node | null;
        while ((node = walker.nextNode())) {
          const text = node.textContent ?? "";
          const at = text.indexOf(word);
          if (at >= 0) {
            const range = document.createRange();
            range.setStart(node, at);
            range.setEnd(node, at + word.length);
            const glyph = range.getBoundingClientRect();
            return { dx: box.left - glyph.left, dy: box.top - glyph.top, dw: box.width - glyph.width };
          }
        }
        return { dx: NaN, dy: NaN, dw: NaN };
      });
    }

    // The `.pages` box changes shape between the unscaled (100 %) and scaled
    // layouts, so a rect measured in one is wrong in the other until it is
    // re-measured: every zoom change must re-measure on its own, with no
    // scroll to prompt it.
    for (const pct of ["150", "50", "100", "200"]) {
      await setZoomInput(page, pct);
      await expect.poll(async () => Math.abs((await offset()).dx), { timeout: 3000 }).toBeLessThanOrEqual(1);
      const o = await offset();
      expect(Math.abs(o.dy)).toBeLessThanOrEqual(1);
      expect(Math.abs(o.dw)).toBeLessThanOrEqual(1);
    }
  });

  test("the fixed selection toolbar stays over its selection under zoom (AC-25, AC-57)", async ({
    page,
  }) => {
    await newDoc(page);
    const content = page.getByTestId("document-content");
    await content.click();
    await page.keyboard.type("one two three four five six seven");

    // Select a couple of words; the toolbar appears on selection.
    await page.keyboard.press(`${MOD}+Home`);
    await page.keyboard.press("Shift+Control+ArrowRight");
    await page.keyboard.press("Shift+Control+ArrowRight");
    const toolbar = page.getByTestId("text-selection-toolbar");
    await expect(toolbar).toBeVisible();

    async function dx() {
      return page.evaluate(() => {
        const tb = document.querySelector('[data-testid="text-selection-toolbar"]')!.getBoundingClientRect();
        const sel = window.getSelection()!.getRangeAt(0).getBoundingClientRect();
        return tb.left + tb.width / 2 - (sel.left + sel.width / 2);
      });
    }

    expect(Math.abs(await dx())).toBeLessThanOrEqual(3);
    // Zoom via the keyboard ladder so the editor keeps focus and the native
    // selection survives — a zoom fires none of the events the toolbar listens
    // to, so it must re-measure on the zoom change itself (AC-57).
    for (let i = 0; i < 3; i++) await page.keyboard.press(`${MOD}+-`); // → 50%
    await expect(page.getByTestId("zoom-percent")).toHaveValue("50%");
    await expect(toolbar).toBeVisible();
    expect(Math.abs(await dx())).toBeLessThanOrEqual(3);

    await page.keyboard.press(`${MOD}+0`);
    for (let i = 0; i < 4; i++) await page.keyboard.press(`${MOD}+=`); // → 200%
    await expect(page.getByTestId("zoom-percent")).toHaveValue("200%");
    await expect(toolbar).toBeVisible();
    expect(Math.abs(await dx())).toBeLessThanOrEqual(3);
  });

  test("caret placement lands where the pointer is at 50% and 200% (AC-24)", async ({ page }) => {
    await newDoc(page);
    await page.getByTestId("document-content").click();
    await page.keyboard.type("one two three four five six seven");

    async function clickFiveAndRead(pct: string) {
      await setZoomInput(page, pct);
      const rect = await page.evaluate(() => {
        const walker = document.createTreeWalker(
          document.querySelector(".document-content")!,
          NodeFilter.SHOW_TEXT,
        );
        let node: Node | null;
        while ((node = walker.nextNode())) {
          const i = (node.textContent ?? "").indexOf("five");
          if (i >= 0) {
            const r = document.createRange();
            r.setStart(node, i);
            r.setEnd(node, i + 4);
            const b = r.getBoundingClientRect();
            return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
          }
        }
        return null;
      });
      await page.mouse.click(rect!.x, rect!.y);
      return page.evaluate(() => {
        const r = window.getSelection()!.getRangeAt(0);
        const t = r.startContainer.textContent ?? "";
        return t.slice(Math.max(0, r.startOffset - 2), r.startOffset + 2);
      });
    }

    // Clicking the middle of "five" lands the caret inside the word at both zooms.
    expect(await clickFiveAndRead("50")).toMatch(/i.?v/);
    expect(await clickFiveAndRead("200")).toMatch(/i.?v/);
  });

  test("Ctrl+wheel zooms the document and prevents the browser's page zoom (AC-13, AC-14)", async ({
    page,
  }) => {
    await newDoc(page);
    await setZoomInput(page, "50");
    const defaultPrevented = await page.evaluate(() => {
      const el = document.querySelector(".document-content") ?? document.body;
      const r = el.getBoundingClientRect();
      const ev = new WheelEvent("wheel", {
        deltaY: -100,
        ctrlKey: true,
        cancelable: true,
        bubbles: true,
        clientX: r.left + r.width / 2,
        clientY: r.top + 50,
      });
      return !el.dispatchEvent(ev);
    });
    expect(defaultPrevented).toBe(true);
    await expect(page.getByTestId("zoom-percent")).toHaveValue("75%");
  });

  test("Web layout hides page chrome and constrains the column to the section content width (AC-38..AC-41)", async ({
    page,
  }) => {
    await newDoc(page);
    await seedMultiPage(page);

    await openLayoutMenu(page);
    await page.getByTestId("layout-web").click();
    await expect(page.locator(".is-web-layout")).toHaveCount(1);

    const info = await page.evaluate(() => {
      const pages = [...document.querySelectorAll('[data-testid="document-page"]')];
      const cs = getComputedStyle(pages[0] as HTMLElement);
      const header = document.querySelector('[data-testid="page-header-band"]');
      const footer = document.querySelector('[data-testid="page-footer-band"]');
      const gap =
        pages.length > 1
          ? pages[1].getBoundingClientRect().top - pages[0].getBoundingClientRect().bottom
          : 0;
      const fullPage = pages[0].getBoundingClientRect().width;
      return {
        boxShadow: cs.boxShadow,
        marginBottom: cs.marginBottom,
        gap,
        headerDisplay: header ? getComputedStyle(header).display : "none",
        footerDisplay: footer ? getComputedStyle(footer).display : "none",
        pageWidth: fullPage,
      };
    });
    // AC-39: frames, gaps, headers and footers hidden.
    expect(info.boxShadow).toBe("none");
    expect(info.marginBottom).toBe("0px");
    expect(Math.abs(info.gap)).toBeLessThanOrEqual(1);
    expect(info.headerDisplay).toBe("none");
    expect(info.footerDisplay).toBe("none");
    // AC-40: content width < full page width (a printed A4 page is ~794px here).
    expect(info.pageWidth).toBeLessThan(700);

    // AC-41: pagination still runs; the indicator reads "Page n of m".
    await expect(page.getByTestId("page-indicator")).toContainText(/Page \d+ of \d+/);
  });

  test("Print layout tiles multiple pages in a centred grid when zoom leaves room (AC-36)", async ({
    page,
  }) => {
    await newDoc(page);
    await seedMultiPage(page);
    await pickZoomPreset(page, 50);

    const grid = await page.evaluate(() => {
      const dc = document.querySelector<HTMLElement>(".document-content")!;
      const cs = getComputedStyle(dc);
      const rects = [...document.querySelectorAll('[data-testid="document-page"]')].map((p) =>
        p.getBoundingClientRect(),
      );
      const top0 = rects[0].top;
      const row1 = rects.filter((r) => Math.abs(r.top - top0) < 5);
      const gLeft = Math.min(...row1.map((r) => r.left));
      const gRight = Math.max(...row1.map((r) => r.left + r.width));
      return {
        isGrid: dc.classList.contains("is-page-grid"),
        display: cs.display,
        justify: cs.justifyContent,
        cols: Number(cs.getPropertyValue("--page-grid-cols").trim()),
        row1: row1.length,
        gap: row1.length > 1 ? row1[1].left - (row1[0].left + row1[0].width) : 0,
        gridCx: (gLeft + gRight) / 2,
        vpCx: window.innerWidth / 2,
      };
    });
    expect(grid.isGrid).toBe(true);
    expect(grid.display).toBe("grid");
    expect(grid.justify).toBe("center");
    expect(grid.cols).toBeGreaterThan(1);
    expect(grid.row1).toBeGreaterThan(1);
    // 24px gap between columns, and the row is centred in the viewport.
    expect(Math.abs(grid.gap - 24)).toBeLessThanOrEqual(3);
    expect(Math.abs(grid.gridCx - grid.vpCx)).toBeLessThanOrEqual(3);
  });

  test("Focus mode hides chrome but keeps the status bar, toggles on Ctrl+Shift+F / Escape, and restores panes (AC-44..AC-48)", async ({
    page,
  }) => {
    await newDoc(page);
    // Open a pane so the capture/restore path is exercised (AC-47).
    await page.getByTestId("menu-title-view").click();
    await page.getByTestId("view-menu-styles-pane").click();
    const stylesPane = page.locator('[data-testid="styles-pane"], .styles-pane').first();
    await expect(stylesPane).toBeVisible();

    await page.getByTestId("document-content").click();
    await page.keyboard.press(`${MOD}+Shift+F`);
    await expect(page.locator(".is-focus-mode")).toHaveCount(1);
    // Menu bar and the open pane hidden; status bar chrome stays (AC-44/AC-45/AC-56).
    await expect(page.getByTestId("menu-bar")).toBeHidden();
    await expect(stylesPane).toBeHidden();
    await expect(page.getByTestId("zoom-control")).toBeVisible();
    await expect(page.getByTestId("page-indicator")).toBeVisible();

    // Escape exits and the pane comes back exactly as it was (AC-46/AC-47).
    await page.keyboard.press("Escape");
    await expect(page.locator(".is-focus-mode")).toHaveCount(0);
    await expect(stylesPane).toBeVisible();

    // Focus is a transient view state — never restored across a reload (AC-48).
    await page.keyboard.press(`${MOD}+Shift+F`);
    await expect(page.locator(".is-focus-mode")).toHaveCount(1);
    await page.reload();
    await expect(page.getByTestId("document-content")).toBeVisible();
    await expect(page.locator(".is-focus-mode")).toHaveCount(0);
  });

  test("zoom and view mode persist per document; Focus is never stored (AC-29..AC-31, AC-49..AC-51)", async ({
    page,
  }) => {
    await newDoc(page);
    const docId = page.url().split("/doc/")[1];

    // Zoom persists and restores (AC-29, AC-30).
    await setZoomInput(page, "175");
    await expect
      .poll(() => page.evaluate((k) => localStorage.getItem(`docu-edit:zoom:${k}`), docId))
      .toBe("175");
    await page.reload();
    await expect(page.getByTestId("zoom-percent")).toHaveValue("175%");

    // View mode persists as "web" / "print" (AC-49).
    await openLayoutMenu(page);
    await page.getByTestId("layout-web").click();
    await expect
      .poll(() => page.evaluate((k) => localStorage.getItem(`docu-edit:view-mode:${k}`), docId))
      .toBe("web");

    // Focus is never written to storage — view-mode stays "web" (AC-50).
    await page.getByTestId("document-content").click();
    await page.keyboard.press(`${MOD}+Shift+F`);
    await expect(page.locator(".is-focus-mode")).toHaveCount(1);
    expect(
      await page.evaluate((k) => localStorage.getItem(`docu-edit:view-mode:${k}`), docId),
    ).toBe("web");
    await page.keyboard.press("Escape");

    // A brand-new document opens at 100% Print layout (AC-31, AC-51).
    await newDoc(page);
    await expect(page.getByTestId("zoom-percent")).toHaveValue("100%");
    await openLayoutMenu(page);
    await expect(page.getByTestId("layout-print")).toHaveAttribute("aria-checked", "true");
  });

  test("print preview always renders at 100% Print layout, restoring chrome from Web layout (AC-52)", async ({
    page,
  }) => {
    await newDoc(page);
    await page.getByTestId("document-content").click();
    await page.keyboard.type("Export me.");

    // Web layout + a non-100% screen zoom: print preview must neutralise both.
    await openLayoutMenu(page);
    await page.getByTestId("layout-web").click();
    await setZoomInput(page, "150");

    await page.getByTestId("document-content").click();
    await page.keyboard.press(`${MOD}+p`);
    await expect(page.getByTestId("print-preview")).toBeVisible();

    const info = await page.evaluate(() => {
      const scaled =
        document.querySelector<HTMLElement>(".pages--scaled") ??
        document.querySelector<HTMLElement>('[data-testid="document-pages"]')!;
      const header = document.querySelector('[data-testid="page-header-band"]');
      return {
        transform: getComputedStyle(scaled).transform,
        headerDisplay: header ? getComputedStyle(header).display : "none",
      };
    });
    // Scale transform neutralised (AC-52) and the header band restored despite
    // Web layout hiding it on screen.
    expect(info.transform === "none" || info.transform === "matrix(1, 0, 0, 1, 0, 0)").toBeTruthy();
    expect(info.headerDisplay).not.toBe("none");
  });
});
