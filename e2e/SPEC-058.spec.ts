import { test, expect, type Page } from "@playwright/test";

// SPEC-058 — hand-authored design system. This QA spec guards the criteria a
// user actually observes: the shell fits its viewport, popups opened by mouse
// dismiss on Escape (the shared UiMenuSurface contract), the dialog chrome is
// modal + focus-trapped + Escape-closable, every interactive control shows the
// 2px --focus-ring, and the sub-32px toolbar controls still meet the 32×32 hit
// target (the criterion that regressed in earlier rounds). The exhaustive
// per-control hit-target sweep lives in SPEC-058.hit-targets.spec.ts.

/** The computed colour --focus-ring resolves to on the current palette — the
 * ring assertions compare against the token, not a baked-in value, so they
 * hold under every design theme (theme-palettes e2e). */
function focusRing(page: Page): Promise<string> {
  return page.evaluate(() => {
    const probe = document.createElement("span");
    probe.style.color = getComputedStyle(document.documentElement)
      .getPropertyValue("--focus-ring")
      .trim();
    document.body.appendChild(probe);
    const value = getComputedStyle(probe).color;
    probe.remove();
    return value;
  });
}

async function newDoc(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

test("app chrome does not scroll horizontally at 1024px or 1280px (AC-59, AC-60)", async ({
  page,
}) => {
  await newDoc(page);
  for (const width of [1024, 1280]) {
    await page.setViewportSize({ width, height: 800 });
    // let the responsive layout settle
    await expect
      .poll(() =>
        page.evaluate(() => {
          const el = document.documentElement;
          return el.scrollWidth - el.clientWidth;
        }),
      )
      .toBeLessThanOrEqual(1);
  }
});

test("a mouse-opened popup dismisses on Escape via the shared surface (AC-39)", async ({
  page,
}) => {
  await newDoc(page);
  const menu = page.getByTestId("text-color-menu");

  // Opened by a real mouse click — focus stays on the trigger, so the popup's
  // dismissal must come from UiMenuSurface's document-level Escape handler.
  await page.getByTestId("text-color-button").click();
  await expect(menu).toBeVisible();

  await page.keyboard.press("Escape");
  await expect(menu).toBeHidden();
});

test("the menu bar's Insert popup also dismisses on Escape (AC-39)", async ({ page }) => {
  await newDoc(page);
  const insert = page.getByRole("menuitem", { name: /insert/i }).first();
  await insert.click();
  const popup = page.getByRole("menu").filter({ hasText: /table|image|shape/i }).first();
  await expect(popup).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(popup).toBeHidden();
});

test("the Page Setup dialog is modal, focus-trapped and Escape-closable (AC-40, AC-42, AC-43)", async ({
  page,
}) => {
  await newDoc(page);
  await page.getByRole("menuitem", { name: /file/i }).first().click();
  await page.getByRole("menuitem", { name: /page setup/i }).first().click();

  const dialog = page.getByRole("dialog");
  await expect(dialog).toBeVisible();
  await expect(dialog).toHaveAttribute("aria-modal", "true");
  await expect(dialog).toHaveAttribute("aria-labelledby", /.+/);

  // Focus is trapped inside the dialog.
  expect(
    await page.evaluate(() => {
      const d = document.querySelector('[role="dialog"]');
      return !!d && d.contains(document.activeElement);
    }),
  ).toBe(true);

  await page.keyboard.press("Escape");
  await expect(page.getByRole("dialog")).toHaveCount(0);
});

test("interactive controls show the 2px --focus-ring at a 2px offset (AC-45)", async ({
  page,
}) => {
  await newDoc(page);
  // AC-45 is about the *keyboard* focus ring (:focus-visible). Establish
  // keyboard modality with a Tab press so programmatic focus renders the ring,
  // exactly as it would for a user navigating by keyboard.
  await page.keyboard.press("Tab");

  // A primitive-based toolbar control.
  const align = page.getByTestId("align-left");
  await align.focus();
  await expect(align).toBeFocused();
  const primitiveRing = await align.evaluate((el) => {
    const cs = getComputedStyle(el);
    return { width: cs.outlineWidth, style: cs.outlineStyle, color: cs.outlineColor, offset: cs.outlineOffset };
  });
  expect(primitiveRing.width).toBe("2px");
  expect(primitiveRing.style).toBe("solid");
  expect(primitiveRing.color).toBe(await focusRing(page));
  expect(primitiveRing.offset).toBe("2px");

  // A hand-rolled status-bar control inherits the same global ring.
  const wc = page.getByTestId("word-count-button");
  await page.keyboard.press("Tab");
  await wc.focus();
  const handRolledRing = await wc.evaluate((el) => {
    const cs = getComputedStyle(el);
    return { width: cs.outlineWidth, color: cs.outlineColor, offset: cs.outlineOffset };
  });
  expect(handRolledRing.width).toBe("2px");
  expect(handRolledRing.color).toBe(await focusRing(page));
  expect(handRolledRing.offset).toBe("2px");
});

test("sub-32px toolbar controls meet the 32×32 hit target (AC-48)", async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 900 });
  await newDoc(page);

  // The controls whose visible box is deliberately narrower than 32px and rely
  // on a centred transparent ::before hit area — the ones that regressed before.
  const ids = [
    "indent-decrease",
    "indent-increase",
    "line-spacing-button",
    "paragraph-direction",
    "paragraph-pagination-button",
    "list-bulleted-menu-button",
    "list-numbered-menu-button",
    // SPEC-061 removed the duplicate toolbar track-changes toggle; the
    // surviving status-bar control is a plain text control, not a toolbar
    // hit-target.
  ];
  for (const id of ids) {
    const control = page.getByTestId(id);
    await expect(control).toBeVisible();
    const { width, height } = await control.evaluate((el) => {
      const box = el.getBoundingClientRect();
      const before = getComputedStyle(el, "::before");
      const hasPseudo = before.content !== "none" && before.content !== "";
      const pw = hasPseudo ? parseFloat(before.width) || 0 : 0;
      const ph = hasPseudo ? parseFloat(before.height) || 0 : 0;
      return { width: Math.max(box.width, pw), height: Math.max(box.height, ph) };
    });
    expect(width, `${id} hit width`).toBeGreaterThanOrEqual(32);
    expect(height, `${id} hit height`).toBeGreaterThanOrEqual(32);
  }

  // The expanded hit areas must not widen the strip past its 1280px budget.
  const overflow = await page
    .getByTestId("toolbar")
    .evaluate((el) => el.scrollWidth > el.clientWidth + 1);
  expect(overflow, "toolbar overflow at 1280px").toBe(false);
});
