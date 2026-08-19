import { test, expect, type Page, type Locator } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-058 AC-48 — every control has a minimum 32×32px hit target at the
// default density. Prior rounds regressed here because the ParagraphToolbar
// cluster and the track-changes toggle sat at 28px tall with no hit-area
// expansion. This spec measures the *effective* hit rectangle exactly as QA
// did: the element's own box, unioned with any transparent ::before hit area
// (the zero-width-cost pattern FontSizeControl/ParagraphToolbar use to stay
// within the 1280px toolbar budget, AC-60). It fixes the viewport at 1280px so
// a regression that widens the strip would surface as an overflow too.

const MIN = 32;

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** The largest 32×32-capable hit rectangle a control offers: its border box,
 *  or a centred ::before hit area when the visible box is deliberately smaller
 *  than the target (so the toolbar stays flush at 1280px). */
async function hitTarget(locator: Locator): Promise<{ width: number; height: number }> {
  return locator.evaluate((el) => {
    const box = el.getBoundingClientRect();
    const before = getComputedStyle(el, "::before");
    const hasPseudo = before.content !== "none" && before.content !== "";
    const pseudoW = hasPseudo ? parseFloat(before.width) || 0 : 0;
    const pseudoH = hasPseudo ? parseFloat(before.height) || 0 : 0;
    return { width: Math.max(box.width, pseudoW), height: Math.max(box.height, pseudoH) };
  });
}

test("the paragraph toolbar controls each meet the 32×32 hit target at 1280px (AC-48)", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1280, height: 900 });
  await createDocument(page);

  const controls = [
    "align-left",
    "align-center",
    "align-right",
    "align-justify",
    "indent-decrease",
    "indent-increase",
    "list-bulleted",
    "list-bulleted-menu-button",
    "list-numbered",
    "list-numbered-menu-button",
    "paragraph-direction",
    "line-spacing-button",
    "paragraph-pagination-button",
  ];

  for (const id of controls) {
    const control = page.getByTestId(id);
    await expect(control).toBeVisible();
    const { width, height } = await hitTarget(control);
    expect(width, `${id} hit width`).toBeGreaterThanOrEqual(MIN);
    expect(height, `${id} hit height`).toBeGreaterThanOrEqual(MIN);
  }

  // The expanded hit areas must not have widened the strip past its budget.
  const overflow = await page
    .getByTestId("toolbar")
    .evaluate((el) => el.scrollWidth > el.clientWidth + 1);
  expect(overflow, "toolbar overflow at 1280px").toBe(false);
});

test("the track-changes toggle and the find bar buttons meet the 32×32 hit target (AC-48)", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1280, height: 900 });
  await createDocument(page);

  const track = page.getByTestId("toolbar-track-changes-toggle");
  await expect(track).toBeVisible();
  const trackHit = await hitTarget(track);
  expect(trackHit.width, "track-changes width").toBeGreaterThanOrEqual(MIN);
  expect(trackHit.height, "track-changes height").toBeGreaterThanOrEqual(MIN);

  await page.getByTestId("document-content").click();
  await page.keyboard.press(`${MOD}+f`);
  await expect(page.getByTestId("find-bar")).toBeVisible();
  // Reveal the replace row so its text buttons render.
  await page.getByTestId("find-expand").click();

  for (const id of ["find-expand", "find-prev", "find-next", "find-close", "find-match-case", "find-replace", "find-replace-all"]) {
    const control = page.getByTestId(id);
    await expect(control).toBeVisible();
    const { width, height } = await hitTarget(control);
    expect(width, `${id} hit width`).toBeGreaterThanOrEqual(MIN);
    expect(height, `${id} hit height`).toBeGreaterThanOrEqual(MIN);
  }
});
