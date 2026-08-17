import { test, expect, type Download, type Locator, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { readFileSync } from "node:fs";
import { MOD } from "./keys";

// SPEC-028 — page and section breaks: content flows across rendered pages
// with a visible gap, the status bar tracks the caret's page, manual
// page/section breaks insert (Ctrl/Cmd+Enter and the Insert-break menu —
// which must be genuinely visible, not clipped by the toolbar) and delete,
// formatting marks show ¶/spaces/tabs/break markers, sections carry their
// own page setup via Apply to, the paragraph pagination rules hold, breaks
// and sections round-trip through .docx/.odt, and pagination stays
// responsive on large documents.


const content = (page: Page) => page.getByTestId("document-content");
const pages = (page: Page) => page.getByTestId("document-page");

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function download(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("download-menu").click();
  const file: Download = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId(`download-${format}`).click(),
  ]).then(([d]) => d);
  const path = await file.path();
  expect(path).toBeTruthy();
  return path!;
}

/** Minimal central-directory walk over a downloaded archive — same helper
 * as e2e/SPEC-027.spec.ts. */
function unzipParts(buffer: Buffer): Map<string, Buffer> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd).toBeGreaterThanOrEqual(0);
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const parts = new Map<string, Buffer>();
  for (let i = 0; i < totalEntries; i++) {
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.toString("utf8", offset + 46, offset + 46 + nameLength);
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const data = buffer.subarray(dataStart, dataStart + compressedSize);
    parts.set(name, method === 8 ? inflateRawSync(data) : Buffer.from(data));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return parts;
}

/** Waits for pagination to settle: the canvas exposes a positive signal
 * (data-pagination-state="clean" once every page uses real measurements),
 * so this never races a fast re-layout the way waiting for a busy
 * indicator would. */
async function waitForPagination(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-pages")).toHaveAttribute(
    "data-pagination-state",
    "clean",
    { timeout },
  );
}

/** Seeds many paragraphs in one operation through the app's own plain-text
 * paste path (one paragraph per line). */
async function pasteText(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.evaluate((payload) => {
    const root = document.querySelector('[data-testid="document-content"]')!;
    const data = new DataTransfer();
    data.setData("text/plain", payload);
    root.dispatchEvent(
      new ClipboardEvent("paste", { bubbles: true, cancelable: true, clipboardData: data }),
    );
  }, text);
}

function paragraphLines(count: number, prefix = "Paragraph"): string {
  return Array.from({ length: count }, (_, i) => `${prefix} ${i} lorem ipsum dolor sit amet`).join(
    "\n",
  );
}

/** Counts an element's rendered visual lines (client rects grouped into
 * horizontal bands). */
async function countLines(locator: Locator): Promise<number> {
  return locator.evaluate((el) => {
    const range = document.createRange();
    range.selectNodeContents(el);
    const rects = Array.from(range.getClientRects())
      .filter((r) => r.height > 0)
      .sort((a, b) => a.top - b.top);
    let lines = 0;
    let bottom = Number.NEGATIVE_INFINITY;
    for (const r of rects) {
      if (r.top >= bottom - 1) {
        lines += 1;
        bottom = r.bottom;
      } else {
        bottom = Math.max(bottom, r.bottom);
      }
    }
    return lines;
  });
}

test("content flows across pages with a visible gap between them (AC-1, AC-2)", async ({ page }) => {
  await createDocument(page);
  await pasteText(page, paragraphLines(150));
  await waitForPagination(page);
  const count = await pages(page).count();
  expect(count).toBeGreaterThanOrEqual(2);
  const first = (await pages(page).nth(0).boundingBox())!;
  const second = (await pages(page).nth(1).boundingBox())!;
  const gap = second.y - (first.y + first.height);
  // AC-2: a 12px gap between consecutive page containers.
  expect(gap).toBeGreaterThanOrEqual(11);
  expect(gap).toBeLessThanOrEqual(13);
  await expect(pages(page).nth(0)).toContainText("Paragraph 0 ");
});

test("the status bar shows Page N of M following the caret (AC-3)", async ({ page }) => {
  await createDocument(page);
  await pasteText(page, paragraphLines(150));
  await waitForPagination(page);
  const total = await pages(page).count();
  await content(page).click();
  await page.keyboard.press(`${MOD}+Home`);
  await expect(page.getByTestId("page-indicator")).toHaveText(`Page 1 of ${total}`);
  await page.keyboard.press(`${MOD}+End`);
  await expect(page.getByTestId("page-indicator")).toHaveText(`Page ${total} of ${total}`);
});

test("Ctrl/Cmd+Enter inserts a manual page break at the caret (AC-5, AC-7)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("HelloWorld");
  for (let i = 0; i < 5; i += 1) {
    await page.keyboard.press("ArrowLeft");
  }
  await page.keyboard.press(`${MOD}+Enter`);
  await waitForPagination(page);
  await expect(page.locator('[data-break-kind="page"]')).toHaveCount(1);
  await expect(pages(page)).toHaveCount(2);
  await expect(pages(page).nth(0)).toContainText("Hello");
  await expect(pages(page).nth(0)).not.toContainText("World");
  await expect(pages(page).nth(1)).toContainText("World");
});

/** Asserts the caret's own paragraph is inside the scrolling canvas. It has
 * to be the paragraph element, not `startContainer.parentElement`: a caret
 * in an empty paragraph reports the paragraph itself as its container, so
 * that shortcut measures the whole page box and passes wherever the page
 * happens to be. */
async function expectCaretOnScreen(page: Page): Promise<void> {
  const box = await page.evaluate(() => {
    const node = getSelection()!.getRangeAt(0).startContainer;
    const paragraph = (node instanceof Element ? node : node.parentElement)!.closest(
      "[data-section][data-block]",
    )!;
    const rect = paragraph.getBoundingClientRect();
    const view = document.querySelector("main")!.getBoundingClientRect();
    return { top: rect.top, bottom: rect.bottom, viewTop: view.top, viewBottom: view.bottom };
  });
  expect(box.top).toBeGreaterThanOrEqual(box.viewTop);
  expect(box.bottom).toBeLessThanOrEqual(box.viewBottom);
}

test("inserting a page break scrolls the new page into view (AC-5)", async ({ page }) => {
  await createDocument(page);
  await pasteText(page, paragraphLines(150));
  await waitForPagination(page);
  await content(page).click();
  await page.keyboard.press(`${MOD}+Home`);
  await page.keyboard.press(`${MOD}+Enter`);
  await waitForPagination(page);
  // The caret is now at the top of page 2 — a page the user must be able to
  // see, or the text they type next lands somewhere off-screen.
  await expect(page.getByTestId("page-indicator")).toContainText("Page 2 of");
  await expectCaretOnScreen(page);
});

test("typing that runs past the fold keeps the caret on screen (AC-1)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  // Enough lines to fill the visible canvas and then some: an edit must
  // never leave the user typing somewhere they cannot see.
  for (let i = 0; i < 40; i += 1) {
    await page.keyboard.type(`Line ${i} of text long enough to occupy a whole line of the page.`);
    await page.keyboard.press("Enter");
  }
  await waitForPagination(page);
  await expectCaretOnScreen(page);
});

test("the page a break ends keeps a paragraph the caret can reach (AC-5)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  // A break with nothing before it: without a paragraph of its own, page 1
  // would render as an empty box no click or arrow key could enter, and
  // every subsequent keystroke would land on page 2.
  await page.keyboard.press(`${MOD}+Enter`);
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(2);
  await expect(pages(page).nth(0).locator("[data-section][data-block]").first()).toHaveCount(1);

  await page.evaluate(() => {
    document.querySelector("main")!.scrollTop = 0;
  });
  const box = (await pages(page).nth(0).boundingBox())!;
  await page.mouse.click(box.x + box.width / 2, box.y + Math.min(box.height / 2, 300));
  await expect(page.getByTestId("page-indicator")).toHaveText("Page 1 of 2");
  await page.keyboard.type("First page");
  await waitForPagination(page);
  await expect(pages(page).nth(0)).toContainText("First page");
  await expect(pages(page).nth(1)).not.toContainText("First page");

  // And the caret still crosses the break in both directions.
  await page.keyboard.press("ArrowDown");
  await expect(page.getByTestId("page-indicator")).toHaveText("Page 2 of 2");
  await page.keyboard.press("ArrowUp");
  await expect(page.getByTestId("page-indicator")).toHaveText("Page 1 of 2");
});

test("the Insert menu is genuinely visible and offers all four breaks (AC-8, AC-9)", async ({
  page,
}) => {
  await createDocument(page);
  // SPEC-049: the breaks moved from an ad-hoc toolbar dropdown into the
  // Insert menu of the application menu bar. The menu bar is not an
  // overflow-clipped container like the toolbar strip was, so the old
  // clip-escape geometry assertions no longer apply — what matters is that
  // every item is visible, inside the viewport, hit-testable and enabled.
  await page.getByTestId("menu-title-insert").click();
  await expect(page.getByTestId("menu-insert")).toBeVisible();

  const items = [
    { id: "break-page", label: "Page break" },
    { id: "break-column", label: "Column break" },
    { id: "break-section-next", label: "Section break (next page)" },
    { id: "break-section-continuous", label: "Section break (continuous)" },
  ];
  const viewport = page.viewportSize()!;
  for (const item of items) {
    const locator = page.getByTestId(item.id);
    await expect(locator).toBeVisible();
    await expect(locator).toContainText(item.label);
    const box = (await locator.boundingBox())!;
    expect(box.height).toBeGreaterThan(10);
    expect(box.y).toBeGreaterThanOrEqual(0);
    expect(box.y + box.height).toBeLessThanOrEqual(viewport.height);
    const hit = await page.evaluate(
      ({ x, y, id }) => {
        const el = document.elementFromPoint(x, y);
        return !!el?.closest(`[data-testid="${id}"]`);
      },
      { x: box.x + box.width / 2, y: box.y + box.height / 2, id: item.id },
    );
    expect(hit).toBe(true);
  }
  // All four items are enabled — the column break inserts a real node even
  // though its layout effect waits for columns (AC-10, AC-12; SPEC-030).
  for (const item of items) {
    await expect(page.getByTestId(item.id)).toBeEnabled();
  }

  // The page-break item actually works from the menu.
  await content(page).click();
  await page.keyboard.type("AlphaOmega");
  for (let i = 0; i < 5; i += 1) {
    await page.keyboard.press("ArrowLeft");
  }
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-page").click();
  await waitForPagination(page);
  await expect(page.locator('[data-break-kind="page"]')).toHaveCount(1);
  await expect(pages(page)).toHaveCount(2);
});

test("Column Break inserts a stored node with no visible layout effect (AC-10, AC-12)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("AlphaOmega");
  for (let i = 0; i < 5; i += 1) {
    await page.keyboard.press("ArrowLeft");
  }
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-column").click();
  await waitForPagination(page);
  // The break node is in the document model and rendered flow...
  const marker = page.locator('[data-break-kind="column"]');
  await expect(marker).toHaveCount(1);
  // ...but has no layout effect: still one page, nothing visible while
  // formatting marks are off (AC-12).
  await expect(pages(page)).toHaveCount(1);
  await expect(marker).not.toBeVisible();
  // With formatting marks on it shows a labelled marker (AC-19).
  await page.keyboard.press(`${MOD}+Shift+P`);
  await expect(marker).toBeVisible();
  const label = await marker.evaluate((el) => getComputedStyle(el, "::before").content);
  expect(label).toContain("Column Break");
});

test("a section break (next page) forces a page; continuous does not (AC-13, AC-15)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("OneTwo");
  for (let i = 0; i < 3; i += 1) {
    await page.keyboard.press("ArrowLeft");
  }
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-section-continuous").click();
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(1);

  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-section-next").click();
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(2);
});

test("deleting a manual break rejoins the content flow (AC-17, AC-18)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("HelloWorld");
  for (let i = 0; i < 5; i += 1) {
    await page.keyboard.press("ArrowLeft");
  }
  await page.keyboard.press(`${MOD}+Enter`);
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(2);
  // The caret sits at the start of the content after the break; one
  // Backspace removes the break, a second rejoins the paragraphs.
  await page.keyboard.press("Backspace");
  await waitForPagination(page);
  await expect(page.locator('[data-break-kind="page"]')).toHaveCount(0);
  await expect(pages(page)).toHaveCount(1);
  await page.keyboard.press("Backspace");
  await expect(content(page)).toHaveText("HelloWorld");
});

test("Ctrl/Cmd+Shift+P shows paragraph marks, spaces, tabs and break markers (AC-19..23)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("a b");
  // A literal tab through the editing engine (Tab itself means indent).
  await page.evaluate(() => {
    const root = document.querySelector('[data-testid="document-content"]')!;
    root.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: "\tc",
      }),
    );
  });
  await page.keyboard.press(`${MOD}+Enter`);
  await waitForPagination(page);

  // Off by default: no space marks, and the break element renders invisible.
  await expect(page.locator(".mark-space")).toHaveCount(0);
  await expect(page.locator('[data-break-kind="page"]').first()).not.toBeVisible();

  await page.keyboard.press(`${MOD}+Shift+P`);
  await expect(page.locator(".mark-space").first()).toBeVisible();
  // SPEC-054: a tab now renders as a sized .doc-tab box; its → formatting mark
  // is a pseudo-element shown only under .marks-on (AC-33), replacing the old
  // .mark-tab span this test looked for before tabs advanced to stops.
  await expect(page.locator(".doc-tab")).toHaveCount(1);
  const tabMark = await page
    .locator(".doc-tab")
    .first()
    .evaluate((el) => getComputedStyle(el, "::before").content);
  expect(tabMark).toContain("→");
  // The pilcrow is a pseudo-element over byte-identical text.
  const pilcrow = await page
    .locator('[data-section="0"][data-block="0"]')
    .first()
    .evaluate((el) => getComputedStyle(el, "::after").content);
  expect(pilcrow).toContain("¶");
  const breakMarker = page.locator('[data-break-kind="page"]').first();
  await expect(breakMarker).toBeVisible();
  const label = await breakMarker.evaluate((el) => getComputedStyle(el, "::before").content);
  expect(label).toContain("Page Break");
  // The document text itself carries no marker characters.
  const text = await content(page).evaluate((el) => el.textContent);
  expect(text).not.toContain("¶");
  expect(text).not.toContain("Page Break");

  // Section markers are distinct, labelled marks-only elements.
  await page.keyboard.press(`${MOD}+End`);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-section-continuous").click();
  const sectionMarker = page.locator('[data-section-marker="continuous"]');
  await expect(sectionMarker).toHaveCount(1);
  const sectionLabel = await sectionMarker.evaluate(
    (el) => getComputedStyle(el, "::before").content,
  );
  expect(sectionLabel).toContain("Section Break (Continuous)");

  // Toggling off hides everything again.
  await page.keyboard.press(`${MOD}+Shift+P`);
  await expect(page.locator(".mark-space")).toHaveCount(0);
  await expect(page.locator('[data-section-marker="continuous"]')).toHaveCount(0);
  await expect(page.locator('[data-break-kind="page"]').first()).not.toBeVisible();
});

test("a section holds its own page size, orientation and margins via Apply to (AC-24..29)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("First section");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-section-next").click();
  await page.keyboard.type("Second section");
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(2);

  // The Apply to control offers exactly the two scopes (AC-27).
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("page-setup-button").click();
  const applyTo = page.getByTestId("page-setup-apply-to");
  await expect(applyTo).toBeVisible();
  const options = await applyTo.locator("option").allTextContents();
  expect(options.map((option) => option.trim())).toEqual(["this section", "whole document"]);

  // A5 landscape + tighter top margin, applied to the caret's section only.
  await applyTo.selectOption("thisSection");
  await page.getByTestId("page-setup-paper-size").selectOption("A5");
  await page.getByTestId("page-setup-orientation").selectOption("landscape");
  await page.getByTestId("page-setup-margin-top").fill("1");
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).not.toBeVisible();
  await waitForPagination(page);

  const styleOf = async (index: number, prop: string) =>
    Number.parseFloat(
      await pages(page)
        .nth(index)
        .evaluate((el, p) => (el as HTMLElement).style.getPropertyValue(p), prop),
    );
  // Page 1 keeps A4 portrait and its 2.54cm margins.
  expect(await styleOf(0, "width")).toBeCloseTo(21, 1);
  expect(await styleOf(0, "height")).toBeCloseTo(29.7, 1);
  // Page 2 (the new section) is A5 landscape with the 1cm top margin.
  expect(await styleOf(1, "width")).toBeCloseTo(21, 1);
  expect(await styleOf(1, "height")).toBeCloseTo(14.8, 1);
  const padTop = await pages(page)
    .nth(1)
    .evaluate((el) => {
      const contentBox = el.querySelector(".page-content") as HTMLElement;
      return contentBox.style.getPropertyValue("padding-top");
    });
  expect(Number.parseFloat(padTop)).toBeCloseTo(1, 2);

  // Whole document: both sections follow the new size (AC-29).
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("page-setup-button").click();
  await page.getByTestId("page-setup-apply-to").selectOption("wholeDocument");
  await page.getByTestId("page-setup-paper-size").selectOption("Letter");
  await page.getByTestId("page-setup-orientation").selectOption("portrait");
  await page.getByTestId("page-setup-apply").click();
  await waitForPagination(page);
  expect(await styleOf(0, "height")).toBeCloseTo(27.94, 1);
  expect(await styleOf(1, "height")).toBeCloseTo(27.94, 1);
});

test("paragraph pagination options are settable; widow/orphan is on by default (AC-30, AC-32, AC-34, AC-36, AC-40)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Some paragraph");
  await page.getByTestId("paragraph-pagination-button").click();
  const menu = page.getByTestId("paragraph-pagination-menu");
  await expect(menu).toBeVisible();
  await expect(page.getByTestId("pagination-keepNext")).toHaveAttribute("aria-checked", "false");
  await expect(page.getByTestId("pagination-keepLines")).toHaveAttribute("aria-checked", "false");
  await expect(page.getByTestId("pagination-pageBreakBefore")).toHaveAttribute(
    "aria-checked",
    "false",
  );
  // AC-36: widow/orphan control defaults ON (AC-40: toggleable off below).
  await expect(page.getByTestId("pagination-widowControl")).toHaveAttribute("aria-checked", "true");

  await page.getByTestId("pagination-keepNext").click();
  await expect(page.getByTestId("pagination-keepNext")).toHaveAttribute("aria-checked", "true");
  await page.getByTestId("pagination-widowControl").click();
  await expect(page.getByTestId("pagination-widowControl")).toHaveAttribute(
    "aria-checked",
    "false",
  );
});

test("page break before forces its paragraph onto a new page (AC-34, AC-35)", async ({ page }) => {
  await createDocument(page);
  await pasteText(page, "Intro\nChapter start");
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(1);
  await page.getByText("Chapter start").click();
  await page.getByTestId("paragraph-pagination-button").click();
  await page.getByTestId("pagination-pageBreakBefore").click();
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(2);
  await expect(pages(page).nth(1)).toContainText("Chapter start");
  await expect(pages(page).nth(0)).not.toContainText("Chapter start");
});

test("a long paragraph splits across pages (2+ lines each side) and keep-lines-together moves it whole (AC-31, AC-33, AC-37..39)", async ({
  page,
}) => {
  await createDocument(page);
  const filler = paragraphLines(40, "Filler");
  const longText = `SPLITSTART ${"word ".repeat(280)}SPLITEND`;
  await pasteText(page, `${filler}\n${longText}`);
  await waitForPagination(page);
  await expect(pages(page).nth(1)).toBeVisible();

  // The long paragraph straddles the page boundary: two slices of one
  // block, each rendering at least 2 lines (widow/orphan, AC-37..39).
  const slices = page.locator('[data-section="0"][data-block="40"]');
  await expect(slices).toHaveCount(2);
  expect(await countLines(slices.nth(0))).toBeGreaterThanOrEqual(2);
  expect(await countLines(slices.nth(1))).toBeGreaterThanOrEqual(2);
  await expect(pages(page).nth(0)).toContainText("SPLITSTART");
  await expect(pages(page).nth(1)).toContainText("SPLITEND");

  // Keep lines together: the paragraph moves to the next page whole
  // (AC-33).
  await page.getByText("SPLITSTART").click();
  await page.getByTestId("paragraph-pagination-button").click();
  await page.getByTestId("pagination-keepLines").click();
  await waitForPagination(page);
  await expect(slices).toHaveCount(1);
  await expect(pages(page).nth(0)).not.toContainText("SPLITSTART");
  await expect(pages(page).nth(1)).toContainText("SPLITSTART");

  // Keep with next (AC-31): the last paragraph of page 1 follows
  // its successor onto page 2 once keep-with-next is set on it.
  await page.keyboard.press("Escape");
  const lastOnPageOne = pages(page)
    .nth(0)
    .locator("[data-section][data-block]")
    .last();
  const lastText = (await lastOnPageOne.textContent())!;
  await lastOnPageOne.click();
  await page.getByTestId("paragraph-pagination-button").click();
  await page.getByTestId("pagination-keepNext").click();
  await waitForPagination(page);
  await expect(pages(page).nth(1)).toContainText(lastText.slice(0, 20));
  await expect(pages(page).nth(0)).not.toContainText(lastText.slice(0, 20));
});

test("manual breaks and section properties round-trip through .docx (AC-11, AC-41, AC-43)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Alpha");
  await page.keyboard.press(`${MOD}+Enter`);
  await page.keyboard.type("Beta");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-section-continuous").click();
  await page.keyboard.type("Gamma");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-column").click();
  await page.keyboard.type("Delta");
  await waitForPagination(page);

  const path = await download(page, "docx");
  const documentXml = unzipParts(readFileSync(path)).get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain('<w:br w:type="page"/>');
  expect(documentXml).toContain('<w:br w:type="column"/>');
  expect(documentXml).toContain("<w:sectPr");
  expect(documentXml).toContain('<w:type w:val="continuous"/>');

  await openFixture(page, path);
  await waitForPagination(page);
  await expect(page.locator('[data-break-kind="page"]')).toHaveCount(1);
  // The column break survives the cycle without affecting the page count.
  await expect(page.locator('[data-break-kind="column"]')).toHaveCount(1);
  await expect(pages(page)).toHaveCount(2);
  await expect(content(page)).toContainText("Gamma");
  await expect(content(page)).toContainText("Delta");
});

test("breaks and section properties round-trip through .odt (AC-11, AC-42, AC-44)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Alpha");
  await page.keyboard.press(`${MOD}+Enter`);
  await page.keyboard.type("Beta");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-section-next").click();
  await page.keyboard.type("Gamma");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-column").click();
  await page.keyboard.type("Delta");
  await waitForPagination(page);
  // Give the new section its own page size so the master page carries
  // something distinctive.
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("page-setup-button").click();
  await page.getByTestId("page-setup-apply-to").selectOption("thisSection");
  await page.getByTestId("page-setup-paper-size").selectOption("A5");
  await page.getByTestId("page-setup-apply").click();
  await waitForPagination(page);

  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const contentXml = parts.get("content.xml")!.toString("utf8");
  const stylesXml = parts.get("styles.xml")!.toString("utf8");
  expect(contentXml).toContain('fo:break-before="page"');
  expect(contentXml).toContain('fo:break-before="column"');
  expect(contentXml).toContain("style:master-page-name=");
  expect(stylesXml).toContain("SectionMaster");

  await openFixture(page, path);
  await waitForPagination(page);
  await expect(page.locator('[data-break-kind="page"]')).toHaveCount(1);
  await expect(page.locator('[data-break-kind="column"]')).toHaveCount(1);
  const count = await pages(page).count();
  expect(count).toBeGreaterThanOrEqual(3);
  const lastHeight = await pages(page)
    .nth(count - 1)
    .evaluate((el) => (el as HTMLElement).style.getPropertyValue("height"));
  expect(Number.parseFloat(lastHeight)).toBeCloseTo(21, 1);
});

test("a new empty document renders one page reading Page 1 of 1 (AC-55, AC-56)", async ({
  page,
}) => {
  await createDocument(page);
  await waitForPagination(page);
  await expect(pages(page)).toHaveCount(1);
  await content(page).click();
  await expect(page.getByTestId("page-indicator")).toHaveText("Page 1 of 1");
});

test("a large document paginates without freezing and mid-document edits keep earlier pages (AC-49..54)", async ({
  page,
}) => {
  test.setTimeout(120_000);
  await createDocument(page);
  await pasteText(page, paragraphLines(3000));
  // The UI stays responsive while pagination settles: typing lands.
  await content(page).click();
  await page.keyboard.press(`${MOD}+Home`);
  await page.keyboard.type("LIVE");
  await expect(content(page)).toContainText("LIVE");
  await waitForPagination(page, 90_000);
  const total = await pages(page).count();
  expect(total).toBeGreaterThan(40);
  await expect(page.getByTestId("page-indicator")).toContainText(`of ${total}`);

  // Far pages are virtualized (AC-49/50's mechanism): most page boxes are
  // empty geometry, so the DOM holds a small fraction of the document's
  // 3000 blocks — mounting them all is what used to blow the settle budget.
  const virtualCount = await page.locator("[data-virtual]").count();
  expect(virtualCount).toBeGreaterThan(total / 2);
  const domBlocks = await page
    .locator('[data-testid="document-page"] [data-section][data-block]')
    .count();
  expect(domBlocks).toBeLessThan(1000);

  // Tag page 1's first paragraph element, edit at the end, and require the
  // exact same connected DOM node afterwards — earlier pages were not
  // re-rendered (AC-52, AC-53).
  await page.evaluate(() => {
    const first = document
      .querySelector('[data-testid="document-page"]')!
      .querySelector("[data-section][data-block]") as HTMLElement & { dataset: DOMStringMap };
    (window as unknown as { __probe?: Element }).__probe = first;
    first.dataset.probe = "1";
  });
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.type("tail edit");
  await waitForPagination(page, 90_000);
  const intact = await page.evaluate(() => {
    const probe = (window as unknown as { __probe?: Element }).__probe;
    return (
      !!probe &&
      probe.isConnected &&
      probe === document.querySelector('[data-testid="document-page"] [data-probe="1"]')
    );
  });
  expect(intact).toBe(true);
  await expect(content(page)).toContainText("tail edit");

  // Scrolling to the middle realizes those pages: what the user can see is
  // always real content, never an empty virtualized box.
  await page.evaluate(() => {
    const main = document.querySelector("main")!;
    main.scrollTop = main.scrollHeight / 2;
  });
  await expect
    .poll(() =>
      page.evaluate(() => {
        const main = document.querySelector("main")!;
        const view = main.getBoundingClientRect();
        const boxes = Array.from(document.querySelectorAll('[data-testid="document-page"]'));
        const visible = boxes.filter((box) => {
          const rect = box.getBoundingClientRect();
          return rect.bottom > view.top && rect.top < view.bottom;
        });
        return (
          visible.length > 0 &&
          visible.every(
            (box) => !box.hasAttribute("data-virtual") && (box.textContent ?? "").length > 0,
          )
        );
      }),
    )
    .toBe(true);
});
