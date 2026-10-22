import { test, expect, type Page } from "@playwright/test";
import { deflateSync } from "node:zlib";
import { MOD } from "./keys";

// SPEC-043 — the accessibility checker: "Tools -> Check accessibility"
// opens a results pane listing issues grouped by severity with a summary
// count and a clean state; selecting an issue navigates to and highlights
// its offending content; each issue explains why it matters and, where a
// mechanical fix exists (add alt text, mark header row, fix heading level),
// offers it; applying a fix re-runs the check and updates the list. The
// detection logic itself (contrast maths, colour-alone, link text, heading
// structure, …) is covered exhaustively by src/lib/accessibilityCheck.spec.ts —
// this file exercises the pane/menu/fix wiring end to end instead of
// re-testing detection rules already proven at the unit level.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const pane = (page: Page) => page.getByTestId("accessibility-pane");
const caretStatus = (page: Page) => page.getByTestId("caret-status");

async function openAccessibilityPane(page: Page): Promise<void> {
  await page.getByTestId("menu-title-tools").click();
  await page.getByTestId("tools-menu-accessibility").click();
  await expect(pane(page)).toBeVisible();
}

function issueCard(page: Page, kindLabel: string) {
  return pane(page).locator(".accessibility-issue", { hasText: kindLabel });
}

/** Types a heading paragraph via the SPEC-021 shortcut, then a body
 * paragraph — mirrors e2e/SPEC-037.spec.ts's own helper. */
async function typeHeadingThenBody(page: Page, headingText: string, level: number, bodyText: string): Promise<void> {
  await page.keyboard.type(headingText);
  await page.keyboard.press(`${MOD}+Alt+${level}`);
  await page.keyboard.press("Enter");
  await page.keyboard.press(`${MOD}+Alt+0`);
  await page.keyboard.type(bodyText);
}

// --- a tiny inline PNG, mirroring e2e/SPEC-024.spec.ts's own encoder so
// this file needs no fixture for a plain "insert an image" step. ----------
const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Buffer): Buffer {
  const typeBuf = Buffer.from(type, "ascii");
  const body = Buffer.concat([typeBuf, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  return Buffer.concat([length, body, crc]);
}

function buildPng(width: number, height: number, [r, g, b]: [number, number, number]): Buffer {
  const rowBytes = width * 3;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    const rowStart = y * (rowBytes + 1);
    raw[rowStart] = 0;
    for (let x = 0; x < width; x++) {
      const px = rowStart + 1 + x * 3;
      raw[px] = r;
      raw[px + 1] = g;
      raw[px + 2] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([
    signature,
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

const SMALL_PNG = buildPng(40, 20, [0x33, 0x77, 0xbb]);

async function insertImage(page: Page): Promise<void> {
  await page.getByTestId("insert-image-file-input").setInputFiles({
    name: "photo.png",
    mimeType: "image/png",
    buffer: SMALL_PNG,
  });
}

test("opens from Tools and shows a clean state on a fresh document (AC-1, AC-29, AC-30)", async ({ page }) => {
  await createDocument(page);
  await openAccessibilityPane(page);
  await expect(page.getByTestId("accessibility-summary")).toHaveText("No accessibility issues found");
  await expect(page.getByTestId("accessibility-clean")).toBeVisible();
  await page.getByTestId("accessibility-pane-close").click();
  await expect(pane(page)).not.toBeVisible();
});

test("groups detected issues by severity with a summary count and per-issue explanations (AC-2, AC-3, AC-26, AC-29)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Intro", 1, "Ordinary body text under the intro.");
  await page.keyboard.press("Enter");
  await typeHeadingThenBody(page, "Details", 3, "Ordinary body text under details.");
  await insertImage(page);

  await openAccessibilityPane(page);
  await expect(page.getByTestId("accessibility-summary")).toHaveText("2 issues found");

  const errorGroup = page.getByTestId("accessibility-group-error");
  await expect(errorGroup).toContainText("Error (1)");
  await expect(errorGroup).toContainText("Image missing alt text");

  const warningGroup = page.getByTestId("accessibility-group-warning");
  await expect(warningGroup).toContainText("Warning (1)");
  await expect(warningGroup).toContainText("Heading level skipped");

  const headingIssue = issueCard(page, "Heading level skipped");
  await expect(headingIssue.locator(".accessibility-issue-explanation")).not.toHaveText("");
});

test("selecting an issue scrolls to and selects (highlights) the offending content (AC-25)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Intro", 1, "Ordinary body text under the intro.");
  await page.keyboard.press("Enter");
  await typeHeadingThenBody(page, "Details", 3, "Ordinary body text under details.");

  await openAccessibilityPane(page);
  await issueCard(page, "Heading level skipped").click();

  await expect(caretStatus(page)).toContainText("paragraph 3");
  const selectionText = await page.evaluate(() => window.getSelection()?.toString() ?? "");
  expect(selectionText).toContain("Details");
});

test('offers "fix heading level" and applying it corrects the level and re-runs the check (AC-27, AC-28)', async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await typeHeadingThenBody(page, "Intro", 1, "Ordinary body text under the intro.");
  await page.keyboard.press("Enter");
  await typeHeadingThenBody(page, "Details", 3, "Ordinary body text under details.");

  await openAccessibilityPane(page);
  const headingIssue = issueCard(page, "Heading level skipped");
  await expect(headingIssue).toBeVisible();
  await headingIssue.getByTestId("accessibility-issue-fix").click();

  await expect(issueCard(page, "Heading level skipped")).toHaveCount(0);
  await content(page).getByText("Details", { exact: true }).click();
  await expect(page.getByTestId("style-picker-current")).toHaveText("Heading 2");
});

test('offers "add alt text" opening the image options dialog, and applying it updates the list (AC-4, AC-27, AC-28)', async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Some text before the image.");
  await insertImage(page);

  await openAccessibilityPane(page);
  const altIssue = issueCard(page, "Image missing alt text");
  await expect(altIssue).toBeVisible();
  await altIssue.getByTestId("accessibility-issue-fix").click();

  await expect(page.getByTestId("image-options-dialog")).toBeVisible();
  await page.getByTestId("image-options-alt-text").fill("A small blue square");
  await page.getByTestId("image-options-apply").click();
  await expect(page.getByTestId("image-options-dialog")).not.toBeVisible();

  // The rescan drops the fixed missing-alt-text issue; this fixture has no
  // headings at all, so the unrelated "no heading structure" Tip is expected
  // to remain — this test is about the alt-text fix specifically.
  await expect(issueCard(page, "Image missing alt text")).toHaveCount(0);
});

test('offers "mark header row" for a table with no header row, and applying it removes the issue (AC-5, AC-27, AC-28)', async ({
  page,
}) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-2-2").click();
  await expect(content(page).locator("table.doc-table")).toHaveCount(1);

  await openAccessibilityPane(page);
  const headerIssue = issueCard(page, "Table has no header row");
  await expect(headerIssue).toBeVisible();
  await headerIssue.getByTestId("accessibility-issue-fix").click();

  await expect(issueCard(page, "Table has no header row")).toHaveCount(0);
});

// --- AC-33: the editor UI itself passes an automated accessibility scan ---
//
// No axe-core (or an axe-core Playwright binding) is available — package.json
// is protected from feature-driven dependency additions (see CLAUDE.md) — so
// this is a hand-rolled equivalent covering axe's own "serious"/"critical"
// checks that apply to a custom contenteditable canvas app: every
// interactive control has an accessible name, no duplicate ids, and a main
// landmark exists. Scoped to the app chrome (menu bar, toolbars, the open
// accessibility pane) rather than the user's own document content — AC-33
// is about the editor UI, not an audit of whatever the user has typed.
test("the main editing screen has no accessible-name, duplicate-id or landmark violations (AC-33)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Sample document text for the accessibility scan.");
  await openAccessibilityPane(page);

  const violations = await page.evaluate(() => {
    const problems: string[] = [];
    const seenIds = new Map<string, number>();
    for (const el of Array.from(document.querySelectorAll("[id]"))) {
      const id = el.id;
      seenIds.set(id, (seenIds.get(id) ?? 0) + 1);
    }
    for (const [id, count] of seenIds) {
      if (count > 1) {
        problems.push(`duplicate id "${id}" (${count} elements)`);
      }
    }

    const documentContent = document.querySelector('[data-testid="document-content"]');
    const controls = document.querySelectorAll(
      'button, a[href], input, select, textarea, [role="button"], [role="menuitem"], [role="menuitemradio"], [role="menuitemcheckbox"]',
    );
    for (const el of Array.from(controls)) {
      if (documentContent?.contains(el)) {
        continue;
      }
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 && rect.height === 0) {
        continue;
      }
      const accessibleName =
        el.getAttribute("aria-label") ??
        (el.getAttribute("aria-labelledby")
          ? document.getElementById(el.getAttribute("aria-labelledby") ?? "")?.textContent
          : null) ??
        el.textContent?.trim() ??
        el.getAttribute("title") ??
        (el instanceof HTMLInputElement && el.labels && el.labels.length > 0 ? el.labels[0].textContent : null);
      if (!accessibleName || accessibleName.trim().length === 0) {
        problems.push(`no accessible name: <${el.tagName.toLowerCase()} data-testid="${el.getAttribute("data-testid") ?? ""}">`);
      }
    }

    if (!document.querySelector('main, [role="main"]')) {
      problems.push("no main landmark");
    }

    for (const img of Array.from(document.querySelectorAll("img"))) {
      if (documentContent?.contains(img)) {
        continue;
      }
      if (!img.hasAttribute("alt")) {
        problems.push(`<img> with no alt attribute: ${img.outerHTML.slice(0, 80)}`);
      }
    }

    return problems;
  });

  expect(violations).toEqual([]);
});
