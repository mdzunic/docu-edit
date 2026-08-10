import { test, expect, type Download, type Page } from "@playwright/test";
import { MOD } from "./keys";
import { inflateRawSync } from "node:zlib";
import { readFileSync } from "node:fs";

// SPEC-032 — multilevel lists and numbering control: the gallery presets,
// the Define dialog's per-level format, heading-style linking, the context
// menu's restart/continue/set-value, custom tab stops (ruler + dialog) and
// the .docx/.odt round trips. Markers are asserted through
// data-list-marker, the attribute the canvas paints from.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const paragraphAt = (page: Page, index: number) =>
  content(page).locator(".doc-paragraph").nth(index);

async function typeParagraphs(page: Page, texts: string[]): Promise<void> {
  await content(page).click();
  for (let i = 0; i < texts.length; i++) {
    if (i > 0) {
      await page.keyboard.press("Enter");
    }
    await page.keyboard.type(texts[i]);
  }
}

async function applyPreset(page: Page, key: string): Promise<void> {
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId(`multilevel-preset-${key}`).click();
}

async function applyStyle(page: Page, styleId: string): Promise<void> {
  await page.getByTestId("style-picker-button").click();
  await page.getByTestId(`style-option-${styleId}`).click();
}

/** Downloads the document in the given format and returns the file path. */
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

async function openFile(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** Minimal central-directory walk over a downloaded archive — the same
 * helper as e2e/SPEC-018.spec.ts. */
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

/** Builds the AC-28 fixture in the editor: list A (3 items), list B
 * (2 items) and list C (2 items with the second restarted), each separated
 * by a plain paragraph. */
async function buildThreeListDocument(page: Page): Promise<void> {
  await createDocument(page);
  await typeParagraphs(page, ["a1", "a2", "a3", "gap1", "b1", "b2", "gap2", "c1", "c2"]);
  // List A: select a1..a3 by clicking a1 then shift-clicking a3.
  await paragraphAt(page, 0).click();
  await paragraphAt(page, 2).click({ modifiers: ["Shift"] });
  await applyPreset(page, "decimal");
  await paragraphAt(page, 4).click();
  await paragraphAt(page, 5).click({ modifiers: ["Shift"] });
  await applyPreset(page, "decimal");
  await paragraphAt(page, 7).click();
  await paragraphAt(page, 8).click({ modifiers: ["Shift"] });
  await applyPreset(page, "decimal");
  // Restart c2 via the context menu.
  await paragraphAt(page, 8).click({ button: "right" });
  await page.getByTestId("context-menu-restart-numbering").click();
  await expect(paragraphAt(page, 8)).toHaveAttribute("data-list-marker", "1.");
}

const THREE_LIST_MARKERS = ["1.", "2.", "3.", "1.", "2.", "1.", "1."];

async function expectThreeListMarkers(page: Page): Promise<void> {
  const indices = [0, 1, 2, 4, 5, 7, 8];
  for (let i = 0; i < indices.length; i++) {
    await expect(paragraphAt(page, indices[i])).toHaveAttribute(
      "data-list-marker",
      THREE_LIST_MARKERS[i],
    );
  }
  await expect(paragraphAt(page, 3)).not.toHaveAttribute("data-list-marker");
  await expect(paragraphAt(page, 6)).not.toHaveAttribute("data-list-marker");
}

// --- gallery presets (AC-1..4) ----------------------------------------------

test("the gallery offers the four presets and the decimal one numbers a selection (AC-1..4)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["one", "two", "three"]);
  await page.getByTestId("menu-title-format").click();
  await expect(page.getByTestId("multilevel-preset-decimal")).toBeVisible();
  await expect(page.getByTestId("multilevel-preset-outline")).toBeVisible();
  await expect(page.getByTestId("multilevel-preset-legal")).toBeVisible();
  await expect(page.getByTestId("multilevel-preset-bulletMixed")).toBeVisible();
  // Close the gallery again (the toggle), then number the whole selection.
  await page.getByTestId("menu-title-format").click();
  await expect(page.getByTestId("menu-format")).toBeHidden();
  await page.keyboard.press(`${MOD}+a`);
  await applyPreset(page, "decimal");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "2.");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "3.");
});

test("the outline preset renders I./II. and demoting shows the deeper level's format (AC-2)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["first", "second"]);
  await page.keyboard.press(`${MOD}+a`);
  await applyPreset(page, "outline");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "I.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "II.");
  // Tab demotes the second item to level 2 -> "A."
  await paragraphAt(page, 1).click();
  await page.keyboard.press("Home");
  await page.keyboard.press("Tab");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "A.");
});

test("the legal preset renders the full path without trailing dots (AC-3) and bullet-mixed mixes markers (AC-4)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["top", "sub"]);
  await page.keyboard.press(`${MOD}+a`);
  await applyPreset(page, "legal");
  await paragraphAt(page, 1).click();
  await page.keyboard.press("Home");
  await page.keyboard.press("Tab");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "1.1");
  // Bullet-mixed: level 1 numbers, level 2 bullets.
  await page.keyboard.press(`${MOD}+a`);
  await applyPreset(page, "bulletMixed");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "•");
});

// --- Define new multilevel list… (AC-5..12) ---------------------------------

test("the Define dialog sets per-level style, prefix/suffix, start, include levels, alignment and positions (AC-5..12)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["alpha", "beta"]);
  await page.keyboard.press(`${MOD}+a`);
  await applyPreset(page, "decimal");
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("multilevel-define-list").click();
  await expect(page.getByTestId("define-list-dialog")).toBeVisible();
  // Level 1: upperRoman in parentheses starting at 5.
  await page.getByTestId("define-list-style").selectOption("upperRoman");
  await page.getByTestId("define-list-prefix").fill("(");
  await page.getByTestId("define-list-suffix").fill(")");
  await page.getByTestId("define-list-start").fill("5");
  await page.getByTestId("define-list-alignment").selectOption("center");
  await page.getByTestId("define-list-number-position").fill("1.5");
  await page.getByTestId("define-list-text-position").fill("2.5");
  // Level 2 offers include-levels 1 and 2 (AC-8).
  await page.getByTestId("define-list-level").selectOption("2");
  const includeOptions = page.getByTestId("define-list-include").locator("option");
  await expect(includeOptions).toHaveCount(2);
  await page.getByTestId("define-list-apply").click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "(V)");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "(VI)");
  // Re-opening shows the persisted values (the same list's own definition).
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("multilevel-define-list").click();
  await expect(page.getByTestId("define-list-style")).toHaveValue("upperRoman");
  await expect(page.getByTestId("define-list-prefix")).toHaveValue("(");
  await expect(page.getByTestId("define-list-suffix")).toHaveValue(")");
  await expect(page.getByTestId("define-list-start")).toHaveValue("5");
  await expect(page.getByTestId("define-list-alignment")).toHaveValue("center");
  await expect(page.getByTestId("define-list-number-position")).toHaveValue("1.5");
  await expect(page.getByTestId("define-list-text-position")).toHaveValue("2.5");
  await page.getByTestId("define-list-cancel").click();
});

test("editing a level's format updates every item at that level in the same list and no other list (AC-17, AC-18)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["a1", "a2", "gap", "b1", "b2"]);
  await paragraphAt(page, 0).click();
  await paragraphAt(page, 1).click({ modifiers: ["Shift"] });
  await applyPreset(page, "decimal");
  await paragraphAt(page, 3).click();
  await paragraphAt(page, 4).click({ modifiers: ["Shift"] });
  await applyPreset(page, "decimal");
  // Edit list A's level 1 to upperRoman via the Define dialog.
  await paragraphAt(page, 0).click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("multilevel-define-list").click();
  await page.getByTestId("define-list-style").selectOption("upperRoman");
  await page.getByTestId("define-list-apply").click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "I.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "II.");
  // List B untouched (AC-18).
  await expect(paragraphAt(page, 3)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 4)).toHaveAttribute("data-list-marker", "2.");
});

// --- heading numbering (AC-13, AC-29..31) -----------------------------------

test("linking numbers Heading 1-3 automatically, including later headings, and survives demotion (AC-13, AC-31)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["Chapter One", "Section A"]);
  await paragraphAt(page, 0).click();
  await applyStyle(page, "Heading1");
  await paragraphAt(page, 1).click();
  await applyStyle(page, "Heading2");
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("multilevel-link-headings").click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "1.1.");
  // A newly typed Heading 1 auto-numbers.
  await paragraphAt(page, 1).click();
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Chapter Two");
  await applyStyle(page, "Heading1");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "2.");
  // Demoting it renumbers at the deeper level (AC-31).
  await applyStyle(page, "Heading2");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "1.2.");
});

test("numbered headings keep numbering when a heading is deleted or moved (AC-29, AC-30)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["One", "Two", "Three"]);
  for (const index of [0, 1, 2]) {
    await paragraphAt(page, index).click();
    await applyStyle(page, "Heading1");
  }
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("multilevel-link-headings").click();
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "3.");
  // Delete the middle heading: select its whole line plus the boundary.
  await paragraphAt(page, 1).click();
  await page.keyboard.press("Home");
  await page.keyboard.press("Shift+ArrowDown");
  await page.keyboard.press("Delete");
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "2.");
  // "Move" Three to the top by retyping it there and restyling (AC-30).
  await paragraphAt(page, 0).click();
  await page.keyboard.press("Home");
  await page.keyboard.type("Zero");
  await page.keyboard.press("Enter");
  await paragraphAt(page, 0).click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "2.");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "3.");
});

// --- context-menu numbering control (AC-14..16) -----------------------------

test("Restart numbering at 1 and Continue previous numbering (AC-14, AC-15)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["one", "two", "three"]);
  await page.keyboard.press(`${MOD}+a`);
  await applyPreset(page, "decimal");
  await paragraphAt(page, 1).click({ button: "right" });
  await page.getByTestId("context-menu-restart-numbering").click();
  await expect(paragraphAt(page, 0)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "1.");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "2.");
  await paragraphAt(page, 1).click({ button: "right" });
  await page.getByTestId("context-menu-continue-numbering").click();
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "2.");
  await expect(paragraphAt(page, 2)).toHaveAttribute("data-list-marker", "3.");
});

test("Set numbering value… sets an explicit start; cancelling changes nothing (AC-16)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["one", "two"]);
  await page.keyboard.press(`${MOD}+a`);
  await applyPreset(page, "decimal");
  // Cancelled prompt: no change.
  page.once("dialog", (dialog) => dialog.dismiss());
  await paragraphAt(page, 1).click({ button: "right" });
  await page.getByTestId("context-menu-set-numbering-value").click();
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "2.");
  // Accepted prompt: the item restarts at 42.
  page.once("dialog", (dialog) => dialog.accept("42"));
  await paragraphAt(page, 1).click({ button: "right" });
  await page.getByTestId("context-menu-set-numbering-value").click();
  await expect(paragraphAt(page, 1)).toHaveAttribute("data-list-marker", "42.");
});

test("the numbering section only appears on list paragraphs (AC-14, AC-15)", async ({ page }) => {
  await createDocument(page);
  await typeParagraphs(page, ["plain"]);
  await paragraphAt(page, 0).click({ button: "right" });
  await expect(page.getByTestId("editor-context-menu")).toBeVisible();
  await expect(page.getByTestId("context-menu-restart-numbering")).toHaveCount(0);
});

// --- custom tab stops (AC-19..25) -------------------------------------------

test("the tab dialog accepts stops of every alignment and leader (AC-19..23, AC-25)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.getByTestId("ruler-tab-selector").dblclick();
  await expect(page.getByTestId("tabs-dialog")).toBeVisible();
  const combos: Array<[string, string, string]> = [
    ["2", "left", ""],
    ["4", "right", "dot"],
    ["6", "center", "hyphen"],
    ["8", "decimal", "underscore"],
  ];
  for (const [position, alignment, leader] of combos) {
    await page.getByTestId("tab-position-input").fill(position);
    await page.getByTestId("tab-alignment-select").selectOption(alignment);
    await page.getByTestId("tab-leader-select").selectOption(leader);
    await page.getByTestId("tab-add").click();
  }
  const entries = page.getByTestId("tab-stop-entry");
  await expect(entries).toHaveCount(4);
  await expect(entries.nth(1)).toContainText("4 cm · right · dot");
  await expect(entries.nth(2)).toContainText("6 cm · centre · hyphen");
  await expect(entries.nth(3)).toContainText("8 cm · decimal · underscore");
  await page.getByTestId("tabs-apply").click();
  // The applied stops appear on the ruler as handles (AC-24, AC-25).
  await expect(page.getByTestId("ruler-tab-stop-0")).toBeVisible();
  await expect(page.getByTestId("ruler-tab-stop-3")).toBeVisible();
});

test("tab stops are settable and removable via the ruler (AC-24)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  const lane = page.getByTestId("ruler-tab-lane");
  await lane.click({ position: { x: 120, y: 5 } });
  await expect(page.getByTestId("ruler-tab-stop-0")).toBeVisible();
  // Clicking the handle removes the stop again.
  await page.getByTestId("ruler-tab-stop-0").click();
  await expect(page.getByTestId("ruler-tab-stop-0")).toHaveCount(0);
});

// --- round trips (AC-26..28) ------------------------------------------------

test("definitions round-trip through .docx as distinct abstractNum entries (AC-26)", async ({ page }) => {
  await buildThreeListDocument(page);
  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  const numbering = parts.get("word/numbering.xml")?.toString("utf8") ?? "";
  expect(numbering.match(/<w:abstractNum /g)?.length).toBe(3);
  expect(numbering).toContain('<w:lvlText w:val="%1."/>');
  expect(numbering).toContain('<w:lvlText w:val="%1.%2."/>');
  expect(numbering).toContain('<w:numFmt w:val="decimal"/>');
});

test("definitions round-trip through .odt as text:list-level-style-number entries (AC-27)", async ({ page }) => {
  await buildThreeListDocument(page);
  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const contentXml = parts.get("content.xml")?.toString("utf8") ?? "";
  expect(contentXml).toContain("<text:list-level-style-number");
  expect(contentXml).toContain('text:display-levels="2"');
});

test("two independent lists and a restarted list survive a .docx save and re-open (AC-28)", async ({ page }) => {
  await buildThreeListDocument(page);
  await expectThreeListMarkers(page);
  const path = await download(page, "docx");
  await openFile(page, path);
  await expectThreeListMarkers(page);
});

test("two independent lists and a restarted list survive an .odt save and re-open (AC-28)", async ({ page }) => {
  await buildThreeListDocument(page);
  await expectThreeListMarkers(page);
  const path = await download(page, "odt");
  await openFile(page, path);
  await expectThreeListMarkers(page);
});
