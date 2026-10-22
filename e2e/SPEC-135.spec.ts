import { test, expect, type Locator, type Page } from "@playwright/test";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzip } from "../src/lib/core/zip";
import { zip } from "../src/lib/core/zipWriter";

// SPEC-135 — notes, comments and hyperlinks on cells (AC-65–AC-71), with the
// real pointer and keyboard.

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

test.use({ viewport: { width: 1440, height: 900 } });

// The links (AC-68, AC-69 and the links of AC-71's .ods read): select C7,
// Ctrl/Cmd+K, link it to the place Data!A10 → the chip's Open activates
// "Data" with A10 selected; Ctrl/Cmd+click on the javascript: link's text
// shows LINK_NOT_OPENABLE_MESSAGE and opens nothing; in the .ods twin A7 and
// B7 carry data-link and following B7 activates "Data" with A1 selected.
test.describe("SPEC-135 links", () => {
  const NOT_OPENABLE = "This link can't be opened from the browser";

  function cellAt(page: Page, ref: string): Locator {
    return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
  }

  function sheetTab(page: Page, name: string): Locator {
    return page.getByTestId("sheet-tab").filter({ hasText: new RegExp(`^\\s*${name}\\s*$`) });
  }

  /** Clicks the centre of a cell through its column and row headers. */
  async function pressCell(page: Page, ref: string): Promise<void> {
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
    const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
    const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
    await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
    await expect(page.getByTestId("name-box")).toHaveValue(ref);
  }

  /** Opens a fixture, recording every window.open the page makes. */
  async function openFixture(page: Page, name: string): Promise<void> {
    await page.addInitScript(() => {
      const calls: unknown[] = [];
      Object.defineProperty(window, "__linkOpens", { value: calls });
      const original = window.open.bind(window);
      window.open = (...args: Parameters<typeof window.open>) => {
        calls.push(args);
        return original(...args);
      };
    });
    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(join(fixtures, name));
    await expect(page).toHaveURL(/\/sheets\/.+/);
    await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
  }

  const opens = (page: Page): Promise<unknown[]> => page.evaluate(() => (window as unknown as { __linkOpens: unknown[] }).__linkOpens);

  test("Ctrl/Cmd+K links C7 to the place Data!A10, and the chip's Open activates Data with A10 selected", async ({ page }) => {
    await openFixture(page, "spec135-annotations.xlsx");
    await expect(cellAt(page, "C7")).not.toHaveAttribute("data-link", "true");
    await pressCell(page, "C7");
    await page.keyboard.press("ControlOrMeta+k");
    const dialog = page.getByTestId("sheet-link-dialog");
    await expect(dialog).toBeVisible();
    await expect(page.getByTestId("sheet-link-text")).toHaveValue("Details");
    await page.getByTestId("sheet-link-kind-place").check();
    await page.getByTestId("sheet-link-sheet").selectOption("Data");
    await page.getByTestId("sheet-link-ref").fill("A10");
    await expect(page.getByTestId("sheet-link-error")).toHaveText("");
    await page.getByTestId("sheet-link-save").click();
    await expect(dialog).toBeHidden();
    await expect(cellAt(page, "C7")).toHaveAttribute("data-link", "true");
    await expect(cellAt(page, "C7")).toHaveText("Details");
    await expect(page.getByTestId("link-chip")).toBeVisible();
    await expect(page.getByTestId("link-chip-target")).toHaveText("Data!A10");
    await page.getByTestId("link-chip-open").click();
    await expect(sheetTab(page, "Data")).toHaveAttribute("aria-selected", "true");
    await expect(page.getByTestId("name-box")).toHaveValue("A10");
    expect(await opens(page)).toEqual([]);
  });

  test("Ctrl/Cmd+click on the javascript: link's text shows the refusal and opens nothing", async ({ page }) => {
    await openFixture(page, "spec135-annotations.xlsx");
    const popups: Page[] = [];
    page.context().on("page", (opened) => popups.push(opened));
    await expect(cellAt(page, "B8")).toHaveAttribute("data-link", "true");
    await pressCell(page, "A1");
    await cellAt(page, "B8").locator(".sheet-cell-link").click({ modifiers: ["ControlOrMeta"] });
    await expect(page.getByTestId("toast").filter({ hasText: NOT_OPENABLE })).toBeVisible();
    // Followed, not added: the selection is still A1 alone.
    await expect(page.getByTestId("name-box")).toHaveValue("A1");
    await expect(page.getByTestId("sheet-selection")).toHaveCount(0);
    expect(await opens(page)).toEqual([]);
    expect(popups).toEqual([]);
    expect(page.context().pages()).toHaveLength(1);
  });

  test("an .ods workbook's links read: A7 and B7 carry data-link, and following B7 activates Data with A1 selected", async ({ page }) => {
    await openFixture(page, "spec135-annotations.ods");
    await expect(cellAt(page, "A7")).toHaveAttribute("data-link", "true");
    await expect(cellAt(page, "B7")).toHaveAttribute("data-link", "true");
    await expect(cellAt(page, "A1")).not.toHaveAttribute("data-link", "true");
    // Leave Data's selection elsewhere, so the follow is what selects A1.
    await sheetTab(page, "Data").click();
    await pressCell(page, "A10");
    await sheetTab(page, "Summary").click();
    await expect(sheetTab(page, "Summary")).toHaveAttribute("aria-selected", "true");
    await cellAt(page, "B7").locator(".sheet-cell-link").click({ modifiers: ["ControlOrMeta"] });
    await expect(sheetTab(page, "Data")).toHaveAttribute("aria-selected", "true");
    await expect(page.getByTestId("name-box")).toHaveValue("A1");
    expect(await opens(page)).toEqual([]);
  });
});

// The notes and comments (AC-65's notes, comments and round trip, AC-66,
// AC-67, AC-70, AC-71's .ods read): B2's note indicator and its popover on
// hover → Shift+F2 on D4, type, Escape commits → the comments pane: reply to
// D2's thread as "You", resolve it → insert a row above row 2, the note now
// on B3 → download, reopen: every note, the reply, the resolved state and the
// fixture's links come back, the VML checkbox shape byte-for-byte; the .ods
// read half; and the human-captured Excel 365 workbook, skipped until
// committed.
test.describe("SPEC-135 notes and comments", () => {
  const ODS_NOTES_ONLY = "OpenDocument spreadsheets support notes only";
  const EXCEL365 = join(fixtures, "spec135-excel365-threads.xlsx");

  function cellAt(page: Page, ref: string): Locator {
    return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
  }

  const noteMark = (page: Page, ref: string): Locator => cellAt(page, ref).getByTestId("sheet-note-indicator");
  const threadMark = (page: Page, ref: string): Locator => cellAt(page, ref).getByTestId("sheet-thread-indicator");

  async function open(page: Page, path: string): Promise<void> {
    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(path);
    await expect(page).toHaveURL(/\/sheets\/.+/);
    await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
  }

  /** Clicks the centre of a cell through its column and row headers. */
  async function selectCell(page: Page, ref: string): Promise<void> {
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
    const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
    const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
    await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
    await expect(page.getByTestId("name-box")).toHaveValue(ref);
  }

  /** Rests the pointer on a cell (any popover shown before gone first, so
   * it cannot sit over the cell). */
  async function hover(page: Page, ref: string): Promise<void> {
    await page.mouse.move(0, 0);
    await expect(page.locator("[data-note-popover]")).toHaveCount(0);
    await cellAt(page, ref).hover();
  }

  /** Shift+F2 on `ref`, types `text`, Escape commits. */
  async function addNote(page: Page, ref: string, text: string): Promise<void> {
    await selectCell(page, ref);
    await page.keyboard.press("Shift+F2");
    const editor = page.getByTestId("sheet-note-editor");
    await expect(editor).toBeFocused();
    await page.keyboard.type(text);
    await page.keyboard.press("Escape");
    await expect(editor).toBeHidden();
  }

  /** The comments pane, opened from the app bar's toggle. */
  async function openComments(page: Page): Promise<Locator> {
    const panel = page.getByTestId("sheet-comments-panel");
    if (!(await panel.isVisible())) {
      await page.getByTestId("sheet-comments-toggle").click();
    }
    await expect(panel).toBeVisible();
    return panel;
  }

  const threadCard = (panel: Locator, opening: string): Locator => panel.locator('[data-testid^="sheet-comment-thread-"]').filter({ hasText: opening });

  /** Replies to D2's thread as "You" and resolves it. */
  async function replyAndResolve(page: Page): Promise<Locator> {
    const panel = await openComments(page);
    const thread = threadCard(panel, "Where does this number come from?");
    await expect(thread).toBeVisible();
    await thread.getByTestId("sheet-comment-reply-input").fill("Noted, thanks.");
    await thread.getByTestId("sheet-comment-reply-submit").click();
    const reply = thread.locator('[data-testid^="sheet-comment-message-"]').filter({ hasText: "Noted, thanks." });
    await expect(reply.getByTestId("sheet-comment-author")).toHaveText("You");
    await thread.getByTestId("sheet-comment-resolve").click();
    await expect(thread).toHaveAttribute("data-resolved", "true");
    await expect(thread.getByTestId("sheet-comment-reopen")).toBeVisible();
    return thread;
  }

  /** The VML part's form-control checkbox shape, verbatim. */
  function checkboxShape(vml: Uint8Array | undefined): string {
    const text = new TextDecoder().decode(vml);
    const match = /<v:shape\b(?:(?!<v:shape\b)[\s\S])*?ObjectType="Checkbox"[\s\S]*?<\/v:shape>/.exec(text);
    expect(match, "the VML part holds the checkbox shape").not.toBeNull();
    return match![0];
  }

  test("B2 shows a note indicator and hovering it shows the note; D2 and E2 show thread markers, the resolved one dimmed", async ({ page }) => {
    await open(page, join(fixtures, "spec135-annotations.xlsx"));
    await expect(noteMark(page, "B2")).toBeVisible();
    await expect(noteMark(page, "A2")).toHaveCount(0);
    await expect(threadMark(page, "D2")).not.toHaveAttribute("data-resolved", "true");
    await expect(threadMark(page, "E2")).toHaveAttribute("data-resolved", "true");
    // The indicator restyles nothing: B2 still shows its value.
    await expect(cellAt(page, "B2")).toHaveText("1200");
    await hover(page, "B2");
    const popover = page.getByTestId("sheet-note-popover");
    await expect(popover).toBeVisible();
    await expect(popover.getByTestId("sheet-note-popover-author")).toHaveText("Alex Kim");
    await expect(popover.getByTestId("sheet-note-popover-text")).toHaveText("Alex Kim:\nQ3 figure, unaudited.");
    // The active cell's note reaches assistive technology through aria-describedby.
    await selectCell(page, "B2");
    const describedBy = (await cellAt(page, "B2").getAttribute("aria-describedby")) ?? "";
    expect(describedBy.split(" ")).toContain("sheet-annotation-description");
    await expect(page.locator("#sheet-annotation-description")).toHaveText(/Q3 figure, unaudited\./);
    // A thread's popover: its opening message, its replies, and the way to the pane.
    await hover(page, "D2");
    const thread = page.getByTestId("sheet-thread-popover");
    await expect(thread).toBeVisible();
    await expect(thread.getByTestId("sheet-thread-popover-text")).toHaveText("Where does this number come from?");
    await expect(thread.getByTestId("sheet-thread-popover-replies")).toHaveText("2 replies");
    await thread.getByTestId("sheet-thread-open").click();
    const panel = page.getByTestId("sheet-comments-panel");
    await expect(panel).toBeVisible();
    await expect(threadCard(panel, "Where does this number come from?")).toBeVisible();
  });

  test("Shift+F2 on D4, type a note, Escape commits it: the indicator appears and hovering shows it", async ({ page }) => {
    await open(page, join(fixtures, "spec135-annotations.xlsx"));
    await expect(noteMark(page, "D4")).toHaveCount(0);
    await addNote(page, "D4", "Check with finance");
    await expect(noteMark(page, "D4")).toBeVisible();
    await hover(page, "D4");
    const popover = page.getByTestId("sheet-note-popover");
    await expect(popover.getByTestId("sheet-note-popover-text")).toHaveText("Check with finance");
    await expect(popover.getByTestId("sheet-note-popover-author")).toHaveText("You");
    // Undo takes the note back off in one step.
    await page.getByTestId("spreadsheet-undo").click();
    await expect(noteMark(page, "D4")).toHaveCount(0);
  });

  test("the comments pane: reply to D2's thread as You and resolve it", async ({ page }) => {
    await open(page, join(fixtures, "spec135-annotations.xlsx"));
    const panel = await openComments(page);
    await expect(panel.getByTestId("sheet-comment-group").first()).toContainText("Summary");
    const thread = await replyAndResolve(page);
    await expect(threadMark(page, "D2")).toHaveAttribute("data-resolved", "true");
    // The unresolved filter now hides both resolved threads.
    await panel.getByTestId("sheet-comment-filter").selectOption("unresolved");
    await expect(thread).toBeHidden();
    await expect(panel.getByTestId("sheet-comment-empty")).toBeVisible();
  });

  test("insert a row above row 2: the note follows to B3; download and reopen keep every note, the reply, the resolved state, the links and the checkbox's VML", async ({ page }) => {
    const sourcePath = join(fixtures, "spec135-annotations.xlsx");
    await open(page, sourcePath);
    await addNote(page, "D4", "Check with finance");
    await replyAndResolve(page);

    await page.getByTestId("sheet-row-header").filter({ hasText: /^2$/ }).click({ button: "right" });
    await page.getByTestId("sheet-header-menu").getByTestId("sheet-insert-before").click();
    await expect(cellAt(page, "B3")).toHaveText("1200");
    await expect(noteMark(page, "B3")).toBeVisible();
    await expect(noteMark(page, "B2")).toHaveCount(0);
    await hover(page, "B3");
    await expect(page.getByTestId("sheet-note-popover-text")).toHaveText("Alex Kim:\nQ3 figure, unaudited.");

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("spreadsheet-download").click();
    const saved = await downloadPromise;
    const path = join(tmpdir(), `spec135-${Date.now()}-annotations.xlsx`);
    await saved.saveAs(path);
    const savedParts = await unzip(new Uint8Array(readFileSync(path)));
    const sourceParts = await unzip(new Uint8Array(readFileSync(sourcePath)));
    expect(checkboxShape(savedParts.get("xl/drawings/vmlDrawing1.vml"))).toBe(checkboxShape(sourceParts.get("xl/drawings/vmlDrawing1.vml")));

    await open(page, path);
    for (const ref of ["B3", "B5", "D5"]) {
      await expect(noteMark(page, ref), ref).toBeVisible();
    }
    await expect(threadMark(page, "D3")).toHaveAttribute("data-resolved", "true");
    await expect(threadMark(page, "E3")).toHaveAttribute("data-resolved", "true");
    // B4's visible note, now B5's, shows as its box.
    await expect(page.getByTestId("sheet-note-box").filter({ hasText: "Always shown" })).toBeVisible();
    await hover(page, "D5");
    await expect(page.getByTestId("sheet-note-popover-text")).toHaveText("Check with finance");
    const panel = await openComments(page);
    const thread = threadCard(panel, "Where does this number come from?");
    await expect(thread).toHaveAttribute("data-resolved", "true");
    await expect(thread.locator('[data-testid^="sheet-comment-message-"]')).toHaveCount(4);
    await expect(thread.locator('[data-testid^="sheet-comment-message-"]').last()).toContainText("Noted, thanks.");
    // The fixture's links, a row lower.
    for (const ref of ["A7", "B7", "A8", "B8", "A9", "B9"]) {
      await expect(cellAt(page, ref), ref).toHaveAttribute("data-link", "true");
    }
  });

  test("an .ods workbook's notes read: B2's indicator and popover, B4's shown box, the pane listing notes and New comment disabled", async ({ page }) => {
    await open(page, join(fixtures, "spec135-annotations.ods"));
    await expect(noteMark(page, "B2")).toBeVisible();
    await hover(page, "B2");
    const popover = page.getByTestId("sheet-note-popover");
    await expect(popover.getByTestId("sheet-note-popover-author")).toHaveText("Alex Kim");
    await expect(popover.getByTestId("sheet-note-popover-text")).toHaveText("Q3 figure, unaudited.");
    const box = page.getByTestId("sheet-note-box").filter({ hasText: "Always shown" });
    await expect(box).toBeVisible();
    await expect(box).toContainText("Second line");
    // The cell menu's New comment is off, saying why.
    await cellAt(page, "C3").click({ button: "right" });
    const comment = page.getByTestId("sheet-cell-menu-comment");
    await expect(comment).toBeDisabled();
    await expect(comment).toHaveAttribute("title", ODS_NOTES_ONLY);
    await page.keyboard.press("Escape");
    const panel = await openComments(page);
    await expect(panel.locator('[data-testid^="sheet-comment-note-"]')).toHaveCount(2);
    await expect(panel.locator('[data-testid^="sheet-comment-thread-"]')).toHaveCount(0);
  });

  test("the Excel 365 workbook opens with its threads, their authors and resolved state, and its legacy note", async ({ page }) => {
    test.skip(!existsSync(EXCEL365), "fixture not yet captured: spec135-excel365-threads.xlsx");
    await open(page, EXCEL365);
    await expect(page.getByTestId("sheet-thread-indicator").first()).toBeVisible();
    await expect(page.getByTestId("sheet-note-indicator").first()).toBeVisible();
    const panel = await openComments(page);
    const threads = panel.locator('[data-testid^="sheet-comment-thread-"]');
    expect(await threads.count()).toBeGreaterThan(0);
    await expect(panel.locator('[data-testid^="sheet-comment-thread-"][data-resolved="true"]').first()).toBeVisible();
    const authors = await panel.getByTestId("sheet-comment-author").allTextContents();
    expect(authors.length).toBeGreaterThan(1);
    expect(authors.every((author) => author.trim() !== "" && !author.startsWith("tc="))).toBe(true);
  });
});

// What review and QA asked of the second round, each with the real pointer
// and keyboard: annotations refused where they could not be saved (AC-11 —
// an .xlsx sheet whose comment parts broke the subset; AC-34 — links on a
// CSV workbook), the note editor on a cell scrolled out of view (AC-23), a
// save over self-closing comment parts (AC-49), and the criteria the first
// QA pass left unexercised: the hyperlink look on insert (AC-35), place
// locations following a structural edit, a rename and a delete (AC-42,
// AC-46), HYPERLINK() cells (AC-40), cut-and-paste and sort carrying
// annotations (AC-44, AC-45), <mentions> re-emitted (AC-13) and the
// sanctioned deltas (AC-59).
test.describe("SPEC-135 annotations: refusals, moves and deltas", () => {
  test.use({ permissions: ["clipboard-read", "clipboard-write"] });

  const SOURCE = join(fixtures, "spec135-annotations.xlsx");
  const NOTES_KEPT = "This sheet's notes and comments are kept exactly as the file has them, so they can't be changed here";
  const CSV_NO_LINKS = "CSV and TSV files can't keep links";
  const DESTINATION_MISSING = "This link's destination doesn't exist";
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();

  function cellAt(page: Page, ref: string): Locator {
    return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
  }
  const noteMark = (page: Page, ref: string): Locator => cellAt(page, ref).getByTestId("sheet-note-indicator");
  const threadMark = (page: Page, ref: string): Locator => cellAt(page, ref).getByTestId("sheet-thread-indicator");
  function sheetTab(page: Page, name: string): Locator {
    return page.getByTestId("sheet-tab").filter({ hasText: new RegExp(`^\\s*${name}\\s*$`) });
  }

  async function open(page: Page, path: string): Promise<void> {
    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(path);
    await expect(page).toHaveURL(/\/sheets\/.+/);
    await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
  }

  /** The fixture with some parts edited, written to a temporary file. */
  async function variant(label: string, edit: (parts: Map<string, Uint8Array>) => void): Promise<string> {
    const parts = await unzip(new Uint8Array(readFileSync(SOURCE)));
    edit(parts);
    const path = join(tmpdir(), `spec135-${label}-${Date.now()}.xlsx`);
    writeFileSync(path, await zip([...parts].map(([name, data]) => ({ name, data }))));
    return path;
  }
  const editPart = (parts: Map<string, Uint8Array>, name: string, edit: (xml: string) => string): void => {
    parts.set(name, encoder.encode(edit(decoder.decode(parts.get(name)))));
  };

  /** Clicks the centre of a cell through its column and row headers. */
  async function selectCell(page: Page, ref: string, modifiers: ("Shift" | "ControlOrMeta")[] = []): Promise<void> {
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
    const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
    const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
    for (const key of modifiers) {
      await page.keyboard.down(key);
    }
    await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
    for (const key of modifiers) {
      await page.keyboard.up(key);
    }
    if (modifiers.length === 0) {
      await expect(page.getByTestId("name-box")).toHaveValue(ref);
    }
  }

  async function hover(page: Page, ref: string): Promise<void> {
    await page.mouse.move(0, 0);
    await expect(page.locator("[data-note-popover]")).toHaveCount(0);
    await cellAt(page, ref).hover();
  }

  /** Shift+F2 on the active cell, types `text`, Escape commits. */
  async function typeNote(page: Page, text: string): Promise<void> {
    await page.keyboard.press("Shift+F2");
    const editor = page.getByTestId("sheet-note-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();
    await page.keyboard.type(text);
    await page.keyboard.press("Escape");
    await expect(editor).toBeHidden();
  }

  /** Ctrl/Cmd+K on the active cell, linking it to a place on `sheet`. */
  async function linkToPlace(page: Page, sheet: string, ref: string): Promise<void> {
    await page.keyboard.press("ControlOrMeta+k");
    await expect(page.getByTestId("sheet-link-dialog")).toBeVisible();
    await page.getByTestId("sheet-link-kind-place").check();
    await page.getByTestId("sheet-link-sheet").selectOption(sheet);
    await page.getByTestId("sheet-link-ref").fill(ref);
    await page.getByTestId("sheet-link-save").click();
    await expect(page.getByTestId("sheet-link-dialog")).toBeHidden();
  }

  async function download(page: Page, label: string, extension = "xlsx"): Promise<{ path: string; parts: Map<string, Uint8Array> }> {
    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("spreadsheet-download").click();
    const saved = await downloadPromise;
    const path = join(tmpdir(), `spec135-${label}-${Date.now()}.${extension}`);
    await saved.saveAs(path);
    return { path, parts: await unzip(new Uint8Array(readFileSync(path))) };
  }

  test("AC-11: an .xlsx sheet whose comment parts broke the subset refuses New note and New comment with the reason; a modelled sheet still takes notes", async ({ page }) => {
    // Two comments on B2: the reader models none of Summary's notes or threads.
    const path = await variant("two-on-b2", (parts) => editPart(parts, "xl/comments1.xml", (xml) => xml.replace(`<comment ref="B4"`, `<comment ref="B2"`)));
    await open(page, path);
    await expect(noteMark(page, "B2")).toHaveCount(0);
    await expect(cellAt(page, "B7")).toHaveAttribute("data-link", "true");
    await cellAt(page, "A10").click({ button: "right" });
    const note = page.getByTestId("sheet-cell-menu-note");
    await expect(note).toBeDisabled();
    await expect(note).toHaveAttribute("title", NOTES_KEPT);
    const comment = page.getByTestId("sheet-cell-menu-comment");
    await expect(comment).toBeDisabled();
    await expect(comment).toHaveAttribute("title", NOTES_KEPT);
    // Its links were modelled, so Link… still applies.
    await expect(page.getByTestId("sheet-cell-menu-link")).toBeEnabled();
    await page.keyboard.press("Escape");
    await selectCell(page, "A10");
    await page.keyboard.press("Shift+F2");
    await expect(page.getByTestId("sheet-note-editor")).toHaveCount(0);
    await page.getByTestId("menu-title-insert").click();
    await expect(page.getByTestId("sheet-menu-insert-note")).toBeDisabled();
    await expect(page.getByTestId("sheet-menu-insert-comment")).toBeDisabled();
    await page.keyboard.press("Escape");
    // "Data" has no comment parts at all: a note goes in, and survives the save.
    await sheetTab(page, "Data").click();
    await selectCell(page, "A1");
    await typeNote(page, "Kept on Data");
    await expect(noteMark(page, "A1")).toBeVisible();
    const saved = await download(page, "two-on-b2-saved");
    const source = await unzip(new Uint8Array(readFileSync(path)));
    expect(decoder.decode(saved.parts.get("xl/comments1.xml"))).toBe(decoder.decode(source.get("xl/comments1.xml")));
    await open(page, saved.path);
    await sheetTab(page, "Data").click();
    await expect(noteMark(page, "A1")).toBeVisible();
  });

  test("AC-23: Shift+F2 on an active cell scrolled out of view brings it back and opens the note editor on it", async ({ page }) => {
    await open(page, SOURCE);
    await selectCell(page, "D4");
    const grid = page.getByTestId("spreadsheet-grid");
    const box = (await grid.boundingBox())!;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.wheel(0, 4000);
    await expect(cellAt(page, "D4")).toHaveCount(0);
    await expect(page.getByTestId("name-box")).toHaveValue("D4");
    await typeNote(page, "Written off screen");
    await expect(noteMark(page, "D4")).toBeVisible();
    await hover(page, "D4");
    await expect(page.getByTestId("sheet-note-popover-text")).toHaveText("Written off screen");
  });

  test("AC-34: a CSV workbook offers no Link… — disabled with its reason, and Ctrl/Cmd+K opens nothing", async ({ page }) => {
    await open(page, join(fixtures, "spec126-excel-utf8.csv"));
    await selectCell(page, "B3");
    await page.keyboard.press("ControlOrMeta+k");
    await expect(page.getByTestId("sheet-link-dialog")).toHaveCount(0);
    await cellAt(page, "B3").click({ button: "right" });
    const link = page.getByTestId("sheet-cell-menu-link");
    await expect(link).toBeDisabled();
    await expect(link).toHaveAttribute("title", CSV_NO_LINKS);
    await page.keyboard.press("Escape");
    await page.getByTestId("menu-title-insert").click();
    await expect(page.getByTestId("sheet-menu-insert-link")).toBeDisabled();
    await expect(page.getByTestId("sheet-menu-insert-link")).toHaveAttribute("title", CSV_NO_LINKS);
    await page.keyboard.press("Escape");
    await expect(cellAt(page, "B3")).not.toHaveAttribute("data-link", "true");
  });

  test("AC-49: a note and a comment added over self-closing <comments/>, <ThreadedComments/> and VML <xml/> save as whole parts and reopen", async ({ page }) => {
    const empty = (root: string) => (xml: string): string => {
      const start = new RegExp(`<${root}\\b[^>]*>`).exec(xml)![0];
      return xml.slice(0, xml.indexOf(start)) + start.replace(/>$/, "/>");
    };
    const path = await variant("self-closing", (parts) => {
      editPart(parts, "xl/comments1.xml", empty("comments"));
      editPart(parts, "xl/threadedComments/threadedComment1.xml", empty("ThreadedComments"));
      editPart(parts, "xl/drawings/vmlDrawing1.vml", empty("xml"));
    });
    await open(page, path);
    await expect(noteMark(page, "B2")).toHaveCount(0);
    await selectCell(page, "D4");
    await typeNote(page, "Over an empty part");
    await selectCell(page, "H12");
    await page.mouse.move(0, 0);
    await cellAt(page, "H12").click({ button: "right" });
    await page.getByTestId("sheet-cell-menu-comment").click();
    const composer = page.getByTestId("sheet-comment-composer");
    await expect(composer).toBeVisible();
    await composer.getByTestId("sheet-comment-composer-input").fill("A thread over an empty part");
    await composer.getByTestId("sheet-comment-composer-submit").click();
    await expect(threadMark(page, "H12")).toBeVisible();
    const saved = await download(page, "self-closing-saved");
    const comments = decoder.decode(saved.parts.get("xl/comments1.xml"));
    expect(comments).toMatch(/<comments [^>]*><authors>[\s\S]*<\/authors><commentList>[\s\S]*<\/commentList><\/comments>$/);
    expect(comments).toContain("Over an empty part");
    expect(decoder.decode(saved.parts.get("xl/threadedComments/threadedComment1.xml"))).toMatch(/<ThreadedComments [^>]*><threadedComment ref="H12"[\s\S]*<\/ThreadedComments>$/);
    expect(decoder.decode(saved.parts.get("xl/drawings/vmlDrawing1.vml"))).toMatch(/<\/xml>$/);
    await open(page, saved.path);
    await expect(noteMark(page, "D4")).toBeVisible();
    await expect(threadMark(page, "H12")).toBeVisible();
    await hover(page, "D4");
    await expect(page.getByTestId("sheet-note-popover-text")).toHaveText("Over an empty part");
  });

  test("AC-35, AC-42, AC-46: an inserted link takes the hyperlink look; its place follows a row insert on its sheet, a rename, and a delete to #REF!", async ({ page }) => {
    await open(page, SOURCE);
    await selectCell(page, "C7");
    await linkToPlace(page, "Data", "A10");
    // AC-35: C7's font colour was automatic — now the theme's hlink (#467886), underlined.
    const look = await cellAt(page, "C7")
      .locator(".sheet-cell-text")
      .evaluate((el) => ({ color: getComputedStyle(el).color, line: getComputedStyle(el).textDecorationLine }));
    expect(look.color).toBe("rgb(70, 120, 134)");
    expect(look.line).toContain("underline");
    // AC-42: a row inserted above row 5 on Data moves the place to Data!A11.
    await sheetTab(page, "Data").click();
    await page.getByTestId("sheet-row-header").filter({ hasText: /^5$/ }).click({ button: "right" });
    await page.getByTestId("sheet-header-menu").getByTestId("sheet-insert-before").click();
    await sheetTab(page, "Summary").click();
    await selectCell(page, "C7");
    await expect(page.getByTestId("link-chip-target")).toHaveText("Data!A11");
    await selectCell(page, "B7");
    await expect(page.getByTestId("link-chip-target")).toHaveText("Data!A1");
    // AC-46: a rename renames the places.
    await sheetTab(page, "Data").dblclick();
    const input = page.getByTestId("sheet-tab-rename-input");
    await input.fill("Facts");
    await input.press("Enter");
    await sheetTab(page, "Summary").click();
    await selectCell(page, "C7");
    await expect(page.getByTestId("link-chip-target")).toHaveText("Facts!A11");
    await page.getByTestId("link-chip-open").click();
    await expect(sheetTab(page, "Facts")).toHaveAttribute("aria-selected", "true");
    await expect(page.getByTestId("name-box")).toHaveValue("A11");
    // ...and a delete turns them to #REF!, which follows nowhere.
    await sheetTab(page, "Facts").click({ button: "right" });
    await page.getByTestId("sheet-tab-menu-delete").click();
    await page.getByTestId("sheet-tab-delete-confirm").click();
    await expect(sheetTab(page, "Facts")).toHaveCount(0);
    await selectCell(page, "C7");
    await expect(page.getByTestId("link-chip-target")).toContainText("#REF!");
    await page.getByTestId("link-chip-open").click();
    await expect(page.getByTestId("toast").filter({ hasText: DESTINATION_MISSING })).toBeVisible();
  });

  test("AC-40: a HYPERLINK() cell behaves as linked — data-link, the chip, Ctrl/Cmd+click — and is never written as a <hyperlink>", async ({ page }) => {
    await open(page, SOURCE);
    await selectCell(page, "C9");
    await page.keyboard.type('=HYPERLINK("#Data!A10","Go")');
    await page.keyboard.press("Enter");
    await expect(cellAt(page, "C9")).toHaveText("Go");
    await expect(cellAt(page, "C9")).toHaveAttribute("data-link", "true");
    await selectCell(page, "C9");
    await expect(page.getByTestId("link-chip-target")).toHaveText("Data!A10");
    await selectCell(page, "A1");
    await cellAt(page, "C9").locator(".sheet-cell-link").click({ modifiers: ["ControlOrMeta"] });
    await expect(sheetTab(page, "Data")).toHaveAttribute("aria-selected", "true");
    await expect(page.getByTestId("name-box")).toHaveValue("A10");
    const saved = await download(page, "hyperlink-formula");
    const sheet = decoder.decode(saved.parts.get("xl/worksheets/sheet1.xml"));
    expect(sheet).toContain("HYPERLINK(");
    expect(sheet).not.toMatch(/<hyperlink ref="C9"/);
  });

  test("AC-44: cut and paste carries a note, the threads and a link to the destination", async ({ page }) => {
    await open(page, SOURCE);
    await selectCell(page, "B2");
    await page.keyboard.press("ControlOrMeta+x");
    await selectCell(page, "F10");
    await page.keyboard.press("ControlOrMeta+v");
    await expect(cellAt(page, "F10")).toHaveText("1200");
    await expect(noteMark(page, "F10")).toBeVisible();
    await expect(noteMark(page, "B2")).toHaveCount(0);
    await selectCell(page, "D2");
    await selectCell(page, "E2", ["Shift"]);
    await page.keyboard.press("ControlOrMeta+x");
    await selectCell(page, "F12");
    await page.keyboard.press("ControlOrMeta+v");
    await expect(threadMark(page, "F12")).toBeVisible();
    await expect(threadMark(page, "G12")).toHaveAttribute("data-resolved", "true");
    await expect(threadMark(page, "D2")).toHaveCount(0);
    await selectCell(page, "A7");
    await page.keyboard.press("ControlOrMeta+x");
    await selectCell(page, "H14");
    await page.keyboard.press("ControlOrMeta+v");
    await expect(cellAt(page, "H14")).toHaveAttribute("data-link", "true");
    await expect(cellAt(page, "A7")).not.toHaveAttribute("data-link", "true");
    await selectCell(page, "H14");
    await expect(page.getByTestId("link-chip-target")).toHaveText("https://example.com/");
  });

  test("AC-45: a sort carries the notes with their rows", async ({ page }) => {
    await open(page, SOURCE);
    // A2:B4 by column B ascending: 800 (B4, the shown note) first, then 1200 (B2's note), the empty row last.
    await selectCell(page, "B2");
    await selectCell(page, "A4", ["Shift"]);
    await page.getByTestId("menu-title-data").click();
    await page.getByTestId("sheet-menu-sort-asc").click();
    await expect(cellAt(page, "B2")).toHaveText("800");
    await expect(cellAt(page, "B3")).toHaveText("1200");
    await expect(page.getByTestId("sheet-note-box").filter({ hasText: "Always shown" })).toHaveAttribute("data-note-ref", "B2");
    await expect(noteMark(page, "B3")).toBeVisible();
    await expect(noteMark(page, "B4")).toHaveCount(0);
    await hover(page, "B3");
    await expect(page.getByTestId("sheet-note-popover-text")).toHaveText("Alex Kim:\nQ3 figure, unaudited.");
  });

  test("AC-13: a message's <mentions> is re-emitted while the message is unedited, through a reply and a note edit", async ({ page }) => {
    const mentions = `<mentions><mention mentionpersonId="{6D1E2F30-4152-4637-8A9B-0C1D2E3F4051}" mentionId="{0F1E2D3C-4B5A-4978-8695-A4B3C2D1E0F9}" startIndex="0" length="7"/></mentions>`;
    const path = await variant("mentions", (parts) =>
      editPart(parts, "xl/threadedComments/threadedComment1.xml", (xml) => {
        const at = xml.indexOf("</text>", xml.indexOf("<threadedComment ", xml.indexOf("<threadedComment ") + 1)) + "</text>".length;
        return xml.slice(0, at) + mentions + xml.slice(at);
      }),
    );
    await open(page, path);
    await selectCell(page, "D4");
    await typeNote(page, "Forces the rebuild");
    await page.getByTestId("sheet-comments-toggle").click();
    const panel = page.getByTestId("sheet-comments-panel");
    const thread = panel.locator('[data-testid^="sheet-comment-thread-"]').filter({ hasText: "Where does this number come from?" });
    await thread.getByTestId("sheet-comment-reply-input").fill("Seen.");
    await thread.getByTestId("sheet-comment-reply-submit").click();
    await expect(thread.locator('[data-testid^="sheet-comment-message-"]')).toHaveCount(4);
    const saved = await download(page, "mentions-saved");
    const threaded = decoder.decode(saved.parts.get("xl/threadedComments/threadedComment1.xml"));
    expect(threaded).toContain(mentions);
    expect(threaded).toContain("Seen.");
  });

  test("AC-59: a moved note's box keeps its size; an .ods cell with two links keeps the first once rewritten", async ({ page }) => {
    await open(page, SOURCE);
    await page.getByTestId("sheet-row-header").filter({ hasText: /^2$/ }).click({ button: "right" });
    await page.getByTestId("sheet-header-menu").getByTestId("sheet-insert-before").click();
    await expect(noteMark(page, "B5")).toBeVisible();
    const saved = await download(page, "moved-box");
    const source = await unzip(new Uint8Array(readFileSync(SOURCE)));
    /** The VML note shape on a (0-based) row. */
    const noteShape = (parts: Map<string, Uint8Array>, row: number): string => {
      const shapes = [...decoder.decode(parts.get("xl/drawings/vmlDrawing1.vml")).matchAll(/<v:shape\b(?:(?!<v:shape\b)[\s\S])*?<\/v:shape>/g)].map((match) => match[0]);
      const shape = shapes.find((each) => each.includes('ObjectType="Note"') && each.includes(`<x:Row>${row}</x:Row>`));
      expect(shape, `a note shape on VML row ${row}`).toBeDefined();
      return shape!;
    };
    const size = (shape: string): string[] => [/width:([\d.]+pt)/.exec(shape)![1]!, /height:([\d.]+pt)/.exec(shape)![1]!];
    // B4's shown note (VML row 3) is B5's now (row 4): the same box size.
    expect(size(noteShape(saved.parts, 4))).toEqual(size(noteShape(source, 3)));
    expect(noteShape(saved.parts, 4)).toContain("<x:Visible/>");

    await open(page, join(fixtures, "spec135-annotations.ods"));
    await selectCell(page, "C9");
    await page.keyboard.type("One only");
    await page.keyboard.press("Enter");
    const ods = await download(page, "two-links", "ods");
    const content = decoder.decode(ods.parts.get("content.xml"));
    const cell = /<table:table-cell\b[^>]*>(?:(?!<\/table:table-cell>)[\s\S])*One only(?:(?!<\/table:table-cell>)[\s\S])*<\/table:table-cell>/.exec(content)![0];
    expect(cell).toContain('xlink:href="https://one.example/"');
    expect(cell).not.toContain("two.example");
    expect(cell.match(/<text:a\b/g)).toHaveLength(1);
  });
});
