import { test, expect, type Locator, type Page } from "@playwright/test";
import { existsSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzip } from "../src/lib/core/zip";

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
