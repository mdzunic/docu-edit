import { test, expect, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";

// SPEC-103 — autosave, drafts and version history for presentations, as one
// session: open the pptx fixture, retitle slide 1, reload — the edit is there
// with no recovery bar; plant a newer draft the way a crash would leave one
// (SPEC-010's page.evaluate + indexedDB technique) and reopen — the bar shows
// over the drafted slides, and Discard brings the committed deck back; then
// edit and download (a commit-point snapshot), open the history panel, name
// that version, preview it, leave the preview, restore it — the stage shows
// the restored slides and the panel lists the restore's own two snapshots.

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE = join(here, "fixtures", "mini-deck.pptx");
const SAVED = "All changes saved";

/** The frames on the MAIN stage (not the filmstrip thumbnails). */
function stageFrames(page: Page) {
  return page.getByTestId("stage-workspace").getByTestId("slide-frame");
}

async function openFixture(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  const id = page.url().match(/\/slides\/([^/?#]+)/)?.[1] ?? "";
  expect(id).not.toBe("");
  return id;
}

/** Retypes the current slide's title (its first stage frame). */
async function retitle(page: Page, text: string): Promise<void> {
  await stageFrames(page).first().dblclick();
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.type(text);
  await page.keyboard.press("Escape");
  await expect(stageFrames(page).first()).toContainText(text);
}

async function waitForSaved(page: Page): Promise<void> {
  await expect(page.getByTestId("deck-save-state")).toHaveText(SAVED, { timeout: 15_000 });
}

/** Writes a draft newer than the committed deck straight into the app's
 * store: the committed shell read back, its slide-1 title rewritten and its
 * last slide dropped — what an unclean close leaves behind. */
async function plantNewerDraft(page: Page, id: string, from: string, to: string): Promise<void> {
  await page.evaluate(
    async ([docId, before, after]) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open("docu-edit"); // the app already created it at its current version
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const committed = await new Promise<{ model: { slides: unknown[] } }>((resolve, reject) => {
        const request = db.transaction("documents", "readonly").objectStore("documents").get(docId);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const model = JSON.parse(JSON.stringify(committed.model).split(before).join(after));
      model.slides = model.slides.slice(0, -1);
      const draft = { id: `draft:${docId}`, docId, savedAt: Date.now(), model, caret: null, scroll: 0, slideIndex: 0 };
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction("documents", "readwrite");
        tx.objectStore("documents").put(draft);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      db.close();
    },
    [id, from, to] as const,
  );
}

test("a deck autosaves, recovers a crash's draft, and snapshots, names, previews and restores its versions", async ({ page }) => {
  const id = await openFixture(page);
  const thumbs = page.getByTestId("filmstrip-slide");
  await expect(thumbs).toHaveCount(4);
  await expect(stageFrames(page).first()).toContainText("Quarterly Review");

  // Edit a title; the header reads Saving… then All changes saved.
  await retitle(page, "Edited title");
  await waitForSaved(page);
  await expect(page.getByTestId("deck-save-state")).toHaveAttribute("data-save-state", "saved");

  // Reload: the edit is there, and nothing asks to be recovered.
  await page.reload();
  await expect(stageFrames(page).first()).toContainText("Edited title");
  await expect(page.getByTestId("recovery-bar")).toBeHidden();
  await waitForSaved(page);

  // A crash leaves a draft newer than the committed deck. Plant one from the
  // library (closing the deck cleanly first), then reopen the deck.
  await page.goto("/");
  await expect(page.getByTestId("library")).toBeVisible();
  await plantNewerDraft(page, id, "Edited title", "Drafted title");
  await page.goto(`/slides/${id}`);
  const bar = page.getByTestId("recovery-bar");
  await expect(bar).toBeVisible();
  await expect(bar).toContainText("Recovered unsaved changes from");
  await expect(stageFrames(page).first()).toContainText("Drafted title");
  await expect(thumbs).toHaveCount(3);

  // Discard: the draft goes and the committed deck returns.
  await page.getByTestId("recovery-discard").click();
  await expect(bar).toBeHidden();
  await expect(thumbs).toHaveCount(4);
  await thumbs.first().click();
  await expect(stageFrames(page).first()).toContainText("Edited title");

  // Edit, add a slide, and download: a commit point, so a snapshot.
  await retitle(page, "Snapshot title");
  await page.getByTestId("presentation-add-slide").click();
  await expect(thumbs).toHaveCount(5);
  const download = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  await download;

  // The history panel lists it (the only five-slide version); name it.
  await page.getByTestId("slides-history-toggle").click();
  const panel = page.getByTestId("slides-history-panel");
  await expect(panel).toBeVisible();
  const snapshot = panel.getByTestId("history-version").filter({ hasText: "5 slides" });
  await expect(snapshot).toHaveCount(1);
  await snapshot.getByTestId("history-name-input").fill("Before rewrite");
  await snapshot.getByTestId("history-name-submit").click();
  const named = panel.getByTestId("history-version").filter({ hasText: "Before rewrite" });
  await expect(named.getByTestId("history-name")).toHaveText("Before rewrite");

  // Move on: retitle slide 1 again.
  await thumbs.first().click();
  await retitle(page, "Later title");

  // Preview the named version: read-only, on the panel's own stage.
  await named.getByTestId("history-preview").click();
  const preview = panel.getByTestId("history-preview-banner");
  await expect(preview).toBeVisible();
  await expect(preview).toContainText("Previewing Before rewrite");
  await expect(preview.getByTestId("slide-stage")).toContainText("Snapshot title");
  await expect(stageFrames(page).first()).toContainText("Later title");

  // Exit preview: the live deck is exactly as it was.
  await panel.getByTestId("history-exit-preview").click();
  await expect(preview).toBeHidden();
  await expect(stageFrames(page).first()).toContainText("Later title");
  await expect(thumbs).toHaveCount(5);

  // Restore it: the stage shows the restored slides, and the panel lists the
  // restore's own two snapshots, newest first, above the named version.
  await named.getByTestId("history-restore").click();
  await expect(stageFrames(page).first()).toContainText("Snapshot title");
  await expect(thumbs).toHaveCount(5);
  const versions = panel.getByTestId("history-version");
  await expect(versions.nth(0).getByTestId("history-trigger")).toHaveText("After restore");
  await expect(versions.nth(1).getByTestId("history-trigger")).toHaveText("Before restore");
  await expect(versions.nth(2).getByTestId("history-name")).toHaveText("Before rewrite");
});
