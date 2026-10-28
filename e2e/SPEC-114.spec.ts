import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { test, expect, type Page } from "@playwright/test";

// SPEC-114 — search inside the library (experimental), in a real browser:
// enable the flag in Labs → a library holding a .docx (word-toc.docx), a
// .pptx (mini-deck.pptx), a text PDF (spec111-outlines.pdf: Helvetica,
// WinAnsi), spec107-photo.jpg and the scanned fixture (spec114-scanned.pdf:
// an image XObject and no text operators) → switch scope to Everything →
// wait for the progress line to finish → search a phrase in the .docx body →
// one result with a snippet → Enter opens /doc/:id with the find bar carrying
// the query and the match selected → a word on a slide → the deck result
// opens /slides/:id at that slide → a PDF's word opens /media/:id with the
// query shown above the viewer and #search= on its URL → the image and the
// scanned fixture are under "Not searched" with their reasons. Nothing
// leaves the browser: no request reaches any origin but the app's own and
// its shell's web fonts, and no script reaches any other origin at all.

const fixture = (name: string): string => fileURLToPath(new URL(`./fixtures/${name}`, import.meta.url));

/** The web-font hosts the shell's own index.html links (and its CSP names
 * in style-src/font-src) — part of the shipped shell, not of any feature
 * (SPEC-113's e2e reads them the same way). */
function shellFontOrigins(): string[] {
  const html = readFileSync(new URL("../index.html", import.meta.url), "utf8");
  const csp = /http-equiv="Content-Security-Policy"\s+content="([\s\S]*?)"/.exec(html)?.[1] ?? "";
  const hosts = csp
    .split(";")
    .map((directive) => directive.trim().split(/\s+/))
    .filter(([name]) => name === "style-src" || name === "font-src")
    .flatMap(([, ...sources]) => sources.filter((source) => source.startsWith("https://")));
  return [...new Set(hosts)];
}

/** Opens a file through the library's picker, then returns to the library. */
async function addToLibrary(page: Page, name: string, lands: RegExp): Promise<void> {
  await page.getByTestId("open-document-input").setInputFiles(fixture(name));
  await expect(page).toHaveURL(lands);
  await page.goto("/");
  await expect(page.getByTestId("library")).toBeVisible();
}

async function enableLibrarySearch(page: Page): Promise<void> {
  await page.getByTestId("library-preferences").click();
  await expect(page.getByTestId("preferences-dialog")).toBeVisible();
  await page.getByTestId("prefs-labs").scrollIntoViewIfNeeded();
  await page.getByTestId("experiment-library-search").check();
  await expect(page.getByTestId("prefs-library-search")).toBeVisible();
  await page.getByTestId("preferences-close").click();
  await expect(page.getByTestId("preferences-dialog")).toBeHidden();
}

test("full-text search: index, find a phrase, open at the match, and say what was not searched", async ({ page, baseURL }) => {
  const appOrigin = new URL(baseURL ?? "http://localhost:5173").origin;
  const allowed = new Set([appOrigin, ...shellFontOrigins()]);
  const foreign: string[] = [];
  const scripted: string[] = [];
  // Every request to another origin is recorded, then refused.
  await page.route(
    (url) => url.origin !== appOrigin,
    async (route) => {
      const request = route.request();
      if (["fetch", "xhr", "eventsource", "websocket"].includes(request.resourceType())) {
        scripted.push(`${request.resourceType()} ${request.url()}`);
      }
      if (!allowed.has(new URL(request.url()).origin)) {
        foreign.push(`${request.resourceType()} ${request.url()}`);
      }
      await route.abort();
    },
  );

  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  // The flag is off: the library is SPEC-067's — no scope control.
  await expect(page.getByTestId("library-search")).toBeVisible({ timeout: 30_000 }).catch(() => undefined);
  await expect(page.getByTestId("search-scope")).toHaveCount(0);

  await addToLibrary(page, "word-toc.docx", /\/doc\/.+/);
  await addToLibrary(page, "mini-deck.pptx", /\/slides\/.+/);
  await addToLibrary(page, "spec111-outlines.pdf", /\/media\/.+/);
  await addToLibrary(page, "spec107-photo.jpg", /\/media\/.+/);
  await addToLibrary(page, "spec114-scanned.pdf", /\/media\/.+/);
  await expect(page.getByTestId("library-row")).toHaveCount(5);

  await enableLibrarySearch(page);
  const scope = page.getByTestId("search-scope");
  await expect(scope).toHaveValue("titles");
  await scope.selectOption("everything");

  // The index builds in the background; the library stays usable meanwhile.
  const progress = page.getByTestId("search-index-progress");
  await expect(progress).toHaveAttribute("role", "status");
  await expect(progress).toHaveText("Indexed 5 of 5", { timeout: 30_000 });
  await expect(progress).toHaveAttribute("data-running", "false");

  // A phrase from the .docx body: one result, with a snippet marking it.
  const field = page.getByTestId("library-search");
  await field.fill('"opening paragraph"');
  const results = page.getByTestId("search-result");
  await expect(results).toHaveCount(1);
  await expect(results.first().getByTestId("search-result-kind")).toHaveText("DOCX");
  await expect(results.first().getByTestId("search-snippet")).toHaveCount(1);
  await expect(results.first().getByTestId("snippet-match")).toHaveText("Opening paragraph");
  // The results stand in for the rows.
  await expect(page.getByTestId("library-row")).toHaveCount(0);

  // Enter opens the document AT the match: the find bar carries the query
  // and the match is the current one.
  await field.press("Enter");
  await expect(page).toHaveURL(/\/doc\/[^?]+\?find=opening(%20|\+)paragraph$/);
  await expect(page.getByTestId("find-bar")).toBeVisible();
  await expect(page.getByTestId("find-query")).toHaveValue("opening paragraph");
  await expect(page.getByTestId("find-count")).toHaveText("1 of 1");
  await expect(page.locator('[data-testid="find-highlights"] .find-highlight[data-current="true"]')).toHaveCount(1);

  // Back in the library, still in Everything: a word on a slide.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(page.getByTestId("search-scope")).toHaveValue("everything");
  await field.fill("berlin");
  await expect(results).toHaveCount(1);
  await expect(results.first().getByTestId("search-result-kind")).toHaveText("PPTX");
  await expect(results.first().getByTestId("snippet-match")).toHaveText("Berlin");
  // The rows' keyboard conventions: ArrowDown into the list, Enter opens.
  await field.press("ArrowDown");
  await expect(results.first()).toBeFocused();
  await page.keyboard.press("Enter");
  await expect(page).toHaveURL(/\/slides\/[^?]+\?find=berlin$/);
  await expect(page.getByTestId("find-bar")).toBeVisible();
  await expect(page.getByTestId("find-query")).toHaveValue("berlin");
  await expect(page.getByTestId("find-count")).toHaveText("1 of 1");
  await expect(page.getByTestId("filmstrip-slide").nth(1)).toHaveAttribute("aria-selected", "true");

  // A PDF's word opens the viewer with the query above it, and #search=.
  // (In-app navigation: the scope is this session's, Titles on a reload.)
  await page.getByTestId("presentation-close").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(page.getByTestId("search-scope")).toHaveValue("everything");
  await field.fill("chapter");
  await expect(results).toHaveCount(1);
  await expect(results.first().getByTestId("search-result-kind")).toHaveText("PDF");
  await results.first().click();
  await expect(page).toHaveURL(/\/media\/[^?]+\?find=chapter$/);
  await expect(page.getByTestId("media-find-query")).toHaveValue("chapter");
  await expect(page.getByTestId("media-find-copy")).toBeVisible();
  await expect(page.getByTestId("media-find-note")).toContainText("only some browsers");
  await expect(page.locator('object[type="application/pdf"]')).toHaveAttribute("data", /^blob:.+#search=chapter$/);

  // What could not be read is listed, with why — never silently absent.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(page.getByTestId("search-scope")).toHaveValue("everything");
  const unsearched = page.getByTestId("search-unsearched");
  await expect(unsearched).toBeVisible();
  await unsearched.locator("summary").click();
  const item = (title: string) => unsearched.getByTestId("search-unsearched-item").filter({ hasText: title });
  await expect(item("spec107-photo").getByTestId("search-unsearched-reason")).toHaveText("image, no text layer");
  await expect(item("spec114-scanned").getByTestId("search-unsearched-reason")).toHaveText("scanned image, no text layer");
  await expect(unsearched.getByTestId("search-unsearched-item")).toHaveCount(2);

  // A reload starts the session over, in the default Titles scope.
  await page.reload();
  await expect(page.getByTestId("search-scope")).toHaveValue("titles");

  expect(foreign).toEqual([]);
  expect(scripted).toEqual([]);
});
