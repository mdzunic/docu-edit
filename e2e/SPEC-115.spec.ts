import { fileURLToPath } from "node:url";
import { test, expect, type Page } from "@playwright/test";

// SPEC-115 — design a presentation from a description (experimental), in a
// real browser. Per SPEC-113 no test here needs a running model: the
// OpenAI-compatible endpoint is stubbed with page.route.
//
// 1. With the deck-designer flag on and a canned valid reply: New ›
//    "Presentation from a description…" → a brief → the consent dialog shows
//    the payload → Send → the plan lists 5 slides → slide 3's title edited,
//    slide 4 deleted → Build → /slides/:id with 4 slides, slide 3 carrying
//    the edited title; the library holds one new deck and the source
//    document is unchanged.
// 2. With the stub answering invalid JSON: exactly one repair request, then
//    the error and the raw text, and no new library entry.
// 3. With the provider off: "Outline this document" over a .docx fixture
//    builds a deck with no network request at all.
// 4. Redesign… on a 3-slide deck with a reply naming only source 2 then
//    source 0: the plan lists 3, 1, 2 with no add or delete control; Apply
//    redesign reorders the deck with every title unchanged; Undo restores
//    1, 2, 3.

const SOURCE_FIXTURE = fileURLToPath(new URL("./fixtures/word-toc.docx", import.meta.url));
const DECK_FIXTURE = fileURLToPath(new URL("./fixtures/spec104-hidden.pptx", import.meta.url));

const STUB_ORIGIN = "http://localhost:11434";
const STUB_BASE_URL = `${STUB_ORIGIN}/v1`;
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "*",
};

const PLAN = {
  title: "Rotterdam migration",
  theme: "Ion",
  size: "16:9",
  slides: [
    { layout: "title", title: "Rotterdam migration", bullets: ["Board update, October"] },
    { layout: "titleAndBody", title: "Where we are", bullets: ["Phase one done", "Phase two underway"] },
    { layout: "titleAndBody", title: "Risks", bullets: ["Vendor lead times"] },
    { layout: "sectionHeader", title: "Costs", bullets: [] },
    { layout: "titleOnly", title: "Questions?", bullets: [], note: "Leave ten minutes." },
  ],
};

/** A chat-completions reply as a server-sent event stream, cut into small
 * deltas the way a slow local model sends them. */
function eventStream(content: string): string {
  const pieces = content.match(/[\s\S]{1,24}/g) ?? [];
  return (
    pieces
      .map((piece) => `data: ${JSON.stringify({ model: "stub-model", choices: [{ index: 0, delta: { content: piece } }] })}\n\n`)
      .join("") + "data: [DONE]\n\n"
  );
}

/** Stubs the model endpoint: the n-th POST (from 0) is answered with
 * `reply(n)`. Returns the bodies of the POSTs that reached it. */
async function stubModel(page: Page, reply: (n: number) => string): Promise<string[]> {
  const posts: string[] = [];
  await page.route(`${STUB_ORIGIN}/**`, async (route) => {
    const request = route.request();
    if (request.method() === "OPTIONS") {
      await route.fulfill({ status: 204, headers: CORS });
      return;
    }
    const n = posts.length;
    posts.push(request.postData() ?? "");
    await route.fulfill({ status: 200, contentType: "text/event-stream", headers: CORS, body: eventStream(reply(n)) });
  });
  return posts;
}

/** Labs: the deck-designer experiment on, and — unless `providerOff` — a
 * local provider pointed at the stub. */
async function enableDesigner(page: Page, providerOff = false): Promise<void> {
  await page.getByTestId("library-preferences").click();
  await expect(page.getByTestId("preferences-dialog")).toBeVisible();
  await page.getByTestId("prefs-labs").scrollIntoViewIfNeeded();
  await page.getByTestId("experiment-deck-designer").check();
  await expect(page.getByTestId("ai-provider")).toBeVisible();
  if (!providerOff) {
    await page.getByTestId("ai-mode").selectOption("local");
    for (const [testid, value] of [
      ["ai-base-url", STUB_BASE_URL],
      ["ai-model", "test-model"],
    ]) {
      await page.getByTestId(testid).fill(value);
      await page.getByTestId(testid).press("Tab");
    }
    await expect(page.getByTestId("ai-origin")).toContainText(STUB_ORIGIN);
  }
  await page.getByTestId("preferences-close").click();
}

/** Opens the .docx source through the library and comes back, returning the
 * document's text as the editor shows it. */
async function addSourceDocument(page: Page): Promise<string> {
  await page.getByTestId("open-document-input").setInputFiles(SOURCE_FIXTURE);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByRole("heading", { name: "Background", level: 2 })).toBeVisible();
  const text = (await page.getByTestId("document-content").innerText()).trim();
  await page.goto("/");
  await expect(page.getByTestId("library-row")).toHaveCount(1);
  return text;
}

async function openDesigner(page: Page): Promise<void> {
  const entry = page.getByTestId("new-deck-from-description");
  await expect(entry).toBeVisible();
  await expect(entry.locator("xpath=..").getByTestId("experiment-badge")).toBeVisible();
  await entry.click();
  await expect(page.getByTestId("deck-designer")).toBeVisible();
  await expect(page.getByTestId("designer-step-brief")).toHaveAttribute("aria-current", "step");
}

async function filmstripTitles(page: Page): Promise<string[]> {
  const labels = await page.getByTestId("filmstrip-slide").evaluateAll((slides) => slides.map((slide) => slide.getAttribute("aria-label") ?? ""));
  return labels.map((label) => label.replace(/^Slide \d+ of \d+ — /, ""));
}

test("a brief becomes a reviewed, edited plan and then a new deck; the source is untouched (AC-82, AC-83)", async ({ page }) => {
  const posts = await stubModel(page, () => JSON.stringify(PLAN));
  await page.goto("/");
  await expect(page.getByTestId("new-document").first()).toBeVisible();
  // Off: no entry point renders.
  await expect(page.getByTestId("new-deck-from-description")).toHaveCount(0);
  const sourceText = await addSourceDocument(page);
  await expect(page.getByTestId("new-deck-from-description")).toHaveCount(0);
  await enableDesigner(page);

  await openDesigner(page);
  await page.getByTestId("designer-brief").fill("A board update on the Rotterdam migration, restrained.");
  await expect(page.getByTestId("designer-length")).toBeVisible();
  await page.getByTestId("designer-source").selectOption({ label: "word-toc (document)" });
  await expect(page.getByTestId("designer-bytes")).toContainText(/What will be sent: [\d,]+ bytes/);
  await page.getByTestId("designer-send").click();

  // SPEC-113's consent dialog carries the exact request body.
  const consent = page.getByTestId("ai-consent");
  await expect(consent).toBeVisible();
  await expect(page.getByTestId("ai-consent-purpose")).toHaveText("Design a deck");
  const payload = JSON.parse((await page.getByTestId("ai-consent-payload").textContent()) ?? "{}");
  const user = payload.messages.find((message: { role: string }) => message.role === "user").content as string;
  expect(user).toContain("Brief: A board update on the Rotterdam migration, restrained.");
  expect(user).toContain("Source document “word-toc”");
  expect(user).toContain("## Background\nMore detail.");
  expect(payload.stream).toBe(true);
  expect(posts).toHaveLength(0);
  await page.getByTestId("ai-consent-once").click();

  // The plan step: 5 slides, editable.
  const rows = page.getByTestId("designer-plan-row");
  await expect(rows).toHaveCount(5);
  expect(posts).toHaveLength(1);
  expect(posts[0]).toBe(JSON.stringify(payload, null, 2));
  await expect(page.getByTestId("designer-step-plan")).toHaveAttribute("aria-current", "step");
  await rows.nth(2).getByTestId("designer-row-title").fill("Risks we accept");
  await rows.nth(3).getByTestId("designer-row-delete").click();
  await expect(rows).toHaveCount(4);
  await page.getByTestId("designer-build").click();

  // /slides/:id opens with 4 slides, the first selected, slide 3 edited.
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(4);
  await expect.poll(() => filmstripTitles(page)).toEqual(["Rotterdam migration", "Where we are", "Risks we accept", "Questions?"]);
  await expect(page.getByTestId("filmstrip-slide").first()).toHaveAttribute("aria-selected", "true");
  expect(posts).toHaveLength(1);

  // The library: one new deck beside the source document, which is unchanged.
  await page.goto("/");
  await expect(page.getByTestId("library-row")).toHaveCount(2);
  await expect(page.getByTestId("row-title")).toHaveText(["Rotterdam migration", "word-toc"]);
  await page.getByTestId("library-row").filter({ hasText: "word-toc" }).dblclick();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByRole("heading", { name: "Background", level: 2 })).toBeVisible();
  expect((await page.getByTestId("document-content").innerText()).trim()).toBe(sourceText);
});

test("an invalid reply gets exactly one repair, then the error and the raw text — and nothing is stored (AC-84)", async ({ page }) => {
  const posts = await stubModel(page, (n) => (n === 0 ? "Sure! Here is your deck: slides about Rotterdam." : "{\"slides\": [oops"));
  await page.goto("/");
  await addSourceDocument(page);
  await enableDesigner(page);
  await openDesigner(page);
  await page.getByTestId("designer-brief").fill("A board update");
  await page.getByTestId("designer-send").click();
  await expect(page.getByTestId("ai-consent")).toBeVisible();
  await page.getByTestId("ai-consent-once").click();

  // The repair is a request of its own, shown before it goes: the reply
  // sent back with the validator's message.
  const repairPayload = page.getByTestId("ai-consent-payload");
  await expect(repairPayload).toContainText("Your previous reply was:");
  const repair = JSON.parse((await repairPayload.textContent()) ?? "{}");
  const repairUser = repair.messages.find((message: { role: string }) => message.role === "user").content as string;
  expect(repairUser).toContain("Your previous reply was:\n<<<\nSure! Here is your deck: slides about Rotterdam.\n>>>");
  expect(repairUser).toContain("It was rejected: reply: not valid JSON");
  await page.getByTestId("ai-consent-once").click();

  await expect(page.getByTestId("designer-error")).toContainText("reply: not valid JSON");
  const raw = page.getByTestId("designer-raw");
  await raw.locator("summary").click();
  await expect(raw).toContainText("{\"slides\": [oops");
  await expect(page.getByTestId("designer-try-again")).toBeVisible();
  await expect(page.getByTestId("designer-plan-row")).toHaveCount(0);
  await expect(page.getByTestId("designer-build")).toBeDisabled();
  expect(posts).toHaveLength(2);

  await page.getByTestId("designer-cancel").click();
  await expect(page.getByTestId("deck-designer")).toHaveCount(0);
  expect(posts).toHaveLength(2);
  await expect(page.getByTestId("library-row")).toHaveCount(1);
  await expect(page.getByTestId("row-title")).toHaveText(["word-toc"]);
});

test("with the provider off, Outline this document builds a deck from a .docx with no network request at all (AC-85)", async ({
  page,
  baseURL,
}) => {
  const appOrigin = new URL(baseURL ?? "http://localhost:5173").origin;
  const scripted: string[] = [];
  page.on("request", (request) => {
    const origin = new URL(request.url()).origin;
    if (origin !== appOrigin && ["fetch", "xhr", "eventsource", "websocket"].includes(request.resourceType())) {
      scripted.push(`${request.resourceType()} ${request.url()}`);
    }
  });
  await page.route(`${STUB_ORIGIN}/**`, async (route) => {
    scripted.push(`model ${route.request().url()}`);
    await route.abort();
  });

  await page.goto("/");
  await addSourceDocument(page);
  await enableDesigner(page, true);
  await openDesigner(page);
  await expect(page.getByTestId("designer-provider-off")).toBeVisible();
  await expect(page.getByTestId("designer-send")).toHaveCount(0);
  await expect(page.getByTestId("designer-outline")).toBeDisabled();
  await page.getByTestId("designer-source").selectOption({ label: "word-toc (document)" });
  await page.getByTestId("designer-outline").click();

  const rows = page.getByTestId("designer-plan-row");
  await expect(rows).toHaveCount(3);
  await expect(rows.nth(1).getByTestId("designer-row-title")).toHaveValue("Introduction");
  await expect(rows.nth(1).getByTestId("designer-row-layout")).toHaveValue("sectionHeader");
  await expect(rows.nth(2).getByTestId("designer-row-bullets")).toHaveValue("More detail.");
  await page.getByTestId("designer-build").click();

  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect.poll(() => filmstripTitles(page)).toEqual(["word-toc", "Introduction", "Background"]);
  expect(scripted).toEqual([]);
  // And the gateway itself was never asked: its session log is empty.
  await page.getByRole("menubar", { name: "Presentation menu" }).getByRole("menuitem", { name: "File", exact: true }).click();
  await page.getByTestId("slides-menu-preferences").click();
  await page.getByTestId("prefs-labs").scrollIntoViewIfNeeded();
  await expect(page.getByTestId("ai-activity-empty")).toBeVisible();
});

test("Redesign… reorders a 3-slide deck as one undoable step and never touches its text (AC-88)", async ({ page }) => {
  const posts = await stubModel(page, () => JSON.stringify({ slides: [{ source: 2 }, { source: 0 }] }));
  await page.goto("/");
  await enableDesigner(page);
  await page.getByTestId("open-document-input").setInputFiles(DECK_FIXTURE);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(3);
  const before = await filmstripTitles(page);
  expect(before.map((label) => label.replace(/ \(hidden\)$/, ""))).toEqual(["Welcome", "Backup data", "Thank you"]);

  const menu = page.getByRole("menubar", { name: "Presentation menu" });
  await menu.getByRole("menuitem", { name: "Slide", exact: true }).click();
  const item = page.getByTestId("deck-redesign");
  await expect(item).toBeEnabled();
  await expect(item.locator("xpath=..").getByTestId("experiment-badge")).toBeVisible();
  await item.click();
  await expect(page.getByTestId("deck-designer")).toBeVisible();

  // Redesign's brief: the prompt and audience/tone — no length, no source.
  await expect(page.getByTestId("designer-audience")).toBeVisible();
  await expect(page.getByTestId("designer-tone")).toBeVisible();
  await expect(page.getByTestId("designer-length")).toHaveCount(0);
  await expect(page.getByTestId("designer-source")).toHaveCount(0);
  await page.getByTestId("designer-brief").fill("Close on the thank-you slide first");
  await page.getByTestId("designer-send").click();
  await expect(page.getByTestId("ai-consent-purpose")).toHaveText("Redesign a deck");
  const payload = (await page.getByTestId("ai-consent-payload").textContent()) ?? "";
  expect(payload).not.toContain("Slide 2 body");
  await page.getByTestId("ai-consent-once").click();

  const rows = page.getByTestId("designer-redesign-row");
  await expect(rows).toHaveCount(3);
  await expect(rows.getByTestId("designer-row-number")).toHaveText(["Slide 3", "Slide 1", "Slide 2"]);
  await expect(page.getByTestId("designer-add-slide")).toHaveCount(0);
  await expect(page.getByTestId("designer-row-delete")).toHaveCount(0);
  await expect(page.getByTestId("deck-designer").locator("input[type=text], textarea")).toHaveCount(0);
  await expect(page.getByTestId("designer-build")).toHaveCount(0);
  await page.getByTestId("designer-apply-redesign").click();
  await expect(page.getByTestId("deck-designer")).toHaveCount(0);
  expect(posts).toHaveLength(1);

  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(3);
  await expect.poll(() => filmstripTitles(page)).toEqual([before[2], before[0], before[1]]);

  await menu.getByRole("menuitem", { name: "Edit", exact: true }).click();
  await page.getByTestId("slides-menu-undo").click();
  await expect.poll(() => filmstripTitles(page)).toEqual(before);
});
