import { readFileSync } from "node:fs";
import { test, expect, type Page } from "@playwright/test";

// SPEC-113 — experimental features and the model gateway, in a real browser.
//
// 1. With Labs off (every experiment at its default), loading the app and
//    using the library sends nothing to any origin but the app's own, and the
//    gateway's code is never even requested from the server.
// 2. With deck-designer on and the endpoint stubbed through page.route (no
//    real model is ever called), the experimental surface carries its badge;
//    "Test connection" — the one request no feature asks for — shows the
//    consent dialog with the exact payload; Cancel sends nothing; "Always"
//    sends exactly the bytes shown, once, and is not asked again; the
//    activity log holds one entry of byte counts and no text, and a reload
//    empties it.
// 3. In cloud mode the key is masked, sent only as a bearer token to the
//    pinned origin, and discarded — with a line saying why — when the base
//    URL moves to another origin.
// 4. The CSP's plain-http loopback allowance is the gateway's alone: a
//    Markdown image or an /open?url= link that redirects onto this computer
//    is not followed there.

/** The app's own origins: the one it is served from, plus the two web-font
 * hosts its own index.html links for the design themes' faces (and its CSP
 * names in style-src/font-src, pinned by src/cspPolicy.spec.ts). Those two
 * are part of the shipped shell, not of any feature — and test 1 still
 * asserts that no script (fetch, XHR, EventSource, WebSocket) reaches them
 * or anything else. */
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

const STUB_ORIGIN = "http://localhost:11434";
const STUB_BASE_URL = `${STUB_ORIGIN}/v1`;
const STUB_REPLY = "stub-reply-7f3a";
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "*",
};

interface SentRequest {
  method: string;
  url: string;
  body: string | null;
  authorization: string | undefined;
}

/** Stubs an OpenAI-compatible endpoint at `origin`, recording every request
 * that reaches it (preflights included). */
async function stubEndpoint(page: Page, origin: string, model = "stub-model"): Promise<SentRequest[]> {
  const sent: SentRequest[] = [];
  await page.route(`${origin}/**`, async (route) => {
    const request = route.request();
    sent.push({
      method: request.method(),
      url: request.url(),
      body: request.postData(),
      authorization: request.headers().authorization,
    });
    if (request.method() === "OPTIONS") {
      await route.fulfill({ status: 204, headers: CORS });
      return;
    }
    await route.fulfill({
      status: 200,
      contentType: "application/json",
      headers: CORS,
      body: JSON.stringify({
        id: "stub-1",
        model,
        choices: [{ index: 0, message: { role: "assistant", content: STUB_REPLY } }],
      }),
    });
  });
  return sent;
}

async function openLabs(page: Page): Promise<void> {
  await page.getByTestId("library-preferences").click();
  await expect(page.getByTestId("preferences-dialog")).toBeVisible();
  await page.getByTestId("prefs-labs").scrollIntoViewIfNeeded();
}

/** Types into a provider field and commits it (the form saves on change). */
async function commit(page: Page, testid: string, value: string): Promise<void> {
  const field = page.getByTestId(testid);
  await field.fill(value);
  await field.press("Tab");
}

async function configureLocalEndpoint(page: Page): Promise<void> {
  await page.getByTestId("experiment-deck-designer").check();
  await expect(page.getByTestId("ai-provider")).toBeVisible();
  await page.getByTestId("ai-mode").selectOption("local");
  await commit(page, "ai-base-url", STUB_BASE_URL);
  await commit(page, "ai-model", "test-model");
  await expect(page.getByTestId("ai-origin")).toHaveText(`Requests go to ${STUB_ORIGIN}, on this computer.`);
}

const posts = (sent: SentRequest[]) => sent.filter((request) => request.method === "POST");

test("with Labs off, the app and its library reach no origin but the app's own and its shell's web fonts (AC-57)", async ({
  page,
  baseURL,
}) => {
  const appOrigin = new URL(baseURL ?? "http://localhost:5173").origin;
  const allowed = new Set([appOrigin, ...shellFontOrigins()]);
  const foreign: string[] = [];
  const scripted: string[] = [];
  const gatewayCode: string[] = [];

  // Every request to another origin is recorded, then refused.
  await page.route(
    (url) => url.origin !== appOrigin,
    async (route) => {
      const request = route.request();
      const origin = new URL(request.url()).origin;
      if (["fetch", "xhr", "eventsource", "websocket"].includes(request.resourceType())) {
        scripted.push(`${request.resourceType()} ${request.url()}`);
      }
      if (!allowed.has(origin)) {
        foreign.push(`${request.resourceType()} ${request.url()}`);
      }
      await route.abort();
    },
  );
  page.on("websocket", (socket) => {
    // Vite's own dev-server HMR socket is the app's origin on ws:.
    if (new URL(socket.url()).host !== new URL(appOrigin).host) {
      foreign.push(`websocket ${socket.url()}`);
    }
  });
  page.on("request", (request) => {
    if (/\/lib\/ai\/modelGateway|\/ModelProviderPanel\.vue|\/ModelConsentDialog\.vue/.test(request.url())) {
      gatewayCode.push(request.url());
    }
  });

  await page.goto("/");
  await expect(page.getByTestId("new-document").first()).toBeVisible();

  // Exercise the library: create a document, write in it, come back, search.
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await page.getByTestId("document-content").click();
  await page.keyboard.type("Rotterdam figures");
  await expect(page.getByTestId("document-content")).toContainText("Rotterdam figures");
  await page.goto("/");
  await expect(page.getByTestId("library-row").first()).toBeVisible();
  await page.getByTestId("library-search").fill("Untitled");
  await page.getByTestId("library-search").fill("");

  // Labs is there, every experiment off, nothing experimental rendered.
  await openLabs(page);
  await expect(page.getByTestId("experiment-library-search")).not.toBeChecked();
  await expect(page.getByTestId("experiment-deck-designer")).not.toBeChecked();
  await expect(page.getByTestId("ai-provider")).toHaveCount(0);
  await expect(page.getByTestId("experiment-badge")).toHaveCount(0);
  await page.getByTestId("preferences-close").click();
  await page.reload();
  await expect(page.getByTestId("library-row").first()).toBeVisible();

  expect(foreign).toEqual([]);
  expect(scripted).toEqual([]);
  expect(gatewayCode).toEqual([]);

  // The trap is live: a deliberate cross-origin request is caught by it.
  await page.evaluate(() => fetch("https://probe.example.test/x").catch(() => undefined));
  expect(foreign).toEqual(["fetch https://probe.example.test/x"]);
});

test("an enabled experiment: badge, consent with the exact payload, Cancel, Always, the activity log (AC-58)", async ({
  page,
}) => {
  const sent = await stubEndpoint(page, STUB_ORIGIN);
  const gatewayCode: string[] = [];
  page.on("request", (request) => {
    if (/\/lib\/ai\/modelGateway/.test(request.url())) {
      gatewayCode.push(request.url());
    }
  });
  await page.goto("/");
  await openLabs(page);

  // Off: absent, its code not loaded. On: the experimental surface appears,
  // badged — and only now is the gateway's code fetched.
  await expect(page.getByTestId("experiment-badge")).toHaveCount(0);
  expect(gatewayCode).toEqual([]);
  await configureLocalEndpoint(page);
  expect(gatewayCode.length).toBeGreaterThan(0);
  const badge = page.getByTestId("ai-provider").getByTestId("experiment-badge");
  await expect(badge).toBeVisible();
  await expect(badge).toHaveText("Experimental");
  await badge.hover();
  await expect(page.getByTestId("experiment-badge-labs")).toBeVisible();

  // The request shows the consent dialog with the exact payload — and
  // Cancel sends nothing at all, not even a preflight.
  await page.getByTestId("ai-test").click();
  const consent = page.getByTestId("ai-consent");
  await expect(consent).toBeVisible();
  await expect(page.getByTestId("ai-consent-destination")).toContainText("This stays on your computer");
  await expect(page.getByTestId("ai-consent-destination")).toContainText(STUB_ORIGIN);
  const cancelledPayload = (await page.getByTestId("ai-consent-payload").textContent()) ?? "";
  expect(JSON.parse(cancelledPayload)).toEqual({
    model: "test-model",
    messages: [{ role: "user", content: "Reply with the single word OK." }],
    max_tokens: 1,
    temperature: 0,
    stream: false,
  });
  await expect(page.getByTestId("ai-consent-bytes")).toHaveText(
    `Exactly this will be sent — ${Buffer.byteLength(cancelledPayload)} bytes:`,
  );
  await page.getByTestId("ai-consent-cancel").click();
  await expect(consent).toHaveCount(0);
  await expect(page.getByTestId("ai-test-result")).toContainText("nothing was sent");
  expect(sent).toEqual([]);
  await expect(page.getByTestId("ai-activity-empty")).toBeVisible();

  // "Always": the bytes shown are the bytes sent, exactly once.
  await page.getByTestId("ai-test").click();
  await expect(consent).toBeVisible();
  const payload = (await page.getByTestId("ai-consent-payload").textContent()) ?? "";
  await page.getByTestId("ai-consent-always").click();
  await expect(consent).toHaveCount(0);
  await expect(page.getByTestId("ai-test-result")).toHaveText(/^Reachable — model stub-model, \d+ ms round trip\.$/);
  expect(posts(sent)).toHaveLength(1);
  expect(posts(sent)[0]).toMatchObject({ url: `${STUB_BASE_URL}/chat/completions`, body: payload });
  expect(posts(sent)[0].authorization).toBeUndefined();

  // The activity log: one entry, byte counts, no text either way.
  const entries = page.getByTestId("ai-activity-entry");
  await expect(entries).toHaveCount(1);
  await expect(entries.first()).toHaveAttribute("data-request-bytes", String(Buffer.byteLength(payload)));
  await expect(entries.first()).toHaveAttribute("data-outcome", "ok");
  expect(Number(await entries.first().getAttribute("data-response-bytes"))).toBeGreaterThan(0);
  const log = (await page.getByTestId("ai-activity").textContent()) ?? "";
  expect(log).toContain(`${Buffer.byteLength(payload)} B`);
  expect(log).not.toContain("Reply with the single word OK");
  expect(log).not.toContain(STUB_REPLY);

  // Not asked again: with the dialog never shown, a second request goes out.
  await expect(page.getByTestId("ai-consent-grant")).toContainText(STUB_ORIGIN);
  await page.getByTestId("ai-test").click();
  await expect.poll(() => posts(sent).length).toBe(2);
  await expect(consent).toHaveCount(0);
  await expect(entries).toHaveCount(2);

  // Session-only: a reload empties the log; the standing consent remains,
  // revocable from Labs.
  await page.reload();
  await openLabs(page);
  await expect(page.getByTestId("ai-activity-empty")).toBeVisible();
  await page.getByTestId("ai-consent-revoke").click();
  await expect(page.getByTestId("ai-consent-grant")).toHaveCount(0);
  await page.getByTestId("ai-test").click();
  await expect(consent).toBeVisible();
  await page.getByTestId("ai-consent-cancel").click();
  expect(posts(sent)).toHaveLength(2);
});

test("cloud mode: a masked key, sent only as a bearer token, discarded when the origin moves (AC-34, AC-39..43)", async ({
  page,
}) => {
  const first = await stubEndpoint(page, "https://models-a.example.test", "hosted-a");
  const second = await stubEndpoint(page, "https://models-b.example.test", "hosted-b");
  await page.goto("/");
  await openLabs(page);
  await page.getByTestId("experiment-deck-designer").check();
  await page.getByTestId("ai-mode").selectOption("local");
  await expect(page.getByTestId("ai-api-key")).toHaveCount(0);
  await page.getByTestId("ai-mode").selectOption("cloud");
  await commit(page, "ai-base-url", "https://models-a.example.test/v1");
  await commit(page, "ai-model", "hosted-model");
  await expect(page.getByTestId("ai-api-key")).toHaveAttribute("type", "password");
  await commit(page, "ai-api-key", "sk-e2e-secret");

  await page.getByTestId("ai-test").click();
  await expect(page.getByTestId("ai-consent-destination")).toHaveText(
    "Your text will be sent to https://models-a.example.test, which is not part of this product.",
  );
  await expect(page.getByTestId("ai-consent-payload")).not.toContainText("sk-e2e-secret");
  await page.getByTestId("ai-consent-once").click();
  await expect(page.getByTestId("ai-test-result")).toHaveText(/^Reachable — model hosted-a, \d+ ms round trip\.$/);
  expect(posts(first)).toHaveLength(1);
  expect(posts(first)[0].authorization).toBe("Bearer sk-e2e-secret");
  await expect(page.getByTestId("ai-activity")).not.toContainText("sk-e2e-secret");
  const prefs = await page.evaluate(() => localStorage.getItem("docu-edit:preferences"));
  expect(prefs).not.toContain("sk-e2e-secret");

  // Move the base URL to another origin: the key goes, visibly, and never
  // reaches the new host.
  await commit(page, "ai-base-url", "https://models-b.example.test/v1");
  await expect(page.getByTestId("ai-api-key")).toHaveValue("");
  await expect(page.getByTestId("ai-api-key-cleared")).toContainText("The API key was removed");
  await page.getByTestId("ai-test").click();
  await page.getByTestId("ai-consent-once").click();
  await expect(page.getByTestId("ai-test-result")).toHaveText(/^Reachable — model hosted-b/);
  expect(posts(second)).toHaveLength(1);
  expect(posts(second)[0].authorization).toBeUndefined();
});

test("a document's fetch follows no redirect onto this computer — the CSP's loopback allowance is the gateway's alone", async ({
  page,
}) => {
  // A service on this computer, which the CSP admits over plain http for a
  // local model server. Nothing a document or a link names may reach it,
  // directly or through a redirect.
  const localService = "http://127.0.0.1:8765";
  const reachedLocal: string[] = [];
  const reachedRedirector: string[] = [];
  page.on("request", (request) => {
    if (request.url().startsWith(localService)) {
      reachedLocal.push(request.url());
    }
  });
  await page.route(`${localService}/**`, (route) =>
    route.fulfill({ status: 200, headers: CORS, contentType: "application/octet-stream", body: "local secrets" }),
  );
  await page.route("https://redirector.example/**", async (route) => {
    reachedRedirector.push(route.request().url());
    await route.fulfill({ status: 302, headers: { ...CORS, location: `${localService}/api/whatever` } });
  });
  await page.route("https://docs.example/notes.md", (route) =>
    route.fulfill({
      status: 200,
      headers: CORS,
      contentType: "text/markdown",
      body: "# Field notes\n\n![chart](https://redirector.example/chart.png)\n",
    }),
  );

  // A Markdown image: the document opens and the image's redirect is not
  // followed. The first hop going out shows the trap is live.
  await page.goto("/open?url=" + encodeURIComponent("https://docs.example/notes.md"));
  await page.getByTestId("open-from-url-confirm-button").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("document-content")).toContainText("Field notes");
  expect(reachedRedirector).toEqual(["https://redirector.example/chart.png"]);
  expect(reachedLocal).toEqual([]);

  // A link itself: refused, saying why.
  await page.goto("/open?url=" + encodeURIComponent("https://redirector.example/report.docx"));
  await page.getByTestId("open-from-url-confirm-button").click();
  await expect(page.getByTestId("open-from-url-reason")).toContainText("redirects to another address");

  expect(reachedRedirector).toEqual(["https://redirector.example/chart.png", "https://redirector.example/report.docx"]);
  expect(reachedLocal).toEqual([]);
});
