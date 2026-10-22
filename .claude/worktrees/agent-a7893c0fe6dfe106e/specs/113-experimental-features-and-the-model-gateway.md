---
id: SPEC-113
title: Experimental features — the flag registry and the model gateway
status: draft
area: platform
complexity: normal
story_points: 5
depends_on: [SPEC-058, SPEC-066]
---
## Why
Two features want to exist that nothing in this product has needed before.
One searches inside every stored document (SPEC-114); the other asks a
language model to design a deck (SPEC-115). They are speculative, they will
change shape, and one of them wants to send the user's words to a server.

None of that is a reason not to build them. It is a reason to build the
thing that makes them safe to ship half-finished: a flag registry that keeps
an unfinished feature out of the product until someone asks for it, and — for
anything involving a model — a single, auditable place where a network
request can be made, with the user's explicit consent, to an endpoint the
user named.

The second half matters more than the first. Today this application reaches the network
in exactly four places, and every one of them is either its own bundled
fonts or a URL the user's own input named — a typed address, or an image a
Markdown file they opened refers to. Everything else happens in the browser,
and that is a property people choose this product for. Adding a model must not quietly
end it. **This file is the normative source for `EXPERIMENTS`, for
`ModelProviderConfig`, for `src/lib/ai/modelGateway.ts` as the only module
permitted to make an AI request, and for the consent and key-handling rules
every model-backed feature obeys.**

## Acceptance criteria
- [ ] `src/lib/experiments.ts` exports `EXPERIMENTS` — a frozen array of `{ id, title, summary, risk }` over a closed `ExperimentId` union — plus `isExperimentEnabled(id)`, `setExperimentEnabled(id, on)` and `onExperimentsChange(listener)`, persisted through the existing preferences store under one `experiments: Record<ExperimentId, boolean>` key with **every experiment defaulting to `false`**; an unknown id in stored preferences is ignored, not resurrected
- [ ] Preferences gains a "Labs" section (`prefs-labs`) listing every experiment with its title, summary, a risk line and a toggle (`experiment-<id>`), under a standing note that experimental features may change or be removed and are not covered by the product's round-trip guarantees; nothing else about SPEC-066 changes
- [ ] A disabled experiment is **absent**, not hidden: its controls do not render, its route (where it has one) redirects to `/` through a router guard, its modules are reached only through a `import()` behind the flag check so a disabled feature's code is not in the initial bundle, and a test asserts each gated entry point renders nothing and loads nothing when its flag is off
- [ ] Every enabled experimental surface carries `ExperimentBadge.vue` (`experiment-badge`, an "Experimental" chip with a tooltip linking to Labs), so a user always knows which part of the product is provisional
- [ ] `src/lib/ai/modelGateway.ts` is the **only** module in the repo that may make a request to a model endpoint, enforced by a new `src/lib/ai/boundary.spec.ts` asserting that no runtime module under `src/` references `fetch`, `EventSource`, `WebSocket` or `XMLHttpRequest` outside a checked-in allowlist of exactly five entries — the gateway, and the four that exist today, each with its reason recorded in the test: `src/lib/openDocumentFromUrl.ts` (a URL the user typed), `src/lib/md/reader.ts` (a remote image a Markdown document references, SPEC-050), and `src/views/document/useFonts.ts` + `usePdfExport.ts` (the app's own bundled font files under `import.meta.env.BASE_URL`, same-origin) — the same architecture-as-test pattern as `core/boundary.spec.ts`, and a list that may only shrink
- [ ] `ModelProviderConfig` is per the Reference: `mode` (`"off"` default, `"local"`, `"cloud"`), `baseUrl`, `model`, `timeoutMs` (default 60 000), `maxOutputTokens` (default 4 096) and `temperature` (default 0.2), stored in preferences — **except the API key**, which is never in the preferences object, never in a preferences export, and never written into any document
- [ ] The gateway speaks the OpenAI-compatible HTTP shape and nothing else: `POST {baseUrl}/chat/completions` and `GET {baseUrl}/models`, so a local Ollama (`http://localhost:11434/v1`), LM Studio (`http://localhost:1234/v1`), a llama.cpp or vLLM server, and any hosted endpoint the user names all work with no per-provider code; **no model is bundled, downloaded or run in-process**, and no provider is preconfigured, recommended or defaulted to
- [ ] Requests are user-initiated, bounded and abortable: `complete(request)` and `stream(request, onDelta)` each take an `AbortSignal`, enforce `timeoutMs` with their own timer, parse SSE incrementally for the streaming form, never retry automatically (a caller may retry once, explicitly — SPEC-115's repair pass is the only one specified), and are never called from a timer, a watcher, a background task or a route guard
- [ ] The destination is pinned: the gateway resolves `baseUrl` to an origin once, refuses any request URL outside it, passes `redirect: "error"` so a redirect cannot move a request — or a key — to another host, refuses a non-`https:` origin unless it is loopback (`localhost`, `127.0.0.1`, `[::1]`), and sends no cookies (`credentials: "omit"`)
- [ ] Consent is per origin and shown before the first request to it: `ModelConsentDialog.vue` (`ai-consent`) names the exact destination origin, says plainly whether it is on this machine ("This stays on your computer") or a third party ("Your text will be sent to <origin>, which is not part of this product"), shows the **exact** request body that will be sent in a scrollable, selectable region (`ai-consent-payload`) with its byte count, and offers `ai-consent-once`, `ai-consent-always` and `ai-consent-cancel`; consent is stored per origin, is revocable in Labs, and is cleared whenever the origin, the model or the feature changes
- [ ] The API key is handled as a credential: entered by the user in a masked field (`ai-api-key`) shown only in `cloud` mode, stored under its own `localStorage` key separate from preferences, never logged, never included in an error message or the activity log, sent only as `Authorization: Bearer` to the pinned origin, and **discarded automatically whenever `baseUrl`'s origin changes**, with the field cleared and a line explaining why — so a key can never follow a URL edit to a new host
- [ ] Failures are typed and specific, never a spinner that stops: `ModelNotConfiguredError`, `ModelConsentDeniedError`, `ModelUnreachableError` (with the likely cause spelled out for loopback — "nothing is listening on <origin>; is your local model running?"), `ModelHttpError` (status plus the first 500 characters of the body), `ModelTimeoutError` and `ModelOutputError`; each maps to a message that says what happened and what to change, and no failure leaves partial state anywhere
- [ ] Nothing is sent that the user did not put in front of the model, and nothing is kept that they cannot see: the gateway makes no telemetry or analytics request of any kind, sends no document content unless the calling feature's own spec says it does and the consent payload shows it, and keeps a session-only activity log (`ai-activity`, in memory, cleared on reload) of `{ time, origin, model, requestBytes, responseBytes, outcome }` — **never** prompt or response text — viewable and clearable from Labs
- [ ] A model's output is data, never an instruction and never markup: `stream`/`complete` return plain strings, callers render them as text (never `v-html`, never into a `<script>`, never as a URL to fetch), and a feature that wants structure asks for JSON and validates it with its own schema — the gateway parses nothing, executes nothing and follows no instruction found in a response
- [ ] "Test connection" (`ai-test`) in Labs does one trivial request and reports the outcome concretely — reachable / model id / round-trip milliseconds, or the typed error above — and is the one place a request happens without a feature asking for it (the consent dialog still applies)
- [ ] Tests: `src/lib/experiments.spec.ts` (defaults off, persistence, unknown ids, listener), `src/lib/ai/modelGateway.spec.ts` against a stubbed `fetch` — request shape, `Authorization` present only in cloud mode, timeout and abort, SSE chunk assembly including a split-mid-event chunk, every error type, the origin pin (an absolute URL outside the origin, a redirect, an `http:` non-loopback origin, all refused), and the key-clearing rule on origin change; `src/lib/ai/boundary.spec.ts`; e2e (`e2e/SPEC-113.spec.ts`): with Labs off, load the app, exercise the library and assert **zero** requests to any non-app origin via `page.route`; enable an experiment and set a stubbed endpoint → the badge appears → a feature request shows the consent dialog with the exact payload → Cancel sends nothing → "Always" sends once and does not ask again → the activity log shows one entry with byte counts and no text

## Out of scope
Bundling, downloading or running a model in the browser (WebGPU, WASM,
`transformers.js` — all would be dependencies `package.json` forbids, and one
of them is a several-hundred-megabyte download). Any provider-specific API:
Anthropic's, OpenAI's beyond the compatible chat shape, or a vendor SDK.
Tool use, function calling, embeddings, image or audio models, and multi-turn
conversation state — the gateway does one bounded request-response, and a
feature that needs more specifies it. Prompt storage, prompt libraries and
prompt sharing. Server-side proxying, rate limiting or key escrow — the key
is the user's and the request goes straight from their browser. Billing,
usage estimation or cost display. Evaluating or comparing model output
quality. A/B experiment infrastructure, remote flag delivery, or any flag
that is not a local user choice — Labs toggles are set by the person using
the app and by nobody else.

## Reference

### Dependencies restated
- `specs/066-preferences.md` — `src/lib/preferences.ts` with
  `PREFERENCES_STORAGE_KEY`, `Preferences`, `PREFERENCE_DEFAULTS`,
  `loadPreferences()`, `savePreference(key, value)`, `resetPreferences()` and
  `onPreferencesChange(listener)`, plus the preferences view's section
  layout. The experiments map and the provider config are two new keys in
  that object; the API key is deliberately **not**.
- `specs/058-…` — the dialog shell (`BaseDialog`), chips and control
  primitives the Labs section and the consent dialog are built from.
- The four `fetch` sites that exist today, which the boundary test allowlists
  **by name and reason** so the rule stays honest rather than becoming a lie
  the moment it is written: `src/lib/openDocumentFromUrl.ts` and
  `src/lib/md/reader.ts` reach arbitrary origins the *user's own input*
  names (a typed URL; an image `src` inside a Markdown file they opened),
  and `src/views/document/useFonts.ts` + `usePdfExport.ts` fetch the app's
  own bundled `fonts/` assets from its own origin. None of them is a model
  request, and none of them may grow into one.

### The registry
```ts
// src/lib/experiments.ts
export type ExperimentId = "library-search" | "deck-designer";   // grows one entry per experimental spec

export interface Experiment {
  id: ExperimentId;
  title: string;
  summary: string;    // one sentence, what it does
  risk: string;       // one sentence, what it costs or might break
}

export const EXPERIMENTS: readonly Experiment[];
export function isExperimentEnabled(id: ExperimentId): boolean;
export function setExperimentEnabled(id: ExperimentId, on: boolean): void;
export function onExperimentsChange(listener: (map: Record<ExperimentId, boolean>) => void): () => void;
```
The union is closed so a feature cannot invent a flag at a call site, and so
the Labs list and the guards cannot drift apart.

### The gateway
```ts
// src/lib/ai/modelGateway.ts
export type ProviderMode = "off" | "local" | "cloud";

export interface ModelProviderConfig {
  mode: ProviderMode;
  baseUrl: string;        // e.g. "http://localhost:11434/v1"
  model: string;          // e.g. "llama3.1:8b"
  timeoutMs: number;      // default 60_000
  maxOutputTokens: number;// default 4_096
  temperature: number;    // default 0.2
}

export interface ModelRequest {
  /** Why the request is being made — shown in the consent dialog and the log. */
  purpose: string;
  system: string;
  user: string;
  /** Ask for a JSON object back. The gateway sets the provider's JSON mode
   *  where the endpoint advertises it and otherwise says so in the prompt;
   *  it never parses or trusts the result — the caller validates. */
  json?: boolean;
  signal?: AbortSignal;
}

export interface ModelResult { text: string; model: string; elapsedMs: number }

export function getProviderConfig(): ModelProviderConfig;
export function complete(request: ModelRequest): Promise<ModelResult>;
export function stream(request: ModelRequest, onDelta: (chunk: string) => void): Promise<ModelResult>;
export function listModels(): Promise<string[]>;
export function testConnection(): Promise<{ model: string; elapsedMs: number }>;
export function activityLog(): readonly ModelActivity[];
export function clearActivityLog(): void;
```

### The rules, and why each one is a rule
- **One module.** Auditing "can this app send my document somewhere" must be a
  single file to read, and a test that keeps it that way. Hence the boundary
  spec, which is the only mechanism that survives a year of feature work.
- **Origin pin, no redirects, no cookies.** A `baseUrl` the user edits is an
  input; a redirect is an input the *server* controls. Following one with a
  bearer token attached hands the key to whoever the endpoint points at.
- **Key dies with the origin.** The commonest way a key leaks is a URL edit
  the user did not think of as a security event. Clearing on origin change
  makes it one, visibly.
- **Consent shows the payload, not a summary.** "This will send your document"
  is a claim; the actual bytes are a fact. A user who wants to check can.
- **The log holds counts, not text.** A log that holds prompts is a second
  copy of the user's private writing, in a place they did not choose.
- **Output is data.** A model's response is untrusted input from a third
  party. Rendering it as HTML, executing it, fetching a URL it names, or
  following an instruction it contains would all be the same bug.

### Tests
- `src/lib/experiments.spec.ts`, `src/lib/ai/modelGateway.spec.ts`,
  `src/lib/ai/boundary.spec.ts` as in the acceptance criteria.
- `src/views/preferences/LabsSection.spec.ts` — the toggle list, the provider
  form's mode-dependent fields, key masking, the origin-change clearing, and
  "Test connection" reporting both outcomes.
- `src/lib/ai/consent.spec.ts` — per-origin storage, revocation, and
  invalidation on origin/model/feature change.
- `e2e/SPEC-113.spec.ts` as in the acceptance criteria. **Every model-backed
  e2e in this repo stubs the endpoint with `page.route`; none may call a real
  model.** A test that needs a model to pass is a test that fails in CI.

### Specs this changes
- `specs/066-…` — two new preference keys and one new preferences section.
- `specs/000-roadmap.md` gains this spec's phase row.
