// The model gateway (SPEC-113) — the ONLY module in this repo that may make a
// request to a model endpoint. Auditing "can this app send my words
// somewhere" is this one file; lib/ai/boundary.spec.ts keeps it that way by
// failing on any other runtime module that reaches the network.
//
// What it does, and nothing more:
// - speaks the OpenAI-compatible HTTP shape and only that — POST
//   {baseUrl}/chat/completions and GET {baseUrl}/models — so a local server
//   or any hosted endpoint the user names works with no per-provider code.
//   No model is bundled, downloaded or run here, and none is defaulted to.
// - pins the destination: baseUrl resolves to one origin per request
//   (lib/ai/origin.ts), every URL is checked against it, `redirect: "error"`
//   stops a redirect moving a request or a key elsewhere, `credentials:
//   "omit"` sends no cookies, and the referrer is withheld.
// - asks first: before a request to an origin without a standing grant, the
//   registered consent prompt (the shell's ModelConsentDialog) is shown the
//   exact body that will be sent; "Always" is stored per origin
//   (lib/ai/consent.ts), and no prompt means no request.
// - sends the API key (lib/ai/apiKey.ts) only in cloud mode, only as
//   `Authorization: Bearer`, only to the origin it was entered for; never
//   logs it and redacts it from any server text it repeats.
// - bounds every request: the caller's AbortSignal, a timer of its own at
//   timeoutMs, and no automatic retry — ever.
// - keeps a session-only activity log of counts — never prompt or response
//   text — and makes no telemetry or analytics request of any kind.
// - treats output as data: complete/stream return plain strings. The gateway
//   reads only the protocol envelope around them; it parses nothing inside,
//   executes nothing, and follows no instruction found in a response. A
//   feature that wants structure asks with `json: true` and validates the
//   string with its own schema.
//
// Requests are user-initiated: a feature calls this from a user's action,
// never from a timer, a watcher, a background task or a route guard. Loaded
// only through an import() behind an experiment's flag (see
// composables/useModelGateway.ts and src/experimentGating.spec.ts).
import { loadPreferences, savePreference } from "../preferences";
import { discardApiKeyUnlessFor, loadApiKey, saveApiKey } from "./apiKey";
import { claimConsent, grantConsent, pruneConsent } from "./consent";
import {
  ModelConsentDeniedError,
  ModelHttpError,
  ModelNotConfiguredError,
  ModelOutputError,
  ModelTimeoutError,
  ModelUnreachableError,
  isAbortError,
} from "./errors";
import { originOf, pinnedUrl, resolveEndpoint, type Endpoint } from "./origin";
import type { ModelProviderConfig, ProviderMode } from "./providerConfig";
import { createSseParser } from "./sse";

export {
  ModelConsentDeniedError,
  ModelGatewayError,
  ModelHttpError,
  ModelNotConfiguredError,
  ModelOutputError,
  ModelTimeoutError,
  ModelUnreachableError,
  describeModelError,
  isAbortError,
} from "./errors";
export type { ModelProviderConfig, ProviderMode };

export interface ModelRequest {
  /** Why the request is being made — a fixed, feature-owned label (never the
   * user's text), shown in the consent dialog and the activity log. It is
   * also the "feature" a standing consent covers: a grant given for one
   * purpose does not carry over to another. */
  purpose: string;
  system: string;
  user: string;
  /** Ask for a JSON object back. The gateway sets the provider's JSON mode
   * where the endpoint advertises it (in its GET /models listing) and
   * otherwise says so in the prompt; it never parses or trusts the result —
   * the caller validates. */
  json?: boolean;
  signal?: AbortSignal;
}

export interface ModelResult {
  text: string;
  model: string;
  elapsedMs: number;
}

export type ModelOutcome = "ok" | "http-error" | "unreachable" | "timeout" | "output-error" | "aborted" | "failed";

/** One request that went out. Counts only — never prompt or response text. */
export interface ModelActivity {
  time: number;
  origin: string;
  model: string;
  purpose: string;
  requestBytes: number;
  responseBytes: number;
  outcome: ModelOutcome;
}

export type ConsentChoice = "once" | "always" | "cancel";

/** What the consent dialog is shown: where the request goes, whether that is
 * this machine, and the exact bytes it will carry. */
export interface ConsentRequest {
  origin: string;
  local: boolean;
  purpose: string;
  model: string;
  method: "GET" | "POST";
  url: string;
  /** The exact request body — the very string that is sent. Empty for GET. */
  body: string;
  bytes: number;
  /** Whether an API key goes along, as the Authorization header. */
  sendsKey: boolean;
}

/** Shows the user a request and resolves with their answer. The signal
 * aborts when the caller cancels while the prompt is open. */
export type ConsentPrompt = (request: ConsentRequest, signal?: AbortSignal) => Promise<ConsentChoice>;

/** The purpose "Test connection" sends under. */
export const TEST_CONNECTION_PURPOSE = "Test connection";
export const LIST_MODELS_PURPOSE = "List models";

/** What the prompt says when JSON is wanted and the endpoint has not
 * advertised a JSON mode. */
export const JSON_INSTRUCTION = "Respond with a single JSON object and nothing else: no prose, no code fences.";

const MAX_ACTIVITY_ENTRIES = 200;

let consentPrompt: ConsentPrompt | null = null;
let activity: readonly ModelActivity[] = Object.freeze([]);
const activityListeners = new Set<() => void>();
/** "origin model" pairs whose /models entry advertised response_format —
 * learnt only from a listModels() the user asked for, never by probing. */
const jsonModeAdvertised = new Set<string>();

/** Registers the consent prompt (the app layer does, once). Without one,
 * every request that needs consent is refused. */
export function setConsentPrompt(prompt: ConsentPrompt | null): void {
  consentPrompt = prompt;
}

// --- config and key -----------------------------------------------------------

export function getProviderConfig(): ModelProviderConfig {
  return { ...loadPreferences().modelProvider };
}

export interface ProviderConfigChange {
  previousOrigin: string | null;
  origin: string | null;
  /** Whether a held API key was discarded because the origin changed. */
  keyDiscarded: boolean;
}

/** Applies a config edit. The key dies with the origin: whenever the edit
 * moves baseUrl to another origin (or to none), the held key is discarded —
 * before anything else can read it — and the caller is told so it can say
 * why. Standing consent no longer matching the origin or model is cleared. */
export function updateProviderConfig(patch: Partial<ModelProviderConfig>): ProviderConfigChange {
  const previous = getProviderConfig();
  const next: ModelProviderConfig = { ...previous, ...patch };
  const previousOrigin = originOf(previous.baseUrl);
  const origin = originOf(next.baseUrl);
  const keyDiscarded = discardApiKeyUnlessFor(origin);
  savePreference("modelProvider", next);
  // Trimmed, as the gateway grants and checks it: stray whitespace typed
  // around the same model id is not a different model.
  pruneConsent(origin, next.model.trim());
  return { previousOrigin, origin, keyDiscarded };
}

/** Stores the API key for the configured origin (an empty key removes it).
 * Refused while the base URL does not resolve: a key is always bound to the
 * origin it was entered for. */
export function setApiKey(key: string): void {
  const origin = originOf(getProviderConfig().baseUrl);
  if (origin === null) {
    throw new ModelNotConfiguredError("Set a valid base URL first: an API key is bound to the endpoint it is entered for.");
  }
  const trimmed = key.trim();
  // A header value must be printable ASCII; refusing here keeps a pasted
  // stray character from surfacing later as a baffling network failure.
  if (trimmed !== "" && !/^[\x21-\x7e]+$/.test(trimmed)) {
    throw new ModelNotConfiguredError("That API key contains spaces or characters a key cannot have. Paste it again.");
  }
  saveApiKey(origin, trimmed);
}

/** The key held for the configured origin, for the masked field — "" when
 * there is none. A key held for any other origin is discarded here. */
export function getApiKey(): string {
  const origin = originOf(getProviderConfig().baseUrl);
  discardApiKeyUnlessFor(origin);
  return loadApiKey(origin) ?? "";
}

// --- the activity log ---------------------------------------------------------

/** This session's requests, oldest first. In memory only: a reload clears it. */
export function activityLog(): readonly ModelActivity[] {
  return activity;
}

export function clearActivityLog(): void {
  activity = Object.freeze([]);
  notifyActivity();
}

export function onActivityLogChange(listener: () => void): () => void {
  activityListeners.add(listener);
  return () => {
    activityListeners.delete(listener);
  };
}

function notifyActivity(): void {
  for (const listener of activityListeners) {
    listener();
  }
}

function recordActivity(entry: ModelActivity): void {
  activity = Object.freeze([...activity, Object.freeze(entry)].slice(-MAX_ACTIVITY_ENTRIES));
  notifyActivity();
}

function outcomeOf(error: unknown): ModelOutcome {
  if (error instanceof ModelHttpError) {
    return "http-error";
  }
  if (error instanceof ModelUnreachableError) {
    return "unreachable";
  }
  if (error instanceof ModelTimeoutError) {
    return "timeout";
  }
  if (error instanceof ModelOutputError) {
    return "output-error";
  }
  return isAbortError(error) ? "aborted" : "failed";
}

// --- preparing a request ------------------------------------------------------

interface Prepared {
  config: ModelProviderConfig;
  endpoint: Endpoint;
  model: string;
  purpose: string;
  method: "GET" | "POST";
  url: string;
  body: string;
}

function utf8Bytes(text: string): number {
  return new TextEncoder().encode(text).byteLength;
}

function abortError(): DOMException {
  return new DOMException("The model request was cancelled.", "AbortError");
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw abortError();
  }
}

/** Resolves the config into one pinned request, or refuses it. The body is
 * built here, once, as the exact string that is shown and then sent —
 * pretty-printed, so what the user reads in the consent dialog IS the
 * payload, byte for byte. */
function prepare(
  method: "GET" | "POST",
  path: string,
  purpose: string,
  payload: ((config: ModelProviderConfig, model: string, endpoint: Endpoint) => object) | null,
  requireModel = true,
): Prepared {
  const config = getProviderConfig();
  if (config.mode === "off") {
    throw new ModelNotConfiguredError("The model provider is off. Choose one in Preferences › Labs › Model provider.");
  }
  const resolved = resolveEndpoint(config.baseUrl);
  if (!resolved.ok) {
    throw new ModelNotConfiguredError(resolved.reason);
  }
  const { endpoint } = resolved;
  const model = config.model.trim();
  if (requireModel && model === "") {
    throw new ModelNotConfiguredError("No model is set. Enter the model id your endpoint serves in Preferences › Labs.");
  }
  const url = pinnedUrl(endpoint, path);
  if (url === null) {
    throw new ModelNotConfiguredError(`Refused ${path}: it is outside ${endpoint.origin}, the only origin requests may go to.`);
  }
  const body = payload === null ? "" : JSON.stringify(payload(config, model, endpoint), null, 2);
  return { config, endpoint, model, purpose, method, url, body };
}

/** The chat-completions body for a feature's request. */
function chatPayload(request: ModelRequest, streaming: boolean) {
  return (config: ModelProviderConfig, model: string, endpoint: Endpoint): object => {
    const wantsJson = request.json === true;
    const jsonMode = wantsJson && jsonModeAdvertised.has(`${endpoint.origin} ${model}`);
    let system = request.system;
    if (wantsJson && !jsonMode) {
      system = system === "" ? JSON_INSTRUCTION : `${system}\n\n${JSON_INSTRUCTION}`;
    }
    const messages = [
      ...(system === "" ? [] : [{ role: "system", content: system }]),
      { role: "user", content: request.user },
    ];
    return {
      model,
      messages,
      max_tokens: config.maxOutputTokens,
      temperature: config.temperature,
      stream: streaming,
      ...(jsonMode ? { response_format: { type: "json_object" } } : {}),
    };
  };
}

/** Asks for consent unless a standing grant covers exactly this origin,
 * model and feature. Resolves only when the request may go. */
async function obtainConsent(prepared: Prepared, signal: AbortSignal | undefined): Promise<void> {
  throwIfAborted(signal);
  const { endpoint, model, purpose } = prepared;
  if (claimConsent(endpoint.origin, model, purpose)) {
    return;
  }
  if (consentPrompt === null) {
    throw new ModelConsentDeniedError(endpoint.origin);
  }
  const choice = await consentPrompt(
    {
      origin: endpoint.origin,
      local: endpoint.loopback,
      purpose,
      model,
      method: prepared.method,
      url: prepared.url,
      body: prepared.body,
      bytes: utf8Bytes(prepared.body),
      sendsKey: prepared.config.mode === "cloud" && loadApiKey(endpoint.origin) !== null,
    },
    signal,
  );
  throwIfAborted(signal);
  if (choice === "always") {
    grantConsent(endpoint.origin, model, purpose);
  } else if (choice !== "once") {
    throw new ModelConsentDeniedError(endpoint.origin);
  }
}

// --- sending ------------------------------------------------------------------

interface ByteCounter {
  bytes: number;
}

/** A body read that failed for any reason but a cancellation: the
 * connection went mid-answer. exchange() reports it as unreachable. */
class ConnectionDropped extends Error {}

/** The next chunk of a body. A cancellation (the caller's abort, the timer)
 * passes through as itself; any other failure is a dropped connection. */
async function readChunk(
  reader: ReadableStreamDefaultReader<Uint8Array>,
): Promise<ReadableStreamReadResult<Uint8Array>> {
  try {
    return await reader.read();
  } catch (error) {
    throw isAbortError(error) ? error : new ConnectionDropped();
  }
}

/** Reads a body as text, counting its bytes; stops after `limit` characters
 * when given (an error body is only ever quoted, so the rest is not read). */
async function readText(response: Response, counter: ByteCounter, limit = Infinity): Promise<string> {
  const reader = response.body?.getReader();
  if (!reader) {
    return "";
  }
  const decoder = new TextDecoder();
  let text = "";
  for (;;) {
    const { done, value } = await readChunk(reader);
    if (done) {
      break;
    }
    counter.bytes += value.byteLength;
    text += decoder.decode(value, { stream: true });
    if (text.length >= limit) {
      await reader.cancel();
      return text;
    }
  }
  return text + decoder.decode();
}

function redact(text: string, key: string | null): string {
  return key ? text.split(key).join("[redacted]") : text;
}

function isSameOrigin(url: string, origin: string): boolean {
  try {
    return new URL(url).origin === origin;
  } catch {
    return false;
  }
}

/** Sends one prepared request and hands the response to `read`. Owns
 * everything every request shares: the timer, the caller's signal, the
 * pinned-origin fetch options, the key header, the redirect and status
 * checks, the mapping of failures to typed errors, and the activity entry. */
async function exchange<T>(
  prepared: Prepared,
  signal: AbortSignal | undefined,
  accept: string,
  read: (response: Response, counter: ByteCounter, key: string | null) => Promise<T>,
): Promise<T> {
  throwIfAborted(signal);
  const { config, endpoint } = prepared;
  const controller = new AbortController();
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, config.timeoutMs);
  const onCallerAbort = () => controller.abort();
  signal?.addEventListener("abort", onCallerAbort, { once: true });

  const counter: ByteCounter = { bytes: 0 };
  let sent = false;
  let outcome: ModelOutcome = "ok";
  try {
    const key = config.mode === "cloud" ? loadApiKey(endpoint.origin) : null;
    const headers: Record<string, string> = { Accept: accept };
    if (prepared.method === "POST") {
      headers["Content-Type"] = "application/json";
    }
    if (key !== null) {
      headers.Authorization = `Bearer ${key}`;
    }
    let response: Response;
    try {
      sent = true;
      response = await fetch(prepared.url, {
        method: prepared.method,
        headers,
        body: prepared.method === "POST" ? prepared.body : undefined,
        redirect: "error",
        credentials: "omit",
        mode: "cors",
        cache: "no-store",
        referrerPolicy: "no-referrer",
        signal: controller.signal,
      });
    } catch (error) {
      // A rejected fetch is no HTTP answer at all: nothing listening, DNS,
      // CORS, or a redirect that `redirect: "error"` refused.
      throw timedOut || signal?.aborted ? error : new ModelUnreachableError(endpoint.origin, endpoint.loopback);
    }
    if (response.redirected || (response.url !== "" && !isSameOrigin(response.url, endpoint.origin))) {
      await response.body?.cancel();
      throw new ModelUnreachableError(endpoint.origin, endpoint.loopback, "redirect");
    }
    if (!response.ok) {
      // Read a key's length past the excerpt, so a key straddling the cut is
      // whole when it is redacted — then ModelHttpError cuts to 500.
      const body = await readText(response, counter, ModelHttpError.BODY_EXCERPT_LENGTH + (key?.length ?? 0));
      throw new ModelHttpError(response.status, redact(body, key));
    }
    return await read(response, counter, key);
  } catch (error) {
    const failure = timedOut
      ? new ModelTimeoutError(config.timeoutMs)
      : signal?.aborted
        ? abortError()
        : error instanceof ConnectionDropped
          ? new ModelUnreachableError(endpoint.origin, endpoint.loopback, "dropped")
          : error;
    outcome = outcomeOf(failure);
    throw failure;
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener("abort", onCallerAbort);
    if (sent) {
      recordActivity({
        time: Date.now(),
        origin: endpoint.origin,
        model: prepared.model,
        purpose: prepared.purpose,
        requestBytes: utf8Bytes(prepared.body),
        responseBytes: counter.bytes,
        outcome,
      });
    }
  }
}

// --- reading the envelope -----------------------------------------------------
// Only the OpenAI-compatible wrapper is read: where the text is, and which
// model answered. The text itself is returned untouched.

function parseJsonEnvelope(text: string): Record<string, unknown> {
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch {
    throw new ModelOutputError("the body is not JSON");
  }
  if (data === null || typeof data !== "object" || Array.isArray(data)) {
    throw new ModelOutputError("the body is not a JSON object");
  }
  return data as Record<string, unknown>;
}

function firstChoice(data: Record<string, unknown>): Record<string, unknown> | null {
  const choices = data.choices;
  const first: unknown = Array.isArray(choices) ? choices[0] : undefined;
  return first !== null && typeof first === "object" ? (first as Record<string, unknown>) : null;
}

function modelOf(data: Record<string, unknown>, fallback: string): string {
  return typeof data.model === "string" && data.model !== "" ? data.model : fallback;
}

function completionText(data: Record<string, unknown>): string {
  const message = firstChoice(data)?.message;
  const content = message !== null && typeof message === "object" ? (message as Record<string, unknown>).content : undefined;
  if (typeof content !== "string") {
    throw new ModelOutputError("no choices[0].message.content text");
  }
  return content;
}

/** Reads an event stream to its end (or `[DONE]`), handing each text delta
 * to `onDelta` as it arrives. */
async function readEventStream(
  response: Response,
  counter: ByteCounter,
  key: string | null,
  fallbackModel: string,
  onDelta: (chunk: string) => void,
): Promise<{ text: string; model: string }> {
  const reader = response.body?.getReader();
  if (!reader) {
    throw new ModelOutputError("an empty stream");
  }
  const decoder = new TextDecoder();
  // Mutated from the parser's callback, so held in one object rather than
  // loose lets the compiler would narrow as never changing.
  const state: { text: string; model: string; finished: boolean; failure: ModelOutputError | null } = {
    text: "",
    model: fallbackModel,
    finished: false,
    failure: null,
  };
  const parser = createSseParser((data) => {
    if (state.finished || state.failure !== null) {
      return;
    }
    if (data.trim() === "[DONE]") {
      state.finished = true;
      return;
    }
    let event: Record<string, unknown>;
    try {
      event = parseJsonEnvelope(data);
    } catch {
      state.failure = new ModelOutputError("a stream event that is not a JSON object");
      return;
    }
    if (event.error !== undefined) {
      const detail =
        typeof event.error === "object" && event.error !== null ? (event.error as Record<string, unknown>).message : event.error;
      const excerpt = redact(String(detail), key).slice(0, ModelHttpError.BODY_EXCERPT_LENGTH);
      state.failure = new ModelOutputError(`the stream reported an error: ${excerpt}`);
      return;
    }
    state.model = modelOf(event, state.model);
    const delta = firstChoice(event)?.delta;
    const content = delta !== null && typeof delta === "object" ? (delta as Record<string, unknown>).content : undefined;
    if (typeof content === "string" && content !== "") {
      state.text += content;
      onDelta(content);
    }
  });
  while (!state.finished && state.failure === null) {
    const { done, value } = await readChunk(reader);
    if (done) {
      break;
    }
    counter.bytes += value.byteLength;
    parser.push(decoder.decode(value, { stream: true }));
  }
  if (state.finished || state.failure !== null) {
    await reader.cancel();
  } else {
    parser.push(decoder.decode());
  }
  parser.end();
  if (state.failure !== null) {
    throw state.failure;
  }
  if (!state.finished && state.text === "") {
    throw new ModelOutputError("the stream ended without any text");
  }
  return { text: state.text, model: state.model };
}

// --- the public requests ------------------------------------------------------

/** One bounded request-response: the whole answer as a plain string. */
export async function complete(request: ModelRequest): Promise<ModelResult> {
  const prepared = prepare("POST", "chat/completions", request.purpose, chatPayload(request, false));
  await obtainConsent(prepared, request.signal);
  const started = performance.now();
  const result = await exchange(prepared, request.signal, "application/json", async (response, counter) => {
    const data = parseJsonEnvelope(await readText(response, counter));
    return { text: completionText(data), model: modelOf(data, prepared.model) };
  });
  return { ...result, elapsedMs: Math.round(performance.now() - started) };
}

/** The same request, streamed: `onDelta` receives each piece of text as the
 * server-sent events arrive, parsed incrementally, and the promise resolves
 * with the whole string. If it rejects, nothing it delivered is a result —
 * the caller discards what onDelta received. An endpoint that ignores
 * `stream: true` and answers with one JSON body is read as such. */
export async function stream(request: ModelRequest, onDelta: (chunk: string) => void): Promise<ModelResult> {
  const prepared = prepare("POST", "chat/completions", request.purpose, chatPayload(request, true));
  await obtainConsent(prepared, request.signal);
  const started = performance.now();
  const result = await exchange(prepared, request.signal, "text/event-stream", async (response, counter, key) => {
    if ((response.headers.get("content-type") ?? "").includes("application/json")) {
      const data = parseJsonEnvelope(await readText(response, counter));
      const text = completionText(data);
      if (text !== "") {
        onDelta(text);
      }
      return { text, model: modelOf(data, prepared.model) };
    }
    return readEventStream(response, counter, key, prepared.model, onDelta);
  });
  return { ...result, elapsedMs: Math.round(performance.now() - started) };
}

/** GET {baseUrl}/models: the model ids the endpoint lists. Also notes which
 * of them advertise a JSON mode, so a later `json: true` request uses it. */
export async function listModels(signal?: AbortSignal): Promise<string[]> {
  const prepared = prepare("GET", "models", LIST_MODELS_PURPOSE, null, false);
  await obtainConsent(prepared, signal);
  return exchange(prepared, signal, "application/json", async (response, counter) => {
    const data = parseJsonEnvelope(await readText(response, counter));
    if (!Array.isArray(data.data)) {
      throw new ModelOutputError("no data array of models");
    }
    const ids: string[] = [];
    for (const entry of data.data as unknown[]) {
      if (entry === null || typeof entry !== "object") {
        continue;
      }
      const { id, supported_parameters: supported } = entry as Record<string, unknown>;
      if (typeof id !== "string") {
        continue;
      }
      ids.push(id);
      if (Array.isArray(supported) && supported.includes("response_format")) {
        jsonModeAdvertised.add(`${prepared.endpoint.origin} ${id}`);
      }
    }
    return ids;
  });
}

/** "Test connection": one trivial chat completion — no document content, one
 * output token — reporting which model answered and the round trip. The one
 * request no feature asks for; consent applies to it like any other. */
export async function testConnection(signal?: AbortSignal): Promise<{ model: string; elapsedMs: number }> {
  const prepared = prepare("POST", "chat/completions", TEST_CONNECTION_PURPOSE, (_config, model) => ({
    model,
    messages: [{ role: "user", content: "Reply with the single word OK." }],
    max_tokens: 1,
    temperature: 0,
    stream: false,
  }));
  await obtainConsent(prepared, signal);
  const started = performance.now();
  const model = await exchange(prepared, signal, "application/json", async (response, counter) => {
    const data = parseJsonEnvelope(await readText(response, counter));
    if (!Array.isArray(data.choices)) {
      throw new ModelOutputError("no choices array");
    }
    return modelOf(data, prepared.model);
  });
  return { model, elapsedMs: Math.round(performance.now() - started) };
}
