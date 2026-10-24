// Where model requests go (SPEC-113): the ModelProviderConfig the gateway
// reads, its defaults and the validator the preferences store runs over its
// `modelProvider` key. A leaf module — preferences.ts imports it, and it
// imports nothing — so the stored config is validated without the gateway
// (the only module that sends anything) ever being loaded.
//
// The API key is deliberately NOT part of this shape. It is a credential,
// kept under its own storage key (lib/ai/apiKey.ts), so it is never in the
// preferences object, never in anything that copies that object, and never
// in a document.

export type ProviderMode = "off" | "local" | "cloud";

export interface ModelProviderConfig {
  /** "off" until the user turns a provider on. "local" never sends a key;
   * "cloud" shows the key field and sends it as a bearer token. */
  mode: ProviderMode;
  /** The OpenAI-compatible base URL, up to and including its version path —
   * requests go to {baseUrl}/chat/completions and {baseUrl}/models. */
  baseUrl: string;
  /** The model id the endpoint knows the model by. */
  model: string;
  timeoutMs: number;
  maxOutputTokens: number;
  temperature: number;
}

export const PROVIDER_MODES: readonly ProviderMode[] = Object.freeze(["off", "local", "cloud"] as ProviderMode[]);

export const MIN_TIMEOUT_MS = 1_000;
export const MAX_TIMEOUT_MS = 600_000;
export const MAX_OUTPUT_TOKENS_LIMIT = 1_000_000;
export const MAX_TEMPERATURE = 2;
const MAX_BASE_URL_LENGTH = 2_048;
const MAX_MODEL_LENGTH = 256;

/** Off, with no endpoint and no model: no provider is preconfigured,
 * recommended or defaulted to. Frozen — it is the store's shared fallback. */
export const PROVIDER_CONFIG_DEFAULTS: ModelProviderConfig = Object.freeze({
  mode: "off",
  baseUrl: "",
  model: "",
  timeoutMs: 60_000,
  maxOutputTokens: 4_096,
  temperature: 0.2,
});

function integerIn(raw: unknown, min: number, max: number): number | undefined {
  return typeof raw === "number" && Number.isInteger(raw) && raw >= min && raw <= max ? raw : undefined;
}

function stringUpTo(raw: unknown, max: number): string | undefined {
  return typeof raw === "string" && raw.length <= max ? raw : undefined;
}

/** Repairs a stored, untrusted config field by field — one bad field falls
 * back to its default without discarding the rest, as the store does for
 * whole preferences. Anything not in the shape (an `apiKey` someone wrote
 * into the stored object, say) is dropped. A non-object rejects. */
export function parseProviderConfig(raw: unknown): ModelProviderConfig | undefined {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    return undefined;
  }
  const record = raw as Record<string, unknown>;
  const defaults = PROVIDER_CONFIG_DEFAULTS;
  const temperature = record.temperature;
  return {
    mode: PROVIDER_MODES.includes(record.mode as ProviderMode) ? (record.mode as ProviderMode) : defaults.mode,
    baseUrl: stringUpTo(record.baseUrl, MAX_BASE_URL_LENGTH) ?? defaults.baseUrl,
    model: stringUpTo(record.model, MAX_MODEL_LENGTH) ?? defaults.model,
    timeoutMs: integerIn(record.timeoutMs, MIN_TIMEOUT_MS, MAX_TIMEOUT_MS) ?? defaults.timeoutMs,
    maxOutputTokens: integerIn(record.maxOutputTokens, 1, MAX_OUTPUT_TOKENS_LIMIT) ?? defaults.maxOutputTokens,
    temperature:
      typeof temperature === "number" && Number.isFinite(temperature) && temperature >= 0 && temperature <= MAX_TEMPERATURE
        ? temperature
        : defaults.temperature,
  };
}
