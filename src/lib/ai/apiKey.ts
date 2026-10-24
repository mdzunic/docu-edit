// The API key, handled as a credential (SPEC-113). It lives under its own
// localStorage key — never inside the preferences object, so nothing that
// reads, copies or resets preferences ever sees it, and never in a document.
//
// The key is stored bound to the origin it was entered for and is only ever
// returned for that origin, so a key can never follow a URL edit to a new
// host. Whenever the configured origin moves — whichever path moved it (the
// Labs form, a preferences reset, another tab) — discardApiKeyUnlessFor
// deletes it outright. Nothing here logs, and nothing here puts the key into
// an error message.
//
// Read only by the gateway (which sends it) and the provider guard (which
// discards it) — lib/ai/boundary.spec.ts holds that.

export const API_KEY_STORAGE_KEY = "docu-edit:ai-api-key";

interface StoredKey {
  origin: string;
  key: string;
}

/** The in-session copy, used when localStorage is blocked so a key entered
 * this session still works until reload. */
let sessionKey: StoredKey | null = null;

function readStored(): StoredKey | null {
  let raw: string | null;
  try {
    raw = localStorage.getItem(API_KEY_STORAGE_KEY);
  } catch {
    // Blocked storage: the key entered this session still applies.
    return sessionKey;
  }
  if (raw === null) {
    return null;
  }
  let parsed: unknown = null;
  try {
    parsed = JSON.parse(raw);
  } catch {
    // Corrupt: handled below as unusable.
  }
  if (parsed !== null && typeof parsed === "object") {
    const { origin, key } = parsed as Record<string, unknown>;
    if (typeof origin === "string" && typeof key === "string" && key !== "") {
      return { origin, key };
    }
  }
  // Nothing usable, so nothing kept.
  clearApiKey();
  return null;
}

/** Forgets the key everywhere it is held. */
export function clearApiKey(): void {
  sessionKey = null;
  try {
    localStorage.removeItem(API_KEY_STORAGE_KEY);
  } catch {
    // Blocked storage held nothing to remove.
  }
}

/** Stores `key` for `origin`, replacing any key held for any origin. An
 * empty key clears it. */
export function saveApiKey(origin: string, key: string): void {
  const trimmed = key.trim();
  if (trimmed === "") {
    clearApiKey();
    return;
  }
  sessionKey = { origin, key: trimmed };
  try {
    localStorage.setItem(API_KEY_STORAGE_KEY, JSON.stringify(sessionKey));
  } catch {
    // Blocked or full: the in-session copy still applies until reload.
  }
}

/** The key entered for `origin`, or null — never a key entered for any
 * other origin. Read-only: a request still in flight for an old origin must
 * not delete the key the user has since entered for the new one, so the
 * discarding is discardApiKeyUnlessFor's, run against the configured origin. */
export function loadApiKey(origin: string | null): string | null {
  const stored = readStored();
  return stored !== null && origin !== null && stored.origin === origin ? stored.key : null;
}

/** Discards the held key unless it was entered for `origin`. Returns whether
 * a key was discarded. */
export function discardApiKeyUnlessFor(origin: string | null): boolean {
  const stored = readStored();
  if (stored === null || (origin !== null && stored.origin === origin)) {
    return false;
  }
  clearApiKey();
  return true;
}
