// The API key, handled as a credential (SPEC-113). It lives under its own
// localStorage key — never inside the preferences object, so nothing that
// reads, copies or resets preferences ever sees it, and never in a document.
//
// The key is stored bound to the origin it was entered for. Reading it for
// any other origin discards it on the spot, so a key can never follow a URL
// edit to a new host — whichever path changed the URL (the Labs form, a
// preferences reset, another tab). Nothing here logs, and nothing here puts
// the key into an error message.
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
  try {
    const raw = localStorage.getItem(API_KEY_STORAGE_KEY);
    if (raw === null) {
      return null;
    }
    const parsed: unknown = JSON.parse(raw);
    if (parsed !== null && typeof parsed === "object") {
      const { origin, key } = parsed as Record<string, unknown>;
      if (typeof origin === "string" && typeof key === "string" && key !== "") {
        return { origin, key };
      }
    }
    // Corrupt: nothing usable, so nothing kept.
    localStorage.removeItem(API_KEY_STORAGE_KEY);
    return null;
  } catch {
    return sessionKey;
  }
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

/** The key entered for `origin`, or null. A key stored for any other origin
 * (or for none) is discarded, not returned. */
export function loadApiKey(origin: string | null): string | null {
  const stored = readStored();
  if (stored === null) {
    return null;
  }
  if (origin === null || stored.origin !== origin) {
    clearApiKey();
    return null;
  }
  return stored.key;
}

/** Whether any key is held at all, for any origin — what the origin-change
 * rule checks before discarding, so the user can be told a key went. */
export function hasStoredApiKey(): boolean {
  return readStored() !== null;
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
