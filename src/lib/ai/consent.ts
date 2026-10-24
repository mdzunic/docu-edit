// Standing consent for model requests (SPEC-113). Consent is per origin:
// before the first request to an origin the user sees the consent dialog
// with the exact payload, and "Always" records a grant here so the next
// request does not ask again. A grant covers one origin, one model and one
// feature (the request's `purpose`), and is cleared whenever any of the
// three changes: a config change prunes it, and a request for a different
// feature clears it before asking again.
//
// Persisted under its own localStorage key, revocable from Labs. Holds no
// prompt, no payload and no key — only who may be sent what kind of request.

export const CONSENT_STORAGE_KEY = "docu-edit:ai-consent";

export interface ConsentGrant {
  origin: string;
  model: string;
  /** The feature the grant covers — the request purpose it was given for. */
  feature: string;
  grantedAt: number;
}

type GrantMap = Record<string, Omit<ConsentGrant, "origin">>;

let sessionGrants: GrantMap = {};
const listeners = new Set<() => void>();

function readGrants(): GrantMap {
  try {
    const raw = localStorage.getItem(CONSENT_STORAGE_KEY);
    if (raw === null) {
      return {};
    }
    const parsed: unknown = JSON.parse(raw);
    if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) {
      return {};
    }
    const grants: GrantMap = {};
    for (const [origin, value] of Object.entries(parsed as Record<string, unknown>)) {
      if (value === null || typeof value !== "object") {
        continue;
      }
      const { model, feature, grantedAt } = value as Record<string, unknown>;
      if (typeof model === "string" && typeof feature === "string" && typeof grantedAt === "number") {
        grants[origin] = { model, feature, grantedAt };
      }
    }
    return grants;
  } catch {
    return { ...sessionGrants };
  }
}

function writeGrants(grants: GrantMap): void {
  sessionGrants = { ...grants };
  try {
    if (Object.keys(grants).length === 0) {
      localStorage.removeItem(CONSENT_STORAGE_KEY);
    } else {
      localStorage.setItem(CONSENT_STORAGE_KEY, JSON.stringify(grants));
    }
  } catch {
    // Blocked or full: the in-session copy still applies until reload.
  }
  for (const listener of listeners) {
    listener();
  }
}

/** Every standing grant, for Labs to list. */
export function consentGrants(): ConsentGrant[] {
  return Object.entries(readGrants()).map(([origin, grant]) => ({ origin, ...grant }));
}

/** Whether `feature` may be sent to `origin` with `model` without asking. */
export function hasConsent(origin: string, model: string, feature: string): boolean {
  const grant = readGrants()[origin];
  return grant !== undefined && grant.model === model && grant.feature === feature;
}

/** Records "Always" for this origin, model and feature, replacing whatever
 * the origin held before. */
export function grantConsent(origin: string, model: string, feature: string, now: number = Date.now()): void {
  writeGrants({ ...readGrants(), [origin]: { model, feature, grantedAt: now } });
}

/** Revokes the origin's standing consent: its next request asks again. */
export function revokeConsent(origin: string): void {
  const grants = readGrants();
  if (!(origin in grants)) {
    return;
  }
  delete grants[origin];
  writeGrants(grants);
}

/** Clears every grant that no longer describes the configured endpoint: any
 * other origin's, and this origin's when the model has changed. Run on every
 * config change and before every consent check. */
export function pruneConsent(origin: string | null, model: string): void {
  const grants = readGrants();
  const kept: GrantMap = {};
  for (const [grantOrigin, grant] of Object.entries(grants)) {
    if (grantOrigin === origin && grant.model === model) {
      kept[grantOrigin] = grant;
    }
  }
  if (Object.keys(kept).length !== Object.keys(grants).length) {
    writeGrants(kept);
  }
}

/** The gateway's check before a request: whether a standing grant covers
 * exactly this origin, model and feature. Any grant that does not — another
 * origin's, or this origin's for another model or another feature — is
 * cleared on the way, so the request asks again. */
export function claimConsent(origin: string, model: string, feature: string): boolean {
  pruneConsent(origin, model);
  if (hasConsent(origin, model, feature)) {
    return true;
  }
  revokeConsent(origin);
  return false;
}

/** Subscribes to grant changes (Labs' revocation list). Returns an
 * unsubscribe. */
export function onConsentChange(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}
