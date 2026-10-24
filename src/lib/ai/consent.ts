// Standing consent for model requests (SPEC-113). Consent is per origin:
// before the first request to an origin the user sees the consent dialog
// with the exact payload, and "Always" records a grant here so the next
// request does not ask again. A grant is scoped to one origin, one model and
// one feature (the request's `purpose`): a request for another feature does
// not match it and asks for its own; a change of origin or model clears the
// origin's grants outright. Declining a request never touches a grant some
// other request was given.
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

/** Per origin: the model its grants were given for, and when each feature
 * was granted. */
type GrantMap = Record<string, { model: string; features: Record<string, number> }>;

let sessionGrants: GrantMap = {};
const listeners = new Set<() => void>();

function copyOf(grants: GrantMap): GrantMap {
  return JSON.parse(JSON.stringify(grants)) as GrantMap;
}

function readGrants(): GrantMap {
  let raw: string | null;
  try {
    raw = localStorage.getItem(CONSENT_STORAGE_KEY);
  } catch {
    // Blocked storage: this session's grants still apply.
    return copyOf(sessionGrants);
  }
  if (raw === null) {
    return {};
  }
  try {
    const parsed: unknown = JSON.parse(raw);
    if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) {
      return {};
    }
    const grants: GrantMap = {};
    for (const [origin, value] of Object.entries(parsed as Record<string, unknown>)) {
      if (value === null || typeof value !== "object") {
        continue;
      }
      const { model, features } = value as Record<string, unknown>;
      if (typeof model !== "string" || features === null || typeof features !== "object") {
        continue;
      }
      const kept = Object.fromEntries(
        Object.entries(features as Record<string, unknown>).filter(([, at]) => typeof at === "number"),
      ) as Record<string, number>;
      if (Object.keys(kept).length > 0) {
        grants[origin] = { model, features: kept };
      }
    }
    return grants;
  } catch {
    // Corrupt: no grant can be trusted, so every request asks again.
    return {};
  }
}

function writeGrants(grants: GrantMap): void {
  sessionGrants = copyOf(grants);
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

/** Every standing grant, one per origin and feature, for Labs to list. */
export function consentGrants(): ConsentGrant[] {
  return Object.entries(readGrants()).flatMap(([origin, { model, features }]) =>
    Object.entries(features).map(([feature, grantedAt]) => ({ origin, model, feature, grantedAt })),
  );
}

/** Whether `feature` may be sent to `origin` with `model` without asking. */
export function hasConsent(origin: string, model: string, feature: string): boolean {
  const grant = readGrants()[origin];
  return grant !== undefined && grant.model === model && Object.hasOwn(grant.features, feature);
}

/** Records "Always" for this origin, model and feature. A grant held for the
 * origin under another model is dropped: it no longer describes it. */
export function grantConsent(origin: string, model: string, feature: string, now: number = Date.now()): void {
  const grants = readGrants();
  const current = grants[origin];
  const features = current !== undefined && current.model === model ? current.features : {};
  writeGrants({ ...grants, [origin]: { model, features: { ...features, [feature]: now } } });
}

/** Revokes standing consent — one feature's, or the whole origin's — so the
 * next request asks again. */
export function revokeConsent(origin: string, feature?: string): void {
  const grants = readGrants();
  const grant = grants[origin];
  if (grant === undefined || (feature !== undefined && !Object.hasOwn(grant.features, feature))) {
    return;
  }
  if (feature !== undefined) {
    delete grant.features[feature];
  }
  if (feature === undefined || Object.keys(grant.features).length === 0) {
    delete grants[origin];
  }
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
 * exactly this origin, model and feature. Grants for another origin or
 * model are cleared on the way; another feature's grant is left alone. */
export function claimConsent(origin: string, model: string, feature: string): boolean {
  pruneConsent(origin, model);
  return hasConsent(origin, model, feature);
}

/** Subscribes to grant changes (Labs' revocation list). Returns an
 * unsubscribe. */
export function onConsentChange(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}
