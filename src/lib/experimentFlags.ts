// The stored shape of the Labs switchboard (SPEC-113): the closed list of
// experiment ids and the one validator the preferences store runs over its
// `experiments` key. A leaf module on purpose — preferences.ts imports it for
// that validator and experiments.ts imports preferences.ts to persist, so
// keeping the ids here (rather than in experiments.ts) is what stops the two
// from forming an import cycle that would leave one of them half-initialised,
// the same reason assemblyPersist.ts sits apart from assembly.ts.
//
// The registry itself — titles, summaries, risk lines — is experiments.ts;
// it is typed over EXPERIMENT_IDS, so the Labs list and the stored flags
// cannot drift apart.

/** Every experiment there is, closed. A feature cannot invent a flag at a
 * call site; it adds its id here (and its copy in experiments.ts). Grows one
 * entry per experimental spec. */
export const EXPERIMENT_IDS = Object.freeze(["library-search", "deck-designer"] as const);

export type ExperimentId = (typeof EXPERIMENT_IDS)[number];

/** One flag per experiment — what the preferences store keeps under its
 * single `experiments` key. */
export type ExperimentFlags = Record<ExperimentId, boolean>;

/** Every experiment off: nothing experimental is in the product until the
 * user asks for it. Frozen, because it is the shared fallback the store
 * hands out whenever the stored map is missing or corrupt. */
export const EXPERIMENT_FLAG_DEFAULTS: ExperimentFlags = Object.freeze(
  Object.fromEntries(EXPERIMENT_IDS.map((id) => [id, false])) as ExperimentFlags,
);

export function isExperimentId(value: unknown): value is ExperimentId {
  return typeof value === "string" && (EXPERIMENT_IDS as readonly string[]).includes(value);
}

/** Repairs a stored, untrusted flag map: every known id is on only when it is
 * stored as exactly `true`; an unknown id is ignored — not carried, not
 * resurrected — so a flag whose experiment no longer exists can never come
 * back to life. A non-object rejects (the store then uses the defaults). */
export function parseExperimentFlags(raw: unknown): ExperimentFlags | undefined {
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    return undefined;
  }
  const record = raw as Record<string, unknown>;
  return Object.fromEntries(EXPERIMENT_IDS.map((id) => [id, record[id] === true])) as ExperimentFlags;
}
