// The experiment registry (SPEC-113): the features that ship behind a Labs
// toggle, and the one way to ask whether one is on. Framework-free — the
// router's guards, the gates in the UI (components/ExperimentGate.vue) and
// any lib module read the same flags through these four exports.
//
// Persisted through the preferences store under its one `experiments` key
// (lib/preferences.ts), every flag false until the user turns it on in Labs.
// A disabled experiment is absent, not hidden: its controls do not render,
// its route (where it has one) redirects to "/", and its code is reached only
// through an import() behind the flag check.
import { loadPreferences, onPreferencesChange, savePreference } from "./preferences";
import { EXPERIMENT_IDS, isExperimentId, type ExperimentFlags, type ExperimentId } from "./experimentFlags";

export { isExperimentId };
export type { ExperimentFlags, ExperimentId };

export interface Experiment {
  id: ExperimentId;
  title: string;
  /** One sentence: what it does. */
  summary: string;
  /** One sentence: what it costs or might break. */
  risk: string;
}

// The copy, keyed by id so the compiler insists every id in the closed union
// has an entry. SPEC-114 and SPEC-115 own their routes and gated UI, and may
// refine their own line here.
const COPY: Record<ExperimentId, Omit<Experiment, "id">> = {
  "library-search": {
    title: "Library search",
    summary: "Search inside every document stored in your library.",
    risk: "Search indexing may be slow on large libraries and its results are not covered by round-trip guarantees.",
  },
  "deck-designer": {
    title: "Deck designer",
    summary: "Ask a language model you configure to design a deck from your notes.",
    risk: "Sends the text you put in front of the model to the endpoint you name; output is untrusted and provisional.",
  },
};

/** Every experiment, in Labs order. Frozen, entries and all. */
export const EXPERIMENTS: readonly Experiment[] = Object.freeze(
  EXPERIMENT_IDS.map((id) => Object.freeze({ id, ...COPY[id] })),
);

/** The experiments that reach the model gateway (lib/ai/modelGateway.ts).
 * Labs shows the model provider settings only while one of these is on, so
 * with all of them off the gateway's code is never even loaded. */
export const MODEL_EXPERIMENTS: readonly ExperimentId[] = Object.freeze(["deck-designer"] as ExperimentId[]);

export function isExperimentEnabled(id: ExperimentId): boolean {
  return isExperimentId(id) && loadPreferences().experiments[id] === true;
}

/** Turns one experiment on or off, persisted and applied live. An id outside
 * the closed union is ignored. */
export function setExperimentEnabled(id: ExperimentId, on: boolean): void {
  if (!isExperimentId(id)) {
    return;
  }
  savePreference("experiments", { ...loadPreferences().experiments, [id]: on === true });
}

/** Subscribes to flag changes — this tab's and, through the preferences
 * store's storage event, another tab's. Fires only when the map actually
 * changes, with a copy of it. Returns an unsubscribe. */
export function onExperimentsChange(listener: (map: Record<ExperimentId, boolean>) => void): () => void {
  let last: ExperimentFlags = loadPreferences().experiments;
  return onPreferencesChange((prefs) => {
    const next = prefs.experiments;
    if (EXPERIMENT_IDS.every((id) => next[id] === last[id])) {
      return;
    }
    last = next;
    listener({ ...next });
  });
}
