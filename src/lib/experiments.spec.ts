// SPEC-113 — the experiment registry: a frozen, closed list with Labs copy,
// every flag off by default, persisted under the preferences store's one
// `experiments` key, unknown stored ids ignored (never resurrected), and a
// listener that fires only when the flag map changes.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  EXPERIMENTS,
  MODEL_EXPERIMENTS,
  isExperimentEnabled,
  onExperimentsChange,
  setExperimentEnabled,
  type ExperimentId,
} from "./experiments";
import { EXPERIMENT_FLAG_DEFAULTS, EXPERIMENT_IDS, parseExperimentFlags } from "./experimentFlags";
import {
  PREFERENCES_STORAGE_KEY,
  PREFERENCE_DEFAULTS,
  loadPreferences,
  resetPreferences,
  resetPreferencesSession,
  savePreference,
} from "./preferences";

function stored(): Record<string, unknown> {
  const raw = localStorage.getItem(PREFERENCES_STORAGE_KEY);
  return raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
}

beforeEach(() => {
  localStorage.clear();
  resetPreferencesSession();
});

describe("the registry", () => {
  it("lists library-search and deck-designer from day one, with their Labs copy", () => {
    expect(EXPERIMENTS.map((experiment) => experiment.id)).toEqual(["library-search", "deck-designer"]);
    expect(EXPERIMENTS[0]).toEqual({
      id: "library-search",
      title: "Library search",
      summary: "Search inside every document stored in your library.",
      risk: "Search indexing may be slow on large libraries and its results are not covered by round-trip guarantees.",
    });
    expect(EXPERIMENTS[1]).toEqual({
      id: "deck-designer",
      title: "Deck designer",
      summary: "Ask a language model you configure to design a deck from your notes.",
      risk: "Sends the text you put in front of the model to the endpoint you name; output is untrusted and provisional.",
    });
  });

  it("is frozen, array and entries alike", () => {
    expect(Object.isFrozen(EXPERIMENTS)).toBe(true);
    for (const experiment of EXPERIMENTS) {
      expect(Object.isFrozen(experiment)).toBe(true);
    }
    expect(() => (EXPERIMENTS as Experiment[]).push(EXPERIMENTS[0])).toThrow();
  });

  it("covers exactly the closed id union the stored flags are validated against", () => {
    expect(EXPERIMENTS.map((experiment) => experiment.id)).toEqual([...EXPERIMENT_IDS]);
    for (const experiment of EXPERIMENTS) {
      expect(experiment.title).not.toBe("");
      expect(experiment.summary).not.toBe("");
      expect(experiment.risk).not.toBe("");
    }
  });

  it("names deck-designer as the one experiment that uses the model gateway", () => {
    expect(MODEL_EXPERIMENTS).toEqual(["deck-designer"]);
    expect(Object.isFrozen(MODEL_EXPERIMENTS)).toBe(true);
  });
});

type Experiment = (typeof EXPERIMENTS)[number];

describe("defaults", () => {
  it("every experiment is off on a fresh profile", () => {
    for (const experiment of EXPERIMENTS) {
      expect(isExperimentEnabled(experiment.id)).toBe(false);
    }
    expect(PREFERENCE_DEFAULTS.experiments).toEqual({ "library-search": false, "deck-designer": false });
    expect(Object.isFrozen(EXPERIMENT_FLAG_DEFAULTS)).toBe(true);
  });

  it("a reset turns every experiment back off", () => {
    setExperimentEnabled("deck-designer", true);
    resetPreferences();
    expect(isExperimentEnabled("deck-designer")).toBe(false);
  });
});

describe("persistence", () => {
  it("lives under the preferences store's single experiments key", () => {
    setExperimentEnabled("library-search", true);
    expect(stored().experiments).toEqual({ "library-search": true, "deck-designer": false });
    // No key of its own: the preferences object is the only place flags live.
    expect(Object.keys(localStorage).filter((key) => key.includes("experiment"))).toEqual([]);
  });

  it("survives a fresh session (a reload) and can be turned off again", () => {
    setExperimentEnabled("deck-designer", true);
    resetPreferencesSession();
    expect(isExperimentEnabled("deck-designer")).toBe(true);
    expect(isExperimentEnabled("library-search")).toBe(false);
    setExperimentEnabled("deck-designer", false);
    expect(isExperimentEnabled("deck-designer")).toBe(false);
    expect(stored().experiments).toEqual({ "library-search": false, "deck-designer": false });
  });

  it("toggling one flag leaves the other and every other preference alone", () => {
    savePreference("theme", "dark");
    setExperimentEnabled("library-search", true);
    setExperimentEnabled("deck-designer", true);
    setExperimentEnabled("library-search", false);
    expect(loadPreferences().experiments).toEqual({ "library-search": false, "deck-designer": true });
    expect(loadPreferences().theme).toBe("dark");
  });

  it("a corrupt or non-boolean stored map repairs to off, without disturbing other fields", () => {
    localStorage.setItem(
      PREFERENCES_STORAGE_KEY,
      JSON.stringify({ theme: "dark", experiments: { "library-search": "yes", "deck-designer": 1 } }),
    );
    expect(loadPreferences().experiments).toEqual({ "library-search": false, "deck-designer": false });
    expect(loadPreferences().theme).toBe("dark");
    localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ experiments: ["deck-designer"] }));
    expect(loadPreferences().experiments).toEqual(EXPERIMENT_FLAG_DEFAULTS);
  });
});

describe("unknown ids", () => {
  it("an unknown id in stored preferences is ignored, and dropped on the next write", () => {
    localStorage.setItem(
      PREFERENCES_STORAGE_KEY,
      JSON.stringify({ experiments: { "time-travel": true, "deck-designer": true } }),
    );
    expect(loadPreferences().experiments).toEqual({ "library-search": false, "deck-designer": true });
    expect(isExperimentEnabled("time-travel" as ExperimentId)).toBe(false);
    // The next write persists only known ids: the stray flag is not resurrected.
    setExperimentEnabled("library-search", true);
    expect(stored().experiments).toEqual({ "library-search": true, "deck-designer": true });
  });

  it("setting an id outside the closed union does nothing", () => {
    setExperimentEnabled("time-travel" as ExperimentId, true);
    expect(stored().experiments).toBeUndefined();
    expect(loadPreferences().experiments).toEqual(EXPERIMENT_FLAG_DEFAULTS);
  });

  it("parseExperimentFlags keeps only known ids, true only when exactly true", () => {
    expect(parseExperimentFlags({ "deck-designer": true, extra: true })).toEqual({
      "library-search": false,
      "deck-designer": true,
    });
    expect(parseExperimentFlags(null)).toBeUndefined();
    expect(parseExperimentFlags("deck-designer")).toBeUndefined();
  });
});

describe("onExperimentsChange", () => {
  it("fires with a copy of the map when a flag changes, and stops after unsubscribe", () => {
    const listener = vi.fn();
    const stop = onExperimentsChange(listener);
    setExperimentEnabled("deck-designer", true);
    expect(listener).toHaveBeenCalledTimes(1);
    expect(listener).toHaveBeenLastCalledWith({ "library-search": false, "deck-designer": true });
    stop();
    setExperimentEnabled("deck-designer", false);
    expect(listener).toHaveBeenCalledTimes(1);
  });

  it("does not fire for a change to any other preference, or a no-op write", () => {
    const listener = vi.fn();
    const stop = onExperimentsChange(listener);
    savePreference("theme", "dark");
    setExperimentEnabled("library-search", false);
    expect(listener).not.toHaveBeenCalled();
    stop();
  });

  it("fires when a reset turns an experiment off", () => {
    setExperimentEnabled("library-search", true);
    const listener = vi.fn();
    const stop = onExperimentsChange(listener);
    resetPreferences();
    expect(listener).toHaveBeenCalledWith({ "library-search": false, "deck-designer": false });
    stop();
  });

  it("fires for a change another tab writes (the storage event)", () => {
    const listener = vi.fn();
    const stop = onExperimentsChange(listener);
    localStorage.setItem(
      PREFERENCES_STORAGE_KEY,
      JSON.stringify({ experiments: { "library-search": true, "deck-designer": false } }),
    );
    window.dispatchEvent(new StorageEvent("storage", { key: PREFERENCES_STORAGE_KEY }));
    expect(listener).toHaveBeenCalledWith({ "library-search": true, "deck-designer": false });
    stop();
  });

  it("hands each listener its own copy", () => {
    const seen: Record<ExperimentId, boolean>[] = [];
    const stop = onExperimentsChange((map) => {
      seen.push(map);
      map["deck-designer"] = false;
    });
    setExperimentEnabled("deck-designer", true);
    expect(isExperimentEnabled("deck-designer")).toBe(true);
    expect(seen).toHaveLength(1);
    stop();
  });
});
