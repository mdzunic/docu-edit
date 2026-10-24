// The live Labs flags for components (SPEC-113) — a reactive mirror of
// lib/experiments.ts, kept current through its onExperimentsChange, so a gate
// renders the moment its experiment is switched on (here or in another tab)
// and disappears the moment it is switched off. A module-level singleton, like
// usePreferences: every gate reads the same object.
import { reactive, readonly } from "vue";
import {
  EXPERIMENTS,
  isExperimentEnabled,
  isExperimentId,
  onExperimentsChange,
  setExperimentEnabled,
  type ExperimentFlags,
} from "../lib/experiments";

function currentFlags(): ExperimentFlags {
  return Object.fromEntries(EXPERIMENTS.map((experiment) => [experiment.id, isExperimentEnabled(experiment.id)])) as ExperimentFlags;
}

const flags = reactive<ExperimentFlags>(currentFlags());

let installed = false;

/** Subscribes exactly once, re-reading first so a change made before the
 * first component asked is not missed. Never unsubscribed: the singleton
 * lives for the whole session. */
function install(): void {
  if (installed) {
    return;
  }
  installed = true;
  Object.assign(flags, currentFlags());
  onExperimentsChange((map) => {
    Object.assign(flags, map);
  });
}

export function useExperiments() {
  install();
  return {
    /** The live flag map. Read-only: switch through `setEnabled`. */
    flags: readonly(flags),
    /** Whether `id` names an experiment that is on. A string, not an
     * ExperimentId, because .vue templates and runtime props pass plain
     * strings; anything outside the closed union is simply off. */
    isEnabled(id: string): boolean {
      return isExperimentId(id) && flags[id] === true;
    },
    setEnabled(id: string, on: boolean): void {
      if (isExperimentId(id)) {
        setExperimentEnabled(id, on);
      }
    },
  };
}
