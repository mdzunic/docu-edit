// The thin Vue wrapper over the framework-free preferences store
// (src/lib/preferences.ts) — SPEC-066. A module-level singleton so every
// view and component reads the same live reactive state: the editor, the
// library (HomeView), the Preferences dialog and each subsystem that reacts
// to a change (theme, ruler, autosave) all see one object.
//
// Components use THIS, never localStorage directly (AC-18, AC-46). Reads go
// through the reactive `preferences` object; writes go through
// `setPreference`, which persists via the store and — through the store's
// change notification — flows straight back into the reactive state, so a
// change made anywhere (this tab or, via the storage event, another tab)
// updates every consumer (AC-7, AC-44).
import { reactive, readonly, type DeepReadonly } from "vue";
import {
  loadPreferences,
  onPreferencesChange,
  resetPreferences as resetPreferencesStore,
  savePreference,
  type Preferences,
} from "../lib/preferences";

// Seeded from the store on first import — which also runs the one-time legacy
// migration (AC-10) as the app starts.
const state = reactive<Preferences>(loadPreferences());

let installed = false;

/** Copies the store's snapshot into the shared reactive state. */
function sync(prefs: Preferences): void {
  Object.assign(state, prefs);
}

/** Registers the store subscription exactly once. The listener is never
 * removed: the singleton lives for the whole app session, the same way
 * useTheme's OS-appearance listener does. */
function install(): void {
  if (installed) {
    return;
  }
  installed = true;
  onPreferencesChange(sync);
}

/** Sets one preference, applying it live everywhere (AC-7). */
function setPreference<K extends keyof Preferences>(key: K, value: Preferences[K]): void {
  savePreference(key, value);
}

/** Restores every preference to its default (AC-28). */
function resetPreferences(): void {
  resetPreferencesStore();
}

export function usePreferences() {
  install();

  return {
    /** The live, reactive preferences. Read-only: mutate through
     * `setPreference` so every write is validated and persisted. */
    preferences: readonly(state) as DeepReadonly<Preferences>,
    setPreference,
    resetPreferences,
  };
}
