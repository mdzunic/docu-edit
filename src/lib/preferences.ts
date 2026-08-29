// The centralized application-preferences store (SPEC-066). One
// localStorage key (`docu-edit:preferences`) holding a single JSON object,
// one declared schema with a default and a validator per field, and one
// tolerant load path: a missing, corrupt or partially-populated stored
// object falls back per field to that field's default, and one bad field
// never discards the rest. This is the single typed module every preference
// is read through — components go through the `usePreferences` composable
// (src/composables/usePreferences.ts), never localStorage directly.
//
// Framework-free and unit-testable: no Vue imports. It mirrors the tolerant,
// storage-guarded shape of the per-key modules it replaces
// (lib/themePrefs.ts, lib/core/autocorrect.ts, lib/core/commentIdentity.ts):
// a blocked or full localStorage (SPEC-003) degrades to in-session
// preferences and shows the storage warning, and never throws.
import type { MeasurementUnit } from "./core/units";
import type { Margins } from "./core/pageSetup";
import { PAPER_SIZE_NAMES } from "./core/pageSetup";
import { AUTOSAVE_DEBOUNCE_MS } from "./core/autosaveScheduler";
import { DEFAULT_LANGUAGE, SUPPORTED_LANGUAGES } from "./core/language";
import { STANDARD_FONT_FAMILIES } from "./core/fonts";
import { MAX_FONT_SIZE_PT, MIN_FONT_SIZE_PT } from "./fonts";
import { DEFAULT_ZOOM, INPUT_MAX, INPUT_MIN } from "./zoom";
import { DEFAULT_COMMENT_AUTHOR, loadCommentAuthor, COMMENT_AUTHOR_STORAGE_KEY } from "./core/commentIdentity";
import { AUTOCORRECT_STORAGE_KEY, loadAutocorrectSettings } from "./core/autocorrect";
import {
  DIM_PAGE_STORAGE_KEY,
  THEME_STORAGE_KEY,
  loadDimPage,
  loadThemePreference,
} from "./themePrefs";

/** The one key every preference lives under (AC-9). */
export const PREFERENCES_STORAGE_KEY = "docu-edit:preferences";

// The SPEC-003 storage-warning hook. documentStore owns the warning banner and
// registers its markStorageUnavailable here at load; keeping this a registered
// callback (rather than importing documentStore) avoids a documentStore ⇄
// preferences import cycle, which would leave this module half-initialised when
// useTheme/usePreferences call loadPreferences at their own module load.
let storageFailureHandler: (() => void) | undefined;

/** Registers the handler run when a preferences write hits a blocked or full
 * localStorage (AC-21) — documentStore wires its warning banner through it. */
export function setPreferencesStorageFailureHandler(handler: () => void): void {
  storageFailureHandler = handler;
}

/** A zoom that is a fixed percentage, or one of the two fit modes (SPEC-063). */
export type DefaultZoom = number | "fit-width" | "fit-page";

/** What a fresh download defaults to when the user has not chosen: the
 * document's own source format, or a fixed target (SPEC-006/007). */
export type DownloadFormatPreference = "source" | "docx" | "odt";

/** The full, typed set of application preferences (SPEC-066). Every field is
 * required here; a stored object that omits, mistypes or corrupts any of
 * them is repaired per field on read (see `loadPreferences`). */
export interface Preferences {
  theme: "system" | "light" | "dark";
  dimPage: boolean;
  measurementUnit: MeasurementUnit;
  defaultZoom: DefaultZoom;
  defaultViewMode: "print" | "web";
  showRuler: boolean;
  showFormattingMarks: boolean;
  smartQuotes: boolean;
  smartDashes: boolean;
  spellCheck: boolean;
  defaultLanguage: string;
  authorName: string;
  defaultFontFamily: string;
  defaultFontSizePt: number;
  defaultPageSizeId: string;
  defaultMarginsId: string;
  autosave: boolean;
  autosaveDelayMs: number;
  keepVersionHistory: boolean;
  defaultDownloadFormat: DownloadFormatPreference;
  showConversionReport: boolean;
  /** Whether the presentation editor's speaker-notes panel is expanded
   * (SPEC-083). Per-device, default collapsed. */
  presentationNotesOpen: boolean;
  /** Whether the presentation editor's slide-transition panel is expanded
   * (SPEC-085). Per-device, default collapsed. */
  presentationTransitionsOpen: boolean;
  /** Whether the presentation editor's animations panel is expanded
   * (SPEC-086). Per-device, default collapsed. */
  presentationAnimationsOpen: boolean;
  /** Snap a dragged drawing to alignment guides (SPEC-072). Default on. */
  snapToGuides: boolean;
  /** Snap a dragged drawing to the grid (SPEC-072). Default off. */
  snapToGrid: boolean;
  /** The snap-to-grid spacing in centimetres (SPEC-072). Default 0.25cm. */
  gridSpacingCm: number;
}

// --- margin presets ---------------------------------------------------------
// SPEC-027 (page setup) defines page sizes but never named margin presets, so
// this spec introduces the small set its "Default margins" preference offers.
// "Normal" is exactly the product's built-in margin default (2.54 cm all
// round, `STANDARD_MARGINS` in documentModel.ts), so a new document created
// with the default preference is byte-identical to one created before.

/** The named margin presets the Preferences dialog offers (SPEC-066). */
export const MARGIN_PRESETS: Record<string, Margins> = {
  Normal: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
  Narrow: { topCm: 1.27, rightCm: 1.27, bottomCm: 1.27, leftCm: 1.27 },
  Moderate: { topCm: 2.54, rightCm: 1.91, bottomCm: 2.54, leftCm: 1.91 },
  Wide: { topCm: 2.54, rightCm: 5.08, bottomCm: 2.54, leftCm: 5.08 },
};

/** Display order for the margin presets. */
export const MARGIN_PRESET_NAMES = ["Normal", "Narrow", "Moderate", "Wide"];

/** The page-size presets a new document may default to — every SPEC-027 named
 * size except the "Custom" sentinel, which is not a preset. */
export const PAGE_SIZE_PRESET_NAMES: string[] = PAPER_SIZE_NAMES.filter((name) => name !== "Custom");

/** The autosave-delay choices the dialog offers (SPEC-010). The default value
 * (the SPEC-010 debounce, 2 s) is included so it is a selectable choice even
 * though the spec's headline list is 1/3/5/10 s. */
export const AUTOSAVE_DELAY_CHOICES: number[] = [1000, AUTOSAVE_DEBOUNCE_MS, 3000, 5000, 10000].filter(
  (value, index, all) => all.indexOf(value) === index,
);

/** The declared default for every field (AC-6, AC-11). These match the
 * product's existing hard-coded defaults exactly, so wiring a consumer to
 * read its preference changes no behaviour until the user changes the
 * preference. */
export const PREFERENCE_DEFAULTS: Preferences = {
  theme: "system",
  dimPage: false,
  measurementUnit: "cm",
  defaultZoom: DEFAULT_ZOOM,
  defaultViewMode: "print",
  showRuler: true,
  showFormattingMarks: false,
  smartQuotes: true,
  smartDashes: true,
  spellCheck: true,
  defaultLanguage: DEFAULT_LANGUAGE,
  authorName: DEFAULT_COMMENT_AUTHOR,
  defaultFontFamily: "Calibri",
  defaultFontSizePt: 11,
  defaultPageSizeId: "A4",
  defaultMarginsId: "Normal",
  autosave: true,
  autosaveDelayMs: AUTOSAVE_DEBOUNCE_MS,
  keepVersionHistory: true,
  defaultDownloadFormat: "source",
  showConversionReport: true,
  presentationNotesOpen: false,
  presentationTransitionsOpen: false,
  presentationAnimationsOpen: false,
  snapToGuides: true,
  snapToGrid: false,
  gridSpacingCm: 0.25,
};

/** The snap-to-grid spacing bounds (SPEC-072): fine enough to be useful, never
 * zero or negative (which would divide by zero when snapping). */
export const MIN_GRID_SPACING_CM = 0.05;
export const MAX_GRID_SPACING_CM = 5;

// --- validators -------------------------------------------------------------
// One validator per field. Each takes the raw stored value (of unknown type)
// and returns the accepted value, or `undefined` to mean "reject — use the
// default". A wrong type, or a value outside its declared set, is rejected;
// the field is then repaired to its default without disturbing any other
// field (AC-13, AC-14).

type Validator<K extends keyof Preferences> = (raw: unknown) => Preferences[K] | undefined;

function isBoolean(raw: unknown): boolean | undefined {
  return typeof raw === "boolean" ? raw : undefined;
}

function oneOf<K extends keyof Preferences>(values: readonly Preferences[K][]): Validator<K> {
  return (raw) => (values.includes(raw as Preferences[K]) ? (raw as Preferences[K]) : undefined);
}

function nonEmptyString(raw: unknown): string | undefined {
  return typeof raw === "string" && raw.trim().length > 0 ? raw : undefined;
}

function inSet(raw: unknown, values: readonly string[]): string | undefined {
  return typeof raw === "string" && values.includes(raw) ? raw : undefined;
}

const VALIDATORS: { [K in keyof Preferences]: Validator<K> } = {
  theme: oneOf<"theme">(["system", "light", "dark"]),
  dimPage: isBoolean,
  measurementUnit: oneOf<"measurementUnit">(["cm", "mm", "in", "pt"]),
  defaultZoom: (raw) => {
    if (raw === "fit-width" || raw === "fit-page") {
      return raw;
    }
    return typeof raw === "number" && Number.isFinite(raw) && raw >= INPUT_MIN && raw <= INPUT_MAX
      ? raw
      : undefined;
  },
  defaultViewMode: oneOf<"defaultViewMode">(["print", "web"]),
  showRuler: isBoolean,
  showFormattingMarks: isBoolean,
  smartQuotes: isBoolean,
  smartDashes: isBoolean,
  spellCheck: isBoolean,
  defaultLanguage: (raw) => inSet(raw, SUPPORTED_LANGUAGES.map((option) => option.tag)),
  authorName: nonEmptyString,
  defaultFontFamily: (raw) => inSet(raw, STANDARD_FONT_FAMILIES),
  defaultFontSizePt: (raw) =>
    typeof raw === "number" && Number.isFinite(raw) && raw >= MIN_FONT_SIZE_PT && raw <= MAX_FONT_SIZE_PT
      ? raw
      : undefined,
  defaultPageSizeId: (raw) => inSet(raw, PAGE_SIZE_PRESET_NAMES),
  defaultMarginsId: (raw) => inSet(raw, MARGIN_PRESET_NAMES),
  autosave: isBoolean,
  autosaveDelayMs: (raw) =>
    typeof raw === "number" && AUTOSAVE_DELAY_CHOICES.includes(raw) ? raw : undefined,
  keepVersionHistory: isBoolean,
  defaultDownloadFormat: oneOf<"defaultDownloadFormat">(["source", "docx", "odt"]),
  showConversionReport: isBoolean,
  presentationNotesOpen: isBoolean,
  presentationTransitionsOpen: isBoolean,
  presentationAnimationsOpen: isBoolean,
  snapToGuides: isBoolean,
  snapToGrid: isBoolean,
  gridSpacingCm: (raw) =>
    typeof raw === "number" && Number.isFinite(raw) && raw >= MIN_GRID_SPACING_CM && raw <= MAX_GRID_SPACING_CM
      ? raw
      : undefined,
};

const PREFERENCE_KEYS = Object.keys(PREFERENCE_DEFAULTS) as (keyof Preferences)[];

/** Repairs a partial/untrusted stored object into a full, valid Preferences:
 * every field is taken from `stored` when its validator accepts it, else from
 * the declared default. This is where AC-11..AC-14 live. */
function validate(stored: Record<string, unknown>): Preferences {
  const result = { ...PREFERENCE_DEFAULTS };
  for (const key of PREFERENCE_KEYS) {
    const validated = VALIDATORS[key](stored[key]);
    if (validated !== undefined) {
      // The per-key generic is erased through the loop; the validator's own
      // return type guarantees the value matches result[key].
      (result[key] as Preferences[typeof key]) = validated;
    }
  }
  return result;
}

// --- storage ----------------------------------------------------------------

/** The last preferences applied this session. Read only as the fallback when
 * localStorage is blocked (AC-19, AC-20): a value set this session still
 * applies even though it could not be persisted. */
let sessionPrefs: Preferences | null = null;

const listeners = new Set<(prefs: Preferences) => void>();

/** Reads and parses the stored object. `getItem` throwing (blocked storage)
 * propagates so the caller can fall back to the in-session copy; corrupt JSON
 * or a non-object payload degrades to an empty object so every field defaults
 * (AC-12). */
function readRaw(): Record<string, unknown> {
  const raw = localStorage.getItem(PREFERENCES_STORAGE_KEY);
  if (!raw) {
    return {};
  }
  try {
    const parsed: unknown = JSON.parse(raw);
    return parsed !== null && typeof parsed === "object" ? (parsed as Record<string, unknown>) : {};
  } catch {
    return {};
  }
}

/** Folds the legacy per-key preferences into `stored` (only for fields the
 * stored object does not already carry) and removes each legacy key it
 * consumed (AC-10). Idempotent: once the legacy keys are gone this is a no-op,
 * so it can run on every load without a "have I migrated?" flag. */
function migrateLegacyKeys(stored: Record<string, unknown>): Record<string, unknown> {
  const merged = { ...stored };
  const consumed: string[] = [];

  const adopt = (legacyKey: string, apply: () => void): void => {
    if (localStorage.getItem(legacyKey) === null) {
      return;
    }
    apply();
    consumed.push(legacyKey);
  };

  adopt(THEME_STORAGE_KEY, () => {
    if (merged.theme === undefined) {
      merged.theme = loadThemePreference();
    }
  });
  adopt(DIM_PAGE_STORAGE_KEY, () => {
    if (merged.dimPage === undefined) {
      merged.dimPage = loadDimPage();
    }
  });
  adopt(AUTOCORRECT_STORAGE_KEY, () => {
    const settings = loadAutocorrectSettings();
    if (merged.smartQuotes === undefined) {
      merged.smartQuotes = settings.smartQuotes;
    }
    if (merged.smartDashes === undefined) {
      merged.smartDashes = settings.dashes;
    }
  });
  adopt(COMMENT_AUTHOR_STORAGE_KEY, () => {
    if (merged.authorName === undefined) {
      merged.authorName = loadCommentAuthor();
    }
  });

  if (consumed.length > 0) {
    // Persist the folded-in object, then drop the legacy keys, so the
    // migration happens exactly once and the old keys never linger.
    try {
      localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify(validate(merged)));
      for (const key of consumed) {
        localStorage.removeItem(key);
      }
    } catch {
      // A blocked write leaves the legacy keys in place to try again next
      // load; the in-session merge still applies.
      storageFailureHandler?.();
    }
  }

  return merged;
}

/** Reads the full, validated preferences. Missing/corrupt/partial storage
 * falls back per field to the default (AC-11..14); a blocked localStorage
 * degrades to the in-session copy (or the defaults) without throwing
 * (AC-19, AC-20). Legacy keys are migrated on first read (AC-10). */
export function loadPreferences(): Preferences {
  try {
    const stored = migrateLegacyKeys(readRaw());
    const prefs = validate(stored);
    sessionPrefs = prefs;
    return { ...prefs };
  } catch {
    return sessionPrefs ? { ...sessionPrefs } : { ...PREFERENCE_DEFAULTS };
  }
}

function persist(prefs: Preferences): void {
  try {
    localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify(prefs));
  } catch {
    // Blocked or full: the in-session copy (set by the caller) still applies,
    // and the SPEC-003 warning banner is shown (AC-21).
    storageFailureHandler?.();
  }
}

function notify(prefs: Preferences): void {
  for (const listener of listeners) {
    listener({ ...prefs });
  }
}

/** Sets one preference, applies it live (in-session and, best-effort, to
 * storage) and notifies every listener (AC-7). An invalid value is coerced to
 * that field's default rather than stored. */
export function savePreference<K extends keyof Preferences>(key: K, value: Preferences[K]): void {
  const current = loadPreferences();
  const validated = VALIDATORS[key](value);
  const next: Preferences = { ...current, [key]: validated !== undefined ? validated : PREFERENCE_DEFAULTS[key] };
  sessionPrefs = next;
  persist(next);
  notify(next);
}

/** Restores every preference to its default (AC-28). Writes only the one
 * preferences key, so documents, drafts, the custom dictionary and version
 * history — each under its own key — are untouched (AC-29..32). */
export function resetPreferences(): void {
  const next = { ...PREFERENCE_DEFAULTS };
  sessionPrefs = next;
  persist(next);
  notify(next);
}

/** Subscribes to preference changes — how the theme controller, the ruler and
 * the autosave scheduler react without polling. Fires for changes made in
 * this tab (via `savePreference`/`resetPreferences`) and, through the
 * `storage` event, for changes made in another tab so two open tabs stay
 * consistent (AC-44). Returns an unsubscribe. */
export function onPreferencesChange(listener: (prefs: Preferences) => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Clears the in-session fallback copy. Used when starting a fresh session
 * and by tests that exercise the blocked-storage fallback in isolation. */
export function resetPreferencesSession(): void {
  sessionPrefs = null;
}

// Cross-tab synchronisation (AC-44): another tab writing the preferences key
// (or clearing storage, key === null) re-reads and notifies this tab's
// listeners. Same-tab writes notify directly from savePreference and never
// fire a storage event, so there is no double-notification.
if (typeof window !== "undefined") {
  window.addEventListener("storage", (event) => {
    if (event.key !== null && event.key !== PREFERENCES_STORAGE_KEY) {
      return;
    }
    notify(loadPreferences());
  });
}
