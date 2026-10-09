// The spreadsheet's number locale (SPEC-118 "The preference"): the ONE place
// the `spreadsheetLocale` preference is resolved into the SpreadsheetLocale
// the grid, the formula bar and the status bar format with —
// lib/spreadsheet/numberFormat never reads preferences; the view hands the
// locale in. The locale is a per-user preference (SPEC-066), never a
// property of a file: changing it re-renders every open sheet (through
// usePreferences' live state, which onPreferencesChange keeps current across
// tabs) and changes no stored or downloaded byte.
import { computed } from "vue";
import { usePreferences } from "../../composables/usePreferences";
import { spreadsheetLocaleFor, type SpreadsheetLocaleTag } from "../../lib/spreadsheet/numberFormat/locale";

export function useSpreadsheetLocale() {
  const { preferences, setPreference } = usePreferences();

  /** The locale the preference names (an unknown one is the default). */
  const locale = computed(() => spreadsheetLocaleFor(preferences.spreadsheetLocale));
  const tag = computed((): SpreadsheetLocaleTag => locale.value.tag);

  /** Chooses the locale — a preference write, never a document edit. */
  function setLocale(next: SpreadsheetLocaleTag): void {
    setPreference("spreadsheetLocale", next);
  }

  return { locale, tag, setLocale };
}

export type SpreadsheetLocaleController = ReturnType<typeof useSpreadsheetLocale>;
