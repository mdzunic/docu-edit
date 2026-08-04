// Colour data and parsing behind SPEC-013's text-colour and highlight
// pickers. Framework-free (no Vue imports) per the src/lib convention.
import { HIGHLIGHT_HEX } from "./odf/properties";

/** The text colour picker's standard palette (SPEC-013 AC): a greyscale
 * ramp plus two value rows over ten hues — the layout users know from
 * mainstream editors. Rows of `TEXT_COLOR_COLUMNS`, lowercase #rrggbb. */
export const TEXT_COLOR_COLUMNS = 10;
export const TEXT_COLOR_PALETTE: string[] = [
  "#000000", "#434343", "#666666", "#999999", "#b7b7b7",
  "#cccccc", "#d9d9d9", "#efefef", "#f3f3f3", "#ffffff",
  "#980000", "#ff0000", "#ff9900", "#ffff00", "#00ff00",
  "#00ffff", "#4a86e8", "#0000ff", "#9900ff", "#ff00ff",
  "#85200c", "#cc0000", "#e69138", "#f1c232", "#6aa84f",
  "#45818e", "#3d85c6", "#0b5394", "#674ea7", "#a64d79",
];

/** Parses the custom colour input (SPEC-013 AC): "#rrggbb", "rrggbb",
 * "#rgb", or "rgb(r, g, b)" — normalised to lowercase "#rrggbb", or
 * undefined for anything unparseable. */
export function parseColorInput(input: string): string | undefined {
  const text = input.trim().toLowerCase();
  const hex6 = /^#?([0-9a-f]{6})$/.exec(text);
  if (hex6) {
    return `#${hex6[1]}`;
  }
  const hex3 = /^#?([0-9a-f]{3})$/.exec(text);
  if (hex3) {
    return `#${hex3[1].split("").map((c) => c + c).join("")}`;
  }
  const rgb = /^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/.exec(text);
  if (rgb) {
    const channels = rgb.slice(1, 4).map(Number);
    if (channels.some((value) => value > 255)) {
      return undefined;
    }
    return `#${channels.map((value) => value.toString(16).padStart(2, "0")).join("")}`;
  }
  return undefined;
}

// --- recent colours ---------------------------------------------------------

/** How many recent colours the picker retains (SPEC-013 AC). */
export const RECENT_COLORS_LIMIT = 8;

const RECENT_COLORS_KEY = "docu-edit:recent-text-colors";

/** The recently applied text colours, most recent first — persisted in
 * localStorage so they survive reloads and apply across documents. Returns
 * [] where storage is unavailable (private mode, quota). */
export function loadRecentColors(): string[] {
  try {
    const raw = localStorage.getItem(RECENT_COLORS_KEY);
    if (!raw) {
      return [];
    }
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed
      .filter((value): value is string => typeof value === "string" && /^#[0-9a-f]{6}$/.test(value))
      .slice(0, RECENT_COLORS_LIMIT);
  } catch {
    return [];
  }
}

/** Records an applied colour and returns the updated list: moved to the
 * front, de-duplicated, capped at RECENT_COLORS_LIMIT. */
export function pushRecentColor(hex: string, current: string[]): string[] {
  const next = [hex, ...current.filter((value) => value !== hex)].slice(0, RECENT_COLORS_LIMIT);
  try {
    localStorage.setItem(RECENT_COLORS_KEY, JSON.stringify(next));
  } catch {
    // Storage full/unavailable: the in-memory list still updates.
  }
  return next;
}

// --- highlight picker -------------------------------------------------------

export interface HighlightPickerItem {
  /** The OOXML colour name the model stores, e.g. "yellow". */
  name: string;
  /** Word's UI label, e.g. "Bright Green". */
  label: string;
  /** The swatch colour. */
  hex: string;
}

// Word's picker order and labels. `white` is deliberately absent: the
// read/write palette has 16 entries but the picker offers 15 (SPEC-013).
const HIGHLIGHT_LABELS: Array<[string, string]> = [
  ["yellow", "Yellow"],
  ["green", "Bright Green"],
  ["cyan", "Turquoise"],
  ["magenta", "Pink"],
  ["blue", "Blue"],
  ["red", "Red"],
  ["darkBlue", "Dark Blue"],
  ["darkCyan", "Teal"],
  ["darkGreen", "Green"],
  ["darkMagenta", "Violet"],
  ["darkRed", "Dark Red"],
  ["darkYellow", "Dark Yellow"],
  ["darkGray", "Gray-50%"],
  ["lightGray", "Gray-25%"],
  ["black", "Black"],
];

/** The 15 Word highlight colours the picker offers, in Word's order
 * (SPEC-013 AC); "No colour" is the picker's 16th, non-colour action. */
export const HIGHLIGHT_PICKER_ITEMS: HighlightPickerItem[] = HIGHLIGHT_LABELS.map(
  ([name, label]) => ({ name, label, hex: HIGHLIGHT_HEX[name] }),
);

/** Typed default for a highlight-select handler parameter in a `.vue`
 * script (espree bars the `string | false` annotation there): the picker
 * emits an OOXML name or `false` for "No colour". */
export function noHighlight(): string | false {
  return false;
}
