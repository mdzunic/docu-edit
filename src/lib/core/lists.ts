// The product-neutral half of bulleted/numbered lists (SPEC-018; multilevel
// definitions SPEC-032): the list-formatting types, marker styles, ordinal
// formatting and level planning. Everything here is independent of the
// document model — a Slides product numbers the list items inside a text
// box with the same marker math. The model-walking half (deriving every
// block's marker from the block sequence) stays in lib/lists.ts.

export type ListType = "bullet" | "number";

/** The marker styles the editor offers (SPEC-018 AC-12): three bullet
 * characters and five number formats. */
export type ListMarkerStyle =
  | "disc"
  | "circle"
  | "square"
  | "decimal"
  | "lowerAlpha"
  | "upperAlpha"
  | "lowerRoman"
  | "upperRoman";

/** A paragraph's membership in a bulleted or numbered list (SPEC-018).
 * List *identity* is adjacency: consecutive list paragraphs form one list,
 * and a non-list paragraph between two list items splits them into two
 * lists whose numbering restarts (AC-10, AC-11) — so no list id is stored.
 * Displayed numbers are always derived by scanning the block sequence
 * (lib/lists.ts), never persisted. */
export interface ListFormatting {
  type: ListType;
  /** 1-based nesting depth, 1..9 (AC-3, AC-4). */
  level: number;
  /** Explicit marker style for this item's level (AC-12, AC-13); absent
   * means the level's default cycle (AC-5, see below). */
  marker?: ListMarkerStyle;
  /** The number the item's list counts from at this level; absent means 1.
   * Carried on every item of an imported list so a restart after an
   * interruption returns to the list's own start value (AC-11, AC-22). */
  startAt?: number;
  /** The level's text indent in cm as defined by an imported numbering
   * definition (AC-20, AC-21); absent means the editor default
   * (level × 1.27 cm, see listIndentCm). */
  indentCm?: number;
  /** References a list definition (SPEC-032). Presence makes this a
   * *defined* list item: its marker text, per-level styles and indents all
   * come from the definition, and its numbering is counted per
   * (defId, listId) instance rather than by adjacency. Deliberately not
   * named `id` (see lib/equality.ts's IGNORED_KEYS rule). */
  defId?: string;
  /** The numbering instance this item counts in (SPEC-032): items sharing
   * (defId, listId) form one sequence that continues across interruptions;
   * a fresh listId is what "Restart numbering" mints. Meaningless without
   * `defId`. */
  listId?: string;
}

/** How a defined level's number is aligned within its marker box (SPEC-032
 * AC "alignment"; `w:lvlJc` / `fo:text-align` on the level). */
export type ListLevelAlignment = "left" | "center" | "right";

/** One level (1..9) of a multilevel list definition (SPEC-032). Every field
 * except `style` is optional with a documented default, so a definition
 * only stores what differs — which keeps save/re-open fixed points exact. */
export interface ListLevelDefinition {
  /** The level's marker style: a number format or a bullet glyph. */
  style: ListMarkerStyle;
  /** Literal text before the number (e.g. "("); absent means "". */
  prefix?: string;
  /** Literal text after the number (e.g. ")" or "."); absent means "". */
  suffix?: string;
  /** How many level components the number shows, ending at this level
   * (AC-8's "1.1" vs "1"): 3 at level 3 renders "1.1.1"; absent means 1. */
  includeLevels?: number;
  /** The level's start value; absent means 1. */
  startAt?: number;
  /** Number alignment; absent means "left". */
  alignment?: ListLevelAlignment;
  /** The marker's position (indent) in cm; absent means the editor default
   * (text position − 0.635 cm, see listDefinitions.ts). */
  numberIndentCm?: number;
  /** The item text's position in cm; absent means level × 1.27 cm. */
  textIndentCm?: number;
}

/** A multilevel list definition (SPEC-032): nine levels of numbering
 * format, referenced by `ListFormatting.defId`. Each list owns its own
 * definition (gallery presets are copied on apply), which is what makes
 * editing a level's format update that list only (AC-17, AC-18). */
export interface ListDefinition {
  /** Identity referenced by `ListFormatting.defId` — not named `id` for
   * the same equality.ts reason as `BookmarkAnchor.name`. */
  defId: string;
  /** Exactly MAX_LIST_LEVEL entries, level 1 first. */
  levels: ListLevelDefinition[];
  /** Paragraph styles linked to this definition's levels (AC-13):
   * `styleId -> level`. Applying a linked style makes the paragraph a
   * defined item of this definition at that level automatically. */
  styleLinks?: Record<string, number>;
}

/** Typed pass-through for a `ListLevelDefinition[]` arriving from a `.vue`
 * component's untyped Array prop (espree bars the annotation in `.vue`
 * scripts). */
export function asListLevelDefinitions(value: unknown): ListLevelDefinition[] {
  return value as ListLevelDefinition[];
}

/** Typed pass-through for a `ListDefinition[]` prop (espree constraint). */
export function asListDefinitions(value: unknown): ListDefinition[] {
  return value as ListDefinition[];
}

/** Typed empty seed for list-definition arrays in `.vue` scripts. */
export function noListDefinitions(): ListDefinition[] {
  return [];
}

/** Typed pass-through for a list type coming from a `.vue` template's plain
 * string (espree constraint). */
export function asListType(value: string): ListType {
  return value as ListType;
}

/** Typed pass-through for a marker style coming from a `.vue` template's
 * plain string (espree constraint). */
export function asListMarkerStyle(value: string): ListMarkerStyle {
  return value as ListMarkerStyle;
}

// --- marker styles and ordinal formatting (SPEC-018) ------------------------

/** Levels run 1..9 (AC-3: Tab at level 9 has no effect). */
export const MAX_LIST_LEVEL = 9;

/** Text indent per level — Word's half inch, the same step as
 * commands.INDENT_STEP_CM and an exact twip value (720). */
export const LIST_INDENT_STEP_CM = 1.27;

/** The marker box hanging before an item's text, in cm (360 twips). */
export const LIST_MARKER_WIDTH_CM = 0.635;

/** The bullet styles AC-12 offers, in dropdown order. */
export const BULLET_MARKER_STYLES: ListMarkerStyle[] = ["disc", "circle", "square"];

/** The number styles AC-12 offers, in dropdown order. */
export const NUMBER_MARKER_STYLES: ListMarkerStyle[] = [
  "decimal",
  "lowerAlpha",
  "upperAlpha",
  "lowerRoman",
  "upperRoman",
];

/** The three-marker cycles AC-5 defines: levels 1–3 use these in order and
 * levels 4–9 repeat them (level N uses entry (N-1) mod 3). */
const BULLET_CYCLE: ListMarkerStyle[] = ["disc", "circle", "square"];
const NUMBER_CYCLE: ListMarkerStyle[] = ["decimal", "lowerAlpha", "lowerRoman"];

export const BULLET_CHARS: Record<string, string> = {
  disc: "•",
  circle: "◦",
  square: "▪",
};

// Bullet characters seen in real files, mapped onto the three supported
// styles: Word writes Symbol/Wingdings private-use glyphs and "o" (Courier
// New) for circles; LibreOffice writes the plain glyphs. Shared by both
// formats' readers.
const BULLET_STYLE_BY_CHAR: Record<string, ListMarkerStyle> = {
  "•": "disc",
  "·": "disc",
  "\uf0b7": "disc", // Symbol-font bullet
  "◦": "circle",
  o: "circle",
  "○": "circle",
  "▪": "square",
  "■": "square",
  "\uf0a7": "square", // Wingdings small square
};

/** The supported bullet style a marker character maps onto, or undefined
 * for a character outside the recognised set (the caller falls back to the
 * level's default cycle). */
export function bulletStyleForChar(char: string): ListMarkerStyle | undefined {
  return BULLET_STYLE_BY_CHAR[char];
}

export function isBulletStyle(style: ListMarkerStyle): boolean {
  return style === "disc" || style === "circle" || style === "square";
}

/** The list type a marker style belongs to — picking a marker from either
 * dropdown also decides the items' type (AC-13). */
export function markerType(style: ListMarkerStyle): ListType {
  return isBulletStyle(style) ? "bullet" : "number";
}

/** AC-5's default marker for a type and level: •/◦/▪ or 1./a./i., cycling
 * every three levels. */
export function defaultMarkerStyle(type: ListType, level: number): ListMarkerStyle {
  const cycle = type === "bullet" ? BULLET_CYCLE : NUMBER_CYCLE;
  return cycle[(Math.max(1, level) - 1) % 3];
}

/** The marker style an item displays: its explicit marker when set and of
 * the item's own type, else the level default. */
export function markerStyleOf(list: ListFormatting): ListMarkerStyle {
  if (list.marker !== undefined && markerType(list.marker) === list.type) {
    return list.marker;
  }
  return defaultMarkerStyle(list.type, list.level);
}

/** An item's text indent in cm: the imported per-level indent when present,
 * else level × 1.27 cm. */
export function listIndentCm(list: ListFormatting): number {
  return list.indentCm ?? list.level * LIST_INDENT_STEP_CM;
}

const ROMAN_DIGITS: Array<[number, string]> = [
  [1000, "m"],
  [900, "cm"],
  [500, "d"],
  [400, "cd"],
  [100, "c"],
  [90, "xc"],
  [50, "l"],
  [40, "xl"],
  [10, "x"],
  [9, "ix"],
  [5, "v"],
  [4, "iv"],
  [1, "i"],
];

function toRoman(n: number): string {
  let value = Math.max(1, Math.floor(n));
  let out = "";
  for (const [digit, letters] of ROMAN_DIGITS) {
    while (value >= digit) {
      out += letters;
      value -= digit;
    }
  }
  return out;
}

/** 1 -> "a", 26 -> "z", 27 -> "aa" — the bijective base-26 alphabetics Word
 * uses for lowerLetter/upperLetter numbering. */
function toAlpha(n: number): string {
  let value = Math.max(1, Math.floor(n));
  let out = "";
  while (value > 0) {
    value -= 1;
    out = String.fromCharCode(97 + (value % 26)) + out;
    value = Math.floor(value / 26);
  }
  return out;
}

/** A number formatted in a marker style, without the "." suffix. */
export function formatOrdinal(style: ListMarkerStyle, ordinal: number): string {
  switch (style) {
    case "lowerAlpha":
      return toAlpha(ordinal);
    case "upperAlpha":
      return toAlpha(ordinal).toUpperCase();
    case "lowerRoman":
      return toRoman(ordinal);
    case "upperRoman":
      return toRoman(ordinal).toUpperCase();
    default:
      return String(ordinal);
  }
}

/** The display text of one marker: the bullet character, or the formatted
 * ordinal followed by "." (AC-5: `1.`, `a.`, `i.`). */
export function markerText(style: ListMarkerStyle, ordinal: number): string {
  if (isBulletStyle(style)) {
    return BULLET_CHARS[style];
  }
  return `${formatOrdinal(style, ordinal)}.`;
}

export function clampListLevel(level: number): number {
  return Math.min(Math.max(Math.round(level), 1), MAX_LIST_LEVEL);
}

/** One level of a list run's exported definition (both writers): the level
 * is defined by the first item in the run at that level; levels no item
 * uses fall back to the run's leading type with the AC-5 default cycle. */
export interface PlannedListLevel {
  type: ListType;
  style: ListMarkerStyle;
  startAt: number;
  indentCm: number;
}

/** Plans the nine level definitions a contiguous run of list items exports
 * as one list definition (a .docx abstractNum / an .odt text:list-style). */
export function planListLevels(items: ListFormatting[]): PlannedListLevel[] {
  const baseType = items[0].type;
  const levels: PlannedListLevel[] = [];
  for (let level = 1; level <= MAX_LIST_LEVEL; level++) {
    const first = items.find((item) => clampListLevel(item.level) === level);
    const type = first ? first.type : baseType;
    levels.push({
      type,
      style: first ? markerStyleOf(first) : defaultMarkerStyle(type, level),
      startAt: first?.startAt ?? 1,
      indentCm: first?.indentCm ?? level * LIST_INDENT_STEP_CM,
    });
  }
  return levels;
}

/** The autoformat prefixes (AC-14, AC-15): the exact paragraph text after
 * the triggering space, mapped to the list type it creates. */
const AUTOFORMAT_PREFIXES: Record<string, ListType> = {
  "- ": "bullet",
  "* ": "bullet",
  "1. ": "number",
  "1) ": "number",
};

/** The list type a paragraph text autoformats into, or null when the text
 * is not exactly one of the trigger prefixes. The caller checks the rest
 * (paragraph was empty before the prefix, caret at its end, not already a
 * list item) — this only recognises the text (AC-14, AC-15, AC-17). */
export function detectListAutoformat(paragraphText: string): ListType | null {
  return AUTOFORMAT_PREFIXES[paragraphText] ?? null;
}
