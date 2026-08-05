// Bulleted/numbered list logic over the document model (SPEC-018),
// framework-free like every lib module. Markers are never stored as text:
// the displayed marker of every item is derived here from the block
// sequence, which is what makes numbering continuous across consecutive
// items (AC-10) and restart after an interruption (AC-11) without any
// bookkeeping in the editing commands.
import type {
  DocumentModel,
  ListFormatting,
  ListMarkerStyle,
  ListType,
  Paragraph,
} from "./documentModel";

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

/** Every block's displayed marker, per section, in block order — "" for
 * non-list paragraphs. One scan of the document implements the numbering
 * semantics: consecutive items at a level count up from the list's start
 * value (AC-10, AC-22), any non-list paragraph resets every counter
 * (AC-11), and returning to a shallower level resets the deeper counters,
 * so a later sub-list numbers from its start again. */
export function listMarkers(model: DocumentModel): string[][] {
  return model.sections.map((section) => {
    // counters[level] is the last ordinal used at that level within the
    // current uninterrupted run of list paragraphs.
    let counters: Array<number | undefined> = [];
    return section.blocks.map((block) => {
      const list = block.list;
      if (!list) {
        counters = [];
        return "";
      }
      const level = clampListLevel(list.level);
      for (let deeper = level + 1; deeper < counters.length; deeper++) {
        counters[deeper] = undefined;
      }
      const style = markerStyleOf(list);
      if (list.type === "bullet") {
        return markerText(style, 0);
      }
      const ordinal = (counters[level] ?? (list.startAt ?? 1) - 1) + 1;
      counters[level] = ordinal;
      return markerText(style, ordinal);
    });
  });
}

export function clampListLevel(level: number): number {
  return Math.min(Math.max(Math.round(level), 1), MAX_LIST_LEVEL);
}

/** The contiguous run of list paragraphs around `block` in its section —
 * what AC-13 calls "the current list". Returns the inclusive [first, last]
 * block range, or null when the block is not a list item. */
export function listRunAround(
  model: DocumentModel,
  section: number,
  block: number,
): { first: number; last: number } | null {
  const blocks = model.sections[section]?.blocks;
  if (!blocks || !blocks[block]?.list) {
    return null;
  }
  let first = block;
  while (first > 0 && blocks[first - 1].list) {
    first--;
  }
  let last = block;
  while (last + 1 < blocks.length && blocks[last + 1].list) {
    last++;
  }
  return { first, last };
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

/** Whether a paragraph is a list item — a null-safe convenience for `.vue`
 * scripts, which espree bars from optional-chaining into typed fields. */
export function isListItem(paragraph: Paragraph | undefined): boolean {
  return !!paragraph?.list;
}
