// Multilevel list definitions (SPEC-032), framework-free like every lib
// module: the gallery presets, the per-level marker composition ("1.1.1",
// "(a)", …) and the instance-counter scan defined lists are numbered by.
// Product-neutral: everything here works over the core list types alone —
// only the model lookup (definitionById) stays in lib/listDefinitions.ts.
//
// Numbering semantics, distinct from SPEC-018's adjacency lists: a defined
// item belongs to a (defId, listId) *instance*. Counters live per instance
// and survive interruptions by non-list content — which is what makes
// "Continue previous numbering" work — while a fresh listId restarts a
// sequence ("Restart numbering", and every newly applied gallery preset).
// Markers are still always derived at render time, never stored, so
// deleting, moving or demoting an item renumbers the rest for free
// (AC-29..31).
import {
  BULLET_CHARS,
  clampListLevel,
  formatOrdinal,
  isBulletStyle,
  LIST_INDENT_STEP_CM,
  LIST_MARKER_WIDTH_CM,
  MAX_LIST_LEVEL,
  type ListDefinition,
  type ListFormatting,
  type ListLevelDefinition,
  type ListMarkerStyle,
} from "./lists";

/** A defined level's effective text position: its own, else the editor
 * default (level × 1.27 cm — the same default SPEC-018 items use). */
export function definedTextIndentCm(def: ListDefinition, level: number): number {
  return def.levels[level - 1]?.textIndentCm ?? level * LIST_INDENT_STEP_CM;
}

/** A defined level's effective number (marker) position: its own, else the
 * marker-box width before the text position. */
export function definedNumberIndentCm(def: ListDefinition, level: number): number {
  return (
    def.levels[level - 1]?.numberIndentCm ??
    Math.max(0, definedTextIndentCm(def, level) - LIST_MARKER_WIDTH_CM)
  );
}

// --- gallery presets (AC-1..4) ----------------------------------------------

export interface ListGalleryPreset {
  key: string;
  /** The menu entry's label. */
  label: string;
  /** A short per-level preview ("1., 1.1., 1.1.1."). */
  preview: string;
}

/** The four gallery presets, in menu order. The levels themselves come from
 * `presetLevels` — fresh copies per apply, so two lists started from the
 * same preset never share a definition (AC-18). */
export const GALLERY_PRESETS: ListGalleryPreset[] = [
  { key: "decimal", label: "1./1.1/1.1.1", preview: "1.  1.1.  1.1.1." },
  { key: "outline", label: "I./A./1.", preview: "I.  A.  1." },
  { key: "legal", label: "Legal — 1.1.1", preview: "1  1.1  1.1.1" },
  { key: "bulletMixed", label: "Bulleted (mixed)", preview: "1.  •  a." },
];

// The I./A./1. preset's style cycle: Word's outline preset for the first
// three levels, then the SPEC-018 default cycle for the rest.
const OUTLINE_STYLES: ListMarkerStyle[] = [
  "upperRoman",
  "upperAlpha",
  "decimal",
  "lowerAlpha",
  "lowerRoman",
  "decimal",
  "lowerAlpha",
  "lowerRoman",
  "decimal",
];

// The bullet-mixed preset alternates numbers and bullets by level — the
// "mixes bullets at some levels and numbers at others" reading of the
// intake's non-blocking ambiguity.
const BULLET_MIXED_STYLES: ListMarkerStyle[] = [
  "decimal",
  "disc",
  "lowerAlpha",
  "circle",
  "lowerRoman",
  "square",
  "decimal",
  "disc",
  "lowerAlpha",
];

/** A preset's nine level definitions, freshly built (never shared). */
export function presetLevels(key: string): ListLevelDefinition[] {
  const levels: ListLevelDefinition[] = [];
  for (let level = 1; level <= MAX_LIST_LEVEL; level++) {
    if (key === "decimal") {
      const def: ListLevelDefinition = { style: "decimal", suffix: "." };
      if (level > 1) {
        def.includeLevels = level;
      }
      levels.push(def);
    } else if (key === "legal") {
      const def: ListLevelDefinition = { style: "decimal" };
      if (level > 1) {
        def.includeLevels = level;
      }
      levels.push(def);
    } else if (key === "bulletMixed") {
      const style = BULLET_MIXED_STYLES[level - 1];
      levels.push(isBulletStyle(style) ? { style } : { style, suffix: "." });
    } else {
      // "outline" and the fallback for an unknown key.
      levels.push({ style: OUTLINE_STYLES[level - 1], suffix: "." });
    }
  }
  return levels;
}

/** The levels the "Define new multilevel list…" dialog starts from when the
 * selection is not already in a defined list: the decimal preset. */
export function defaultDefinitionLevels(): ListLevelDefinition[] {
  return presetLevels("decimal");
}

/** The heading-numbering definition "Number headings 1–3" links (AC-13):
 * the decimal outline preset with Heading 1–3 linked to levels 1–3. */
export function headingLinkedLevels(): ListLevelDefinition[] {
  return presetLevels("decimal");
}

/** The style links "Number headings 1–3" creates. */
export function headingStyleLinks(): Record<string, number> {
  return { Heading1: 1, Heading2: 2, Heading3: 3 };
}

// --- marker composition ------------------------------------------------------

/** The list type a defined level implies. */
export function definedLevelType(def: ListDefinition, level: number): "bullet" | "number" {
  const style = def.levels[clampListLevel(level) - 1]?.style ?? "decimal";
  return isBulletStyle(style) ? "bullet" : "number";
}

/** The marker text a defined item renders, from the per-level ordinals of
 * its instance: `prefix + components + suffix`, where the components are
 * the last `includeLevels` levels' ordinals — each formatted in its own
 * level's style — joined with "." (AC-5..8). A bullet level renders its
 * glyph between the prefix and suffix instead. */
export function definedMarkerText(
  def: ListDefinition,
  level: number,
  ordinals: readonly number[],
): string {
  const clamped = clampListLevel(level);
  const levelDef = def.levels[clamped - 1] ?? { style: "decimal" };
  const prefix = levelDef.prefix ?? "";
  const suffix = levelDef.suffix ?? "";
  if (isBulletStyle(levelDef.style)) {
    return `${prefix}${BULLET_CHARS[levelDef.style]}${suffix}`;
  }
  const include = Math.min(Math.max(Math.round(levelDef.includeLevels ?? 1), 1), clamped);
  const parts: string[] = [];
  for (let component = clamped - include + 1; component <= clamped; component++) {
    const componentDef = def.levels[component - 1];
    const style =
      componentDef && !isBulletStyle(componentDef.style) ? componentDef.style : "decimal";
    parts.push(formatOrdinal(style, ordinals[component - 1] ?? componentDef?.startAt ?? 1));
  }
  return `${prefix}${parts.join(".")}${suffix}`;
}

// --- instance counters -------------------------------------------------------

/** Counters for every defined-list instance seen so far in one scan:
 * instance key -> last ordinal used per level (index level-1). Created per
 * numbering context (the document body; each table cell) and threaded
 * through `markersForBlocks` (lib/lists.ts). */
export type DefinedListCounters = Map<string, Array<number | undefined>>;

export function createDefinedListCounters(): DefinedListCounters {
  return new Map();
}

/** An instance's counter key. */
export function instanceKey(defId: string, listId: string | undefined): string {
  return `${defId} ${listId ?? ""}`;
}

/** Advances an instance's counters for one item and returns its marker.
 * Deeper counters reset (returning to a shallower level restarts its
 * sub-numbering); an explicit `startAt` on the item restarts its level at
 * that value — "Set numbering value…" (AC-16) and the restart carried by an
 * imported `w:startOverride` / `text:start-value`. */
export function markerForDefinedItem(
  def: ListDefinition,
  list: ListFormatting,
  counters: DefinedListCounters,
): string {
  const level = clampListLevel(list.level);
  const key = instanceKey(list.defId ?? "", list.listId);
  let levels = counters.get(key);
  if (!levels) {
    levels = [];
    counters.set(key, levels);
  }
  for (let deeper = level; deeper < MAX_LIST_LEVEL; deeper++) {
    levels[deeper] = undefined;
  }
  // Materialise every shallower component the marker may show, so a
  // document that starts at level 2 still renders "1.1" and later items
  // continue consistently from the same implied ancestors.
  const ordinals: number[] = [];
  for (let component = 1; component < level; component++) {
    const value = levels[component - 1] ?? def.levels[component - 1]?.startAt ?? 1;
    levels[component - 1] = value;
    ordinals[component - 1] = value;
  }
  const previous = levels[level - 1];
  const ordinal =
    list.startAt !== undefined
      ? list.startAt
      : previous === undefined
        ? (def.levels[level - 1]?.startAt ?? 1)
        : previous + 1;
  levels[level - 1] = ordinal;
  ordinals[level - 1] = ordinal;
  return definedMarkerText(def, level, ordinals);
}
