// Typed vocabulary and event readers for SPEC-097's slide text controls — the
// list popover (SlideListPopover.vue), the Text panel (TextPanel.vue) and the
// slide link dialog (SlideLinkDialog.vue). The espree convention bars type
// syntax in `.vue` scripts, so the option lists, the numeric event reader and
// the link-target assembly live here (the slideToolbarState.ts pattern).
import { markerText, type ListMarkerStyle } from "../../lib/core/lists";
import { CHARACTER_CATEGORIES, displayName, isGlyphless } from "../../lib/core/specialCharacters";
import { isDangerousUrl } from "../../lib/core/urls";
import type { SlideLinkTarget } from "../../lib/presentation/slideLinks";

/** One `list-style-{style}` button: the style, its marker previewed through
 * `markerText(style, 1)` and an accessible name. */
export interface ListStyleChoice {
  style: ListMarkerStyle;
  preview: string;
  label: string;
}

const LIST_STYLE_LABELS: Record<ListMarkerStyle, string> = {
  disc: "Filled bullet",
  circle: "Hollow bullet",
  square: "Square bullet",
  decimal: "Numbers",
  lowerAlpha: "Lowercase letters",
  upperAlpha: "Uppercase letters",
  lowerRoman: "Lowercase roman numerals",
  upperRoman: "Uppercase roman numerals",
};

/** core/lists.ts's eight marker styles, in its declared order. */
export const LIST_STYLE_CHOICES: ListStyleChoice[] = (Object.keys(LIST_STYLE_LABELS) as ListMarkerStyle[]).map((style) => ({
  style,
  preview: markerText(style, 1),
  label: LIST_STYLE_LABELS[style],
}));

/** One custom-bullet glyph from the special-character catalogue. */
export interface BulletCharChoice {
  char: string;
  name: string;
}

/** The custom-bullet grid: core/specialCharacters.ts's arrows and symbols
 * (glyphless code points dropped) — a compact, single-grapheme set that reads
 * as a list marker. */
export const BULLET_CHAR_CHOICES: BulletCharChoice[] = CHARACTER_CATEGORIES.filter(
  (category) => category.id === "arrows" || category.id === "symbols",
).flatMap((category) =>
  category.chars
    .filter((entry) => !isGlyphless(entry.cp))
    .map((entry) => ({ char: String.fromCodePoint(entry.cp), name: displayName(entry.name) })),
);

/** text-autofit's options, labelled exactly as the spec's e2e reads them. */
export const AUTOFIT_OPTIONS = [
  { value: "none", label: "Do not autofit" },
  { value: "shrink", label: "Shrink text on overflow" },
  { value: "resizeShape", label: "Resize shape to fit text" },
];

/** text-anchor's options. */
export const ANCHOR_OPTIONS = [
  { value: "top", label: "Top" },
  { value: "middle", label: "Middle" },
  { value: "bottom", label: "Bottom" },
];

/** text-line-spacing-mode's options — CustomSpacingDialog.vue's vocabulary. */
export const LINE_SPACING_MODE_OPTIONS = [
  { value: "multiple", label: "Multiple" },
  { value: "exact", label: "Exactly" },
  { value: "atLeast", label: "At least" },
];

/** The four inset inputs, `text-inset-{side}`. */
export const INSET_SIDES = [
  { side: "left", label: "Left" },
  { side: "top", label: "Top" },
  { side: "right", label: "Right" },
  { side: "bottom", label: "Bottom" },
];

/** The numeric value of an input's change event; NaN when blank/unparseable,
 * which every controller command ignores. */
export function eventNumber(event: Event): number {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || target.value.trim() === "") return Number.NaN;
  return Number(target.value);
}

/** The string value of an input/select change event, or "". */
export function eventString(event: Event): string {
  const target = event.target;
  return target instanceof HTMLInputElement || target instanceof HTMLSelectElement ? target.value : "";
}

/** A number-input value: the number, or "" so a mixed selection shows blank. */
export function inputValue(value: unknown): string {
  return typeof value === "number" && Number.isFinite(value) ? String(value) : "";
}

/** The link dialog's inline URL error ("" when acceptable or blank). */
export function linkUrlError(url: string): string {
  const trimmed = url.trim();
  return trimmed !== "" && isDangerousUrl(trimmed) ? "This address type isn't allowed in links" : "";
}

/** The dialog's draft as a link target: a chosen slide wins over a URL. */
export function linkTargetOf(url: string, slide: string, tooltip: string): SlideLinkTarget {
  const tip = tooltip.trim() ? { tooltip: tooltip.trim() } : {};
  const number = Number(slide);
  if (slide !== "" && Number.isInteger(number) && number > 0) return { slide: number, ...tip };
  return { href: url.trim(), ...tip };
}

/** The dialog's prefill from the controller's untyped `initial`. */
export function linkDraftOf(initial: unknown): { url: string; slide: string; tooltip: string } {
  const target = (initial ?? {}) as SlideLinkTarget;
  return {
    url: target.href ?? "",
    slide: target.slide ? String(target.slide) : "",
    tooltip: target.tooltip ?? "",
  };
}
