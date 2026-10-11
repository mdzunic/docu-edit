// Conditional-format quick presets (SPEC-129 "Presets") — THIS FILE is the
// normative source for `cf-add`'s gallery (CF_ADD_PRESETS, each resolving to
// a full CfRuleDraft the editor opens on — a value preset's value fields
// start blank, e.g. "Greater than…" seeds `formulas: [""]` for the user to
// fill in, per the rule editor's `cf-value-1`) and `cf-format`'s six
// differential-style presets plus "Custom format…" (FORMAT_PRESETS).
// Framework-free and pure, so it unit-tests without mounting a component.
import { ODS_ONLY_ICON_SETS, ICON_SET_NAMES, defaultIconThresholds, iconSetSize, type CfDataBar, type CfRuleDraft, type IconSetName } from "../../lib/spreadsheet/conditionalFormats";
import { ICON_SET_LABELS } from "../../lib/spreadsheet/cfIcons";
import type { Dxf } from "../../lib/spreadsheet/conditionalFormats";

export interface CfAddPreset {
  testid: string;
  label: string;
  draft(): CfRuleDraft;
}

/** A gallery data bar as Excel 2010+ writes one: automatic minimum and
 * maximum, lengths 0–100%, red negative bars, a black automatic axis, and a
 * border in the bar's own colour on the gradient style. */
export function galleryDataBar(color: string, gradient: boolean): CfDataBar {
  return {
    min: { type: "autoMin" },
    max: { type: "autoMax" },
    color,
    gradient,
    ...(gradient ? { border: color, negativeBorderColor: "#FF0000" } : {}),
    negativeColor: "#FF0000",
    axis: "automatic",
    axisColor: "#000000",
    direction: "context",
    minLengthPct: 0,
    maxLengthPct: 100,
    showValue: true,
  };
}

/** The six data-bar colours `cf-add` offers, each as a gradient and a solid
 * bar (SPEC-129 "Presets"). */
const DATA_BAR_COLORS: readonly { id: string; label: string; color: string }[] = [
  { id: "blue", label: "Blue", color: "#638EC6" },
  { id: "green", label: "Green", color: "#63BE7B" },
  { id: "red", label: "Red", color: "#FF555A" },
  { id: "orange", label: "Orange", color: "#FFB628" },
  { id: "light-blue", label: "Light Blue", color: "#008AEF" },
  { id: "purple", label: "Purple", color: "#D6007B" },
];

/** The six colour scales (min, 50th percentile, max), in `cf-add`'s order. */
const COLOR_SCALES: readonly { label: string; stops: { pos: number; color: string }[] }[] = [
  {
    label: "Green - Yellow - Red Color Scale",
    stops: [
      { pos: 0, color: "#63BE7B" },
      { pos: 50, color: "#FFEB84" },
      { pos: 100, color: "#F8696B" },
    ],
  },
  {
    label: "Red - Yellow - Green Color Scale",
    stops: [
      { pos: 0, color: "#F8696B" },
      { pos: 50, color: "#FFEB84" },
      { pos: 100, color: "#63BE7B" },
    ],
  },
  {
    label: "Green - White - Red Color Scale",
    stops: [
      { pos: 0, color: "#63BE7B" },
      { pos: 50, color: "#FCFCFF" },
      { pos: 100, color: "#F8696B" },
    ],
  },
  {
    label: "Red - White - Green Color Scale",
    stops: [
      { pos: 0, color: "#F8696B" },
      { pos: 50, color: "#FCFCFF" },
      { pos: 100, color: "#63BE7B" },
    ],
  },
  {
    label: "White - Red 2-Color Scale",
    stops: [
      { pos: 0, color: "#FFFFFF" },
      { pos: 100, color: "#F8696B" },
    ],
  },
  {
    label: "Green - White 2-Color Scale",
    stops: [
      { pos: 0, color: "#63BE7B" },
      { pos: 100, color: "#FFFFFF" },
    ],
  },
];

function colorScaleDraft(stops: readonly { pos: number; color: string }[]): CfRuleDraft {
  return {
    kind: "colorScale",
    stops: stops.map(({ pos, color }) => ({ at: pos === 0 ? { type: "min" } : pos === 100 ? { type: "max" } : { type: "percentile", value: String(pos) }, color })),
  };
}

function iconSetDraft(set: IconSetName): CfRuleDraft {
  const count = iconSetSize(set);
  return { kind: "iconSet", icons: { set, thresholds: defaultIconThresholds(count), reverse: false, showValue: true } };
}

/** `cf-add`'s gallery, in the order "Presets" lists it. A value preset
 * (Greater than…, Text that contains…, …) seeds its type and operator with
 * blank values — the editor opens ready for `cf-value-1`. */
export const CF_ADD_PRESETS: readonly CfAddPreset[] = [
  { testid: "cf-preset-greater-than", label: "Greater than…", draft: () => ({ kind: "cellIs", operator: "greaterThan", formulas: [""] }) },
  { testid: "cf-preset-less-than", label: "Less than…", draft: () => ({ kind: "cellIs", operator: "lessThan", formulas: [""] }) },
  { testid: "cf-preset-between", label: "Between…", draft: () => ({ kind: "cellIs", operator: "between", formulas: ["", ""] }) },
  { testid: "cf-preset-equal", label: "Equal to…", draft: () => ({ kind: "cellIs", operator: "equal", formulas: [""] }) },
  { testid: "cf-preset-text-contains", label: "Text that contains…", draft: () => ({ kind: "text", operator: "contains", text: "" }) },
  { testid: "cf-preset-date", label: "A date occurring…", draft: () => ({ kind: "timePeriod", period: "today" }) },
  { testid: "cf-preset-duplicates", label: "Duplicate values…", draft: () => ({ kind: "duplicates", unique: false }) },
  { testid: "cf-preset-top10", label: "Top 10 items…", draft: () => ({ kind: "top10", bottom: false, percent: false, rank: 10 }) },
  { testid: "cf-preset-bottom10", label: "Bottom 10 items…", draft: () => ({ kind: "top10", bottom: true, percent: false, rank: 10 }) },
  { testid: "cf-preset-above-average", label: "Above average", draft: () => ({ kind: "aboveAverage", below: false, equalAverage: false }) },
  { testid: "cf-preset-below-average", label: "Below average", draft: () => ({ kind: "aboveAverage", below: true, equalAverage: false }) },
  ...DATA_BAR_COLORS.flatMap((entry) => [
    {
      testid: `cf-preset-data-bar-${entry.id}-gradient`,
      label: `${entry.label} Data Bar (Gradient)`,
      draft: (): CfRuleDraft => ({ kind: "dataBar", bar: galleryDataBar(entry.color, true) }),
    },
    {
      testid: `cf-preset-data-bar-${entry.id}-solid`,
      label: `${entry.label} Data Bar (Solid)`,
      draft: (): CfRuleDraft => ({ kind: "dataBar", bar: galleryDataBar(entry.color, false) }),
    },
  ]),
  ...COLOR_SCALES.map((scale, index) => ({
    testid: `cf-preset-color-scale-${index + 1}`,
    label: scale.label,
    draft: (): CfRuleDraft => colorScaleDraft(scale.stops),
  })),
  ...ICON_SET_NAMES.filter((name) => !ODS_ONLY_ICON_SETS.has(name)).map((name) => ({
    testid: `cf-preset-icon-set-${name}`,
    label: ICON_SET_LABELS[name],
    draft: (): CfRuleDraft => iconSetDraft(name),
  })),
  { testid: "cf-preset-new-rule", label: "New rule…", draft: () => ({ kind: "cellIs", operator: "greaterThan", formulas: [""] }) },
];

export interface CfFormatPreset {
  testid: string;
  label: string;
  dxf: Dxf;
}

/** `cf-format`'s six built-in presets (SPEC-129 "Presets"); "Custom format…"
 * is handled separately — it opens FormatCellsDialog's differential mode
 * rather than naming a fixed Dxf. */
export const CF_FORMAT_PRESETS: readonly CfFormatPreset[] = [
  {
    testid: "cf-format-preset-light-red-fill-dark-red-text",
    label: "Light red fill with dark red text",
    dxf: { fill: { kind: "solid", color: "#FFC7CE" }, font: { color: "#9C0006" } },
  },
  {
    testid: "cf-format-preset-yellow-fill-dark-yellow-text",
    label: "Yellow fill with dark yellow text",
    dxf: { fill: { kind: "solid", color: "#FFEB9C" }, font: { color: "#9C5700" } },
  },
  {
    testid: "cf-format-preset-green-fill-dark-green-text",
    label: "Green fill with dark green text",
    dxf: { fill: { kind: "solid", color: "#C6EFCE" }, font: { color: "#006100" } },
  },
  { testid: "cf-format-preset-light-red-fill", label: "Light red fill", dxf: { fill: { kind: "solid", color: "#FFC7CE" } } },
  { testid: "cf-format-preset-red-text", label: "Red text", dxf: { font: { color: "#9C0006" } } },
  {
    testid: "cf-format-preset-red-border",
    label: "Red border",
    dxf: {
      borders: {
        top: { style: "thin", color: "#9C0006" },
        right: { style: "thin", color: "#9C0006" },
        bottom: { style: "thin", color: "#9C0006" },
        left: { style: "thin", color: "#9C0006" },
      },
    },
  },
];

/** The value presets, which open the rule editor pre-set to their type and
 * operator rather than applying directly (SPEC-129 "Presets"), and New
 * rule…, which opens it blank. */
export const EDITOR_PRESETS: ReadonlySet<string> = new Set([
  "cf-preset-greater-than",
  "cf-preset-less-than",
  "cf-preset-between",
  "cf-preset-equal",
  "cf-preset-text-contains",
  "cf-preset-date",
  "cf-preset-top10",
  "cf-preset-bottom10",
  "cf-preset-new-rule",
]);

/** `cf-add`'s menu sections, in Excel's order. */
export const CF_ADD_PRESET_GROUPS: readonly { label: string; presets: readonly CfAddPreset[] }[] = [
  { label: "Highlight cells rules", prefixes: ["greater-than", "less-than", "between", "equal", "text-contains", "date", "duplicates"] },
  { label: "Top/bottom rules", prefixes: ["top10", "bottom10", "above-average", "below-average"] },
  { label: "Data bars", prefixes: ["data-bar-"] },
  { label: "Colour scales", prefixes: ["color-scale-"] },
  { label: "Icon sets", prefixes: ["icon-set-"] },
  { label: "", prefixes: ["new-rule"] },
].map(({ label, prefixes }) => ({
  label,
  presets: CF_ADD_PRESETS.filter((preset) => {
    const id = preset.testid.slice("cf-preset-".length);
    return prefixes.some((prefix) => (prefix.endsWith("-") ? id.startsWith(prefix) : id === prefix));
  }),
}));

/** The preset with `testid`. */
export function addPreset(testid: string): CfAddPreset | undefined {
  return CF_ADD_PRESETS.find((preset) => preset.testid === testid);
}
