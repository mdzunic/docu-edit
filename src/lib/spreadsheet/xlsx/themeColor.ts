// The SpreadsheetML colour codec (SPEC-117 "Colours"): every xlsx colour
// element — `color`, `fgColor`, `bgColor`, `tabColor`, a gradient
// `stop/color` — to a core ModelColor, and the inverse SPEC-121 (tab colours)
// and SPEC-122 (the style writer) call. THIS FILE is the only spreadsheet
// module that reads a ThemeColorRef's transform fields: core/themeGuard.spec.ts
// allow-lists lib/spreadsheet/(xlsx|ods)/theme*.ts for it, and everything
// else resolves a colour through core/theme.ts's resolveColor.
//
// - An element maps by its FIRST present attribute in the order auto → theme
//   → indexed → rgb: `auto="1"` is automatic (absent); `theme="n"` is a
//   SheetColorRef; `indexed="n"` the file's `<indexedColors>` override, else
//   the legacy palette (styles.ts INDEXED_COLORS), 64 (system foreground)
//   automatic and 65 (system background) white; `rgb="AARRGGBB"` lowercase
//   `#rrggbb` with the alpha ignored (a 6-digit value is accepted as-is).
// - Excel's theme-index swap: SpreadsheetML numbers the scheme light-first —
//   0 lt1, 1 dk1, 2 lt2, 3 dk2 — then 4–9 accent1–accent6, 10 hlink,
//   11 folHlink; any other index is ignored.
// - `tint` t is Excel's HSL-luminance tint, expressed in the core transforms
//   resolveColor applies (lumMod, then lumOff): t < 0 → lumMod 1 + t;
//   t > 0 → lumMod 1 − t, lumOff t; t = 0 → nothing. The attribute's source
//   text rides along as `xlsxTint`, so an unchanged colour writes back byte
//   for byte (resolveColor ignores it). A tint on an rgb or indexed colour is
//   not modelled — Excel only writes tints on theme colours.
import { DEFAULT_THEME, isThemeColor, resolveColor, type DocumentTheme, type ModelColor, type ThemeColorSlot } from "../../core/theme";
import { scanAttr, type XmlScanAttribute } from "../../core/xmlScanner";
import type { SheetColorRef } from "../model";
import { INDEXED_COLORS } from "../styles";

/** SpreadsheetML's theme indices, in order: Excel swaps the dark/light
 * pairs against DrawingML's scheme order (0 is lt1, 1 is dk1, …). */
export const XLSX_THEME_SLOTS: readonly ThemeColorSlot[] = [
  "lt1",
  "dk1",
  "lt2",
  "dk2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

/** The slot a `theme="n"` names, or undefined for any index outside 0–11. */
export function xlsxThemeSlot(index: number): ThemeColorSlot | undefined {
  return Number.isInteger(index) && index >= 0 ? XLSX_THEME_SLOTS[index] : undefined;
}

/** The `theme="n"` index a slot is written with. */
export function xlsxThemeIndex(slot: ThemeColorSlot): number {
  return XLSX_THEME_SLOTS.indexOf(slot);
}

/** Reads one unprefixed attribute's value (undefined when absent) — the one
 * shape a DOM element and a scanner event both offer the readers. */
export type AttributeGetter = (name: string) => string | undefined;

/** The attribute getter of a DOM element. */
export function elementAttributes(element: Element): AttributeGetter {
  return (name) => element.getAttribute(name) ?? undefined;
}

/** The attribute getter of a scanner open event's attributes. */
export function scannedAttributes(attrs: readonly XmlScanAttribute[]): AttributeGetter {
  return (name) => scanAttr(attrs, name);
}

/** The lumMod/lumOff transforms an Excel tint stands for (t clamped to
 * −1..1): t < 0 darkens, `lumMod: 1 + t`; t > 0 lightens, `lumMod: 1 − t,
 * lumOff: t`; 0 is no transform. */
export function tintTransforms(tint: number): { lumMod?: number; lumOff?: number } {
  const t = Math.min(1, Math.max(-1, tint));
  if (t < 0) {
    return { lumMod: 1 + t };
  }
  if (t > 0) {
    return { lumMod: 1 - t, lumOff: t };
  }
  return {};
}

function parseNumber(text: string | undefined): number | undefined {
  if (text === undefined || text.trim() === "") {
    return undefined;
  }
  const value = Number(text);
  return Number.isFinite(value) ? value : undefined;
}

function parseIndex(text: string | undefined): number | undefined {
  const value = parseNumber(text);
  return value !== undefined && Number.isInteger(value) && value >= 0 ? value : undefined;
}

function isOn(value: string | undefined): boolean {
  return value === "1" || value === "true";
}

function themeColor(themeText: string, tintAttr: string | undefined): SheetColorRef | undefined {
  const index = parseIndex(themeText);
  const slot = index === undefined ? undefined : xlsxThemeSlot(index);
  if (slot === undefined) {
    return undefined;
  }
  const ref: SheetColorRef = { theme: slot };
  const tint = parseNumber(tintAttr);
  if (tint !== undefined) {
    Object.assign(ref, tintTransforms(tint));
    ref.xlsxTint = tintAttr;
  }
  return ref;
}

function indexedColor(indexText: string, indexedColors: readonly string[] | undefined): string | undefined {
  const index = parseIndex(indexText);
  if (index === undefined) {
    return undefined;
  }
  if (index < 64) {
    return indexedColors?.[index] ?? INDEXED_COLORS[index];
  }
  // 64 is the system foreground (automatic, i.e. absent); 65 the system
  // background.
  return index === 65 ? "#ffffff" : undefined;
}

/** `AARRGGBB` (alpha ignored) or `RRGGBB` → lowercase `#rrggbb`. */
export function xlsxRgbToHex(rgb: string): string | undefined {
  const text = rgb.trim();
  if (/^[0-9A-Fa-f]{8}$/.test(text)) {
    return `#${text.slice(2).toLowerCase()}`;
  }
  if (/^[0-9A-Fa-f]{6}$/.test(text)) {
    return `#${text.toLowerCase()}`;
  }
  return undefined;
}

/** A SpreadsheetML colour element (CT_Color) → a ModelColor, by its first
 * present attribute in the order auto → theme → indexed → rgb; undefined is
 * automatic (or an attribute that names nothing). `indexedColors` is the
 * workbook's override palette, when it has one. */
export function xlsxColorFrom(attr: AttributeGetter, indexedColors?: readonly string[]): ModelColor | undefined {
  if (isOn(attr("auto")?.trim())) {
    return undefined;
  }
  const theme = attr("theme");
  if (theme !== undefined) {
    return themeColor(theme, attr("tint"));
  }
  const indexed = attr("indexed");
  if (indexed !== undefined) {
    return indexedColor(indexed, indexedColors);
  }
  const rgb = attr("rgb");
  return rgb === undefined ? undefined : xlsxRgbToHex(rgb);
}

/** The attributes of a SpreadsheetML colour element, as plain values. */
export interface XlsxColorAttributes {
  auto?: string;
  theme?: string;
  tint?: string;
  indexed?: string;
  rgb?: string;
}

/** `xlsxColorFrom` over an attribute record. */
export function xlsxColorToModel(attrs: XlsxColorAttributes, indexedColors?: readonly string[]): ModelColor | undefined {
  return xlsxColorFrom((name) => attrs[name as keyof XlsxColorAttributes], indexedColors);
}

/** `xlsxColorFrom` over a DOM colour element (styles.xml). */
export function xlsxColorFromElement(element: Element, indexedColors?: readonly string[]): ModelColor | undefined {
  return xlsxColorFrom(elementAttributes(element), indexedColors);
}

/** `xlsxColorFrom` over a scanned colour element (worksheets, shared
 * strings). */
export function xlsxColorFromScan(attrs: readonly XmlScanAttribute[], indexedColors?: readonly string[]): ModelColor | undefined {
  return xlsxColorFrom(scannedAttributes(attrs), indexedColors);
}

// --- write (SPEC-121's tab colour, SPEC-122's style writer) ------------------

/** Tolerance for "lumMod + lumOff sum to 1" (floating-point arithmetic on the
 * read side never lands exactly). */
const SUM_EPSILON = 1e-9;

/** The `tint` a reference's transforms are, as attribute text: undefined for
 * no tint, null when no tint expresses them exactly — anything but no
 * transform, a lone lumMod in 0..1 (t = lumMod − 1) or a lumMod + lumOff
 * summing to 1 with lumOff in 0..1 (t = lumOff), or a reference carrying
 * shade/tint. The source `xlsxTint` text wins while it still stands for the
 * same transforms. */
function tintText(ref: SheetColorRef): string | undefined | null {
  if (ref.shade !== undefined || ref.tint !== undefined) {
    return null;
  }
  const lumMod = ref.lumMod;
  const lumOff = ref.lumOff;
  let t: number;
  if (lumMod === undefined && lumOff === undefined) {
    t = 0;
  } else if (lumMod === undefined || lumOff === undefined) {
    if (lumMod === undefined || lumMod < 0 || lumMod > 1) {
      return null;
    }
    t = lumMod - 1;
  } else {
    if (lumOff < 0 || lumOff > 1 || Math.abs(lumMod + lumOff - 1) > SUM_EPSILON) {
      return null;
    }
    t = lumOff;
  }
  if (ref.xlsxTint !== undefined) {
    const source = parseNumber(ref.xlsxTint);
    if (source !== undefined) {
      const transforms = tintTransforms(source);
      if (transforms.lumMod === lumMod && transforms.lumOff === lumOff) {
        return ref.xlsxTint;
      }
    }
  }
  // Fifteen significant digits drop the noise of `lumMod − 1` without moving
  // the colour.
  return t === 0 ? undefined : String(Number(t.toPrecision(15)));
}

/** A literal `#rrggbb` as Excel's `FFRRGGBB`. */
function argbOf(hex: string): string {
  const digits = hex.replace(/^#/, "");
  return `FF${/^[0-9A-Fa-f]{6}$/.test(digits) ? digits.toUpperCase() : "000000"}`;
}

/** The SpreadsheetML colour attributes for a model colour, leading space
 * included — ` theme="4" tint="0.39997558519241921"`, ` rgb="FFFF0000"`.
 * A theme reference no tint expresses exactly is written as the `rgb` it
 * resolves to against `theme` (DEFAULT_THEME when the workbook has none). */
export function modelColorToXlsxAttrs(color: ModelColor, theme: DocumentTheme | undefined): string {
  if (!isThemeColor(color)) {
    return ` rgb="${argbOf(color)}"`;
  }
  const ref: SheetColorRef = color;
  const tint = tintText(ref);
  if (tint === null) {
    return ` rgb="${argbOf(resolveColor(ref, theme ?? DEFAULT_THEME))}"`;
  }
  return ` theme="${xlsxThemeIndex(ref.theme)}"${tint === undefined ? "" : ` tint="${tint}"`}`;
}
