// The .docx side of document themes (SPEC-077): the WordprocessingML theme
// vocabulary — `@w:themeColor`/`@w:themeTint`/`@w:themeShade` (and their
// `themeFill` twins) on runs, borders and shading, `@w:asciiTheme` on fonts —
// plus the `word/theme/themeN.xml` part codec: parsing `a:clrScheme`/
// `a:fontScheme` into the shared DocumentTheme (AC-25/26) and splicing an
// edited scheme back into a preserved part byte-conservatively (AC-34/35).
//
// The model vocabulary (DocumentTheme, ThemeColorRef, resolveColor) is the
// shipped src/lib/core/theme.ts (AC-1) — nothing theme-shaped is defined
// here beyond the wire mapping.
import { parseXml } from "../core/xml";
import { THEME_COLOR_SLOTS } from "../core/theme";
import { readThemePart, themeSlotColorOf } from "../core/themeDrawingMl";
import type { DocumentTheme, ThemeColorRef, ThemeColorSlot } from "../core/theme";
import type { DocTheme } from "../documentModel";

/** The relationship type a theme part hangs off the main document part. */
export const THEME_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

/** The theme part's content type, for the [Content_Types].xml override. */
export const THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml";

const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

// --- @w:themeColor names (ST_ThemeColor) -------------------------------------

/** The name this product writes for each slot. Word accepts both the
 * dark1/light1 and text1/background1 spellings; reading maps both onto the
 * theme-space slots, writing emits these. */
const SLOT_TO_ATTR: Record<ThemeColorSlot, string> = {
  dk1: "dark1",
  lt1: "light1",
  dk2: "dark2",
  lt2: "light2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hyperlink",
  folHlink: "followedHyperlink",
};

const ATTR_TO_SLOT: Record<string, ThemeColorSlot> = {
  dark1: "dk1",
  light1: "lt1",
  dark2: "dk2",
  light2: "lt2",
  // The document-space spellings map through the default clrSchemeMapping
  // (bg1 -> lt1, t1 -> dk1, bg2 -> lt2, t2 -> dk2) — this product models no
  // clrSchemeMapping override, matching what Word writes by default.
  text1: "dk1",
  background1: "lt1",
  text2: "dk2",
  background2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
};

/** A `@w:themeColor` attribute value's slot, or undefined for `none` and
 * anything unrecognised (the run then keeps the literal it has today). */
export function themeSlotFromAttr(value: string | null | undefined): ThemeColorSlot | undefined {
  return value ? ATTR_TO_SLOT[value] : undefined;
}

/** The `@w:themeColor` value written for a slot. */
export function themeAttrForSlot(slot: ThemeColorSlot): string {
  return SLOT_TO_ATTR[slot];
}

// --- tint/shade arithmetic (the "Format mapping — .docx" table) --------------
//
// @w:themeTint/@w:themeShade carry the fraction as a hex byte: "99" means
// 0x99/255 = 0.6. The model's ThemeColorRef holds 0..1 fractions.

/** A tint/shade attribute byte ("00".."FF") as a 0..1 fraction, or
 * undefined for a malformed value. */
export function themeByteToFraction(value: string | null | undefined): number | undefined {
  if (!value || !/^[0-9a-fA-F]{1,2}$/.test(value)) {
    return undefined;
  }
  return Number.parseInt(value, 16) / 255;
}

/** A 0..1 fraction as the two-digit uppercase hex byte the attribute wants. */
export function fractionToThemeByte(fraction: number): string {
  const byte = Math.round(Math.min(1, Math.max(0, fraction)) * 255);
  return byte.toString(16).toUpperCase().padStart(2, "0");
}

/** A ThemeColorRef from the three `@w:themeColor`-family attributes, or
 * undefined when the colour attribute names no modelled slot (AC-27). */
export function themeRefFromAttrs(
  themeColor: string | null | undefined,
  tint: string | null | undefined,
  shade: string | null | undefined,
): ThemeColorRef | undefined {
  const slot = themeSlotFromAttr(themeColor);
  if (slot === undefined) {
    return undefined;
  }
  const ref: ThemeColorRef = { theme: slot };
  const tintFraction = themeByteToFraction(tint);
  if (tintFraction !== undefined) {
    ref.tint = tintFraction;
  }
  const shadeFraction = themeByteToFraction(shade);
  if (shadeFraction !== undefined) {
    ref.shade = shadeFraction;
  }
  return ref;
}

/** The `w:themeColor`/`w:themeTint`/`w:themeShade` attribute string for a
 * reference (leading space included), or null for a reference that has no
 * WordprocessingML spelling (D8: lumMod/lumOff carry no `w:theme*`
 * attribute) — the caller then writes the resolved literal alone (AC-32).
 * `family` picks the attribute family: "themeColor" for w:color and border
 * elements, "themeFill" for the w:shd fill. */
export function themeAttrsXml(ref: ThemeColorRef, family: "themeColor" | "themeFill"): string | null {
  if (ref.lumMod !== undefined || ref.lumOff !== undefined) {
    return null;
  }
  let attrs = ` w:${family}="${themeAttrForSlot(ref.theme)}"`;
  if (ref.tint !== undefined) {
    attrs += ` w:${family === "themeFill" ? "themeFillTint" : "themeTint"}="${fractionToThemeByte(ref.tint)}"`;
  }
  if (ref.shade !== undefined) {
    attrs += ` w:${family === "themeFill" ? "themeFillShade" : "themeShade"}="${fractionToThemeByte(ref.shade)}"`;
  }
  return attrs;
}

// --- @w:asciiTheme / @w:hAnsiTheme (AC-28) -----------------------------------

/** The font reference an `asciiTheme`/`hAnsiTheme` attribute names —
 * majorHAnsi/majorAscii -> major, minorHAnsi/minorAscii -> minor; every
 * other value (minorEastAsia, majorBidi, ...) is not modelled (AC-28). */
export function fontThemeFromAttr(value: string | null | undefined): "major" | "minor" | undefined {
  if (value === "majorHAnsi" || value === "majorAscii") {
    return "major";
  }
  if (value === "minorHAnsi" || value === "minorAscii") {
    return "minor";
  }
  return undefined;
}

// --- the theme part: read (AC-25, AC-26) -------------------------------------

/** Parses a `word/theme/themeN.xml` part into the document's theme (AC-25),
 * recording the part path in `partName`. Undefined when the XML is not a
 * usable theme part (no clrScheme) — the document then has no modelled
 * theme and the part stays a preserved opaque part. The parsing itself is
 * core's readThemePart (SPEC-117 moved it there for the spreadsheet
 * product); this keeps the word processor's entry point and its partName. */
export function readDocxTheme(themeXml: string, partName: string): DocTheme | undefined {
  const theme = readThemePart(themeXml);
  return theme ? { ...theme, partName } : undefined;
}

// --- the theme part: splice (AC-34, AC-35) -----------------------------------

const CLR_SCHEME_RE = /<a:clrScheme(?:\s[^>]*)?>[\s\S]*?<\/a:clrScheme>/;

function slotElementRe(slot: string): RegExp {
  return new RegExp(`<a:${slot}>[\\s\\S]*?</a:${slot}>|<a:${slot}/>`);
}

/** A rebuilt `a:clrScheme`: a slot whose colour is unchanged keeps the
 * source's element byte-verbatim (an `a:sysClr` stays an `a:sysClr`); a
 * changed slot writes `a:srgbClr` (AC-35). The scheme's `name` attribute
 * becomes the theme's name. */
function clrSchemeXml(theme: DocumentTheme, sourceScheme: string): string {
  const slots = THEME_COLOR_SLOTS.map((slot) => {
    const sourceSlot = slotElementRe(slot).exec(sourceScheme)?.[0];
    if (sourceSlot !== undefined) {
      const sourceColorRoot = parseXml(
        `<a:root xmlns:a="${DRAWINGML_NS}">${sourceSlot}</a:root>`,
      ).documentElement.firstElementChild;
      if (sourceColorRoot && themeSlotColorOf(sourceColorRoot) === theme.colors[slot]) {
        return sourceSlot;
      }
    }
    return `<a:${slot}><a:srgbClr val="${theme.colors[slot].slice(1).toUpperCase()}"/></a:${slot}>`;
  });
  return `<a:clrScheme name="${escapeAttr(theme.name)}">${slots.join("")}</a:clrScheme>`;
}

function escapeAttr(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll('"', "&quot;");
}

/** Splices an edited theme into a preserved theme part: `a:clrScheme`, the
 * two `a:latin` typefaces inside `a:fontScheme` and the root `a:theme` name
 * attribute (where AC-25 reads the theme's name back from) change;
 * `a:fmtScheme`, `a:effectStyleLst`, `a:bgFillStyleLst`, `a:objectDefaults`,
 * `a:extraClrSchemeLst`, namespace declarations and element order survive
 * byte-verbatim (AC-34). Returns the input unchanged when it holds no
 * clrScheme to replace. */
export function spliceDocxThemeXml(sourceXml: string, theme: DocumentTheme): string {
  const source = CLR_SCHEME_RE.exec(sourceXml);
  if (!source) {
    return sourceXml;
  }
  // A function replacer so `$`-sequences in the generated XML are not
  // interpreted as replacement patterns.
  let out = sourceXml.replace(CLR_SCHEME_RE, () => clrSchemeXml(theme, source[0]));
  // The root's name attribute is where readDocxTheme recovers the theme's
  // name (AC-25), so an edited theme's name must land here or a save →
  // reopen silently reverts it (AC-52).
  out = out.replace(/(<a:theme\b[^>]*?\sname=")[^"]*(")/, (_, before: string, after: string) => before + escapeAttr(theme.name) + after);
  const spliceFont = (tag: string, typeface: string): void => {
    // Attribute-only splice: panose and the script-specific font children
    // survive untouched.
    const re = new RegExp(`(<a:${tag}>[\\s\\S]*?<a:latin\\b[^>]*typeface=")[^"]*(")`);
    out = out.replace(re, (_, before: string, after: string) => before + escapeAttr(typeface) + after);
  };
  spliceFont("majorFont", theme.fonts.major);
  spliceFont("minorFont", theme.fonts.minor);
  return out;
}
