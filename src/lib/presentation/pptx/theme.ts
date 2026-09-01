// The pptx theme wire mapping (SPEC-094) — the normative module for how a
// `DeckTheme` rides in PresentationML: reading `a:clrScheme`/`a:fontScheme`
// into the shared model, capturing `a:schemeClr` as a `ThemeColorRef`, the
// clrMap reverse mapping the writer emits slide-space names through (including
// the `p:clrMapOvr` opt-out), and the in-place theme-part splice that rewrites
// ONLY `a:clrScheme` and the major/minor `a:latin` typefaces while every other
// byte of `ppt/theme/themeN.xml` survives verbatim (SPEC-008's rule).
import { DRAWINGML_NAMESPACE, firstA } from "../../core/drawingMl";
import { escapeXmlAttribute } from "../../core/xml";
import {
  resolveColor,
  THEME_COLOR_SLOTS,
  type DocumentTheme,
  type ModelFont,
  type ThemeColorRef,
  type ThemeColorSlot,
} from "../../core/theme";
import { parseClrMap, parseClrScheme } from "./color";
import type { DeckTheme } from "../model";

const SLOT_SET = new Set<string>(THEME_COLOR_SLOTS);

/** The transforms a reference may carry, in the canonical (wire) order. */
const REF_TRANSFORMS = ["lumMod", "lumOff", "shade", "tint"] as const;

// --- read -------------------------------------------------------------------

/** Reads the deck theme from a theme part root + its master (SPEC-094):
 * scheme colours (missing slots fall back to `fallback`), major/minor latin
 * fonts, the master's `p:clrMap` and the part path. Returns undefined when the
 * part declares no colour scheme at all — `model.theme` then stays absent and
 * every colour stays a literal (AC-7). */
export function readDeckTheme(
  themeRoot: Element | undefined,
  masterRoot: Element | undefined,
  partName: string,
  fallback: DocumentTheme,
): DeckTheme | undefined {
  if (!themeRoot) {
    return undefined;
  }
  const scheme = parseClrScheme(themeRoot);
  if (Object.keys(scheme).length === 0) {
    return undefined;
  }
  const colors = { ...fallback.colors };
  for (const slot of THEME_COLOR_SLOTS) {
    if (scheme[slot]) {
      colors[slot] = scheme[slot];
    }
  }
  const elements = firstA(themeRoot, "themeElements");
  const clrScheme = elements && firstA(elements, "clrScheme");
  const fontScheme = elements && firstA(elements, "fontScheme");
  const latinOf = (which: "majorFont" | "minorFont"): string | undefined => {
    const font = fontScheme && firstA(fontScheme, which);
    const latin = font && firstA(font, "latin");
    return latin?.getAttribute("typeface") || undefined;
  };
  const theme: DeckTheme = {
    name: clrScheme?.getAttribute("name") || themeRoot.getAttribute("name") || "Custom",
    colors,
    fonts: {
      major: latinOf("majorFont") ?? fallback.fonts.major,
      minor: latinOf("minorFont") ?? fallback.fonts.minor,
    },
    partName,
  };
  const map = masterRoot ? parseClrMap(masterRoot) : {};
  if (Object.keys(map).length > 0) {
    theme.colorMap = map;
  }
  return theme;
}

/** Captures an `a:schemeClr` element as a `ThemeColorRef` (AC-14), or
 * undefined when the value must stay a literal: an unknown slot, any transform
 * outside lumMod/lumOff/shade/tint, a repeated transform, or transforms out of
 * the canonical order (the only order `resolveColor` applies and the writer
 * emits — capturing anything else would change the rendered colour, AC-15). */
export function colorRefFrom(el: Element, map: Record<string, string>): ThemeColorRef | undefined {
  if (el.localName !== "schemeClr") {
    return undefined;
  }
  const val = el.getAttribute("val");
  if (!val) {
    return undefined;
  }
  const slot = map[val] ?? val;
  if (!SLOT_SET.has(slot)) {
    return undefined;
  }
  const ref: ThemeColorRef = { theme: slot as ThemeColorSlot };
  let lastIndex = -1;
  for (const child of Array.from(el.children)) {
    if (child.namespaceURI !== DRAWINGML_NAMESPACE) {
      return undefined;
    }
    const index = REF_TRANSFORMS.indexOf(child.localName as (typeof REF_TRANSFORMS)[number]);
    if (index === -1 || index <= lastIndex) {
      return undefined; // out-of-subset, repeated or reordered transform
    }
    const raw = child.getAttribute("val");
    if (raw === null || !Number.isFinite(Number(raw))) {
      return undefined;
    }
    ref[REF_TRANSFORMS[index]] = Number(raw) / 100000;
    lastIndex = index;
  }
  return ref;
}

/** Maps an `a:latin@typeface` to a theme font reference (AC-16). */
export function fontRefFrom(typeface: string): ModelFont | undefined {
  if (typeface === "+mj-lt") {
    return { theme: "major" };
  }
  if (typeface === "+mn-lt") {
    return { theme: "minor" };
  }
  return undefined;
}

// --- write: slide-space names -----------------------------------------------

/** The `p:clrMap` attribute order the reverse lookup scans, so a slot mapped
 * by several attributes resolves deterministically. */
const CLR_MAP_ATTRS = [
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
];

/** The slide-space name a scheme slot writes as (AC-18): the clrMap attribute
 * whose value is the slot (`dk1`→`tx1` under the default map), falling back to
 * the slot name when nothing maps to it. */
export function slideSpaceName(slot: ThemeColorSlot, map: Record<string, string> | undefined): string {
  if (map) {
    for (const attr of CLR_MAP_ATTRS) {
      if (map[attr] === slot) {
        return attr;
      }
    }
    for (const [attr, value] of Object.entries(map)) {
      if (value === slot) {
        return attr;
      }
    }
  }
  return slot;
}

/** The colour map a slide's refs write through (AC-22/23): the deck map by
 * default; a preserved `p:clrMapOvr` carrying `a:overrideClrMapping` supplies
 * its own; one this module cannot parse opts the slide out — every ref on it
 * writes as an `a:srgbClr` literal so the rendered colour never lies. */
export function slideClrMapMode(
  preservedClrMapOvr: string | undefined,
  deckMap: Record<string, string> | undefined,
): Record<string, string> | "literal" {
  if (preservedClrMapOvr === undefined) {
    return deckMap ?? {};
  }
  // The preserved string may carry serialiser-added xmlns attributes.
  if (/<a:masterClrMapping\b[^>]*\/?>/.test(preservedClrMapOvr)) {
    return deckMap ?? {};
  }
  const override = /<a:overrideClrMapping((?:\s+[\w:-]+="[^"]*")*)\s*\/?>/.exec(preservedClrMapOvr);
  if (override) {
    const map: Record<string, string> = {};
    for (const attr of override[1].matchAll(/([\w:-]+)="([^"]*)"/g)) {
      if (!attr[1].startsWith("xmlns")) {
        map[attr[1]] = attr[2];
      }
    }
    if (Object.keys(map).length > 0) {
      return map;
    }
  }
  return "literal";
}

/** How one slide's colours serialise: through a map (refs as `a:schemeClr`)
 * or literals only (the `p:clrMapOvr` opt-out). */
export interface SlideColorWrite {
  theme: DocumentTheme;
  /** "literal" = every ref resolves to `a:srgbClr` (AC-23). */
  mode: Record<string, string> | "literal";
}

/** A `ThemeColorRef` as its wire element (AC-17): `<a:schemeClr val="{name}">`
 * with `<a:lumMod val="60000"/>`-style children in the canonical order
 * lumMod, lumOff, shade, tint. */
export function schemeClrXml(ref: ThemeColorRef, map: Record<string, string> | undefined): string {
  const name = slideSpaceName(ref.theme, map);
  const children = REF_TRANSFORMS.map((transform) => {
    const value = ref[transform];
    return value === undefined ? "" : `<a:${transform} val="${Math.round(value * 100000)}"/>`;
  }).join("");
  if (children.length === 0) {
    return `<a:schemeClr val="${escapeXmlAttribute(name)}"/>`;
  }
  return `<a:schemeClr val="${escapeXmlAttribute(name)}">${children}</a:schemeClr>`;
}

// --- write: the theme-part splice -------------------------------------------

const hexUpper = (color: string): string => color.replace(/^#/, "").toUpperCase();

/** The canonical `a:clrScheme` for a theme (AC-25). `sourceClrScheme` (the
 * part's current element text) supplies each slot's source FORM: a slot that
 * was `a:sysClr` keeps its `val` with `lastClr` updated; everything else — and
 * every slot of a theme with no source form — writes `a:srgbClr`. */
export function clrSchemeXml(theme: DocumentTheme, sourceClrScheme?: string): string {
  const entries = THEME_COLOR_SLOTS.map((slot) => {
    const hex = hexUpper(theme.colors[slot]);
    const sourceEntry = sourceClrScheme
      ? new RegExp(`<a:${slot}>\\s*<a:sysClr\\s+val="([^"]*)"[^>]*/>`).exec(sourceClrScheme)
      : null;
    const inner = sourceEntry
      ? `<a:sysClr val="${escapeXmlAttribute(sourceEntry[1])}" lastClr="${hex}"/>`
      : `<a:srgbClr val="${hex}"/>`;
    return `<a:${slot}>${inner}</a:${slot}>`;
  }).join("");
  return `<a:clrScheme name="${escapeXmlAttribute(theme.name)}">${entries}</a:clrScheme>`;
}

const CLR_SCHEME_RE = /<a:clrScheme(?:\s[^>]*)?>[\s\S]*?<\/a:clrScheme>/;

/** Rewrites `ppt/theme/themeN.xml` IN PLACE for an edited theme (AC-25/26):
 * the `a:clrScheme` element is replaced with the canonical scheme (source
 * forms kept for `dk1`/`lt1`-style `a:sysClr` slots) and the major/minor
 * `a:latin@typeface` attributes are updated — nothing else moves, so
 * `a:fmtScheme`, `a:effectStyleLst`, `a:bgFillStyleLst`, `a:objectDefaults`
 * and `a:extraClrSchemeLst` survive byte-verbatim. Returns the input unchanged
 * when the part carries no `a:clrScheme` (nothing to splice into). */
export function spliceThemeXml(sourceXml: string, theme: DocumentTheme): string {
  const source = CLR_SCHEME_RE.exec(sourceXml);
  if (!source) {
    return sourceXml;
  }
  let out = sourceXml.replace(CLR_SCHEME_RE, () => clrSchemeXml(theme, source[0]));
  const spliceLatin = (which: "majorFont" | "minorFont", family: string): void => {
    out = out.replace(new RegExp(`(<a:${which}>[\\s\\S]*?<a:latin\\b[^>]*typeface=")[^"]*(")`), (_m, pre, post) => {
      return `${pre}${escapeXmlAttribute(family)}${post}`;
    });
  };
  spliceLatin("majorFont", theme.fonts.major);
  spliceLatin("minorFont", theme.fonts.minor);
  return out;
}

/** The literal a ref falls back to under the opt-out mode. */
export function refLiteral(ref: ThemeColorRef, theme: DocumentTheme): string {
  return resolveColor(ref, theme);
}
