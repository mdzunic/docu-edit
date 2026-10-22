// DrawingML theme-colour references (SPEC-077 AC-29/31): `a:schemeClr`
// parsing and emission for hosts that carry no colour map of their own —
// the word processor's shapes, connectors and chart parts. The presentation
// product keeps its own map-aware codec (presentation/pptx/theme.ts); this
// one maps the slide-space spellings (tx1/bg1/tx2/bg2) onto the theme-space
// slots for the model, carries the source's spelling on the ref, and emits
// that spelling back on write — the theme-space name for product-created
// refs, so an unedited shape round-trips byte-identically (AC-31).
//
// It also reads the theme part itself (SPEC-117 "Theme"): `readThemePart`
// is docx/theme.ts's readDocxTheme body moved here verbatim, so the
// spreadsheet product (which may import only core) reads an xlsx workbook's
// `xl/theme/theme1.xml` through the same code the word processor does;
// readDocxTheme now delegates to it.
import { THEME_COLOR_SLOTS } from "./theme";
import type { DocumentTheme, ThemeColorRef, ThemeColorSlot } from "./theme";
import { parseXml } from "./xml";

const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

const SLOT_SET = new Set<string>(THEME_COLOR_SLOTS);

/** The slide-space spellings and their theme-space slots (AC-29). */
const SLIDE_SPACE: Record<string, ThemeColorSlot> = {
  tx1: "dk1",
  bg1: "lt1",
  tx2: "dk2",
  bg2: "lt2",
};

/** A slide-space `a:schemeClr@val` spelling a source file used. */
export type SchemeSpelling = "tx1" | "bg1" | "tx2" | "bg2";

function isSchemeSpelling(val: string | null): val is SchemeSpelling {
  return val !== null && SLIDE_SPACE[val] !== undefined;
}

/** A `ThemeColorRef` read from DrawingML: `spelling` carries the source's
 * slide-space slot name (tx1/bg1/tx2/bg2) so a re-save emits the byte the
 * file arrived with (AC-31); absent — every product-created ref — the
 * theme-space name writes. The extra field rides inside the model's
 * `ModelColor` untouched: `resolveColor` ignores it and `{ ...ref }` copies
 * it, so it survives editing and `equalStructural` round trips. */
export interface DrawingMlColorRef extends ThemeColorRef {
  spelling?: SchemeSpelling;
}

/** The transforms a reference models, in the canonical order they are
 * applied and emitted (D3). */
const REF_TRANSFORMS = ["lumMod", "lumOff", "shade", "tint"] as const;

/** The slot an `a:schemeClr@val` names, accepting both the theme-space
 * (dk1/lt1/dk2/lt2) and slide-space (tx1/bg1/tx2/bg2) spellings; undefined
 * for `phClr` and anything else (AC-29). */
export function schemeSlotFromVal(val: string | null | undefined): ThemeColorSlot | undefined {
  if (!val) {
    return undefined;
  }
  if (SLOT_SET.has(val)) {
    return val as ThemeColorSlot;
  }
  return SLIDE_SPACE[val];
}

/** A ThemeColorRef from an `a:schemeClr` element, or undefined when the
 * element is not one the model can carry — an unknown slot (phClr), a child
 * outside the four modelled transforms, or transforms repeated/out of
 * canonical order — in which case the caller keeps the literal it resolves
 * today (AC-29). Transform children are 0..1 fractions (wire val/100000). */
export function schemeClrRefFrom(el: Element): DrawingMlColorRef | undefined {
  if (el.localName !== "schemeClr") {
    return undefined;
  }
  const val = el.getAttribute("val");
  const slot = schemeSlotFromVal(val);
  if (slot === undefined) {
    return undefined;
  }
  const ref: DrawingMlColorRef = { theme: slot };
  if (isSchemeSpelling(val)) {
    ref.spelling = val;
  }
  let lastIndex = -1;
  for (const child of Array.from(el.children)) {
    if (child.namespaceURI !== DRAWINGML_NS) {
      return undefined;
    }
    const index = REF_TRANSFORMS.indexOf(child.localName as (typeof REF_TRANSFORMS)[number]);
    if (index === -1 || index <= lastIndex) {
      return undefined;
    }
    lastIndex = index;
    const raw = child.getAttribute("val");
    if (raw === null || raw === "" || Number.isNaN(Number(raw))) {
      return undefined;
    }
    ref[REF_TRANSFORMS[index]] = Number(raw) / 100000;
  }
  return ref;
}

/** An `a:schemeClr` element for a reference: the source's slot spelling when
 * the ref carries one (and it still names this ref's slot — a hand-edited
 * mismatch falls back to the slot), else the theme-space name; the transform
 * children in canonical order (AC-31). Self-closing when the reference
 * carries no transform. */
export function schemeClrXmlFor(ref: DrawingMlColorRef): string {
  const val = ref.spelling !== undefined && SLIDE_SPACE[ref.spelling] === ref.theme ? ref.spelling : ref.theme;
  const children = REF_TRANSFORMS.filter((name) => ref[name] !== undefined)
    .map((name) => `<a:${name} val="${Math.round((ref[name] as number) * 100000)}"/>`)
    .join("");
  return children === "" ? `<a:schemeClr val="${val}"/>` : `<a:schemeClr val="${val}">${children}</a:schemeClr>`;
}

// --- the theme part: read (SPEC-077 AC-25/26, moved here by SPEC-117) -------

function firstByLocalName(parent: Element, localName: string): Element | undefined {
  for (const child of Array.from(parent.children)) {
    if (child.localName === localName && child.namespaceURI === DRAWINGML_NS) {
      return child;
    }
  }
  return undefined;
}

function descendant(root: Element, ...path: string[]): Element | undefined {
  let current: Element | undefined = root;
  for (const name of path) {
    current = current ? firstByLocalName(current, name) : undefined;
  }
  return current;
}

/** One clrScheme slot's colour: `a:srgbClr@val`, or `a:sysClr@lastClr` with
 * the windowText/window fallbacks when `lastClr` is absent (AC-26). */
export function themeSlotColorOf(slotEl: Element | undefined): string | undefined {
  if (!slotEl) {
    return undefined;
  }
  const srgb = firstByLocalName(slotEl, "srgbClr");
  if (srgb) {
    const val = srgb.getAttribute("val");
    return val && /^[0-9a-fA-F]{6}$/.test(val) ? `#${val.toLowerCase()}` : undefined;
  }
  const sys = firstByLocalName(slotEl, "sysClr");
  if (sys) {
    const last = sys.getAttribute("lastClr");
    if (last && /^[0-9a-fA-F]{6}$/.test(last)) {
      return `#${last.toLowerCase()}`;
    }
    const val = sys.getAttribute("val");
    return val === "window" ? "#ffffff" : "#000000";
  }
  return undefined;
}

/** Parses a DrawingML theme part (`a:theme`) into a DocumentTheme: the
 * `a:clrScheme` slots, the `a:fontScheme` major/minor `a:latin` typefaces and
 * the root's name. Undefined when the XML is not a usable theme part (no
 * clrScheme). */
export function readThemePart(themeXml: string): DocumentTheme | undefined {
  let doc: Document;
  try {
    doc = parseXml(themeXml);
  } catch {
    return undefined;
  }
  const root = doc.documentElement;
  if (root.localName !== "theme") {
    return undefined;
  }
  const clrScheme = descendant(root, "themeElements", "clrScheme");
  if (!clrScheme) {
    return undefined;
  }
  const colors = {} as Record<ThemeColorSlot, string>;
  let sawAny = false;
  for (const slot of THEME_COLOR_SLOTS) {
    const color = themeSlotColorOf(firstByLocalName(clrScheme, slot));
    if (color !== undefined) {
      colors[slot] = color;
      sawAny = true;
    } else {
      // A missing slot falls back to black — a real theme always has all
      // twelve, so this only cushions a truncated part.
      colors[slot] = "#000000";
    }
  }
  if (!sawAny) {
    return undefined;
  }
  const fontScheme = descendant(root, "themeElements", "fontScheme");
  const major = fontScheme && descendant(fontScheme, "majorFont", "latin")?.getAttribute("typeface");
  const minor = fontScheme && descendant(fontScheme, "minorFont", "latin")?.getAttribute("typeface");
  return {
    name: root.getAttribute("name") ?? "Custom",
    colors,
    fonts: { major: major || "Calibri Light", minor: minor || "Calibri" },
  };
}
