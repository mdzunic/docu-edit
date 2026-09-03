// DrawingML theme-colour references (SPEC-077 AC-29/31): `a:schemeClr`
// parsing and emission for hosts that carry no colour map of their own —
// the word processor's shapes, connectors and chart parts. The presentation
// product keeps its own map-aware codec (presentation/pptx/theme.ts); this
// one maps the slide-space spellings (tx1/bg1/tx2/bg2) straight onto the
// theme-space slots and always emits the theme-space name.
import { THEME_COLOR_SLOTS } from "./theme";
import type { ThemeColorRef, ThemeColorSlot } from "./theme";

const DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

const SLOT_SET = new Set<string>(THEME_COLOR_SLOTS);

/** The slide-space spellings and their theme-space slots (AC-29). */
const SLIDE_SPACE: Record<string, ThemeColorSlot> = {
  tx1: "dk1",
  bg1: "lt1",
  tx2: "dk2",
  bg2: "lt2",
};

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
export function schemeClrRefFrom(el: Element): ThemeColorRef | undefined {
  if (el.localName !== "schemeClr") {
    return undefined;
  }
  const slot = schemeSlotFromVal(el.getAttribute("val"));
  if (slot === undefined) {
    return undefined;
  }
  const ref: ThemeColorRef = { theme: slot };
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

/** An `a:schemeClr` element for a reference: the theme-space slot name and
 * the transform children in canonical order (AC-31). Self-closing when the
 * reference carries no transform. */
export function schemeClrXmlFor(ref: ThemeColorRef): string {
  const children = REF_TRANSFORMS.filter((name) => ref[name] !== undefined)
    .map((name) => `<a:${name} val="${Math.round((ref[name] as number) * 100000)}"/>`)
    .join("");
  return children === ""
    ? `<a:schemeClr val="${ref.theme}"/>`
    : `<a:schemeClr val="${ref.theme}">${children}</a:schemeClr>`;
}
