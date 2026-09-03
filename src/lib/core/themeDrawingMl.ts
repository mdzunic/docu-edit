// DrawingML theme-colour references (SPEC-077 AC-29/31): `a:schemeClr`
// parsing and emission for hosts that carry no colour map of their own —
// the word processor's shapes, connectors and chart parts. The presentation
// product keeps its own map-aware codec (presentation/pptx/theme.ts); this
// one maps the slide-space spellings (tx1/bg1/tx2/bg2) onto the theme-space
// slots for the model, carries the source's spelling on the ref, and emits
// that spelling back on write — the theme-space name for product-created
// refs, so an unedited shape round-trips byte-identically (AC-31).
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
