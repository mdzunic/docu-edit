// PresentationML colour resolution (SPEC-079). A DrawingML colour element
// (`a:srgbClr`, `a:sysClr`, `a:schemeClr`) resolves to a lowercase `#rrggbb`
// through the slide master's `p:clrMap` into the theme's `a:clrScheme`, with
// the `lumMod`/`lumOff`/`shade`/`tint` transforms applied in document order.
// Pure and DOM-reading only — direct unit tests pin the transform maths
// against known PowerPoint values (color.spec.ts).
import { DRAWINGML_NAMESPACE, firstA } from "../../core/drawingMl";
import { applyTransform, hexToRgb, rgbToHex } from "../../core/theme";

/** The resolved theme + master mapping a colour is looked up through. */
export interface ColorContext {
  /** Theme `a:clrScheme` entries: "dk1"/"lt1"/"dk2"/"lt2"/"accent1".."accent6"
   * /"hlink"/"folHlink" → "#rrggbb". */
  scheme: Record<string, string>;
  /** Master `p:clrMap`: the slide colour names ("bg1"/"tx1"/"bg2"/"tx2") →
   * their theme scheme entry ("lt1"/"dk1"/…). */
  map: Record<string, string>;
}

const P_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main";

/** Reads a bare colour element (`a:srgbClr`/`a:sysClr`) into a `#rrggbb`, with
 * NO transforms — used to populate the scheme table. */
function bareColor(el: Element): string | undefined {
  if (el.localName === "srgbClr") {
    const val = el.getAttribute("val");
    return val ? `#${val.toLowerCase()}` : undefined;
  }
  if (el.localName === "sysClr") {
    const last = el.getAttribute("lastClr");
    return last ? `#${last.toLowerCase()}` : undefined;
  }
  return undefined;
}

/** Builds the scheme table from a theme part's `a:theme` root element. */
export function parseClrScheme(themeRoot: Element): Record<string, string> {
  const scheme: Record<string, string> = {};
  const elements = firstA(themeRoot, "themeElements");
  const clrScheme = elements && firstA(elements, "clrScheme");
  if (!clrScheme) {
    return scheme;
  }
  for (const entry of Array.from(clrScheme.children)) {
    if (entry.namespaceURI !== DRAWINGML_NAMESPACE) {
      continue;
    }
    const inner = Array.from(entry.children).find((c) => c.namespaceURI === DRAWINGML_NAMESPACE);
    const color = inner ? bareColor(inner) : undefined;
    if (color) {
      scheme[entry.localName] = color;
    }
  }
  return scheme;
}

/** Builds the `p:clrMap` table from a slide master's root element. */
export function parseClrMap(masterRoot: Element): Record<string, string> {
  const map: Record<string, string> = {};
  const clrMap = Array.from(masterRoot.children).find(
    (c) => c.namespaceURI === P_NAMESPACE && c.localName === "clrMap",
  );
  if (!clrMap) {
    return map;
  }
  for (const attr of Array.from(clrMap.attributes)) {
    map[attr.name] = attr.value;
  }
  return map;
}

// --- colour maths -----------------------------------------------------------
// Hoisted to core/theme.ts (SPEC-094's enabling move) so `resolveColor` shares
// the exact transform semantics; this module re-imports what it needs.

/** Resolves a DrawingML colour element to a lowercase `#rrggbb`, applying its
 * child transforms in document order. Returns undefined for a colour it
 * cannot resolve (unknown scheme entry, no attributes) — the caller then
 * falls back (e.g. to the built-in default). */
export function resolveColorElement(el: Element, ctx: ColorContext): string | undefined {
  let base: string | undefined;
  if (el.localName === "srgbClr" || el.localName === "sysClr") {
    base = bareColor(el);
  } else if (el.localName === "schemeClr") {
    const val = el.getAttribute("val");
    if (val) {
      const mapped = ctx.map[val] ?? val;
      base = ctx.scheme[mapped] ?? ctx.scheme[val];
    }
  }
  if (!base) {
    return undefined;
  }
  let rgb = hexToRgb(base);
  for (const child of Array.from(el.children)) {
    if (child.namespaceURI !== DRAWINGML_NAMESPACE) {
      continue;
    }
    const raw = child.getAttribute("val");
    if (raw === null) {
      continue;
    }
    rgb = applyTransform(rgb, child.localName, Number(raw) / 100000);
  }
  return rgbToHex(rgb);
}

/** Resolves the first colour element inside a fill container (`a:solidFill`)
 * to a `#rrggbb`, or undefined when absent/unresolvable. */
export function resolveFirstColor(container: Element, ctx: ColorContext): string | undefined {
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI !== DRAWINGML_NAMESPACE) {
      continue;
    }
    if (["srgbClr", "sysClr", "schemeClr"].includes(child.localName)) {
      return resolveColorElement(child, ctx);
    }
  }
  return undefined;
}
