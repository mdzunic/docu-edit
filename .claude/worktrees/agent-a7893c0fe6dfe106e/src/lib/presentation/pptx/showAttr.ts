// The pptx wire mapping of SPEC-104's `Slide.hidden`: `<p:sld show="0">`.
// `@show` lives in `preserved["@root"]` — the whole `p:sld` attribute string
// (leading-space-prefixed, SPEC-081), which is also the byte source of every
// other root attribute — so the slot is MIRRORED, never consumed: the reader
// derives `hidden` from it and leaves the string alone, and the serialiser
// splices just this one attribute in place (the SPEC-093
// `showSpecialPlsOnTitleSld` precedent). Every other attribute keeps its bytes
// and position, so an untouched slide still serialises byte-identically.

/** `show="…"` as a whole attribute — the leading whitespace is part of the
 * match so `showMasterSp`/`showSpecialPlsOnTitleSld` can never match it. */
const SHOW_ATTRIBUTE = /\sshow="([^"]*)"/;

function isHiddenValue(value: string): boolean {
  return value === "0" || value === "false";
}

/** Whether a `p:sld` attribute string marks the slide hidden: `show="0"` or
 * `show="false"` (xsd:boolean). Absent, or any other value, is shown. */
export function rootAttrsHidden(rootAttrs: string | undefined): boolean {
  const match = rootAttrs ? SHOW_ATTRIBUTE.exec(rootAttrs) : null;
  return match !== null && isHiddenValue(match[1]);
}

/** The `p:sld` attribute string with `show` set for `hidden`: a hidden slide
 * already carrying a hidden value is returned untouched; a shown value is
 * replaced in place; an absent attribute is appended after the last one. A
 * shown slide loses a hidden `show` (with its leading space) and keeps any
 * other value as it was. Idempotent in both directions. */
export function spliceShowAttr(rootAttrs: string, hidden: boolean): string {
  const match = SHOW_ATTRIBUTE.exec(rootAttrs);
  if (hidden) {
    if (!match) {
      return `${rootAttrs} show="0"`;
    }
    if (isHiddenValue(match[1])) {
      return rootAttrs;
    }
    return `${rootAttrs.slice(0, match.index)} show="0"${rootAttrs.slice(match.index + match[0].length)}`;
  }
  if (!match || !isHiddenValue(match[1])) {
    return rootAttrs;
  }
  return rootAttrs.slice(0, match.index) + rootAttrs.slice(match.index + match[0].length);
}
