// The ODF theme wire mapping (SPEC-094, hoisted for SPEC-077 AC-37). ODF has
// no first-class theme part, so the scheme rides as named `de-theme-*` styles
// under `office:styles` in styles.xml — valid ODF, ignorable by Impress and
// Writer, read back by name:
//   - `de-theme-{slot}`            family "graphic", `draw:fill-color` (fills parent it)
//   - `de-theme-text-{slot}`       family "text", `fo:color` (run colours parent it)
//   - `de-theme-paragraph-{slot}`  family "paragraph", `fo:background-color`
//                                  (paragraph shading parents it — the text
//                                  product's block only, SPEC-077)
//   - `de-theme-cell-{slot}`       family "table-cell", `fo:background-color`
//                                  (cell shading parents it — likewise)
//   - `de-theme-major/minor`       family "text", the heading/body font
//   - `de-theme`                   family "graphic", the theme name carrier
// A content ref rides `style:parent-style-name` on the automatic style beside
// the resolved literal (other readers see the right colour); a transformed ref
// has no faithful parent, so it writes the literal alone and re-opens as one —
// the stated loss, never hidden. Both products share this module; the deck
// block (odpThemeBlockXml) and the document block (odtThemeBlockXml) differ
// only in the extra paragraph/cell carrier families the text product needs.
import { parseXml, XmlParseError, escapeXmlAttribute } from "./xml";
import { THEME_COLOR_SLOTS, isThemeColor, themeSlotLabel } from "./theme";
import type { DocumentTheme, ModelColor, ThemeColorSlot } from "./theme";

const STYLE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const OFFICE_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const FO_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

/** The slot a `style:parent-style-name` marks, or undefined for a non-theme
 * parent — the readers' ref capture (every carrier family form). */
export function themeParentSlot(parent: string | undefined): ThemeColorSlot | undefined {
  if (!parent) {
    return undefined;
  }
  const match = /^de-theme-(?:text-|paragraph-|cell-)?([a-zA-Z0-9]+)$/.exec(parent);
  if (!match) {
    return undefined;
  }
  const slot = match[1];
  return (THEME_COLOR_SLOTS as string[]).includes(slot) ? (slot as ThemeColorSlot) : undefined;
}

/** The slot of an UNTRANSFORMED colour reference (no lumMod/lumOff/shade/
 * tint) — the only kind `style:parent-style-name` can carry faithfully
 * (SPEC-077 AC-38). A literal or a transformed reference returns undefined:
 * the writer then emits the resolved literal alone, the stated loss
 * (AC-39). */
export function untransformedThemeSlot(color: ModelColor | undefined): ThemeColorSlot | undefined {
  if (!isThemeColor(color)) {
    return undefined;
  }
  return color.lumMod === undefined && color.lumOff === undefined && color.shade === undefined && color.tint === undefined
    ? color.theme
    : undefined;
}

/** Whether a style name belongs to the minted `de-theme-*` carrier block —
 * the styles readThemeStyles recovers into `model.theme` and the writers'
 * splice regenerates (THEME_STYLE_RE strips the same prefix). Carrier styles
 * are the theme's wire form, not part of the document's own style catalogue,
 * so the odf reader skips them when building `model.styles`. */
export function isThemeCarrierStyleName(name: string): boolean {
  return name === "de-theme" || name.startsWith("de-theme-");
}

/** The parent style name an untransformed colour ref writes on a graphic
 * automatic style. */
export function themeGraphicParent(slot: ThemeColorSlot): string {
  return `de-theme-${slot}`;
}

/** The parent style name an untransformed colour ref writes on a text
 * automatic style. */
export function themeTextParent(slot: ThemeColorSlot): string {
  return `de-theme-text-${slot}`;
}

/** The parent style name an untransformed paragraph-shading ref writes on a
 * paragraph automatic style (SPEC-077). */
export function themeParagraphParent(slot: ThemeColorSlot): string {
  return `de-theme-paragraph-${slot}`;
}

/** The parent style name an untransformed cell-shading ref writes on a
 * table-cell automatic style (SPEC-077). */
export function themeCellParent(slot: ThemeColorSlot): string {
  return `de-theme-cell-${slot}`;
}

/** The canonical `de-theme-*` block for a deck theme (SPEC-094 AC-30) — what
 * the odp golden test pins. Byte-for-byte the block SPEC-094 shipped. */
export function odpThemeBlockXml(theme: DocumentTheme): string {
  const name = `<style:style style:name="de-theme" style:family="graphic" style:display-name="${escapeXmlAttribute(theme.name)}"/>`;
  const slots = THEME_COLOR_SLOTS.map((slot) => {
    const hex = theme.colors[slot];
    const label = escapeXmlAttribute(themeSlotLabel(slot));
    return (
      `<style:style style:name="de-theme-${slot}" style:family="graphic" style:display-name="${label}">` +
      `<style:graphic-properties draw:fill="solid" draw:fill-color="${hex}"/></style:style>` +
      `<style:style style:name="de-theme-text-${slot}" style:family="text" style:display-name="${label}">` +
      `<style:text-properties fo:color="${hex}"/></style:style>`
    );
  }).join("");
  const fonts = (["major", "minor"] as const)
    .map((which) => {
      const family = escapeXmlAttribute(theme.fonts[which]);
      return (
        `<style:style style:name="de-theme-${which}" style:family="text">` +
        `<style:text-properties style:font-name="${family}" fo:font-family="${family}"/></style:style>`
      );
    })
    .join("");
  return name + slots + fonts;
}

/** The text product's `de-theme-*` block (SPEC-077 AC-37): the shared deck
 * block plus the paragraph and table-cell carrier styles the word processor's
 * shading references parent. */
export function odtThemeBlockXml(theme: DocumentTheme): string {
  const carriers = THEME_COLOR_SLOTS.map((slot) => {
    const hex = theme.colors[slot];
    const label = escapeXmlAttribute(themeSlotLabel(slot));
    return (
      `<style:style style:name="de-theme-paragraph-${slot}" style:family="paragraph" style:display-name="${label}">` +
      `<style:paragraph-properties fo:background-color="${hex}"/></style:style>` +
      `<style:style style:name="de-theme-cell-${slot}" style:family="table-cell" style:display-name="${label}">` +
      `<style:table-cell-properties fo:background-color="${hex}"/></style:style>`
    );
  }).join("");
  return odpThemeBlockXml(theme) + carriers;
}

const THEME_STYLE_RE =
  /<style:style style:name="de-theme[^"]*"[^>]*\/>|<style:style style:name="de-theme[^"]*"[^>]*>[\s\S]*?<\/style:style>/g;

/** Rewrites styles.xml's theme block in place: any existing `de-theme-*`
 * styles are removed and `blockXml` spliced into `office:styles` — every
 * other byte of styles.xml survives. Returns the input unchanged when there
 * is no `office:styles` element to splice into. */
export function spliceThemeStyles(stylesXml: string, blockXml: string): string {
  const stripped = stylesXml.replace(THEME_STYLE_RE, "");
  if (/<\/office:styles>/.test(stripped)) {
    return stripped.replace(/<\/office:styles>/, `${blockXml}</office:styles>`);
  }
  if (/<office:styles\/>/.test(stripped)) {
    return stripped.replace(/<office:styles\/>/, `<office:styles>${blockXml}</office:styles>`);
  }
  return stylesXml;
}

/** Reads a theme back from styles.xml's `de-theme-*` block, or undefined
 * when the document carries none — `model.theme` stays absent and every
 * colour stays a literal. Unknown or extension theme markup (e.g.
 * LibreOffice's own block) is ignored here and preserved verbatim by the
 * writers' splice discipline. */
export function readThemeStyles(
  stylesXmlText: string,
  fallback: DocumentTheme,
): (DocumentTheme & { partName: string }) | undefined {
  let doc: Document;
  try {
    doc = parseXml(stylesXmlText);
  } catch (error) {
    if (error instanceof XmlParseError) {
      return undefined;
    }
    throw error;
  }
  const styles = doc.getElementsByTagNameNS(OFFICE_NAMESPACE, "styles")[0];
  if (!styles) {
    return undefined;
  }
  const colors = { ...fallback.colors };
  const fonts = { ...fallback.fonts };
  let themeName: string | undefined;
  let sawSlot = false;
  for (const el of Array.from(styles.children)) {
    if (el.namespaceURI !== STYLE_NAMESPACE || el.localName !== "style") {
      continue;
    }
    const name = el.getAttributeNS(STYLE_NAMESPACE, "name") ?? "";
    if (name === "de-theme") {
      themeName = el.getAttributeNS(STYLE_NAMESPACE, "display-name") ?? undefined;
      continue;
    }
    const slotMatch = /^de-theme-([a-zA-Z0-9]+)$/.exec(name);
    if (slotMatch && (THEME_COLOR_SLOTS as string[]).includes(slotMatch[1])) {
      const props = el.getElementsByTagNameNS(STYLE_NAMESPACE, "graphic-properties")[0];
      const hex = props?.getAttributeNS(DRAW_NAMESPACE, "fill-color");
      if (hex && /^#[0-9a-fA-F]{6}$/.test(hex)) {
        colors[slotMatch[1] as ThemeColorSlot] = hex.toLowerCase();
        sawSlot = true;
      }
      continue;
    }
    if (name === "de-theme-major" || name === "de-theme-minor") {
      const props = el.getElementsByTagNameNS(STYLE_NAMESPACE, "text-properties")[0];
      const family = props?.getAttributeNS(STYLE_NAMESPACE, "font-name") ?? props?.getAttributeNS(FO_NAMESPACE, "font-family");
      if (family) {
        fonts[name === "de-theme-major" ? "major" : "minor"] = family;
        sawSlot = true;
      }
    }
  }
  if (!sawSlot) {
    return undefined;
  }
  return { name: themeName ?? "Custom", colors, fonts, partName: "styles.xml" };
}
