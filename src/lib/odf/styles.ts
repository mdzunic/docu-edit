// Builds a style catalog from an ODF package's two style sources — automatic
// styles in content.xml (the "P1"/"T1" styles direct formatting is expressed
// through) and named styles in styles.xml ("Standard", "Heading_20_1", ...) —
// and resolves a style name's `style:parent-style-name` chain into effective
// paragraph/run formatting, the ODF counterpart of docx/styles.ts.
import type {
  NamedStyle,
  PageBorders,
  ParagraphFormatting,
  RunFormatting,
  SectionColumns,
  TableStyleBanding,
} from "../documentModel";
import type { ListDefinition, ListMarkerStyle, ListType } from "../core/lists";
import type { Margins, PageSize } from "../core/pageSetup";
import {
  DEFAULT_COLUMN_GAP_CM,
  MAX_COLUMN_COUNT,
  mergeParagraphFormatting,
  mergeRunFormatting,
} from "../documentModel";
import { childElements, firstChildElement, serializeElement } from "../core/xml";
import { isThemeCarrierStyleName, themeParentSlot } from "../core/themeOdf";
import { bulletStyleForChar, clampListLevel, defaultMarkerStyle } from "../core/lists";
import { foAttr, styleAttr, svgAttr, OFFICE_NAMESPACE, STYLE_NAMESPACE, TEXT_NAMESPACE, textAttr } from "../core/odfXml";
import {
  cleanFontFamily,
  paragraphPropertiesOf,
  parseParagraphProperties,
  parseTextProperties,
  textPropertiesOf,
  type FontFaceMap,
  type OdfTextProperties,
} from "./properties";
import { odfColorToCss, odfLengthToCm, odfLengthToPt } from "./units";
import { parseTableTemplates } from "./tableTemplate";

/** The XSL-FO compatible namespace — `foAttr` reads it, but a style
 * element's own attribute needs the URI directly (SPEC-076). */
const FO_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
import { DE_BORDER_NAMESPACE, parseBorderShorthand } from "./borders";
import { borderSideVisible } from "../core/borders";

interface OdfStyleDefinition {
  name: string;
  /** `style:display-name` — the user-facing name when the style:name had to
   * be encoded ("Heading_20_1" displays as "Heading 1"). */
  displayName?: string;
  parentName?: string;
  paragraphFormatting: ParagraphFormatting;
  textProperties: OdfTextProperties;
  /** `style:default-outline-level` — makes the style a heading style. */
  outlineLevel?: number;
  /** `style:master-page-name` on the style:style element itself — a
   * paragraph using the style starts on a new page under that master page,
   * i.e. it begins a new section (SPEC-028 AC-33). */
  masterPageName?: string;
  /** `style:next-style-name` (SPEC-042 AC-3/AC-4's "style for following
   * paragraph") — the ODF counterpart of docx's `w:next`. */
  nextStyleName?: string;
  /** `fo:break-before="column"` on the style's paragraph-properties. Kept
   * off ParagraphFormatting — the model has no column-break formatting flag
   * (columns are SPEC-030) — so an AUTOMATIC style carrying it materialises
   * a column BreakBlock on open instead (SPEC-028 AC-10/11); on a named
   * style it stays unmodelled. */
  columnBreakBefore?: boolean;
  /** `style:page-number` on the style's paragraph-properties (SPEC-029
   * AC-22): an explicit positive integer restarts page numbering at this
   * paragraph; "auto" (continue) and absent both leave this unset. Like
   * `masterPageName`, only meaningful on an AUTOMATIC style — the section-
   * starting paragraph's own. */
  pageNumberStart?: number;
}

/** Styles of one family ("paragraph" or "text"). Automatic and named styles
 * live in separate maps because their names are separate namespaces in ODF —
 * a lookup tries automatic (content.xml) first, then named (styles.xml). */
interface StyleFamily {
  automatic: Map<string, OdfStyleDefinition>;
  named: Map<string, OdfStyleDefinition>;
}

/** One level of a `text:list-style` (SPEC-018): the marker it draws and
 * the text indent its list-level properties define. SPEC-032 widens it with
 * the multilevel-definition fields (prefix/suffix, display levels,
 * alignment, marker position) — see listStyleIsComposite. */
export interface OdfListLevelDef {
  type: ListType;
  marker: ListMarkerStyle;
  startAt: number;
  indentCm?: number;
  /** style:num-prefix, when present and non-empty. */
  prefix?: string;
  /** style:num-suffix, when present and non-empty. */
  suffix?: string;
  /** text:display-levels — stored whenever the ATTRIBUTE is present, even
   * at its default of 1: our own SPEC-032 writer always emits it, which is
   * the discriminator that keeps a re-opened definition composite. */
  displayLevels?: number;
  /** fo:text-align on the level's list-level-properties. */
  alignment?: "left" | "center" | "right";
  /** The marker's own position (label-alignment margin-left + text-indent),
   * in cm. */
  numberIndentCm?: number;
}

/** A named `text:list-style`'s levels, keyed by 1-based level. */
export type OdfListStyle = Map<number, OdfListLevelDef>;

/** One master page's page geometry, resolved through its
 * `style:page-layout-name` (SPEC-028 AC-33). Either field may be absent
 * when the layout omits it — the section then falls back to the document
 * defaults. */
export interface OdfMasterPageLayout {
  pageSize?: PageSize;
  margins?: Margins;
  /** The page layout's `style:columns` as a section column layout
   * (SPEC-030); absent for a single-column page. */
  columns?: SectionColumns;
  /** The page layout's border (SPEC-074 AC-31): `fo:border-*` with
   * `fo:padding-*` as the offset. */
  pageBorders?: PageBorders;
  /** The page layout's `fo:background-color` (SPEC-074 AC-31). */
  pageBackground?: string;
  /** `<style:header>`/`<style:header-left>`/`<style:footer>`/
   * `<style:footer-left>` — direct children of `<style:master-page>`
   * (SPEC-029 AC-9, AC-26). The `-left` variants are ODF's own
   * representation of "different odd and even pages" (their mere presence
   * means it is on, AC-12) — there is no separate boolean flag to read.
   * Content parsing happens in reader.ts (parseTextBlock et al.), not here,
   * to avoid a styles.ts -> reader.ts import cycle. */
  headerEl?: Element;
  headerLeftEl?: Element;
  footerEl?: Element;
  footerLeftEl?: Element;
}

/** SPEC-032 read state the reader threads through the parse via the
 * catalog (the one object every parsing helper already receives): the
 * document-wide list-instance id mint and the ListDefinition registry
 * composite list styles reconstruct into. Absent on catalogs that never
 * parse body content (the styles.xml-scoped one). */
export interface OdfListReadState {
  nextListId(): string;
  definitionFor(styleName: string | undefined): ListDefinition;
  definitions: ListDefinition[];
}

/** A named style of a family SPEC-042 does not model (anything but
 * "paragraph"/"text" — "table", "graphic", "list", …), captured only well
 * enough to list/preserve it (AC-15..18): identity plus verbatim source
 * XML, never resolved into formatting. */
export interface OdfOtherStyle {
  name: string;
  displayName?: string;
  parentName?: string;
  nextStyleName?: string;
  rawXml: string;
  /** SPEC-076 AC-69: the banding a `table:table-template` of this name
   * declares. A table style carrying one is fully modelled — banding is the
   * only conditional formatting the model has — so it is NOT re-emitted
   * verbatim and Modify works on it (AC-71's ODF half). */
  banding?: TableStyleBanding;
  /** The style's own `style:table-properties/fo:background-color` — the base
   * cell shading SPEC-042 already round-trips. */
  cellShadingFill?: string;
}

export interface OdfStyleCatalog {
  paragraphStyles: StyleFamily;
  textStyles: StyleFamily;
  /** Named styles of every other family, by name (SPEC-042) — see
   * OdfOtherStyle. */
  otherStyles: Map<string, OdfOtherStyle>;
  /** `text:list-style` definitions by name — named styles from styles.xml,
   * shadowed by automatic ones from content.xml (SPEC-018). */
  listStyles: Map<string, OdfListStyle>;
  /** SPEC-032 list-instance read state; see OdfListReadState. */
  listRead?: OdfListReadState;
  /** From `style:default-style` (family "paragraph") in styles.xml — the
   * base of every resolution chain, like docx's docDefaults. */
  defaultParagraphFormatting: ParagraphFormatting;
  defaultRunFormatting: RunFormatting;
  /** The default style's language (SPEC-035), held separately from
   * `defaultRunFormatting` — see the lift-out comment in
   * parseOdfStyleCatalog (mirrors docx/styles.ts's docDefaultLanguage). */
  defaultLanguage?: string;
  /** The document's default tab interval (SPEC-054): `style:tab-stop-distance`
   * on the default paragraph style. Undefined means the model's 1.27 cm
   * default applies. */
  defaultTabStopCm?: number;
  /** From the first `style:master-page`'s `style:page-layout` in styles.xml. */
  pageSize?: PageSize;
  margins?: Margins;
  /** The first master page's `style:columns` (SPEC-030) — the leading
   * section's column layout when it names no master of its own. */
  columns?: SectionColumns;
  /** Every `style:master-page`'s resolved geometry by master-page name
   * (SPEC-028 AC-33) — the lookup a section-starting paragraph's
   * `style:master-page-name` reference resolves against. The first master
   * page also stays folded into `pageSize`/`margins` above, keeping
   * single-section documents exactly as before. */
  masterPageLayouts: Map<string, OdfMasterPageLayout>;
  /** The first `style:master-page`'s own name (SPEC-029) — the lookup key
   * into `masterPageLayouts` for the document's own header/footer content
   * when no paragraph names a master explicitly (the common single-section
   * case; `pageSize`/`margins` fold that master's geometry in already, but
   * its header/footer elements need the object itself). */
  firstMasterName?: string;
}

function emptyFamily(): StyleFamily {
  return { automatic: new Map(), named: new Map() };
}

/** Collects `office:font-face-decls` (present in both content.xml and
 * styles.xml) into a name → family map for `style:font-name` resolution. */
function collectFontFaces(root: Element | undefined, fontFaces: FontFaceMap): void {
  if (!root) {
    return;
  }
  const decls = firstChildElement(root, OFFICE_NAMESPACE, "font-face-decls");
  if (!decls) {
    return;
  }
  for (const face of childElements(decls, STYLE_NAMESPACE, "font-face")) {
    const name = styleAttr(face, "name");
    if (!name) {
      continue;
    }
    const familyValue = svgAttr(face, "font-family");
    const family = familyValue ? cleanFontFamily(familyValue) : undefined;
    fontFaces.set(name, family ?? name);
  }
}

function parseStyleDefinition(styleEl: Element, fontFaces: FontFaceMap): OdfStyleDefinition | undefined {
  const name = styleAttr(styleEl, "name");
  if (!name) {
    return undefined;
  }
  const paragraphProps = paragraphPropertiesOf(styleEl);
  const definition: OdfStyleDefinition = {
    name,
    parentName: styleAttr(styleEl, "parent-style-name") ?? undefined,
    paragraphFormatting: parseParagraphProperties(paragraphProps),
    textProperties: parseTextProperties(textPropertiesOf(styleEl), fontFaces),
  };
  if (paragraphProps && foAttr(paragraphProps, "break-before") === "column") {
    definition.columnBreakBefore = true;
  }
  const displayName = styleAttr(styleEl, "display-name");
  if (displayName) {
    definition.displayName = displayName;
  }
  const outlineLevel = styleAttr(styleEl, "default-outline-level");
  if (outlineLevel) {
    const level = Number(outlineLevel);
    if (Number.isInteger(level) && level > 0) {
      definition.outlineLevel = level;
    }
  }
  const masterPageName = styleAttr(styleEl, "master-page-name");
  if (masterPageName) {
    definition.masterPageName = masterPageName;
  }
  const nextStyleName = styleAttr(styleEl, "next-style-name");
  if (nextStyleName) {
    definition.nextStyleName = nextStyleName;
  }
  if (paragraphProps) {
    const pageNumber = styleAttr(paragraphProps, "page-number");
    if (pageNumber !== null && pageNumber !== "auto") {
      const parsed = Number(pageNumber);
      if (Number.isInteger(parsed) && parsed >= 0) {
        definition.pageNumberStart = parsed;
      }
    }
  }
  return definition;
}

function collectStyles(
  container: Element | undefined,
  fontFaces: FontFaceMap,
  into: (family: string, definition: OdfStyleDefinition) => void,
): void {
  if (!container) {
    return;
  }
  for (const styleEl of childElements(container, STYLE_NAMESPACE, "style")) {
    const family = styleAttr(styleEl, "family");
    if (family !== "paragraph" && family !== "text") {
      continue;
    }
    const definition = parseStyleDefinition(styleEl, fontFaces);
    if (definition) {
      into(family, definition);
    }
  }
}

// style:num-format values for the number marker styles, per ODF 19.500.
const NUMBER_STYLE_BY_FORMAT: Record<string, ListMarkerStyle> = {
  "1": "decimal",
  a: "lowerAlpha",
  A: "upperAlpha",
  i: "lowerRoman",
  I: "upperRoman",
};

/** The level's text indent from `style:list-level-properties`: the
 * label-alignment form's fo:margin-left when present, else the older
 * text:space-before + text:min-label-width sum. */
function parseListLevelIndent(levelEl: Element): number | undefined {
  const props = firstChildElement(levelEl, STYLE_NAMESPACE, "list-level-properties");
  if (!props) {
    return undefined;
  }
  const labelAlignment = firstChildElement(props, STYLE_NAMESPACE, "list-level-label-alignment");
  const marginLeft = labelAlignment ? odfLengthToCm(foAttr(labelAlignment, "margin-left")) : undefined;
  if (marginLeft !== undefined) {
    return marginLeft;
  }
  const spaceBefore = odfLengthToCm(textAttr(props, "space-before"));
  const minLabelWidth = odfLengthToCm(textAttr(props, "min-label-width"));
  if (spaceBefore === undefined && minLabelWidth === undefined) {
    return undefined;
  }
  return (spaceBefore ?? 0) + (minLabelWidth ?? 0);
}

/** The level's marker position from `style:list-level-properties`: the
 * label-alignment form's margin-left + text-indent (the indent is negative
 * for the usual hanging marker). */
function parseListLevelNumberIndent(levelEl: Element): number | undefined {
  const props = firstChildElement(levelEl, STYLE_NAMESPACE, "list-level-properties");
  if (!props) {
    return undefined;
  }
  const labelAlignment = firstChildElement(props, STYLE_NAMESPACE, "list-level-label-alignment");
  if (!labelAlignment) {
    return undefined;
  }
  const marginLeft = odfLengthToCm(foAttr(labelAlignment, "margin-left"));
  const textIndent = odfLengthToCm(foAttr(labelAlignment, "text-indent"));
  if (marginLeft === undefined || textIndent === undefined) {
    return undefined;
  }
  return Math.max(0, marginLeft + textIndent);
}

/** Parses one `text:list-style` element's bullet/number levels. Image
 * levels (and anything else) are skipped — an undefined level falls back
 * to defaults at the use site. */
function parseListStyle(listStyleEl: Element): OdfListStyle {
  const levels: OdfListStyle = new Map();
  for (const levelEl of Array.from(listStyleEl.children)) {
    if (levelEl.namespaceURI !== TEXT_NAMESPACE) {
      continue;
    }
    const levelValue = Number(textAttr(levelEl, "level") ?? "");
    if (!Number.isInteger(levelValue) || levelValue < 1) {
      continue;
    }
    const level = clampListLevel(levelValue);
    let def: OdfListLevelDef;
    if (levelEl.localName === "list-level-style-bullet") {
      const char = textAttr(levelEl, "bullet-char") ?? "";
      def = {
        type: "bullet",
        marker: bulletStyleForChar(char) ?? defaultMarkerStyle("bullet", level),
        startAt: 1,
      };
    } else if (levelEl.localName === "list-level-style-number") {
      const format = styleAttr(levelEl, "num-format") ?? "1";
      const start = Number(textAttr(levelEl, "start-value") ?? "1");
      def = {
        type: "number",
        marker: NUMBER_STYLE_BY_FORMAT[format] ?? "decimal",
        startAt: Number.isFinite(start) && start >= 0 ? start : 1,
      };
      const displayLevels = textAttr(levelEl, "display-levels");
      if (displayLevels !== null) {
        const parsed = Number(displayLevels);
        if (Number.isInteger(parsed) && parsed >= 1) {
          def.displayLevels = Math.min(parsed, level);
        }
      }
    } else {
      continue;
    }
    const prefix = styleAttr(levelEl, "num-prefix");
    if (prefix !== null && prefix !== "") {
      def.prefix = prefix;
    }
    const suffix = styleAttr(levelEl, "num-suffix");
    if (suffix !== null && suffix !== "") {
      def.suffix = suffix;
    }
    const props = firstChildElement(levelEl, STYLE_NAMESPACE, "list-level-properties");
    const align = props ? foAttr(props, "text-align") : null;
    if (align === "center" || align === "right") {
      def.alignment = align;
    }
    const indentCm = parseListLevelIndent(levelEl);
    if (indentCm !== undefined) {
      def.indentCm = indentCm;
    }
    const numberIndentCm = parseListLevelNumberIndent(levelEl);
    if (numberIndentCm !== undefined) {
      def.numberIndentCm = numberIndentCm;
    }
    if (!levels.has(level)) {
      levels.set(level, def);
    }
  }
  return levels;
}

/** Whether a list style is a full multilevel definition (SPEC-032) rather
 * than SPEC-018's simple per-level markers. Keyed on the fields SPEC-018's
 * own writer never produces: text:display-levels (our defined writer emits
 * it on every number level, default included, precisely as this marker),
 * any prefix, a number suffix other than ".", or any bullet suffix. Simple
 * styles keep importing as adjacency lists, so pre-SPEC-032 documents are
 * read exactly as before. */
export function listStyleIsComposite(style: OdfListStyle | undefined): boolean {
  if (!style) {
    return false;
  }
  for (const def of style.values()) {
    if (def.displayLevels !== undefined || def.prefix !== undefined || def.alignment !== undefined) {
      return true;
    }
    if (def.type === "number" && def.suffix !== undefined && def.suffix !== ".") {
      return true;
    }
    if (def.type === "bullet" && def.suffix !== undefined) {
      return true;
    }
  }
  return false;
}

function collectListStyles(container: Element | undefined, catalog: OdfStyleCatalog): void {
  if (!container) {
    return;
  }
  for (const listStyleEl of childElements(container, TEXT_NAMESPACE, "list-style")) {
    const name = styleAttr(listStyleEl, "name");
    if (name) {
      catalog.listStyles.set(name, parseListStyle(listStyleEl));
    }
  }
}

/** A `style:page-layout-properties` element's page geometry — width/height
 * (with the defensive landscape swap) and margins. */
function parsePageLayoutProperties(props: Element): OdfMasterPageLayout {
  const layout: OdfMasterPageLayout = {};

  let widthCm = odfLengthToCm(foAttr(props, "page-width"));
  let heightCm = odfLengthToCm(foAttr(props, "page-height"));
  if (widthCm !== undefined && heightCm !== undefined) {
    // ODF stores landscape pages with width/height already swapped;
    // normalise defensively like the .docx reader does for w:orient.
    if (styleAttr(props, "print-orientation")?.toLowerCase() === "landscape" && widthCm < heightCm) {
      [widthCm, heightCm] = [heightCm, widthCm];
    }
    layout.pageSize = { widthCm, heightCm };
  }

  const sharedMargin = odfLengthToCm(foAttr(props, "margin"));
  const topCm = odfLengthToCm(foAttr(props, "margin-top")) ?? sharedMargin;
  const rightCm = odfLengthToCm(foAttr(props, "margin-right")) ?? sharedMargin;
  const bottomCm = odfLengthToCm(foAttr(props, "margin-bottom")) ?? sharedMargin;
  const leftCm = odfLengthToCm(foAttr(props, "margin-left")) ?? sharedMargin;
  if (topCm !== undefined && rightCm !== undefined && bottomCm !== undefined && leftCm !== undefined) {
    layout.margins = { topCm, rightCm, bottomCm, leftCm };
  }

  const columns = parsePageColumns(props, layout);
  if (columns) {
    layout.columns = columns;
  }

  // SPEC-074 AC-31: the page border and background. The offset rides on
  // `fo:padding-*`; the reference edge and the applies-to restriction have no
  // ODF attribute and come back from this product's extension namespace (see
  // odf/borders.ts), defaulting to what ODF itself implies.
  const shorthand = foAttr(props, "border");
  const paddingAll = odfLengthToPt(foAttr(props, "padding"));
  const pageBorders: PageBorders = {};
  let offsetPt: number | undefined;
  for (const edge of ["top", "bottom", "left", "right"] as const) {
    const side = parseBorderShorthand(foAttr(props, `border-${edge}`) ?? shorthand);
    // An explicit `none` is a present edge for a PARAGRAPH — that is how a
    // paragraph cancels the border its style would give it. A page layout
    // inherits from nothing, so `fo:border="none"` here means exactly what
    // an absent attribute means: no page border. Recording it would give a
    // document that has none a `pageBorders` object (and an inert border
    // layer on every page) it did not have before.
    if (!side || !borderSideVisible(side)) {
      continue;
    }
    pageBorders[edge] = side;
    offsetPt = offsetPt ?? odfLengthToPt(foAttr(props, `padding-${edge}`)) ?? paddingAll;
  }
  if (pageBorders.top || pageBorders.bottom || pageBorders.left || pageBorders.right) {
    if (offsetPt !== undefined) {
      pageBorders.offsetPt = offsetPt;
    }
    pageBorders.measureFrom =
      props.getAttributeNS(DE_BORDER_NAMESPACE, "page-border-measure-from") === "text" ? "text" : "page";
    const appliesTo = props.getAttributeNS(DE_BORDER_NAMESPACE, "page-border-applies-to");
    if (appliesTo === "firstPage" || appliesTo === "exceptFirstPage") {
      pageBorders.appliesTo = appliesTo;
    }
    layout.pageBorders = pageBorders;
  }
  const pageBackground = odfColorToCss(foAttr(props, "background-color"));
  if (pageBackground !== undefined) {
    layout.pageBackground = pageBackground;
  }
  return layout;
}

/** `style:columns` inside a page layout (SPEC-030 AC-17): `fo:column-count`,
 * `fo:column-gap`, and `style:column` children whose `style:rel-width`
 * relative weights fix the unequal columns' proportions. A `style:column-sep`
 * child marks "line between columns". Single-column (count ≤ 1) or absent is
 * left unmodelled. */
function parsePageColumns(props: Element, layout: OdfMasterPageLayout): SectionColumns | undefined {
  const cols = firstChildElement(props, STYLE_NAMESPACE, "columns");
  if (!cols) {
    return undefined;
  }
  const rawCount = Number(foAttr(cols, "column-count") ?? "1");
  const count = Number.isFinite(rawCount)
    ? Math.max(1, Math.min(MAX_COLUMN_COUNT, Math.round(rawCount)))
    : 1;
  if (count <= 1) {
    return undefined;
  }
  const gapCm = odfLengthToCm(foAttr(cols, "column-gap")) ?? DEFAULT_COLUMN_GAP_CM;
  const separator = firstChildElement(cols, STYLE_NAMESPACE, "column-sep") !== undefined;
  const columnEls = childElements(cols, STYLE_NAMESPACE, "column");
  if (columnEls.length === count) {
    const weights = columnEls.map((col) => {
      const raw = styleAttr(col, "rel-width") ?? "";
      const value = Number.parseFloat(raw.replace(/\*$/, ""));
      return Number.isFinite(value) ? value : 0;
    });
    const total = weights.reduce((sum, w) => sum + w, 0);
    const contentWidthCm =
      layout.pageSize && layout.margins
        ? layout.pageSize.widthCm - layout.margins.leftCm - layout.margins.rightCm - gapCm * (count - 1)
        : undefined;
    if (weights.every((w) => w > 0) && total > 0 && contentWidthCm && contentWidthCm > 0) {
      const widthsCm = weights.map((w) => (w / total) * contentWidthCm);
      return { count, equalWidth: false, gapCm, separator, widthsCm };
    }
  }
  return { count, equalWidth: true, gapCm, separator };
}

function parsePageLayout(stylesRoot: Element, catalog: OdfStyleCatalog): void {
  const automaticStyles = firstChildElement(stylesRoot, OFFICE_NAMESPACE, "automatic-styles");
  const masterStyles = firstChildElement(stylesRoot, OFFICE_NAMESPACE, "master-styles");
  if (!automaticStyles || !masterStyles) {
    return;
  }

  const layouts = new Map<string, Element>();
  for (const layoutEl of childElements(automaticStyles, STYLE_NAMESPACE, "page-layout")) {
    const name = styleAttr(layoutEl, "name");
    if (name) {
      layouts.set(name, layoutEl);
    }
  }

  // Every master page's geometry is catalogued by name — the lookup a
  // section-starting style:master-page-name resolves (SPEC-028 AC-33). The
  // FIRST master page is the one body text starts on — the ODF equivalent
  // of the body-level sectPr the .docx reader applies — and stays folded
  // into the catalog's document-level pageSize/margins exactly as before.
  let first = true;
  for (const masterPage of childElements(masterStyles, STYLE_NAMESPACE, "master-page")) {
    const masterName = styleAttr(masterPage, "name");
    const layoutName = styleAttr(masterPage, "page-layout-name");
    const layout = layoutName ? layouts.get(layoutName) : undefined;
    const props = layout ? firstChildElement(layout, STYLE_NAMESPACE, "page-layout-properties") : undefined;
    const geometry: OdfMasterPageLayout = props ? parsePageLayoutProperties(props) : {};
    geometry.headerEl = firstChildElement(masterPage, STYLE_NAMESPACE, "header");
    geometry.headerLeftEl = firstChildElement(masterPage, STYLE_NAMESPACE, "header-left");
    geometry.footerEl = firstChildElement(masterPage, STYLE_NAMESPACE, "footer");
    geometry.footerLeftEl = firstChildElement(masterPage, STYLE_NAMESPACE, "footer-left");
    if (masterName) {
      catalog.masterPageLayouts.set(masterName, geometry);
    }
    if (first) {
      first = false;
      if (geometry.pageSize) {
        catalog.pageSize = geometry.pageSize;
      }
      if (geometry.margins) {
        catalog.margins = geometry.margins;
      }
      if (geometry.columns) {
        catalog.columns = geometry.columns;
      }
      catalog.firstMasterName = masterName ?? undefined;
    }
  }
}

/** Parses the styles of an ODF package: automatic styles from content.xml,
 * named styles / defaults / page layout from styles.xml. `stylesDoc` may be
 * missing (a degenerate package) — content still opens without it. */
export function parseOdfStyleCatalog(contentDoc: Document, stylesDoc: Document | undefined): OdfStyleCatalog {
  const catalog: OdfStyleCatalog = {
    paragraphStyles: emptyFamily(),
    textStyles: emptyFamily(),
    otherStyles: new Map(),
    listStyles: new Map(),
    defaultParagraphFormatting: {},
    defaultRunFormatting: {},
    masterPageLayouts: new Map(),
  };

  const fontFaces: FontFaceMap = new Map();
  collectFontFaces(stylesDoc?.documentElement, fontFaces);
  collectFontFaces(contentDoc.documentElement, fontFaces);

  const stylesRoot = stylesDoc?.documentElement;
  if (stylesRoot) {
    const officeStyles = firstChildElement(stylesRoot, OFFICE_NAMESPACE, "styles");
    if (officeStyles) {
      for (const defaultEl of childElements(officeStyles, STYLE_NAMESPACE, "default-style")) {
        if (styleAttr(defaultEl, "family") !== "paragraph") {
          continue;
        }
        const defaultParagraphProps = paragraphPropertiesOf(defaultEl);
        catalog.defaultParagraphFormatting = parseParagraphProperties(defaultParagraphProps);
        // SPEC-054: the document default tab interval is an attribute of the
        // default paragraph style's properties, not a per-paragraph value, so
        // it is lifted onto defaultStyles.defaultTabStopCm rather than kept in
        // defaultParagraphFormatting (which bakes onto every paragraph).
        const tabDistanceCm = defaultParagraphProps
          ? odfLengthToCm(styleAttr(defaultParagraphProps, "tab-stop-distance"))
          : undefined;
        if (tabDistanceCm !== undefined && tabDistanceCm > 0) {
          catalog.defaultTabStopCm = tabDistanceCm;
        }
        // A percentage font size on the default style has nothing to resolve
        // against, so only the absolute formatting is kept.
        catalog.defaultRunFormatting = parseTextProperties(textPropertiesOf(defaultEl), fontFaces).formatting;
        // SPEC-035: lifted out of defaultRunFormatting for the same reason
        // as docx/styles.ts's docDefaultLanguage — resolveParagraphStyle
        // below uses defaultRunFormatting as the base of every paragraph's
        // resolved formatting, baked directly onto each run on read, and a
        // document-default language should not become an explicit `lang`
        // on every run just because the document declared one. The default
        // lives on DocumentModel.defaultStyles.language instead (see
        // reader.ts); a run only carries `lang` when its own or a named
        // style's text-properties set one.
        if (catalog.defaultRunFormatting.lang !== undefined) {
          catalog.defaultLanguage = catalog.defaultRunFormatting.lang;
          const { lang: _lang, ...rest } = catalog.defaultRunFormatting;
          catalog.defaultRunFormatting = rest;
        }
      }
      collectStyles(officeStyles, fontFaces, (family, definition) => {
        const target = family === "paragraph" ? catalog.paragraphStyles : catalog.textStyles;
        target.named.set(definition.name, definition);
      });
      // SPEC-042 AC-15..18: named styles of every other family ("table",
      // "graphic", …) are listed/preserved verbatim rather than resolved —
      // this app has no formatting model for them (out of scope: banding).
      const tableTemplates = parseTableTemplates(officeStyles);
      for (const styleEl of childElements(officeStyles, STYLE_NAMESPACE, "style")) {
        const family = styleAttr(styleEl, "family");
        // A generated table-cell style belongs to a template, not to the
        // catalogue — parseTableTemplates already folded it into the banding
        // it describes (SPEC-076 AC-68).
        if (family === "paragraph" || family === "text" || family === "table-cell" || !family) {
          continue;
        }
        const name = styleAttr(styleEl, "name");
        if (!name) {
          continue;
        }
        const other: OdfOtherStyle = { name, rawXml: serializeElement(styleEl) };
        // SPEC-076: the table:table-template of the same name, if any.
        const banding = tableTemplates.get(name);
        if (family === "table" && banding !== undefined) {
          other.banding = banding;
          const tableProps = firstChildElement(styleEl, STYLE_NAMESPACE, "table-properties");
          const fill = tableProps?.getAttributeNS(FO_NAMESPACE, "background-color") ?? undefined;
          if (fill) {
            other.cellShadingFill = fill;
          }
        }
        const displayName = styleAttr(styleEl, "display-name");
        if (displayName) {
          other.displayName = displayName;
        }
        const parentName = styleAttr(styleEl, "parent-style-name");
        if (parentName) {
          other.parentName = parentName;
        }
        const nextStyleName = styleAttr(styleEl, "next-style-name");
        if (nextStyleName) {
          other.nextStyleName = nextStyleName;
        }
        catalog.otherStyles.set(name, other);
      }
      collectListStyles(officeStyles, catalog);
    }
    parsePageLayout(stylesRoot, catalog);
  }

  const contentAutomatic = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "automatic-styles");
  collectStyles(contentAutomatic, fontFaces, (family, definition) => {
    const target = family === "paragraph" ? catalog.paragraphStyles : catalog.textStyles;
    target.automatic.set(definition.name, definition);
  });
  // Automatic list styles shadow named ones of the same name (collected
  // after, so the Map insertion overwrites).
  collectListStyles(contentAutomatic, catalog);

  return catalog;
}

/** Resolves a list style's level definition for an item at `level`: the
 * exact level when defined, else a nearest-shallower-level fallback that
 * keeps the list's type with the level's default marker cycle. */
export function resolveListLevel(
  style: OdfListStyle | undefined,
  level: number,
): OdfListLevelDef | undefined {
  if (!style) {
    return undefined;
  }
  const exact = style.get(level);
  if (exact) {
    return exact;
  }
  for (let shallower = level - 1; shallower >= 1; shallower--) {
    const def = style.get(shallower);
    if (def) {
      return { type: def.type, marker: defaultMarkerStyle(def.type, level), startAt: 1 };
    }
  }
  return undefined;
}

/** The `style:parent-style-name` chain for a style name, root ancestor
 * first, with a cycle guard. Automatic styles shadow named ones. */
function styleChain(family: StyleFamily, styleName: string | undefined): OdfStyleDefinition[] {
  const chain: OdfStyleDefinition[] = [];
  const seen = new Set<string>();
  let currentName = styleName;
  while (currentName && !seen.has(currentName)) {
    seen.add(currentName);
    const style = family.automatic.get(currentName) ?? family.named.get(currentName);
    if (!style) {
      break;
    }
    chain.unshift(style);
    currentName = style.parentName;
  }
  return chain;
}

/** Layers one style's text properties over formatting inherited so far,
 * resolving a percentage font size ("130%") against the inherited size —
 * the reason resolution happens on the chain rather than per style. */
function applyTextProperties(current: RunFormatting, props: OdfTextProperties): RunFormatting {
  const merged = mergeRunFormatting(current, props.formatting);
  if (
    props.fontSizePercent !== undefined
    && props.formatting.fontSizePt === undefined
    && current.fontSizePt !== undefined
  ) {
    merged.fontSizePt = (current.fontSizePt * props.fontSizePercent) / 100;
  }
  return merged;
}

export interface ResolvedOdfParagraphStyle {
  paragraphFormatting: ParagraphFormatting;
  runFormatting: RunFormatting;
  headingLevel?: number;
}

/** Resolves a paragraph's `text:style-name` from the defaults up through the
 * parent chain into effective formatting — the ODF counterpart of docx
 * `resolveStyle`. Every property in ODF flows through styles (there is no
 * direct formatting in content), so this is the paragraph's full formatting. */
export function resolveParagraphStyle(
  catalog: OdfStyleCatalog,
  styleName: string | undefined,
): ResolvedOdfParagraphStyle {
  let paragraphFormatting = catalog.defaultParagraphFormatting;
  let runFormatting = catalog.defaultRunFormatting;
  let headingLevel: number | undefined;

  // `fo:border="none"` and `fo:background-color="transparent"` are taken at
  // their word here — an explicit "no box"/"no fill" that stops a style's
  // being inherited (SPEC-074 AC-17). The .docx reader instead drops such a
  // value where the style chain had nothing to override; ODF cannot make
  // that comparison, because the chain a paragraph resolves through does not
  // contain the common style its box would come from: this product's own
  // writer names a paragraph's automatic style after its style id (see
  // odf/writer's paragraphNameFor), and styleChain's automatic-before-named
  // lookup then stops at that shadowing entry. Reading the attribute
  // literally is also the more faithful reading — an ODF automatic style IS
  // direct formatting, so a border stated there was stated deliberately.
  for (const style of styleChain(catalog.paragraphStyles, styleName)) {
    paragraphFormatting = mergeParagraphFormatting(paragraphFormatting, style.paragraphFormatting);
    runFormatting = applyTextProperties(runFormatting, style.textProperties);
    if (style.outlineLevel !== undefined) {
      headingLevel = style.outlineLevel;
    }
  }

  // SPEC-077 AC-40: an automatic style parented on `de-theme-paragraph-{slot}`
  // carries a shading REFERENCE — the fo:background-color it also declares is
  // the resolved literal the writer put beside the link for other readers.
  const automatic = styleName ? catalog.paragraphStyles.automatic.get(styleName) : undefined;
  const slot = themeParentSlot(automatic?.parentName);
  if (slot && automatic?.paragraphFormatting.shadingFill !== undefined) {
    paragraphFormatting = { ...paragraphFormatting, shadingFill: { theme: slot } };
  }

  return { paragraphFormatting, runFormatting, headingLevel };
}

/** An ODF style:name with its `_XX_`-encoded characters decoded — the
 * fallback display name when the style carries no style:display-name
 * ("Heading_20_1" -> "Heading 1"). */
function decodeStyleName(name: string): string {
  return name.replace(/_([0-9a-fA-F]{2})_/g, (_, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  );
}

/** The catalog's *named* paragraph styles as the model's style catalogue
 * (SPEC-019), in styles.xml order. Automatic styles are deliberately not
 * included: an automatic style is a paragraph's direct formatting, already
 * baked into the model, and content.xml's automatic styles are regenerated
 * on every save — carrying them would break the unedited round trip's
 * structural equality (SPEC-008). A named style's own indents are stored as
 * parsed (logical-as-LTR); a style that itself turns on RTL swaps them the
 * same way the reader does after chain resolution. */
export function catalogToNamedStyles(catalog: OdfStyleCatalog): NamedStyle[] {
  const styles: NamedStyle[] = [];
  for (const definition of catalog.paragraphStyles.named.values()) {
    // The de-theme-* carrier styles (SPEC-077) are the theme's wire form,
    // recovered separately into model.theme by readThemeStyles — never part
    // of the document's own style catalogue (the writers regenerate the
    // block from model.theme, so listing them would double-write it).
    if (isThemeCarrierStyleName(definition.name)) {
      continue;
    }
    const style: NamedStyle = {
      id: definition.name,
      name: definition.displayName ?? decodeStyleName(definition.name),
    };
    if (definition.parentName !== undefined) {
      style.basedOn = definition.parentName;
    }
    if (definition.nextStyleName !== undefined) {
      style.next = definition.nextStyleName;
    }
    const paragraphFormatting = { ...definition.paragraphFormatting };
    if (paragraphFormatting.direction === "rtl") {
      const start = paragraphFormatting.indentStartCm;
      const end = paragraphFormatting.indentEndCm;
      if (end !== undefined) {
        paragraphFormatting.indentStartCm = end;
      } else {
        delete paragraphFormatting.indentStartCm;
      }
      if (start !== undefined) {
        paragraphFormatting.indentEndCm = start;
      } else {
        delete paragraphFormatting.indentEndCm;
      }
    }
    if (Object.keys(paragraphFormatting).length > 0) {
      style.paragraphFormatting = paragraphFormatting;
    }
    if (Object.keys(definition.textProperties.formatting).length > 0) {
      style.runFormatting = definition.textProperties.formatting;
    }
    if (definition.textProperties.fontSizePercent !== undefined) {
      style.fontSizePercent = definition.textProperties.fontSizePercent;
    }
    if (definition.outlineLevel !== undefined) {
      style.headingLevel = definition.outlineLevel;
    }
    styles.push(style);
  }
  // SPEC-042: named text (character) styles — the ODF family for
  // SPEC-042's character styles, resolved the same way (resolveTextStyle)
  // but never included here before this spec, since a run's direct
  // formatting already carried the resolved effect and nothing referenced
  // the style by id.
  for (const definition of catalog.textStyles.named.values()) {
    if (isThemeCarrierStyleName(definition.name)) {
      continue; // A de-theme-* carrier — see the paragraph loop above.
    }
    const style: NamedStyle = {
      id: definition.name,
      name: definition.displayName ?? decodeStyleName(definition.name),
      type: "character",
    };
    if (definition.parentName !== undefined) {
      style.basedOn = definition.parentName;
    }
    if (Object.keys(definition.textProperties.formatting).length > 0) {
      style.runFormatting = definition.textProperties.formatting;
    }
    if (definition.textProperties.fontSizePercent !== undefined) {
      style.fontSizePercent = definition.textProperties.fontSizePercent;
    }
    styles.push(style);
  }
  // SPEC-042 AC-15..18: styles of every other family, listed/preserved
  // verbatim — see OdfOtherStyle.
  for (const other of catalog.otherStyles.values()) {
    if (isThemeCarrierStyleName(other.name)) {
      continue; // A de-theme-* carrier — see the paragraph loop above.
    }
    const style: NamedStyle = {
      id: other.name,
      name: other.displayName ?? decodeStyleName(other.name),
      type: "table",
    };
    // SPEC-076 AC-69, AC-71: a style whose conditional formatting is banding
    // is modelled; anything else stays verbatim, as SPEC-042 left it.
    if (other.banding !== undefined) {
      style.banding = other.banding;
      if (other.cellShadingFill !== undefined) {
        style.cellShadingFill = other.cellShadingFill;
      }
    } else {
      style.unmodeled = true;
      style.rawXml = other.rawXml;
    }
    if (other.parentName !== undefined) {
      style.basedOn = other.parentName;
    }
    if (other.nextStyleName !== undefined) {
      style.next = other.nextStyleName;
    }
    styles.push(style);
  }
  return styles;
}

/** SPEC-028: the two paragraph-anchored break/section signals a body
 * paragraph's AUTOMATIC style may carry — `fo:break-before="page"` (a manual
 * page break, AC-32) and `style:master-page-name` (a section start, AC-33).
 * Automatic styles only, deliberately: they are the paragraph's direct
 * formatting, so a break-before there is the manual break the writer
 * emitted; the same attribute on a NAMED style (e.g. a heading style that
 * always starts a page) is style formatting and stays the model's
 * `pageBreakBefore` flag via ordinary chain resolution. */
export interface OdfParagraphSignals {
  pageBreakBefore: boolean;
  /** `fo:break-before="column"` (SPEC-028 AC-10/11): materialises a column
   * BreakBlock before the paragraph — a separate signal from
   * pageBreakBefore because the model has no column-break formatting flag
   * to canonicalise into (columns are SPEC-030). */
  columnBreakBefore: boolean;
  masterPageName?: string;
  /** `style:page-number` (SPEC-029 AC-22) — see OdfStyleDefinition. */
  pageNumberStart?: number;
}

export function paragraphAutoStyleSignals(
  catalog: OdfStyleCatalog,
  styleName: string | undefined,
): OdfParagraphSignals {
  const definition = styleName ? catalog.paragraphStyles.automatic.get(styleName) : undefined;
  return {
    pageBreakBefore: definition?.paragraphFormatting.pageBreakBefore === true,
    columnBreakBefore: definition?.columnBreakBefore === true,
    masterPageName: definition?.masterPageName,
    pageNumberStart: definition?.pageNumberStart,
  };
}

/** Resolves a span's `text:style-name` (family "text") over the formatting
 * inherited from its paragraph or enclosing span. */
export function resolveTextStyle(
  catalog: OdfStyleCatalog,
  styleName: string | undefined,
  inherited: RunFormatting,
): RunFormatting {
  let formatting = inherited;
  for (const style of styleChain(catalog.textStyles, styleName)) {
    formatting = applyTextProperties(formatting, style.textProperties);
  }
  // SPEC-077 AC-40: an automatic style parented on `de-theme-text-{slot}`
  // carries a colour REFERENCE — the fo:color it also declares is the
  // resolved literal the writer put beside the link for other readers.
  const automatic = styleName ? catalog.textStyles.automatic.get(styleName) : undefined;
  const slot = themeParentSlot(automatic?.parentName);
  if (slot && automatic?.textProperties.formatting.color !== undefined) {
    formatting = { ...formatting, color: { theme: slot } };
  }
  return formatting;
}
