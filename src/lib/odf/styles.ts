// Builds a style catalog from an ODF package's two style sources — automatic
// styles in content.xml (the "P1"/"T1" styles direct formatting is expressed
// through) and named styles in styles.xml ("Standard", "Heading_20_1", ...) —
// and resolves a style name's `style:parent-style-name` chain into effective
// paragraph/run formatting, the ODF counterpart of docx/styles.ts.
import type {
  ListMarkerStyle,
  ListType,
  Margins,
  NamedStyle,
  PageSize,
  ParagraphFormatting,
  RunFormatting,
} from "../documentModel";
import { childElements, firstChildElement } from "../docx/xml";
import { mergeParagraphFormatting, mergeRunFormatting } from "../docx/properties";
import { bulletStyleForChar, clampListLevel, defaultMarkerStyle } from "../lists";
import { foAttr, styleAttr, svgAttr, OFFICE_NAMESPACE, STYLE_NAMESPACE, TEXT_NAMESPACE, textAttr } from "./xml";
import {
  cleanFontFamily,
  paragraphPropertiesOf,
  parseParagraphProperties,
  parseTextProperties,
  textPropertiesOf,
  type FontFaceMap,
  type OdfTextProperties,
} from "./properties";
import { odfLengthToCm } from "./units";

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
}

/** Styles of one family ("paragraph" or "text"). Automatic and named styles
 * live in separate maps because their names are separate namespaces in ODF —
 * a lookup tries automatic (content.xml) first, then named (styles.xml). */
interface StyleFamily {
  automatic: Map<string, OdfStyleDefinition>;
  named: Map<string, OdfStyleDefinition>;
}

/** One level of a `text:list-style` (SPEC-018): the marker it draws and
 * the text indent its list-level properties define. */
export interface OdfListLevelDef {
  type: ListType;
  marker: ListMarkerStyle;
  startAt: number;
  indentCm?: number;
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
}

export interface OdfStyleCatalog {
  paragraphStyles: StyleFamily;
  textStyles: StyleFamily;
  /** `text:list-style` definitions by name — named styles from styles.xml,
   * shadowed by automatic ones from content.xml (SPEC-018). */
  listStyles: Map<string, OdfListStyle>;
  /** From `style:default-style` (family "paragraph") in styles.xml — the
   * base of every resolution chain, like docx's docDefaults. */
  defaultParagraphFormatting: ParagraphFormatting;
  defaultRunFormatting: RunFormatting;
  /** From the first `style:master-page`'s `style:page-layout` in styles.xml. */
  pageSize?: PageSize;
  margins?: Margins;
  /** Every `style:master-page`'s resolved geometry by master-page name
   * (SPEC-028 AC-33) — the lookup a section-starting paragraph's
   * `style:master-page-name` reference resolves against. The first master
   * page also stays folded into `pageSize`/`margins` above, keeping
   * single-section documents exactly as before. */
  masterPageLayouts: Map<string, OdfMasterPageLayout>;
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
  const definition: OdfStyleDefinition = {
    name,
    parentName: styleAttr(styleEl, "parent-style-name") ?? undefined,
    paragraphFormatting: parseParagraphProperties(paragraphPropertiesOf(styleEl)),
    textProperties: parseTextProperties(textPropertiesOf(styleEl), fontFaces),
  };
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
    } else {
      continue;
    }
    const indentCm = parseListLevelIndent(levelEl);
    if (indentCm !== undefined) {
      def.indentCm = indentCm;
    }
    if (!levels.has(level)) {
      levels.set(level, def);
    }
  }
  return levels;
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
  return layout;
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
    const geometry = props ? parsePageLayoutProperties(props) : {};
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
        catalog.defaultParagraphFormatting = parseParagraphProperties(paragraphPropertiesOf(defaultEl));
        // A percentage font size on the default style has nothing to resolve
        // against, so only the absolute formatting is kept.
        catalog.defaultRunFormatting = parseTextProperties(textPropertiesOf(defaultEl), fontFaces).formatting;
      }
      collectStyles(officeStyles, fontFaces, (family, definition) => {
        const target = family === "paragraph" ? catalog.paragraphStyles : catalog.textStyles;
        target.named.set(definition.name, definition);
      });
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

  for (const style of styleChain(catalog.paragraphStyles, styleName)) {
    paragraphFormatting = mergeParagraphFormatting(paragraphFormatting, style.paragraphFormatting);
    runFormatting = applyTextProperties(runFormatting, style.textProperties);
    if (style.outlineLevel !== undefined) {
      headingLevel = style.outlineLevel;
    }
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
    const style: NamedStyle = {
      id: definition.name,
      name: definition.displayName ?? decodeStyleName(definition.name),
    };
    if (definition.parentName !== undefined) {
      style.basedOn = definition.parentName;
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
  masterPageName?: string;
}

export function paragraphAutoStyleSignals(
  catalog: OdfStyleCatalog,
  styleName: string | undefined,
): OdfParagraphSignals {
  const definition = styleName ? catalog.paragraphStyles.automatic.get(styleName) : undefined;
  return {
    pageBreakBefore: definition?.paragraphFormatting.pageBreakBefore === true,
    masterPageName: definition?.masterPageName,
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
  return formatting;
}
