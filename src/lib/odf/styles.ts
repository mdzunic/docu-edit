// Builds a style catalog from an ODF package's two style sources — automatic
// styles in content.xml (the "P1"/"T1" styles direct formatting is expressed
// through) and named styles in styles.xml ("Standard", "Heading_20_1", ...) —
// and resolves a style name's `style:parent-style-name` chain into effective
// paragraph/run formatting, the ODF counterpart of docx/styles.ts.
import type { Margins, PageSize, ParagraphFormatting, RunFormatting } from "../documentModel";
import { childElements, firstChildElement } from "../docx/xml";
import { mergeParagraphFormatting, mergeRunFormatting } from "../docx/properties";
import { foAttr, styleAttr, svgAttr, OFFICE_NAMESPACE, STYLE_NAMESPACE } from "./xml";
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
  parentName?: string;
  paragraphFormatting: ParagraphFormatting;
  textProperties: OdfTextProperties;
  /** `style:default-outline-level` — makes the style a heading style. */
  outlineLevel?: number;
}

/** Styles of one family ("paragraph" or "text"). Automatic and named styles
 * live in separate maps because their names are separate namespaces in ODF —
 * a lookup tries automatic (content.xml) first, then named (styles.xml). */
interface StyleFamily {
  automatic: Map<string, OdfStyleDefinition>;
  named: Map<string, OdfStyleDefinition>;
}

export interface OdfStyleCatalog {
  paragraphStyles: StyleFamily;
  textStyles: StyleFamily;
  /** From `style:default-style` (family "paragraph") in styles.xml — the
   * base of every resolution chain, like docx's docDefaults. */
  defaultParagraphFormatting: ParagraphFormatting;
  defaultRunFormatting: RunFormatting;
  /** From the first `style:master-page`'s `style:page-layout` in styles.xml. */
  pageSize?: PageSize;
  margins?: Margins;
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
  const outlineLevel = styleAttr(styleEl, "default-outline-level");
  if (outlineLevel) {
    const level = Number(outlineLevel);
    if (Number.isInteger(level) && level > 0) {
      definition.outlineLevel = level;
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

  // The first master page is the one body text starts on — the ODF
  // equivalent of the body-level sectPr the .docx reader applies.
  const masterPage = firstChildElement(masterStyles, STYLE_NAMESPACE, "master-page");
  const layoutName = masterPage ? styleAttr(masterPage, "page-layout-name") : null;
  const layout = layoutName ? layouts.get(layoutName) : undefined;
  const props = layout ? firstChildElement(layout, STYLE_NAMESPACE, "page-layout-properties") : undefined;
  if (!props) {
    return;
  }

  let widthCm = odfLengthToCm(foAttr(props, "page-width"));
  let heightCm = odfLengthToCm(foAttr(props, "page-height"));
  if (widthCm !== undefined && heightCm !== undefined) {
    // ODF stores landscape pages with width/height already swapped;
    // normalise defensively like the .docx reader does for w:orient.
    if (styleAttr(props, "print-orientation")?.toLowerCase() === "landscape" && widthCm < heightCm) {
      [widthCm, heightCm] = [heightCm, widthCm];
    }
    catalog.pageSize = { widthCm, heightCm };
  }

  const sharedMargin = odfLengthToCm(foAttr(props, "margin"));
  const topCm = odfLengthToCm(foAttr(props, "margin-top")) ?? sharedMargin;
  const rightCm = odfLengthToCm(foAttr(props, "margin-right")) ?? sharedMargin;
  const bottomCm = odfLengthToCm(foAttr(props, "margin-bottom")) ?? sharedMargin;
  const leftCm = odfLengthToCm(foAttr(props, "margin-left")) ?? sharedMargin;
  if (topCm !== undefined && rightCm !== undefined && bottomCm !== undefined && leftCm !== undefined) {
    catalog.margins = { topCm, rightCm, bottomCm, leftCm };
  }
}

/** Parses the styles of an ODF package: automatic styles from content.xml,
 * named styles / defaults / page layout from styles.xml. `stylesDoc` may be
 * missing (a degenerate package) — content still opens without it. */
export function parseOdfStyleCatalog(contentDoc: Document, stylesDoc: Document | undefined): OdfStyleCatalog {
  const catalog: OdfStyleCatalog = {
    paragraphStyles: emptyFamily(),
    textStyles: emptyFamily(),
    defaultParagraphFormatting: {},
    defaultRunFormatting: {},
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
    }
    parsePageLayout(stylesRoot, catalog);
  }

  const contentAutomatic = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "automatic-styles");
  collectStyles(contentAutomatic, fontFaces, (family, definition) => {
    const target = family === "paragraph" ? catalog.paragraphStyles : catalog.textStyles;
    target.automatic.set(definition.name, definition);
  });

  return catalog;
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
