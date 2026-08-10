// Parses word/styles.xml into a lookup table, and resolves a paragraph's
// `w:pStyle` chain (following `w:basedOn`) into effective paragraph/run
// formatting — this is what lets built-in heading styles (AC-25) render at
// their defined size and weight even when a paragraph carries no direct
// formatting of its own.
import type { NamedStyle, ParagraphFormatting, RunFormatting, StyleType } from "../documentModel";
import {
  childElements,
  firstChildElement,
  serializeElement,
  wordAttr as attr,
  WORD_NAMESPACE,
} from "./xml";
import {
  mergeParagraphFormatting,
  mergeRunFormatting,
  parseOutlineLevel,
  parseParagraphProperties,
  parseRunProperties,
} from "./properties";

/** `w:type` values SPEC-042 fully models — `"paragraph"` (SPEC-019) and
 * `"character"` (`w:rStyle`). `"table"`/`"numbering"` styles, and any
 * `"paragraph"`/`"character"` style carrying a child the reader does not
 * recognise (e.g. `w:tblStylePr` banding regions), come through as
 * `unmodeled` instead — see `parseStylesCatalog`. */
const MODELED_TYPES = new Set(["paragraph", "character"]);

interface StyleDefinition {
  id: string;
  name?: string;
  /** The raw `w:type` value ("paragraph", "character", "table",
   * "numbering") — narrowed to a model `StyleType` only in
   * `catalogToNamedStyles`, since "numbering" has no model equivalent. */
  type: string;
  basedOn?: string;
  next?: string;
  paragraphFormatting: ParagraphFormatting;
  runFormatting: RunFormatting;
  headingLevel?: number;
  unmodeled?: boolean;
  rawXml?: string;
}

function modelStyleType(rawType: string): StyleType {
  if (rawType === "character") {
    return "character";
  }
  // "table" and the rare legacy "numbering" family both group under
  // "table" in the pane — neither is a style the editor fully models
  // (see MODELED_TYPES), so the grouping only matters for display.
  return rawType === "paragraph" ? "paragraph" : "table";
}

export interface StyleCatalog {
  styles: Map<string, StyleDefinition>;
  docDefaultParagraphFormatting: ParagraphFormatting;
  docDefaultRunFormatting: RunFormatting;
  /** `docDefaults/rPrDefault/rPr/w:lang`, held separately from
   * `docDefaultRunFormatting` (SPEC-035) — see the lift-out comment in
   * parseStylesCatalog. */
  docDefaultLanguage?: string;
}

const HEADING_PATTERN = /^heading\s*([1-9])$/i;

function headingLevelFromName(...names: (string | undefined)[]): number | undefined {
  for (const name of names) {
    if (!name) {
      continue;
    }
    const match = HEADING_PATTERN.exec(name.trim());
    if (match) {
      return Number(match[1]);
    }
  }
  return undefined;
}

/** Parses `word/styles.xml`. Returns an empty catalog (not an error) if the
 * document has no styles part — direct formatting still works without one. */
export function parseStylesCatalog(stylesDoc: Document | undefined): StyleCatalog {
  const catalog: StyleCatalog = {
    styles: new Map(),
    docDefaultParagraphFormatting: {},
    docDefaultRunFormatting: {},
  };
  if (!stylesDoc?.documentElement) {
    return catalog;
  }

  const docDefaults = firstChildElement(stylesDoc.documentElement, WORD_NAMESPACE, "docDefaults");
  if (docDefaults) {
    const rPrDefault = firstChildElement(docDefaults, WORD_NAMESPACE, "rPrDefault");
    const rPr = rPrDefault ? firstChildElement(rPrDefault, WORD_NAMESPACE, "rPr") : undefined;
    catalog.docDefaultRunFormatting = parseRunProperties(rPr);
    // SPEC-035: the document-default language is lifted out of
    // docDefaultRunFormatting rather than left in it. resolveStyle below
    // uses docDefaultRunFormatting as the base of every paragraph/run's
    // resolved formatting, which the reader then bakes directly onto each
    // run (see docx/reader.ts's collectRuns) — exactly right for fontFamily/
    // fontSizePt, but it would stamp an explicit w:lang-equivalent onto
    // *every* run in the model just because the document declared a default,
    // making `RunFormatting.lang` presence meaningless as an "explicit
    // override" signal for the language selector/status bar and adding an
    // unexpected key to every run's formatting. The document default lives
    // on DocumentModel.defaultStyles.language instead (see reader.ts); a
    // run only carries `lang` when its own or a named style's rPr sets one.
    if (catalog.docDefaultRunFormatting.lang !== undefined) {
      catalog.docDefaultLanguage = catalog.docDefaultRunFormatting.lang;
      const { lang: _lang, ...rest } = catalog.docDefaultRunFormatting;
      catalog.docDefaultRunFormatting = rest;
    }

    const pPrDefault = firstChildElement(docDefaults, WORD_NAMESPACE, "pPrDefault");
    const pPr = pPrDefault ? firstChildElement(pPrDefault, WORD_NAMESPACE, "pPr") : undefined;
    catalog.docDefaultParagraphFormatting = parseParagraphProperties(pPr);
  }

  for (const styleEl of childElements(stylesDoc.documentElement, WORD_NAMESPACE, "style")) {
    const id = attr(styleEl, "styleId");
    if (!id) {
      continue;
    }
    const rawType = attr(styleEl, "type") ?? "paragraph";
    const nameEl = firstChildElement(styleEl, WORD_NAMESPACE, "name");
    const name = nameEl ? (attr(nameEl, "val") ?? undefined) : undefined;
    const basedOnEl = firstChildElement(styleEl, WORD_NAMESPACE, "basedOn");
    const basedOn = basedOnEl ? (attr(basedOnEl, "val") ?? undefined) : undefined;
    const nextEl = firstChildElement(styleEl, WORD_NAMESPACE, "next");
    const next = nextEl ? (attr(nextEl, "val") ?? undefined) : undefined;

    // SPEC-042: only paragraph and character styles are modeled — table
    // banding and legacy "numbering" styles are listed/preserved verbatim
    // instead (AC-15..18), never resolved or edited.
    if (!MODELED_TYPES.has(rawType)) {
      catalog.styles.set(id, {
        id,
        name,
        type: rawType,
        basedOn,
        next,
        paragraphFormatting: {},
        runFormatting: {},
        unmodeled: true,
        rawXml: serializeElement(styleEl),
      });
      continue;
    }

    const pPr = firstChildElement(styleEl, WORD_NAMESPACE, "pPr");
    const rPr = firstChildElement(styleEl, WORD_NAMESPACE, "rPr");

    catalog.styles.set(id, {
      id,
      name,
      type: rawType,
      basedOn,
      next,
      paragraphFormatting: rawType === "paragraph" ? parseParagraphProperties(pPr) : {},
      runFormatting: parseRunProperties(rPr),
      headingLevel: rawType === "paragraph" ? (parseOutlineLevel(pPr) ?? headingLevelFromName(id, name)) : undefined,
    });
  }

  return catalog;
}

/** The catalog's paragraph styles as the model's style catalogue
 * (SPEC-019): each style's *own* properties plus the basedOn link, in
 * styles.xml order — the exact data resolveNamedStyle (lib/styles.ts)
 * re-resolves and the writers re-emit. Empty formatting objects are elided
 * so the entries compare structurally against catalogue entries built
 * elsewhere (the writer's changed-style detection). */
export function catalogToNamedStyles(catalog: StyleCatalog): NamedStyle[] {
  const styles: NamedStyle[] = [];
  for (const definition of catalog.styles.values()) {
    const style: NamedStyle = { id: definition.id, name: definition.name ?? definition.id };
    const type = modelStyleType(definition.type);
    if (type !== "paragraph") {
      style.type = type;
    }
    if (definition.basedOn !== undefined) {
      style.basedOn = definition.basedOn;
    }
    if (definition.next !== undefined) {
      style.next = definition.next;
    }
    if (definition.unmodeled) {
      style.unmodeled = true;
      style.rawXml = definition.rawXml;
      styles.push(style);
      continue;
    }
    if (Object.keys(definition.paragraphFormatting).length > 0) {
      style.paragraphFormatting = definition.paragraphFormatting;
    }
    if (Object.keys(definition.runFormatting).length > 0) {
      style.runFormatting = definition.runFormatting;
    }
    if (definition.headingLevel !== undefined) {
      style.headingLevel = definition.headingLevel;
    }
    styles.push(style);
  }
  return styles;
}

/** Resolved effective formatting for a paragraph carrying `w:pStyle` =
 * `styleId`, from `docDefaults` up through the `w:basedOn` chain (root
 * ancestor first) with no per-paragraph direct formatting applied yet —
 * the caller layers the paragraph's own `w:pPr`/`w:rPr` on top. */
export interface ResolvedStyle {
  paragraphFormatting: ParagraphFormatting;
  runFormatting: RunFormatting;
  headingLevel?: number;
}

export function resolveStyle(catalog: StyleCatalog, styleId: string | undefined): ResolvedStyle {
  let paragraphFormatting = catalog.docDefaultParagraphFormatting;
  let runFormatting = catalog.docDefaultRunFormatting;
  let headingLevel: number | undefined;

  const chain: StyleDefinition[] = [];
  const seen = new Set<string>();
  let currentId = styleId;
  while (currentId && !seen.has(currentId)) {
    seen.add(currentId);
    const style = catalog.styles.get(currentId);
    if (!style) {
      break;
    }
    chain.unshift(style);
    currentId = style.basedOn;
  }

  for (const style of chain) {
    paragraphFormatting = mergeParagraphFormatting(paragraphFormatting, style.paragraphFormatting);
    runFormatting = mergeRunFormatting(runFormatting, style.runFormatting);
    if (style.headingLevel !== undefined) {
      headingLevel = style.headingLevel;
    }
  }

  return { paragraphFormatting, runFormatting, headingLevel };
}
