// Parses word/styles.xml into a lookup table, and resolves a paragraph's
// `w:pStyle` chain (following `w:basedOn`) into effective paragraph/run
// formatting — this is what lets built-in heading styles (AC-25) render at
// their defined size and weight even when a paragraph carries no direct
// formatting of its own.
import type { ParagraphFormatting, RunFormatting } from "../documentModel";
import { childElements, firstChildElement, wordAttr as attr, WORD_NAMESPACE } from "./xml";
import {
  mergeParagraphFormatting,
  mergeRunFormatting,
  parseOutlineLevel,
  parseParagraphProperties,
  parseRunProperties,
} from "./properties";

interface StyleDefinition {
  id: string;
  name?: string;
  basedOn?: string;
  paragraphFormatting: ParagraphFormatting;
  runFormatting: RunFormatting;
  headingLevel?: number;
}

export interface StyleCatalog {
  styles: Map<string, StyleDefinition>;
  docDefaultParagraphFormatting: ParagraphFormatting;
  docDefaultRunFormatting: RunFormatting;
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

    const pPrDefault = firstChildElement(docDefaults, WORD_NAMESPACE, "pPrDefault");
    const pPr = pPrDefault ? firstChildElement(pPrDefault, WORD_NAMESPACE, "pPr") : undefined;
    catalog.docDefaultParagraphFormatting = parseParagraphProperties(pPr);
  }

  for (const styleEl of childElements(stylesDoc.documentElement, WORD_NAMESPACE, "style")) {
    if (attr(styleEl, "type") !== "paragraph") {
      continue;
    }
    const id = attr(styleEl, "styleId");
    if (!id) {
      continue;
    }
    const nameEl = firstChildElement(styleEl, WORD_NAMESPACE, "name");
    const name = nameEl ? (attr(nameEl, "val") ?? undefined) : undefined;
    const basedOnEl = firstChildElement(styleEl, WORD_NAMESPACE, "basedOn");
    const basedOn = basedOnEl ? (attr(basedOnEl, "val") ?? undefined) : undefined;
    const pPr = firstChildElement(styleEl, WORD_NAMESPACE, "pPr");
    const rPr = firstChildElement(styleEl, WORD_NAMESPACE, "rPr");

    catalog.styles.set(id, {
      id,
      name,
      basedOn,
      paragraphFormatting: parseParagraphProperties(pPr),
      runFormatting: parseRunProperties(rPr),
      headingLevel: parseOutlineLevel(pPr) ?? headingLevelFromName(id, name),
    });
  }

  return catalog;
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
