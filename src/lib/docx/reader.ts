// Orchestrates the .docx -> DocumentModel pipeline: unzip the OPC package,
// walk the package/document relationships to find the real part names (per
// AC-3, never assume "word/document.xml" is where it is), map w:p/w:r onto
// the document model, resolve w:pStyle against styles.xml, and apply the
// section's page setup.
import type { Block, DefaultStyles, Paragraph, Run, RunFormatting, Section } from "../documentModel";
import { createDefaultStyles } from "../documentModel";
import { ZipError, unzip } from "./zip";
import {
  firstChildElement,
  parseRelationships,
  parseXml,
  relsPathFor,
  resolvePartPath,
  wordAttr,
  WORD_NAMESPACE,
  XmlParseError,
} from "./xml";
import { parseStylesCatalog, resolveStyle, type StyleCatalog } from "./styles";
import {
  mergeParagraphFormatting,
  mergeRunFormatting,
  parseOutlineLevel,
  parseParagraphProperties,
  parseRunProperties,
} from "./properties";
import { twipsToCm } from "./units";

export class DocxOpenError extends Error {}
export class LegacyDocError extends Error {}

export const INVALID_DOCX_MESSAGE = "This file isn't a Word document we can open";
export const LEGACY_DOC_MESSAGE = "Old .doc files aren't supported — save as .docx first";

const OFFICE_DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STYLES_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

// The Compound File Binary Format signature every legacy .doc (Word 97-2003)
// file starts with — used to reject them with a specific, actionable error
// instead of the generic "not a ZIP" message a naive open would produce.
const DOC_MAGIC_BYTES = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function isLegacyDoc(bytes: Uint8Array): boolean {
  if (bytes.length < DOC_MAGIC_BYTES.length) {
    return false;
  }
  return DOC_MAGIC_BYTES.every((byte, index) => bytes[index] === byte);
}

export interface ParsedDocxDocument {
  sections: Section[];
  defaultStyles: DefaultStyles;
}

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

function parseRun(rEl: Element, inheritedFormatting: RunFormatting): Run | undefined {
  const rPr = firstChildElement(rEl, WORD_NAMESPACE, "rPr");
  const formatting = mergeRunFormatting(inheritedFormatting, parseRunProperties(rPr));

  let text = "";
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    if (child.localName === "t") {
      text += child.textContent ?? "";
    } else if (child.localName === "tab") {
      text += "\t";
    } else if (child.localName === "br" || child.localName === "cr") {
      text += "\n";
    }
  }

  if (text.length === 0) {
    return undefined;
  }
  return { text, formatting };
}

function collectRuns(container: Element, inheritedFormatting: RunFormatting, runs: Run[]): void {
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    if (child.localName === "r") {
      const run = parseRun(child, inheritedFormatting);
      if (run) {
        runs.push(run);
      }
    } else if (child.localName === "hyperlink") {
      // Hyperlinks (SPEC-025) wrap w:r children; not modelling the link
      // itself yet, but its text must still reach the document per the
      // round-trip/no-content-loss principle.
      collectRuns(child, inheritedFormatting, runs);
    }
  }
}

function parseParagraph(pEl: Element, styleCatalog: StyleCatalog): Paragraph {
  const pPr = firstChildElement(pEl, WORD_NAMESPACE, "pPr");
  const pStyleEl = pPr ? firstChildElement(pPr, WORD_NAMESPACE, "pStyle") : undefined;
  const styleId = pStyleEl ? (wordAttr(pStyleEl, "val") ?? undefined) : undefined;

  const resolved = resolveStyle(styleCatalog, styleId);
  const formatting = mergeParagraphFormatting(resolved.paragraphFormatting, parseParagraphProperties(pPr));
  const headingLevel = parseOutlineLevel(pPr) ?? resolved.headingLevel;

  const runs: Run[] = [];
  collectRuns(pEl, resolved.runFormatting, runs);

  const paragraph: Paragraph = { type: "paragraph", runs, formatting };
  if (styleId) {
    paragraph.styleId = styleId;
  }
  if (headingLevel !== undefined) {
    paragraph.headingLevel = headingLevel;
  }
  return paragraph;
}

function applySectionProperties(sectPr: Element | undefined, base: DefaultStyles): DefaultStyles {
  if (!sectPr) {
    return base;
  }

  let pageSize = base.pageSize;
  const pgSz = firstChildElement(sectPr, WORD_NAMESPACE, "pgSz");
  if (pgSz) {
    const w = wordAttr(pgSz, "w");
    const h = wordAttr(pgSz, "h");
    if (w && h) {
      let widthTwips = Number(w);
      let heightTwips = Number(h);
      const orient = wordAttr(pgSz, "orient")?.toLowerCase();
      if (orient === "landscape" && widthTwips < heightTwips) {
        [widthTwips, heightTwips] = [heightTwips, widthTwips];
      }
      pageSize = { widthCm: twipsToCm(widthTwips), heightCm: twipsToCm(heightTwips) };
    }
  }

  let margins = base.margins;
  const pgMar = firstChildElement(sectPr, WORD_NAMESPACE, "pgMar");
  if (pgMar) {
    const top = wordAttr(pgMar, "top");
    const right = wordAttr(pgMar, "right");
    const bottom = wordAttr(pgMar, "bottom");
    const left = wordAttr(pgMar, "left");
    if (top && right && bottom && left) {
      margins = {
        topCm: twipsToCm(Number(top)),
        rightCm: twipsToCm(Number(right)),
        bottomCm: twipsToCm(Number(bottom)),
        leftCm: twipsToCm(Number(left)),
      };
    }
  }

  return { ...base, pageSize, margins };
}

/** Parses a `.docx` file's bytes into the sections/defaultStyles this
 * project's document model needs. Throws `LegacyDocError` for a binary
 * Word 97-2003 `.doc` file, or `DocxOpenError` for anything else that
 * isn't a valid, openable `.docx` package. */
export async function parseDocx(data: ArrayBuffer | Uint8Array): Promise<ParsedDocxDocument> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

  if (isLegacyDoc(bytes)) {
    throw new LegacyDocError(LEGACY_DOC_MESSAGE);
  }

  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (error instanceof ZipError) {
      throw new DocxOpenError(INVALID_DOCX_MESSAGE);
    }
    throw error;
  }

  const readPart = (path: string): string | undefined => {
    const raw = parts.get(path);
    return raw ? decodeUtf8(raw) : undefined;
  };

  let mainPartPath = "word/document.xml";
  const rootRelsText = readPart("_rels/.rels");
  if (rootRelsText) {
    const rootRels = parseRelationships(rootRelsText);
    const officeDocRel = Array.from(rootRels.values()).find(
      (rel) => rel.type === OFFICE_DOCUMENT_RELATIONSHIP,
    );
    if (officeDocRel) {
      mainPartPath = resolvePartPath("", officeDocRel.target);
    }
  }

  const documentXmlText = readPart(mainPartPath);
  if (!documentXmlText) {
    throw new DocxOpenError(INVALID_DOCX_MESSAGE);
  }

  let documentXml: Document;
  try {
    documentXml = parseXml(documentXmlText);
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new DocxOpenError(INVALID_DOCX_MESSAGE);
    }
    throw error;
  }

  const body = firstChildElement(documentXml.documentElement, WORD_NAMESPACE, "body");
  if (!body) {
    throw new DocxOpenError(INVALID_DOCX_MESSAGE);
  }

  let stylesDoc: Document | undefined;
  const mainRelsText = readPart(relsPathFor(mainPartPath));
  if (mainRelsText) {
    const mainRels = parseRelationships(mainRelsText);
    const stylesRel = Array.from(mainRels.values()).find((rel) => rel.type === STYLES_RELATIONSHIP);
    if (stylesRel) {
      const stylesText = readPart(resolvePartPath(mainPartPath, stylesRel.target));
      if (stylesText) {
        stylesDoc = parseXml(stylesText);
      }
    }
  }
  const styleCatalog = parseStylesCatalog(stylesDoc);

  const blocks: Block[] = [];
  let sectPr: Element | undefined;
  for (const child of Array.from(body.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    if (child.localName === "p") {
      blocks.push(parseParagraph(child, styleCatalog));
    } else if (child.localName === "sectPr") {
      sectPr = child;
    }
  }

  const fallbackStyles = createDefaultStyles();
  const defaultStyles = applySectionProperties(sectPr, {
    ...fallbackStyles,
    fontFamily: styleCatalog.docDefaultRunFormatting.fontFamily ?? fallbackStyles.fontFamily,
    fontSizePt: styleCatalog.docDefaultRunFormatting.fontSizePt ?? fallbackStyles.fontSizePt,
  });

  return { sections: [{ blocks }], defaultStyles };
}
