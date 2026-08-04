// Orchestrates the .docx -> DocumentModel pipeline: unzip the OPC package,
// walk the package/document relationships to find the real part names (per
// AC-3, never assume "word/document.xml" is where it is), map w:p/w:r onto
// the document model, resolve w:pStyle against styles.xml, and apply the
// section's page setup.
import type {
  Block,
  DefaultStyles,
  OpaqueBlock,
  OpaqueInline,
  OpaquePart,
  Paragraph,
  Run,
  RunFormatting,
  Section,
  SourceXml,
} from "../documentModel";
import { createDefaultStyles } from "../documentModel";
import { bytesToBase64 } from "../opaque";
import { ZipError, unzip } from "./zip";
import {
  firstChildElement,
  parseRelationships,
  parseXml,
  relsPathFor,
  resolvePartPath,
  serializeElement,
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
const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

// Transient markup Word itself regenerates on every open — spell/grammar
// error ranges and rendered pagination markers. Preserving these as opaque
// nodes would freeze stale derived data into the file, so they are the one
// deliberate exception to SPEC-008's "carry everything through" rule.
const TRANSIENT_PARAGRAPH_CHILDREN = new Set(["proofErr"]);
const MODELLED_RUN_CHILDREN = new Set(["rPr", "t", "tab", "br", "cr", "lastRenderedPageBreak"]);

/** True when a run carries content the model cannot express (w:drawing,
 * w:object, w:pict, field chars, …) — such a run is preserved whole as an
 * opaque inline node rather than flattened to text. */
function runHasUnmodelledContent(rEl: Element): boolean {
  return Array.from(rEl.children).some(
    (child) => child.namespaceURI !== WORD_NAMESPACE || !MODELLED_RUN_CHILDREN.has(child.localName),
  );
}

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
  /** Every package part the reader did not model (SPEC-008), in package
   * order — written back byte-identical on a save to .docx. */
  opaque: OpaquePart[];
  /** Verbatim XML captured from modelled parts (the body-level w:sectPr). */
  sourceXml: SourceXml;
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

function collectRuns(
  container: Element,
  inheritedFormatting: RunFormatting,
  runs: Run[],
  opaqueInlines: OpaqueInline[],
): void {
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "pPr") {
      continue; // Modelled separately by parseParagraph.
    }
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "r") {
      if (runHasUnmodelledContent(child)) {
        // A drawing/object/field run is preserved whole — splitting it into
        // "the text we understood" plus "the rest" could not round-trip.
        opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
        continue;
      }
      const run = parseRun(child, inheritedFormatting);
      if (run) {
        runs.push(run);
      }
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "hyperlink") {
      // Hyperlinks (SPEC-025) wrap w:r children; not modelling the link
      // itself yet, but its text must still reach the document per the
      // round-trip/no-content-loss principle.
      collectRuns(child, inheritedFormatting, runs, opaqueInlines);
    } else if (
      child.namespaceURI === WORD_NAMESPACE &&
      TRANSIENT_PARAGRAPH_CHILDREN.has(child.localName)
    ) {
      // Regenerated by Word on open; deliberately not preserved.
    } else {
      // Anything else (w:sdt, w:fldSimple, bookmarks, other namespaces) is
      // an unmodelled inline element: preserved verbatim, anchored to its
      // position between the runs (SPEC-008).
      opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
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
  const opaqueInlines: OpaqueInline[] = [];
  collectRuns(pEl, resolved.runFormatting, runs, opaqueInlines);

  const paragraph: Paragraph = { type: "paragraph", runs, formatting };
  if (styleId) {
    paragraph.styleId = styleId;
  }
  if (headingLevel !== undefined) {
    paragraph.headingLevel = headingLevel;
  }
  if (opaqueInlines.length > 0) {
    paragraph.opaqueInlines = opaqueInlines;
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
  let corePartPath: string | undefined;
  const rootRelsText = readPart("_rels/.rels");
  if (rootRelsText) {
    const rootRels = parseRelationships(rootRelsText);
    const officeDocRel = Array.from(rootRels.values()).find(
      (rel) => rel.type === OFFICE_DOCUMENT_RELATIONSHIP,
    );
    if (officeDocRel) {
      mainPartPath = resolvePartPath("", officeDocRel.target);
    }
    // Core properties are rebuilt on save (they carry the edited title), so
    // the original core part is the one part besides the main document that
    // is not retained verbatim.
    const coreRel = Array.from(rootRels.values()).find(
      (rel) => rel.type === CORE_PROPERTIES_RELATIONSHIP,
    );
    if (coreRel) {
      corePartPath = resolvePartPath("", coreRel.target);
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
  const opaqueBlocks: OpaqueBlock[] = [];
  let sectPr: Element | undefined;
  for (const child of Array.from(body.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "p") {
      blocks.push(parseParagraph(child, styleCatalog));
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "sectPr") {
      sectPr = child;
    } else {
      // Tables, block-level content controls, bookmarks, anything from
      // another namespace: unmodelled body children are preserved verbatim,
      // anchored to their position among the paragraphs (SPEC-008).
      opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
    }
  }

  const fallbackStyles = createDefaultStyles();
  const defaultStyles = applySectionProperties(sectPr, {
    ...fallbackStyles,
    fontFamily: styleCatalog.docDefaultRunFormatting.fontFamily ?? fallbackStyles.fontFamily,
    fontSizePt: styleCatalog.docDefaultRunFormatting.fontSizePt ?? fallbackStyles.fontSizePt,
  });

  const section: Section = { blocks };
  if (opaqueBlocks.length > 0) {
    section.opaqueBlocks = opaqueBlocks;
  }

  // Everything the pipeline above did not consume-and-rebuild is retained
  // verbatim for the lifetime of the document (SPEC-008 AC): styles.xml and
  // the relationship/content-type parts included, since a same-format save
  // reuses them as-is.
  const opaque: OpaquePart[] = [];
  for (const [name, data] of parts) {
    if (name !== mainPartPath && name !== corePartPath) {
      opaque.push({ name, data: bytesToBase64(data) });
    }
  }

  const sourceXml: SourceXml = {};
  if (sectPr) {
    sourceXml.docxSectPr = serializeElement(sectPr);
  }

  return { sections: [section], defaultStyles, opaque, sourceXml };
}
