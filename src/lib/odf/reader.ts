// Orchestrates the .odt -> DocumentModel pipeline: unzip the ODF package,
// detect the format from the package's `mimetype` entry (never the file
// extension), parse content.xml and styles.xml, map text:p/text:h/text:span
// onto the document model with ODF whitespace handling (text:s, text:tab,
// text:line-break), resolve style chains, and apply the master page's layout.
import type { Block, DefaultStyles, Paragraph, Run, RunFormatting, Section } from "../documentModel";
import { createDefaultStyles } from "../documentModel";
import { ZipError, unzip } from "../docx/zip";
import { firstChildElement, parseXml, XmlParseError } from "../docx/xml";
import { OFFICE_NAMESPACE, TEXT_NAMESPACE, textAttr } from "./xml";
import { parseOdfStyleCatalog, resolveParagraphStyle, resolveTextStyle, type OdfStyleCatalog } from "./styles";

export class OdtOpenError extends Error {}
export class FlatOdfError extends Error {}

export const INVALID_ODT_MESSAGE = "This file isn't an OpenDocument text file we can open";
export const FLAT_ODF_MESSAGE = "Flat ODF files aren't supported yet";

/** The package mimetype identifying an OpenDocument text document. */
export const ODT_MIMETYPE = "application/vnd.oasis.opendocument.text";

export interface ParsedOdtDocument {
  sections: Section[];
  defaultStyles: DefaultStyles;
}

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

/** A flat ODF file (.fodt) is the whole document as one XML file: its root
 * element is `office:document`. Recognised only to reject it with a specific
 * message rather than a generic "can't open". */
function isFlatOdf(bytes: Uint8Array): boolean {
  const text = decodeUtf8(bytes);
  if (!text.trimStart().startsWith("<")) {
    return false;
  }
  let doc: Document;
  try {
    doc = parseXml(text);
  } catch {
    return false;
  }
  const root = doc.documentElement;
  return root.namespaceURI === OFFICE_NAMESPACE && root.localName === "document";
}

/** Classifies a file's bytes by content, per AC "format is detected from the
 * mimetype entry rather than the extension": "package" for a ZIP whose
 * `mimetype` entry declares OpenDocument text, "flat" for single-XML ODF
 * (.fodt), "none" for anything else (the .docx path handles those). */
export async function detectOdf(data: ArrayBuffer | Uint8Array): Promise<"package" | "flat" | "none"> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (error instanceof ZipError) {
      return isFlatOdf(bytes) ? "flat" : "none";
    }
    throw error;
  }
  const mimetype = parts.get("mimetype");
  return mimetype && decodeUtf8(mimetype).trim() === ODT_MIMETYPE ? "package" : "none";
}

// --- inline content --------------------------------------------------------

/** A contiguous stretch of same-formatted text while walking a paragraph.
 * Pieces reference the shared resolved formatting object, so adjacent
 * content under the same formatting merges by identity; each piece becomes
 * an independent Run (with its own formatting copy) once the paragraph is
 * done. */
interface TextPiece {
  text: string;
  formatting: RunFormatting;
}

interface InlineAccumulator {
  pieces: TextPiece[];
  /** Last character emitted into the paragraph ("" at paragraph start) —
   * what literal XML whitespace collapses against. */
  lastChar: string;
}

function appendText(acc: InlineAccumulator, text: string, formatting: RunFormatting): void {
  if (text.length === 0) {
    return;
  }
  const last = acc.pieces[acc.pieces.length - 1];
  if (last && last.formatting === formatting) {
    last.text += text;
  } else {
    acc.pieces.push({ text, formatting });
  }
  acc.lastChar = text[text.length - 1];
}

const WHITESPACE_RUN = /[ \t\r\n]+/g;

/** Appends a literal text node with ODF whitespace handling (6.1.2): runs of
 * XML whitespace collapse to a single space, and a space at the start of a
 * paragraph or after other whitespace (including explicit text:s/text:tab
 * content) is dropped. Explicit whitespace elements bypass this entirely —
 * that split is what makes the result match what LibreOffice shows. */
function appendLiteralText(acc: InlineAccumulator, raw: string, formatting: RunFormatting): void {
  let text = raw.replace(WHITESPACE_RUN, " ");
  if (text.startsWith(" ") && (acc.lastChar === "" || " \t\n".includes(acc.lastChar))) {
    text = text.slice(1);
  }
  appendText(acc, text, formatting);
}

function collectInline(
  container: Element,
  formatting: RunFormatting,
  catalog: OdfStyleCatalog,
  acc: InlineAccumulator,
): void {
  for (const node of Array.from(container.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE || node.nodeType === Node.CDATA_SECTION_NODE) {
      appendLiteralText(acc, (node as CharacterData).data, formatting);
      continue;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }
    const el = node as Element;
    if (el.namespaceURI !== TEXT_NAMESPACE) {
      continue;
    }
    switch (el.localName) {
      case "s": {
        // <text:s text:c="3"/> is three explicit (non-collapsing) spaces.
        const count = Number(textAttr(el, "c") ?? "1");
        appendText(acc, " ".repeat(Number.isInteger(count) && count > 0 ? count : 1), formatting);
        break;
      }
      case "tab":
        appendText(acc, "\t", formatting);
        break;
      case "line-break":
        appendText(acc, "\n", formatting);
        break;
      case "span":
        collectInline(el, resolveTextStyle(catalog, textAttr(el, "style-name") ?? undefined, formatting), catalog, acc);
        break;
      case "a":
        // Hyperlinks wrap inline content; the link itself isn't modelled yet
        // (SPEC-025) but its text must reach the document, matching the
        // .docx reader's w:hyperlink handling.
        collectInline(el, formatting, catalog, acc);
        break;
      default:
        // Other inline markup (bookmarks, soft page breaks, ...) carries no
        // body text of its own and is skipped.
        break;
    }
  }
}

function parseTextBlock(el: Element, catalog: OdfStyleCatalog, isHeading: boolean): Paragraph {
  const styleName = textAttr(el, "style-name") ?? undefined;
  const resolved = resolveParagraphStyle(catalog, styleName);

  const acc: InlineAccumulator = { pieces: [], lastChar: "" };
  collectInline(el, resolved.runFormatting, catalog, acc);
  // Each run gets its own formatting object — runs must stay independently
  // mutable for the editing specs even where they shared a resolved style.
  const runs: Run[] = acc.pieces
    .filter((piece) => piece.text.length > 0)
    .map((piece) => ({ text: piece.text, formatting: { ...piece.formatting } }));

  // Fresh object per paragraph — the resolver may hand back the catalog's
  // shared default when the style chain is empty.
  const paragraph: Paragraph = { type: "paragraph", runs, formatting: { ...resolved.paragraphFormatting } };
  if (styleName) {
    paragraph.styleId = styleName;
  }
  if (isHeading) {
    // text:outline-level is 1-based; per ODF a text:h without one is level 1
    // unless its style chain says otherwise (style:default-outline-level).
    const attrValue = Number(textAttr(el, "outline-level") ?? "");
    const level = Number.isInteger(attrValue) && attrValue > 0 ? attrValue : resolved.headingLevel;
    paragraph.headingLevel = level ?? 1;
  }
  return paragraph;
}

// --- package ---------------------------------------------------------------

function parsePart(text: string): Document {
  try {
    return parseXml(text);
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new OdtOpenError(INVALID_ODT_MESSAGE);
    }
    throw error;
  }
}

/** Parses a `.odt` file's bytes into the sections/defaultStyles this
 * project's document model needs. Throws `FlatOdfError` for a flat XML ODF
 * (.fodt) file, or `OdtOpenError` for anything else that isn't a valid,
 * openable OpenDocument text package. */
export async function parseOdt(data: ArrayBuffer | Uint8Array): Promise<ParsedOdtDocument> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(bytes);
  } catch (error) {
    if (error instanceof ZipError) {
      if (isFlatOdf(bytes)) {
        throw new FlatOdfError(FLAT_ODF_MESSAGE);
      }
      throw new OdtOpenError(INVALID_ODT_MESSAGE);
    }
    throw error;
  }

  const contentBytes = parts.get("content.xml");
  if (!contentBytes) {
    throw new OdtOpenError(INVALID_ODT_MESSAGE);
  }
  const contentDoc = parsePart(decodeUtf8(contentBytes));

  const stylesBytes = parts.get("styles.xml");
  const stylesDoc = stylesBytes ? parsePart(decodeUtf8(stylesBytes)) : undefined;

  const catalog = parseOdfStyleCatalog(contentDoc, stylesDoc);

  const body = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "body");
  const textRoot = body ? firstChildElement(body, OFFICE_NAMESPACE, "text") : undefined;
  if (!textRoot) {
    throw new OdtOpenError(INVALID_ODT_MESSAGE);
  }

  const blocks: Block[] = [];
  for (const child of Array.from(textRoot.children)) {
    if (child.namespaceURI !== TEXT_NAMESPACE) {
      continue;
    }
    if (child.localName === "p") {
      blocks.push(parseTextBlock(child, catalog, false));
    } else if (child.localName === "h") {
      blocks.push(parseTextBlock(child, catalog, true));
    }
    // Tables, lists, frames, ... are later specs — skipped on import.
  }

  const fallback = createDefaultStyles();
  const defaultStyles: DefaultStyles = {
    pageSize: catalog.pageSize ?? fallback.pageSize,
    margins: catalog.margins ?? fallback.margins,
    fontFamily: catalog.defaultRunFormatting.fontFamily ?? fallback.fontFamily,
    fontSizePt: catalog.defaultRunFormatting.fontSizePt ?? fallback.fontSizePt,
  };

  return { sections: [{ blocks }], defaultStyles };
}
