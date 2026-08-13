// Orchestrates the .odt -> DocumentModel pipeline: unzip the ODF package,
// detect the format from the package's `mimetype` entry (never the file
// extension), parse content.xml and styles.xml, map text:p/text:h/text:span
// onto the document model with ODF whitespace handling (text:s, text:tab,
// text:line-break), resolve style chains, and apply the master page's layout.
import type {
  Block,
  BookmarkAnchor,
  BorderLineStyle,
  BorderSide,
  CellBorders,
  CellVerticalAlign,
  CommentAnchor,
  CommentMessage,
  CommentThread,
  DefaultStyles,
  DocumentProperties,
  FieldKind,
  FieldSpec,
  ImageWrapMode,
  ListDefinition,
  ListFormatting,
  ListLevelDefinition,
  MediaPart,
  NamedStyle,
  NoteContent,
  NoteKind,
  NoteRef,
  NoteSettings,
  OpaqueBlock,
  OpaqueInline,
  OpaquePart,
  PageNumberFormat,
  Paragraph,
  ParagraphFormatting,
  Run,
  RunFormatting,
  RunTrackedChanges,
  Section,
  SourceXml,
  Table,
  TableAlignment,
  TableCell,
  TableRow,
  TableWidthMode,
  TableOfContents,
  TocEntry,
  TrackedChangeInfo,
} from "../documentModel";
import {
  createDefaultStyles,
  FIELD_PLACEHOLDER_TEXT,
  NOTE_PLACEHOLDER_TEXT,
} from "../documentModel";
import { bytesToBase64 } from "../opaque";
import { ZipError, unzip } from "../docx/zip";
import { cmToTwips } from "../docx/units";
import { firstChildElement, parseXml, serializeElement, XmlParseError } from "../docx/xml";
import { sanitizeHref } from "../links";
import { parseChipText } from "../mergeChips";
import { createMessageId, parseReplyParagraphText } from "../comments";
import { blockText } from "../textEditing";
import {
  clampListLevel,
  defaultMarkerStyle,
  isBulletStyle,
  LIST_INDENT_STEP_CM,
  LIST_MARKER_WIDTH_CM,
} from "../lists";
import { collectGraphicWrapStyles, parseDrawFrame } from "./drawing";
import { parseShapeElement } from "./shape";
import { odfColorToCss, odfLengthToCm, odfLengthToPt, odfPercentage } from "./units";
import { parseOdfCustomProperties, parseOdfMeta } from "./meta";
import {
  DC_NAMESPACE,
  DRAW_NAMESPACE,
  OFFICE_NAMESPACE,
  STYLE_NAMESPACE,
  TABLE_NAMESPACE,
  TEXT_NAMESPACE,
  foAttr,
  officeAttr,
  styleAttr,
  tableAttr,
  textAttr,
  xlinkAttr,
} from "./xml";
import {
  catalogToNamedStyles,
  listStyleIsComposite,
  paragraphAutoStyleSignals,
  parseOdfStyleCatalog,
  resolveListLevel,
  resolveParagraphStyle,
  resolveTextStyle,
  type OdfListLevelDef,
  type OdfListStyle,
  type OdfMasterPageLayout,
  type OdfStyleCatalog,
} from "./styles";

export class OdtOpenError extends Error {}
export class FlatOdfError extends Error {}

export const INVALID_ODT_MESSAGE = "This file isn't an OpenDocument text file we can open";
export const FLAT_ODF_MESSAGE = "Flat ODF files aren't supported yet";

/** The package mimetype identifying an OpenDocument text document. */
export const ODT_MIMETYPE = "application/vnd.oasis.opendocument.text";
/** An .ott template's mimetype entry (SPEC-046 AC-10) — same package shape
 * as .odt, ODF's own "this is a template" flag. `parseOdt` reads it exactly
 * like a plain .odt: opening one always creates a new document (see
 * openDocument.ts's `assembleDocumentModel`, which mints a fresh id for
 * every open), so no separate parse path is needed. */
export const OTT_MIMETYPE = "application/vnd.oasis.opendocument.text-template";

export interface ParsedOdtDocument {
  sections: Section[];
  defaultStyles: DefaultStyles;
  /** The *named* paragraph styles of styles.xml as the model's style
   * catalogue (SPEC-019); empty when the package defines none. Automatic
   * styles are direct formatting and stay baked (see catalogToNamedStyles). */
  styles: NamedStyle[];
  /** Every package part the reader did not model (SPEC-008), in package
   * order — written back byte-identical on a save to .odt. */
  opaque: OpaquePart[];
  /** Verbatim XML captured from content.xml that the model cannot express:
   * font-face declarations and non-paragraph/text automatic styles that
   * preserved opaque nodes may reference. */
  sourceXml: SourceXml;
  /** Media parts modelled Image blocks reference (SPEC-024), one per
   * distinct source media part. */
  media: MediaPart[];
  /** Multilevel list definitions (SPEC-032) reconstructed from composite
   * text:list-styles the body actually references; empty when none. */
  listDefinitions: ListDefinition[];
  /** Footnote/endnote contents (SPEC-031) the body's reference marks link
   * to; empty when the document holds no text:note elements. */
  notes: NoteContent[];
  /** Footnote numbering settings from styles.xml's notes-configuration. */
  footnoteSettings?: NoteSettings;
  /** Endnote numbering settings from styles.xml's notes-configuration. */
  endnoteSettings?: NoteSettings;
  /** Comment threads (SPEC-038) from office:annotation elements anywhere in
   * the body; empty when the document holds none. */
  comments: CommentThread[];
  /** Whether settings.xml declared `RecordChanges` true (SPEC-039) —
   * round-trips onto DocumentModel.trackChangesRecording unchanged. */
  trackChangesRecording: boolean;
  /** Document properties and metadata (SPEC-041) read from meta.xml; `{}`
   * when the package had none. */
  properties: DocumentProperties;
}

// Transient markup LibreOffice regenerates on every open (rendered
// pagination) — the one deliberate exception to "carry everything through";
// preserving it would freeze stale derived data into the file.
const TRANSIENT_INLINE_ELEMENTS = new Set(["soft-page-break"]);

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
  const declared = mimetype && decodeUtf8(mimetype).trim();
  return declared === ODT_MIMETYPE || declared === OTT_MIMETYPE ? "package" : "none";
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
  /** The nearest enclosing `text:span`'s own `text:style-name` (SPEC-042) —
   * the ODF counterpart of the .docx reader's `w:rStyle` capture, same "raw
   * reference, resolved later" convention `parseTextBlock`'s own
   * `paragraph.styleId` already uses for `text:style-name` on the paragraph
   * itself. */
  styleId?: string;
  /** Marks this piece as a page-number/page-count field (SPEC-029) rather
   * than literal text. */
  field?: FieldSpec;
  /** Marks this piece as a footnote/endnote reference mark (SPEC-031). */
  note?: NoteRef;
  /** Pending tracked-change state (SPEC-039). */
  changes?: RunTrackedChanges;
}

// SPEC-031: the note contents collected while the body parses. Set by
// parseOdt around its body walk so a text:note anywhere in the flow — list
// items and table cells included — funnels into one document-level list;
// null outside a body parse (header/footer content), where a note stays an
// opaque inline exactly as before this spec.
let odtNoteSink: NoteContent[] | null = null;
let odtNoteCounter = 0;

// SPEC-038: comment threads collected while the body parses — set by
// parseOdt around its body walk, mirroring odtNoteSink; null outside a body
// parse (header/footer content), where an annotation stays an opaque inline
// exactly as before this spec.
let odtCommentSink: CommentThread[] | null = null;
let odtCommentCounter = 0;

// SPEC-039: the document's text:tracked-changes definitions, parsed once
// before the body walk and consulted by text:change-start/-end/-change
// marks anywhere in it — the changed-region counterpart of odtCommentSink,
// except definitions are collected up front (text:tracked-changes is one
// block near the top of office:text) rather than while walking.
interface OdtChangedRegion extends TrackedChangeInfo {
  kind: "insert" | "delete" | "formatChange";
  /** For a deletion (SPEC-039 AC-16): the deleted content stored verbatim
   * inside text:deletion, since ODF's body flow carries only a point
   * text:change marker for a deletion, not the text itself. Parsed as
   * plain runs (no nested tracked changes inside a deletion's own stored
   * copy — not a shape this reader/writer produce or need to round-trip). */
  deletedRuns?: Run[];
}
let odtChangedRegions: Map<string, OdtChangedRegion> | null = null;
// changeId -> the piece index text:change-start recorded it at, so the
// matching text:change-end can stamp every piece appended in between.
let odtPendingChangeStarts: Map<string, number> | null = null;

interface InlineAccumulator {
  pieces: TextPiece[];
  /** Last character emitted into the paragraph ("" at paragraph start) —
   * what literal XML whitespace collapses against. Opaque nodes are
   * transparent to it, matching ODF whitespace handling (elements do not
   * interrupt collapsing). */
  lastChar: string;
  /** Unmodelled inline elements, anchored by the piece index they precede. */
  opaqueInlines: OpaqueInline[];
  /** Set after an opaque node so the following text starts a new piece —
   * merging across it would break the node's anchor. */
  mergeBarrier: boolean;
  /** Point bookmarks (SPEC-025), anchored at their character offset. */
  bookmarks: BookmarkAnchor[];
  /** Comment anchor points (SPEC-038), anchored at their character offset. */
  commentAnchors: CommentAnchor[];
}

function accumulatedTextLength(acc: InlineAccumulator): number {
  return acc.pieces.reduce((sum, piece) => sum + piece.text.length, 0);
}

/** The run formatting a `text:a`'s children inherit: its xlink:href folded
 * in as linkAnchor (a "#fragment" target) or linkHref (external, sanitized
 * — a dangerous scheme leaves the text unlinked, AC-17). A fresh object, so
 * pieces inside the link never identity-merge with pieces outside it. */
function linkFormatting(el: Element, inherited: RunFormatting): RunFormatting {
  const href = xlinkAttr(el, "href");
  if (href === null || href.length === 0) {
    return inherited;
  }
  if (href.startsWith("#")) {
    return { ...inherited, linkAnchor: href.slice(1) };
  }
  const sanitized = sanitizeHref(href);
  return sanitized !== null ? { ...inherited, linkHref: sanitized } : inherited;
}

function appendText(
  acc: InlineAccumulator,
  text: string,
  formatting: RunFormatting,
  styleId?: string,
): void {
  if (text.length === 0) {
    return;
  }
  const last = acc.pieces[acc.pieces.length - 1];
  if (!acc.mergeBarrier && last && last.formatting === formatting && last.styleId === styleId) {
    last.text += text;
  } else {
    acc.pieces.push(styleId !== undefined ? { text, formatting, styleId } : { text, formatting });
  }
  acc.mergeBarrier = false;
  acc.lastChar = text[text.length - 1];
}

/** `style:num-format` (shared with list numbering, SPEC-018) narrowed to the
 * three formats SPEC-029's page-number fields offer — case-folded, since
 * this app never emits the uppercase variants and reads either back the
 * same way (see documentModel.ts's PageNumberFormat doc comment). */
function fieldNumberFormat(el: Element): PageNumberFormat {
  const format = styleAttr(el, "num-format");
  if (format === "i" || format === "I") {
    return "lowerRoman";
  }
  if (format === "a" || format === "A") {
    return "lowerAlpha";
  }
  return "decimal";
}

/** A `text:page-number`/`text:page-count` field (SPEC-029 AC-25, AC-26) as
 * its own text piece — a field is always a merge barrier (like an opaque
 * node) so it never coalesces into a neighbouring plain-text run. */
function appendField(
  acc: InlineAccumulator,
  kind: FieldKind,
  format: PageNumberFormat,
  formatting: RunFormatting,
): void {
  acc.pieces.push({ text: FIELD_PLACEHOLDER_TEXT, formatting, field: { kind, format } });
  acc.mergeBarrier = true;
  acc.lastChar = FIELD_PLACEHOLDER_TEXT;
}

const WHITESPACE_RUN = /[ \t\r\n]+/g;

/** Appends a literal text node with ODF whitespace handling (6.1.2): runs of
 * XML whitespace collapse to a single space, and a space at the start of a
 * paragraph or after other whitespace (including explicit text:s/text:tab
 * content) is dropped. Explicit whitespace elements bypass this entirely —
 * that split is what makes the result match what LibreOffice shows. */
function appendLiteralText(
  acc: InlineAccumulator,
  raw: string,
  formatting: RunFormatting,
  styleId?: string,
): void {
  let text = raw.replace(WHITESPACE_RUN, " ");
  if (text.startsWith(" ") && (acc.lastChar === "" || " \t\n".includes(acc.lastChar))) {
    text = text.slice(1);
  }
  appendText(acc, text, formatting, styleId);
}

const UNKNOWN_COMMENT_AUTHOR = "Unknown";

/** The stable model thread id for a source annotation's `office:name`
 * (SPEC-038) — matched between its `office:annotation` start and
 * `office:annotation-end`. Model ids from this reader (and the .docx
 * reader's docxCommentId) are opaque strings, so no shared scheme is
 * needed across formats — only within one parsed document. */
function odtCommentId(name: string): string {
  return `oc-${name}`;
}

/** Records an `office:annotation`/`office:annotation-end` as a comment
 * anchor point at the current text offset (SPEC-038) — the ranged
 * counterpart of the bookmark case above, except an annotation START also
 * carries the whole comment's content (author, date, messages) inline,
 * ODF's own shape (see documentModel.ts's CommentThread doc comment: ODF
 * has no separate "comments part" the way `.docx` does). Outside a body
 * parse (odtCommentSink === null — header/footer content) the element
 * stays an opaque inline, the pre-SPEC-038 behaviour, mirroring the `note`
 * case's same guard. */
function collectAnnotation(el: Element, acc: InlineAccumulator, catalog: OdfStyleCatalog): void {
  const isEnd = el.localName === "annotation-end";
  if (odtCommentSink === null) {
    acc.opaqueInlines.push({ xml: serializeElement(el), runIndex: acc.pieces.length });
    acc.mergeBarrier = true;
    return;
  }
  const name = officeAttr(el, "name");
  if (isEnd) {
    if (name) {
      acc.commentAnchors.push({
        threadId: odtCommentId(name),
        boundary: "end",
        offset: accumulatedTextLength(acc),
      });
    }
    return;
  }
  odtCommentCounter += 1;
  const threadId = odtCommentId(name ?? `annotation-${odtCommentCounter}`);
  const creatorEl = firstChildElement(el, DC_NAMESPACE, "creator");
  const dateEl = firstChildElement(el, DC_NAMESPACE, "date");
  const author = (creatorEl?.textContent ?? "").trim() || UNKNOWN_COMMENT_AUTHOR;
  const createdAt = (dateEl?.textContent ?? "").trim() || new Date(0).toISOString();
  const texts = Array.from(el.children)
    .filter((child) => child.namespaceURI === TEXT_NAMESPACE && child.localName === "p")
    .map((pEl) => blockText(parseTextBlock(pEl, catalog, false)));
  const messages: CommentMessage[] = [
    { messageId: createMessageId(), author, text: texts[0] ?? "", createdAt },
  ];
  for (const text of texts.slice(1)) {
    const reply = parseReplyParagraphText(text);
    if (reply) {
      messages.push({ messageId: createMessageId(), ...reply });
      continue;
    }
    const last = messages[messages.length - 1];
    last.text = last.text.length > 0 ? `${last.text}\n${text}` : text;
  }
  odtCommentSink.push({ threadId, resolved: false, messages });
  acc.commentAnchors.push({ threadId, boundary: "start", offset: accumulatedTextLength(acc) });
}

const UNKNOWN_TRACK_CHANGE_AUTHOR = "Unknown";

/** The model change id for a source `text:changed-region`'s `text:id`
 * (SPEC-039) — the identity, deliberately: unlike docx (which renumbers
 * every tracked-change id from scratch on every save, so the reader can
 * safely prefix), the .odt writer reuses the model's own changeId as
 * text:id verbatim (odtChangeXmlId is a sanitize-only pass). Adding a
 * prefix here would grow on every open/save cycle instead of reaching a
 * fixed point. */
function odtChangeId(sourceId: string): string {
  return sourceId;
}

/** A changed-region's own author/date, from its `text:insertion`/
 * `text:deletion`/`text:format-change` child's nested `office:change-info`. */
function odtChangeInfo(el: Element, changeId: string): TrackedChangeInfo {
  const info = firstChildElement(el, OFFICE_NAMESPACE, "change-info");
  const creator = info ? firstChildElement(info, DC_NAMESPACE, "creator") : undefined;
  const date = info ? firstChildElement(info, DC_NAMESPACE, "date") : undefined;
  return {
    changeId,
    author: (creator?.textContent ?? "").trim() || UNKNOWN_TRACK_CHANGE_AUTHOR,
    date: (date?.textContent ?? "").trim() || new Date(0).toISOString(),
  };
}

/** A `text:deletion`'s verbatim stored content (SPEC-039 AC-16) as plain
 * runs: inline content (a mid-paragraph deletion) is collected directly;
 * a whole deleted paragraph's `text:p`/`text:h` children each contribute
 * their runs, joined by "\n" so nothing is lost even though the model
 * keeps a deletion's text inline in one run flow rather than as separate
 * paragraphs. */
function parseOdtDeletionContent(deletionEl: Element, catalog: OdfStyleCatalog): Run[] {
  const paragraphChildren = Array.from(deletionEl.children).filter(
    (child) => child.namespaceURI === TEXT_NAMESPACE && (child.localName === "p" || child.localName === "h"),
  );
  if (paragraphChildren.length > 0) {
    const runs: Run[] = [];
    paragraphChildren.forEach((pEl, index) => {
      if (index > 0) {
        runs.push({ text: "\n", formatting: {} });
      }
      runs.push(...parseTextBlock(pEl, catalog, false).runs);
    });
    return runs;
  }
  const acc: InlineAccumulator = {
    pieces: [],
    lastChar: "",
    opaqueInlines: [],
    mergeBarrier: false,
    bookmarks: [],
    commentAnchors: [],
  };
  collectInline(deletionEl, {}, catalog, acc);
  return acc.pieces
    .filter((piece) => piece.text.length > 0)
    .map((piece) => ({ text: piece.text, formatting: { ...piece.formatting } }));
}

/** Parses `office:text`'s `text:tracked-changes` block (SPEC-039), if
 * present — the changed-region definitions text:change-start/-end/-change
 * marks anywhere in the body reference by id. Absent (no tracked changes
 * in the source) is the common case and costs nothing. */
function parseTrackedChangesRegions(textRoot: Element, catalog: OdfStyleCatalog): Map<string, OdtChangedRegion> {
  const regions = new Map<string, OdtChangedRegion>();
  const trackedChangesEl = firstChildElement(textRoot, TEXT_NAMESPACE, "tracked-changes");
  if (!trackedChangesEl) {
    return regions;
  }
  for (const regionEl of Array.from(trackedChangesEl.children)) {
    if (regionEl.namespaceURI !== TEXT_NAMESPACE || regionEl.localName !== "changed-region") {
      continue;
    }
    const sourceId = textAttr(regionEl, "id");
    if (!sourceId) {
      continue;
    }
    const changeId = odtChangeId(sourceId);
    const insertionEl = firstChildElement(regionEl, TEXT_NAMESPACE, "insertion");
    const deletionEl = firstChildElement(regionEl, TEXT_NAMESPACE, "deletion");
    const formatChangeEl = firstChildElement(regionEl, TEXT_NAMESPACE, "format-change");
    if (insertionEl) {
      regions.set(sourceId, { ...odtChangeInfo(insertionEl, changeId), kind: "insert" });
    } else if (deletionEl) {
      regions.set(sourceId, {
        ...odtChangeInfo(deletionEl, changeId),
        kind: "delete",
        deletedRuns: parseOdtDeletionContent(deletionEl, catalog),
      });
    } else if (formatChangeEl) {
      regions.set(sourceId, { ...odtChangeInfo(formatChangeEl, changeId), kind: "formatChange" });
    }
  }
  return regions;
}

/** Stamps every piece in `[start, end)` with `region`'s tracked-change
 * marker (SPEC-039) — the retroactive-marking counterpart of the .docx
 * reader's collectTrackedRun, since ODF brackets a change with a start/end
 * pair rather than wrapping it in one element. Formatting changes carry no
 * "previous formatting" here (ODF's text:format-change has no standard
 * slot for it, unlike w:rPrChange) — a rejected formatting change from an
 * .odt source falls back to clearing the marker only (documented gap). */
function stampChangedPieces(pieces: TextPiece[], start: number, end: number, region: OdtChangedRegion): void {
  for (let i = start; i < end && i < pieces.length; i++) {
    const piece = pieces[i];
    if (region.kind === "insert") {
      piece.changes = { ...piece.changes, insert: region };
    } else if (region.kind === "delete") {
      piece.changes = { ...piece.changes, delete: region };
    } else {
      piece.changes = {
        ...piece.changes,
        formatChange: { ...region, previous: {} },
      };
    }
  }
}

function collectInline(
  container: Element,
  formatting: RunFormatting,
  catalog: OdfStyleCatalog,
  acc: InlineAccumulator,
  styleId?: string,
): void {
  for (const node of Array.from(container.childNodes)) {
    if (node.nodeType === Node.TEXT_NODE || node.nodeType === Node.CDATA_SECTION_NODE) {
      appendLiteralText(acc, (node as CharacterData).data, formatting, styleId);
      continue;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }
    const el = node as Element;
    if (
      el.namespaceURI === OFFICE_NAMESPACE &&
      (el.localName === "annotation" || el.localName === "annotation-end")
    ) {
      collectAnnotation(el, acc, catalog);
      continue;
    }
    if (el.namespaceURI !== TEXT_NAMESPACE) {
      // Frames, change marks, … — unmodelled inline elements from other
      // namespaces are preserved verbatim, anchored to their position
      // between the runs (SPEC-008).
      acc.opaqueInlines.push({ xml: serializeElement(el), runIndex: acc.pieces.length });
      acc.mergeBarrier = true;
      continue;
    }
    switch (el.localName) {
      case "s": {
        // <text:s text:c="3"/> is three explicit (non-collapsing) spaces.
        const count = Number(textAttr(el, "c") ?? "1");
        appendText(acc, " ".repeat(Number.isInteger(count) && count > 0 ? count : 1), formatting, styleId);
        break;
      }
      case "tab":
        appendText(acc, "\t", formatting, styleId);
        break;
      case "line-break":
        appendText(acc, "\n", formatting, styleId);
        break;
      case "span": {
        // SPEC-042: capturing a run-style reference the way the .docx
        // reader captures w:rStyle would need distinguishing a genuine
        // named-character-style reference from ODF's own per-run automatic
        // "direct formatting" style — both take the identical shape
        // (text:style-name pointing at a family="text" style; LibreOffice
        // itself parents an automatic style at the named one when direct
        // formatting sits on top of a real character style, so even
        // "is this name in the *named* map" is not a reliable signal). This
        // app's own writer still round-trips an *applied* character style's
        // baked formatting correctly (resolveCharacterStyle bakes it at
        // apply time); only the `styleId` reference itself — needed for a
        // later Modify/Rename to keep affecting this text after a save and
        // reopen — does not survive an .odt round trip. Threading `styleId`
        // through stays inert (always undefined) until that is resolved.
        collectInline(
          el,
          resolveTextStyle(catalog, textAttr(el, "style-name") ?? undefined, formatting),
          catalog,
          acc,
          styleId,
        );
        break;
      }
      case "a":
        // Hyperlinks (SPEC-025): the wrapper's target becomes link
        // formatting on its inline content, matching the .docx reader's
        // w:hyperlink handling; the element itself is not preserved.
        collectInline(el, linkFormatting(el, formatting), catalog, acc, styleId);
        break;
      case "bookmark":
      case "bookmark-start": {
        // Point bookmarks (SPEC-025); a ranged text:bookmark-start anchors
        // at its start position and its matching -end is dropped — the
        // writer re-emits a point text:bookmark.
        const name = textAttr(el, "name");
        if (name) {
          acc.bookmarks.push({ name, offset: accumulatedTextLength(acc) });
        }
        break;
      }
      case "bookmark-end":
        break;
      case "change-start": {
        // A tracked change's opening bracket (SPEC-039): records where in
        // `acc.pieces` it started so the matching change-end can stamp
        // everything appended in between. mergeBarrier keeps the marked
        // span from merging into whatever text preceded it.
        const changeId = textAttr(el, "change-id");
        if (changeId && odtPendingChangeStarts) {
          odtPendingChangeStarts.set(changeId, acc.pieces.length);
        }
        acc.mergeBarrier = true;
        break;
      }
      case "change-end": {
        const changeId = textAttr(el, "change-id");
        const start = changeId ? odtPendingChangeStarts?.get(changeId) : undefined;
        const region = changeId ? odtChangedRegions?.get(changeId) : undefined;
        if (start !== undefined && region) {
          stampChangedPieces(acc.pieces, start, acc.pieces.length, region);
        }
        if (changeId) {
          odtPendingChangeStarts?.delete(changeId);
        }
        acc.mergeBarrier = true;
        break;
      }
      case "change": {
        // A point deletion marker (SPEC-039 AC-16): the deleted text lives
        // only in the changed-region's own text:deletion, not the body
        // flow — inserted here inline (marked deleted), never lost. This
        // app's own writer always uses change-start/-end pairs for text-
        // level deletions instead (keeping the text inline directly, no
        // separate storage needed), so this path is for a foreign file's
        // structural deletion only.
        const changeId = textAttr(el, "change-id");
        const region = changeId ? odtChangedRegions?.get(changeId) : undefined;
        if (region?.kind === "delete" && region.deletedRuns) {
          for (const run of region.deletedRuns) {
            acc.pieces.push({ ...run, changes: { ...run.changes, delete: region } });
          }
          acc.mergeBarrier = true;
          acc.lastChar = region.deletedRuns.at(-1)?.text.at(-1) ?? acc.lastChar;
        }
        break;
      }
      case "page-number":
        // SPEC-029 AC-25, AC-26: the current page's live number.
        appendField(acc, "page", fieldNumberFormat(el), formatting);
        break;
      case "page-count":
        // SPEC-029 AC-25, AC-26: the document's total page count.
        appendField(acc, "numpages", fieldNumberFormat(el), formatting);
        break;
      case "placeholder": {
        // A mail-merge chip (SPEC-046): its content is the same canonical
        // text the .docx MERGEFIELD instruction quotes. A placeholder
        // whose content isn't valid chip syntax stays what it was — an
        // opaque inline (SPEC-008) — same as an unrecognised w:fldSimple.
        const chip = textAttr(el, "placeholder-type") === "text" ? parseChipText(el.textContent ?? "") : null;
        if (chip) {
          acc.pieces.push({ text: FIELD_PLACEHOLDER_TEXT, formatting, field: chip });
          acc.mergeBarrier = true;
          acc.lastChar = FIELD_PLACEHOLDER_TEXT;
        } else {
          acc.opaqueInlines.push({ xml: serializeElement(el), runIndex: acc.pieces.length });
          acc.mergeBarrier = true;
        }
        break;
      }
      case "note": {
        // Footnotes/endnotes (SPEC-031 AC-23/24): the citation becomes a
        // superscript marker run (numbers are derived, so the citation text
        // itself is dropped) and the note-body's paragraphs become the
        // note's own content. Outside a body parse the element stays an
        // opaque inline, the pre-SPEC-031 behaviour.
        if (odtNoteSink === null) {
          acc.opaqueInlines.push({ xml: serializeElement(el), runIndex: acc.pieces.length });
          acc.mergeBarrier = true;
          break;
        }
        const kind: NoteKind = textAttr(el, "note-class") === "endnote" ? "endnote" : "footnote";
        odtNoteCounter += 1;
        const noteId = textAttr(el, "id") ?? `odt-note-${odtNoteCounter}`;
        const bodyEl = firstChildElement(el, TEXT_NAMESPACE, "note-body");
        const noteBlocks: Block[] = [];
        if (bodyEl) {
          for (const child of Array.from(bodyEl.children)) {
            if (
              child.namespaceURI === TEXT_NAMESPACE &&
              (child.localName === "p" || child.localName === "h")
            ) {
              noteBlocks.push(parseTextBlock(child, catalog, child.localName === "h"));
            }
          }
        }
        odtNoteSink.push({
          noteId,
          kind,
          blocks: noteBlocks.length > 0 ? noteBlocks : [{ type: "paragraph", runs: [] }],
        });
        acc.pieces.push({
          text: NOTE_PLACEHOLDER_TEXT,
          formatting: { ...formatting, vertAlign: "superscript" },
          note: { noteId, kind },
        });
        acc.mergeBarrier = true;
        acc.lastChar = NOTE_PLACEHOLDER_TEXT;
        break;
      }
      default:
        if (TRANSIENT_INLINE_ELEMENTS.has(el.localName)) {
          // Regenerated by the producer on open; deliberately not preserved.
          break;
        }
        // Fields, notes, … — unmodelled text-namespace elements are
        // preserved verbatim at their anchored position (SPEC-008).
        acc.opaqueInlines.push({ xml: serializeElement(el), runIndex: acc.pieces.length });
        acc.mergeBarrier = true;
        break;
    }
  }
}

/** ODF's fo:margin-left/right are physical, so the property parser fills
 * the logical slots as if the paragraph were LTR. Once the style chain is
 * resolved, an effective RTL writing mode means the physical left edge is
 * the logical END — so the two slots swap; an explicit "ltr" (recorded only
 * to win the chain merge) is dropped, leaving "rtl" the model's one stored
 * direction (SPEC-014). */
function normalizeResolvedDirection(formatting: ParagraphFormatting): void {
  if (formatting.direction === "rtl") {
    const start = formatting.indentStartCm;
    const end = formatting.indentEndCm;
    if (end !== undefined) {
      formatting.indentStartCm = end;
    } else {
      delete formatting.indentStartCm;
    }
    if (start !== undefined) {
      formatting.indentEndCm = start;
    } else {
      delete formatting.indentEndCm;
    }
  } else if (formatting.direction === "ltr") {
    delete formatting.direction;
  }
}

function parseTextBlock(el: Element, catalog: OdfStyleCatalog, isHeading: boolean): Paragraph {
  const styleName = textAttr(el, "style-name") ?? undefined;
  const resolved = resolveParagraphStyle(catalog, styleName);

  const acc: InlineAccumulator = {
    pieces: [],
    lastChar: "",
    opaqueInlines: [],
    mergeBarrier: false,
    bookmarks: [],
    commentAnchors: [],
  };
  collectInline(el, resolved.runFormatting, catalog, acc);
  // Each run gets its own formatting object — runs must stay independently
  // mutable for the editing specs even where they shared a resolved style.
  const runs: Run[] = acc.pieces
    .filter((piece) => piece.text.length > 0)
    .map((piece) => ({
      text: piece.text,
      formatting: { ...piece.formatting },
      ...(piece.styleId ? { styleId: piece.styleId } : {}),
      ...(piece.field ? { field: piece.field } : {}),
      ...(piece.note ? { note: piece.note } : {}),
      ...(piece.changes ? { changes: piece.changes } : {}),
    }));

  // Fresh object per paragraph — the resolver may hand back the catalog's
  // shared default when the style chain is empty.
  const formatting: ParagraphFormatting = { ...resolved.paragraphFormatting };
  normalizeResolvedDirection(formatting);
  const paragraph: Paragraph = { type: "paragraph", runs, formatting };
  if (styleName) {
    paragraph.styleId = styleName;
  }
  if (acc.opaqueInlines.length > 0) {
    paragraph.opaqueInlines = acc.opaqueInlines;
  }
  if (acc.bookmarks.length > 0) {
    paragraph.bookmarks = acc.bookmarks;
  }
  if (acc.commentAnchors.length > 0) {
    paragraph.commentAnchors = acc.commentAnchors;
  }
  // SPEC-039 AC-16: the paragraph's own pending markInsert/markDelete —
  // see writer.ts's paragraphMarkChangeXml — trails the paragraph's own
  // content as a point text:change (unambiguous against a row/cell marker,
  // which always LEADS a paragraph and is stripped before this point — see
  // stripLeadingChangeMarkers). Two can coexist (an author's split pending
  // right before another's merge), so every trailing text:change is read,
  // not just the very last child.
  const trailingChildren = Array.from(el.children);
  for (let i = trailingChildren.length - 1; i >= 0; i--) {
    const candidate = trailingChildren[i];
    if (candidate.namespaceURI !== TEXT_NAMESPACE || candidate.localName !== "change") {
      break;
    }
    const changeId = textAttr(candidate, "change-id");
    const region = changeId ? odtChangedRegions?.get(changeId) : undefined;
    if (region?.kind === "insert") {
      paragraph.changes = { ...paragraph.changes, markInsert: region };
    } else if (region?.kind === "delete") {
      paragraph.changes = { ...paragraph.changes, markDelete: region };
    }
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

/** Context a body-level text-block parse needs to resolve a `draw:frame`
 * into actual image bytes (SPEC-024). */
export interface OdfDrawingContext {
  parts: Map<string, Uint8Array>;
  graphicWrapStyles: Map<string, ImageWrapMode>;
  contentDoc: Document;
  stylesDoc: Document | undefined;
}

/** collectInline run over a detached clone of `nodes` — reused so a
 * mid-paragraph image split (parseBodyTextBlockOrSplit) gets exactly the
 * same whitespace-collapsing/span-resolution behaviour as an ordinary
 * paragraph, without re-implementing collectInline's state machine for a
 * node subset. Cloning is safe (the source tree is never mutated) and cheap
 * at this scale (paragraph-sized fragments only). */
function collectInlineSegment(
  nodes: Node[],
  formatting: RunFormatting,
  catalog: OdfStyleCatalog,
  ownerDocument: Document,
): InlineAccumulator {
  const acc: InlineAccumulator = {
    pieces: [],
    lastChar: "",
    opaqueInlines: [],
    mergeBarrier: false,
    bookmarks: [],
    commentAnchors: [],
  };
  const temp = ownerDocument.createElementNS(TEXT_NAMESPACE, "text:tmp");
  for (const node of nodes) {
    temp.appendChild(node.cloneNode(true));
  }
  collectInline(temp, formatting, catalog, acc);
  return acc;
}

interface BodyTextBlockResult {
  blocks: Block[];
  media: MediaPart[];
}

// --- table of contents (SPEC-033) -------------------------------------------

/** One `text:table-of-content`'s index-body paragraph as a TOC entry
 * (SPEC-033 AC-19): the `text:a`'s href (an internal "#anchor" target) and
 * text, split at its `text:tab` into the heading text and the cached page
 * number. A paragraph with no `text:a` at all (the empty-TOC placeholder)
 * yields null. */
function tocEntryFromParagraph(pEl: Element): TocEntry | null {
  const link = firstChildElement(pEl, TEXT_NAMESPACE, "a");
  if (!link) {
    return null;
  }
  const href = xlinkAttr(link, "href") ?? "";
  const anchor = href.startsWith("#") ? href.slice(1) : href;
  let text = "";
  let pageNumberText = "";
  let sawTab = false;
  for (const node of Array.from(link.childNodes)) {
    if (node.nodeType === node.TEXT_NODE) {
      const value = node.textContent ?? "";
      if (sawTab) {
        pageNumberText += value;
      } else {
        text += value;
      }
    } else if (node.nodeType === node.ELEMENT_NODE) {
      const el = node as Element;
      if (el.namespaceURI === TEXT_NAMESPACE && el.localName === "tab") {
        sawTab = true;
      } else if (el.namespaceURI === TEXT_NAMESPACE && el.localName === "span") {
        if (sawTab) {
          pageNumberText += el.textContent ?? "";
        } else {
          text += el.textContent ?? "";
        }
      }
    }
  }
  const styleVal = textAttr(pEl, "style-name") ?? "";
  const levelMatch = /^Contents_20_([1-6])$/.exec(styleVal);
  const pageNumber = /^\d+$/.test(pageNumberText) ? Number(pageNumberText) : null;
  return { text, level: levelMatch ? Number(levelMatch[1]) : 1, anchor, pageNumber };
}

/** Parses a `text:table-of-content` element (SPEC-033 AC-19, AC-20): every
 * index-body paragraph becomes one entry (see tocEntryFromParagraph); the
 * included levels are the entries' own distinct levels, falling back to the
 * source config's entry templates (a foreign producer's index whose body
 * happens to be empty) and finally to the default top-three levels. */
function parseTableOfContent(el: Element): TableOfContents {
  const body = firstChildElement(el, TEXT_NAMESPACE, "index-body");
  const entries: TocEntry[] = [];
  if (body) {
    for (const child of Array.from(body.children)) {
      if (child.namespaceURI !== TEXT_NAMESPACE || child.localName !== "p") {
        continue;
      }
      const entry = tocEntryFromParagraph(child);
      if (entry) {
        entries.push(entry);
      }
    }
  }
  let levels = Array.from(new Set(entries.map((entry) => entry.level))).sort((a, b) => a - b);
  if (levels.length === 0) {
    const source = firstChildElement(el, TEXT_NAMESPACE, "table-of-content-source");
    const templateLevels = source
      ? Array.from(source.children)
          .filter(
            (child) =>
              child.namespaceURI === TEXT_NAMESPACE &&
              child.localName === "table-of-content-entry-template",
          )
          .map((child) => Number(textAttr(child, "outline-level")))
          .filter((level) => Number.isInteger(level) && level >= 1 && level <= 6)
      : [];
    levels = templateLevels.length > 0 ? templateLevels : [1, 2, 3];
  }
  const hasPageNumbers = entries.some((entry) => entry.pageNumber !== null);
  return {
    type: "toc",
    options: { levels, showPageNumbers: hasPageNumbers, hyperlinksOnly: !hasPageNumbers },
    entries,
  };
}

/** Parses one body-level `text:p`/`text:h` into one or more blocks
 * (SPEC-024): a paragraph containing a modelled `draw:frame` picture splits
 * around each image, since the document model addresses an image as its own
 * top-level block (see documentModel.ts's Image doc comment) — mirroring
 * how a `table:table` already breaks the block sequence, just with the split
 * happening here instead of being structural in the source XML. Every
 * resulting paragraph piece keeps the source block's own style/heading
 * level. A frame nested inside a `text:span`/`text:a` is not detected (real
 * producers place image frames as direct paragraph children, never inside a
 * run) — such a paragraph parses exactly as parseTextBlock always did, so
 * this function is a pure superset, never a behaviour change for the common
 * case. Used only for section-level body content, not list items or table
 * cells — images are not modelled inside either (out of scope, SPEC-024). */
function parseBodyTextBlockOrSplit(
  el: Element,
  catalog: OdfStyleCatalog,
  isHeading: boolean,
  ctx: OdfDrawingContext,
): BodyTextBlockResult {
  const frameChildren = Array.from(el.children).filter(
    (child) => child.namespaceURI === DRAW_NAMESPACE && child.localName === "frame",
  );
  const hasImage = frameChildren.some((frame) => parseDrawFrame(frame, ctx.parts, ctx.graphicWrapStyles) !== null);
  const shapeTextCtx = { parseParagraph: (pEl: Element) => parseTextBlock(pEl, catalog, pEl.localName === "h") };
  const hasShape = Array.from(el.children).some(
    (child) =>
      child.namespaceURI === DRAW_NAMESPACE &&
      (child.localName === "custom-shape" || child.localName === "line") &&
      parseShapeElement(child, ctx.contentDoc, ctx.stylesDoc, shapeTextCtx) !== null,
  );
  if (!hasImage && !hasShape) {
    return { blocks: [parseTextBlock(el, catalog, isHeading)], media: [] };
  }

  const styleName = textAttr(el, "style-name") ?? undefined;
  const resolved = resolveParagraphStyle(catalog, styleName);
  let headingLevel: number | undefined;
  if (isHeading) {
    const attrValue = Number(textAttr(el, "outline-level") ?? "");
    headingLevel = (Number.isInteger(attrValue) && attrValue > 0 ? attrValue : resolved.headingLevel) ?? 1;
  }

  const blocks: Block[] = [];
  const media: MediaPart[] = [];
  let segment: Node[] = [];
  const flush = () => {
    if (segment.length === 0) {
      return;
    }
    const acc = collectInlineSegment(segment, resolved.runFormatting, catalog, el.ownerDocument);
    segment = [];
    const runs: Run[] = acc.pieces
      .filter((piece) => piece.text.length > 0)
      .map((piece) => ({
        text: piece.text,
        formatting: { ...piece.formatting },
        ...(piece.styleId ? { styleId: piece.styleId } : {}),
        ...(piece.field ? { field: piece.field } : {}),
      }));
    if (
      runs.length === 0 &&
      acc.opaqueInlines.length === 0 &&
      acc.bookmarks.length === 0 &&
      acc.commentAnchors.length === 0
    ) {
      return;
    }
    const formatting: ParagraphFormatting = { ...resolved.paragraphFormatting };
    normalizeResolvedDirection(formatting);
    const paragraph: Paragraph = { type: "paragraph", runs, formatting };
    if (styleName) {
      paragraph.styleId = styleName;
    }
    if (acc.opaqueInlines.length > 0) {
      paragraph.opaqueInlines = acc.opaqueInlines;
    }
    if (acc.bookmarks.length > 0) {
      paragraph.bookmarks = acc.bookmarks;
    }
    if (acc.commentAnchors.length > 0) {
      paragraph.commentAnchors = acc.commentAnchors;
    }
    if (headingLevel !== undefined) {
      paragraph.headingLevel = headingLevel;
    }
    blocks.push(paragraph);
  };
  for (const node of Array.from(el.childNodes)) {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const frameEl = node as Element;
      if (frameEl.namespaceURI === DRAW_NAMESPACE && frameEl.localName === "frame") {
        const parsed = parseDrawFrame(frameEl, ctx.parts, ctx.graphicWrapStyles);
        if (parsed) {
          flush();
          blocks.push(parsed.image);
          media.push(parsed.media);
          continue;
        }
      }
      if (
        frameEl.namespaceURI === DRAW_NAMESPACE &&
        (frameEl.localName === "custom-shape" || frameEl.localName === "line")
      ) {
        const parsedShape = parseShapeElement(frameEl, ctx.contentDoc, ctx.stylesDoc, shapeTextCtx);
        if (parsedShape) {
          flush();
          blocks.push(parsedShape.shape);
          continue;
        }
      }
    }
    segment.push(node);
  }
  flush();
  return { blocks, media };
}

// --- lists (SPEC-018) ------------------------------------------------------

/** An item's model list formatting from its resolved level definition, with
 * default-elision so an editor-default list re-imports without noise: the
 * marker is stored only off the AC-5 cycle, startAt only when not 1, and
 * the indent only when it differs from level × 1.27 cm (compared in twips,
 * the shared cross-format grid, so float noise cannot leak in). */
function buildListFormatting(def: OdfListLevelDef | undefined, level: number): ListFormatting {
  const type = def?.type ?? "bullet";
  const list: ListFormatting = { type, level };
  if (!def) {
    return list;
  }
  if (def.marker !== defaultMarkerStyle(type, level)) {
    list.marker = def.marker;
  }
  if (type === "number" && def.startAt !== 1) {
    list.startAt = def.startAt;
  }
  if (def.indentCm !== undefined && cmToTwips(def.indentCm) !== level * 720) {
    list.indentCm = def.indentCm;
  }
  return list;
}

// --- multilevel definitions and list instances (SPEC-032) -------------------

const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";

/** The instance a top-level `text:list` was assigned by planListInstances:
 * whether it parses as a DEFINED list (instance-numbered, SPEC-032) or the
 * SPEC-018 adjacency form, and — when defined — which listId its items
 * share. */
interface PlannedListInstance {
  defined: boolean;
  listId: string;
}

/** Installs the SPEC-032 read state on the catalog — the document-wide
 * list-instance id mint and the ListDefinition registry composite list
 * styles reconstruct into. */
function installListReadState(catalog: OdfStyleCatalog): void {
  const defsByStyle = new Map<string, ListDefinition>();
  const definitions: ListDefinition[] = [];
  let counter = 0;
  catalog.listRead = {
    nextListId(): string {
      counter += 1;
      return `l${counter}`;
    },
    definitionFor(styleName: string | undefined): ListDefinition {
      const key = styleName ?? "";
      let def = defsByStyle.get(key);
      if (!def) {
        def = definitionFromOdfListStyle(
          styleName,
          styleName ? catalog.listStyles.get(styleName) : undefined,
        );
        defsByStyle.set(key, def);
        definitions.push(def);
      }
      return def;
    },
    definitions,
  };
}

/** Reconstructs a ListDefinition from a text:list-style's levels — the
 * inverse of the writer's definedListLevelXml, with the same default
 * elision as commands.normalizeDefinitionLevels so a save/re-open cycle is
 * a fixed point. */
function definitionFromOdfListStyle(
  styleName: string | undefined,
  style: OdfListStyle | undefined,
): ListDefinition {
  const levels: ListLevelDefinition[] = [];
  for (let level = 1; level <= 9; level++) {
    const raw = style?.get(level);
    const resolved = raw ?? resolveListLevel(style, level);
    const levelDef: ListLevelDefinition = {
      style: resolved?.marker ?? defaultMarkerStyle("number", level),
    };
    if (resolved?.prefix !== undefined) {
      levelDef.prefix = resolved.prefix;
    }
    if (resolved?.suffix !== undefined) {
      levelDef.suffix = resolved.suffix;
    }
    if (resolved !== undefined && resolved.type === "number" && resolved.startAt !== 1) {
      levelDef.startAt = resolved.startAt;
    }
    if (raw?.displayLevels !== undefined && raw.displayLevels > 1) {
      levelDef.includeLevels = Math.min(raw.displayLevels, level);
    }
    if (raw?.alignment !== undefined) {
      levelDef.alignment = raw.alignment;
    }
    const textIndentCm = raw?.indentCm;
    if (
      textIndentCm !== undefined &&
      cmToTwips(textIndentCm) !== cmToTwips(level * LIST_INDENT_STEP_CM)
    ) {
      levelDef.textIndentCm = textIndentCm;
    }
    const effectiveText = levelDef.textIndentCm ?? level * LIST_INDENT_STEP_CM;
    if (
      raw?.numberIndentCm !== undefined &&
      cmToTwips(raw.numberIndentCm) !== cmToTwips(Math.max(0, effectiveText - LIST_MARKER_WIDTH_CM))
    ) {
      levelDef.numberIndentCm = raw.numberIndentCm;
    }
    levels.push(levelDef);
  }
  return { defId: styleName !== undefined && styleName !== "" ? styleName : "list", levels };
}

/** Plans list instances for one container's top-level `text:list` children
 * BEFORE they are parsed — a continuation may reference a list that has to
 * become instance-numbered retroactively, so a single forward parse cannot
 * decide. Chains link through text:continue-list (by xml:id) or
 * text:continue-numbering (previous list of the same style); every chain
 * shares one listId. A list parses as DEFINED when its style is a full
 * multilevel definition, when its chain has more than one segment (an
 * interrupted instance), or when it touches another list directly (two
 * instances back to back — the restart case the previous attempt lost,
 * AC-28). Isolated simple lists keep the SPEC-018 adjacency form, so
 * pre-SPEC-032 documents read exactly as before. */
function planListInstances(
  container: Element,
  catalog: OdfStyleCatalog,
): Map<Element, PlannedListInstance> {
  const plan = new Map<Element, PlannedListInstance>();
  const listRead = catalog.listRead;
  if (!listRead) {
    return plan;
  }
  interface ListInfo {
    el: Element;
    styleName: string;
    token: number;
    adjacent: boolean;
  }
  const infos: ListInfo[] = [];
  const tokenSizes: number[] = [];
  const tokenByXmlId = new Map<string, number>();
  const lastTokenByStyle = new Map<string, number>();
  let previous: ListInfo | null = null;
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "sequence-decls") {
      continue; // Leading declarations, not content.
    }
    if (child.namespaceURI !== TEXT_NAMESPACE || child.localName !== "list") {
      previous = null;
      continue;
    }
    const styleName = textAttr(child, "style-name") ?? "";
    const continueList = textAttr(child, "continue-list");
    const continueNumbering = textAttr(child, "continue-numbering") === "true";
    let token: number | undefined;
    if (continueList !== null && tokenByXmlId.has(continueList)) {
      token = tokenByXmlId.get(continueList);
    } else if (continueNumbering && lastTokenByStyle.has(styleName)) {
      token = lastTokenByStyle.get(styleName);
    }
    if (token === undefined) {
      token = tokenSizes.length;
      tokenSizes.push(0);
    }
    tokenSizes[token] += 1;
    const info: ListInfo = { el: child, styleName, token, adjacent: previous !== null };
    if (previous !== null) {
      previous.adjacent = true;
    }
    infos.push(info);
    const xmlId = child.getAttributeNS(XML_NAMESPACE, "id") ?? child.getAttribute("xml:id");
    if (xmlId) {
      tokenByXmlId.set(xmlId, token);
    }
    lastTokenByStyle.set(styleName, token);
    previous = info;
  }
  // Defined-ness is per chain: any qualifying segment upgrades the whole
  // chain (its earlier segments must carry the same instance).
  const tokenDefined = new Map<number, boolean>();
  for (const info of infos) {
    const composite = listStyleIsComposite(
      info.styleName !== "" ? catalog.listStyles.get(info.styleName) : undefined,
    );
    if (composite || info.adjacent || tokenSizes[info.token] > 1) {
      tokenDefined.set(info.token, true);
    }
  }
  const tokenListIds = new Map<number, string>();
  for (const info of infos) {
    if (tokenDefined.get(info.token) !== true) {
      plan.set(info.el, { defined: false, listId: "" });
      continue;
    }
    let listId = tokenListIds.get(info.token);
    if (listId === undefined) {
      listId = listRead.nextListId();
      tokenListIds.set(info.token, listId);
    }
    plan.set(info.el, { defined: true, listId });
  }
  return plan;
}

/** The instance parameter a planned top-level list parses under, or
 * undefined for the SPEC-018 path. */
function definedInstanceFor(
  listEl: Element,
  plan: Map<Element, PlannedListInstance>,
  catalog: OdfStyleCatalog,
): { def: ListDefinition; listId: string } | undefined {
  const planned = plan.get(listEl);
  if (!planned?.defined || !catalog.listRead) {
    return undefined;
  }
  const styleName = textAttr(listEl, "style-name") ?? undefined;
  return { def: catalog.listRead.definitionFor(styleName), listId: planned.listId };
}

/** Parses a `text:list` (and its nested sub-lists) into consecutive list
 * paragraphs (AC-21): nesting depth is the model's level, and the list's
 * `text:list-style` levels supply type, marker, start and indent. Nested
 * lists without their own style name inherit the ancestor's, per ODF.
 * Under an `instance` (SPEC-032), items become defined members of the
 * instance's definition instead, with a `text:start-value` restart carried
 * on the item's first paragraph. */
function parseListElement(
  listEl: Element,
  catalog: OdfStyleCatalog,
  depth: number,
  inheritedStyleName: string | undefined,
  blocks: Block[],
  opaqueBlocks: OpaqueBlock[],
  instance?: { def: ListDefinition; listId: string },
): void {
  const styleName = textAttr(listEl, "style-name") ?? inheritedStyleName;
  const level = clampListLevel(depth);
  for (const item of Array.from(listEl.children)) {
    const isItem =
      item.namespaceURI === TEXT_NAMESPACE &&
      (item.localName === "list-item" || item.localName === "list-header");
    if (!isItem) {
      opaqueBlocks.push({ xml: serializeElement(item), blockIndex: blocks.length });
      continue;
    }
    let firstParagraphOfItem = true;
    for (const child of Array.from(item.children)) {
      if (
        child.namespaceURI === TEXT_NAMESPACE &&
        (child.localName === "p" || child.localName === "h")
      ) {
        const paragraph = parseTextBlock(child, catalog, child.localName === "h");
        if (instance) {
          const list: ListFormatting = {
            type: isBulletStyle(instance.def.levels[level - 1]?.style ?? "decimal")
              ? "bullet"
              : "number",
            level,
            defId: instance.def.defId,
            listId: instance.listId,
          };
          if (firstParagraphOfItem) {
            const startAttr = textAttr(item, "start-value");
            const startValue = startAttr !== null ? Number(startAttr) : NaN;
            if (Number.isInteger(startValue) && startValue >= 0) {
              list.startAt = startValue;
            }
          }
          paragraph.list = list;
        } else {
          paragraph.list = buildListFormatting(
            resolveListLevel(styleName ? catalog.listStyles.get(styleName) : undefined, level),
            level,
          );
        }
        firstParagraphOfItem = false;
        blocks.push(paragraph);
      } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "list") {
        parseListElement(child, catalog, depth + 1, styleName, blocks, opaqueBlocks, instance);
      } else {
        // Unmodelled list-item content (tables, frames, …) is preserved at
        // the body level, anchored where the item sat (SPEC-008).
        opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
      }
    }
  }
}

// --- tables (SPEC-022) -----------------------------------------------------

/** Column widths declared by `table-column`-family styles, from both parts'
 * automatic styles: style name -> `style:column-width` in cm. */
function collectTableColumnWidths(contentDoc: Document, stylesDoc: Document | undefined): Map<string, number> {
  const widths = new Map<string, number>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "table-column") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "table-column-properties");
      const width = props ? odfLengthToCm(styleAttr(props, "column-width")) : undefined;
      if (name && width !== undefined) {
        widths.set(name, width);
      }
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return widths;
}

// Repeat expansion cap: `table:number-columns-repeated` is also how ODF
// producers pad to the sheet-like maximum; anything past this is not a text
// document's table and stays opaque.
const MAX_TABLE_REPEAT = 512;

function repeatCount(el: Element, attr: string): number | null {
  const raw = tableAttr(el, attr);
  if (raw === null) {
    return 1;
  }
  const count = Number(raw);
  if (!Number.isInteger(count) || count < 1 || count > MAX_TABLE_REPEAT) {
    return null;
  }
  return count;
}

/** ODF's `fo:border`-family shorthand ("0.5pt solid #ff0000") or "none"/
 * "hidden" (SPEC-023). */
function parseBorderShorthand(value: string | null): BorderSide | undefined {
  if (!value) {
    return undefined;
  }
  const trimmed = value.trim();
  if (trimmed === "none" || trimmed === "hidden") {
    return { style: "none", widthPt: 0, color: "#000000" };
  }
  const match = /^(\S+)\s+(\S+)\s+(#[0-9a-fA-F]{6})$/.exec(trimmed);
  if (!match) {
    return undefined;
  }
  const widthPt = odfLengthToPt(match[1]) ?? 1;
  const style = mapOdfBorderStyle(match[2]);
  const color = odfColorToCss(match[3]) ?? "#000000";
  return { style, widthPt, color };
}

/** ODF's CSS2-derived border style keywords collapse to the model's small
 * enum (SPEC-023), the same simplification the .docx reader applies. */
function mapOdfBorderStyle(word: string): BorderLineStyle {
  const w = word.toLowerCase();
  if (w === "none" || w === "hidden") {
    return "none";
  }
  if (w === "double") {
    return "double";
  }
  if (w === "dashed") {
    return "dashed";
  }
  if (w === "dotted") {
    return "dotted";
  }
  return "single"; // "solid", "groove", "ridge", "inset", "outset", …
}

/** A `style:table-cell-properties` element's SPEC-023 fields: per-side
 * borders (falling back to the `fo:border` shorthand for any side without
 * its own `fo:border-*`), background shading, padding and vertical
 * alignment. */
interface ParsedCellStyle {
  borders?: CellBorders;
  shadingFill?: string;
  paddingCm?: number;
  verticalAlign?: CellVerticalAlign;
}

function parseCellStyleProperties(props: Element): ParsedCellStyle {
  const shorthand = foAttr(props, "border");
  const top = parseBorderShorthand(foAttr(props, "border-top") ?? shorthand);
  const bottom = parseBorderShorthand(foAttr(props, "border-bottom") ?? shorthand);
  const left = parseBorderShorthand(foAttr(props, "border-left") ?? shorthand);
  const right = parseBorderShorthand(foAttr(props, "border-right") ?? shorthand);
  const result: ParsedCellStyle = {};
  if (top || bottom || left || right) {
    const borders: CellBorders = {};
    if (top) {
      borders.top = top;
    }
    if (bottom) {
      borders.bottom = bottom;
    }
    if (left) {
      borders.left = left;
    }
    if (right) {
      borders.right = right;
    }
    result.borders = borders;
  }
  const shadingFill = odfColorToCss(foAttr(props, "background-color"));
  if (shadingFill !== undefined) {
    result.shadingFill = shadingFill;
  }
  const paddingCm = odfLengthToCm(foAttr(props, "padding-top") ?? foAttr(props, "padding"));
  if (paddingCm !== undefined) {
    result.paddingCm = paddingCm;
  }
  const vAlign = styleAttr(props, "vertical-align");
  if (vAlign === "top" || vAlign === "bottom") {
    result.verticalAlign = vAlign;
  } else if (vAlign === "middle") {
    result.verticalAlign = "center";
  }
  return result;
}

/** A `style:style` (family "table") element's SPEC-023 width fields:
 * `style:rel-width` is a percentage of the text area; its absence means a
 * fixed width (the `columnWidthsCm` sum) — ODF text tables have no distinct
 * "size to content" mode the way `w:tblLayout="autofit"` does, so `"auto"`
 * is reachable only by setting it in the editor, not by reading one. */
interface ParsedTableStyle {
  widthMode?: TableWidthMode;
  widthPercent?: number;
}

/** The three style families a table references, collected once per document
 * (mirroring collectTableColumnWidths' scan-both-parts pattern) and threaded
 * through parseTableElement/parseTableCellElement. */
interface TableStyleTables {
  columnWidths: Map<string, number>;
  cellStyles: Map<string, ParsedCellStyle>;
  rowMinHeights: Map<string, number>;
  tableStyles: Map<string, ParsedTableStyle>;
}

function collectTableCellStyles(contentDoc: Document, stylesDoc: Document | undefined): Map<string, ParsedCellStyle> {
  const styles = new Map<string, ParsedCellStyle>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "table-cell") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "table-cell-properties");
      if (name && props) {
        styles.set(name, parseCellStyleProperties(props));
      }
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return styles;
}

function collectTableRowMinHeights(contentDoc: Document, stylesDoc: Document | undefined): Map<string, number> {
  const heights = new Map<string, number>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "table-row") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "table-row-properties");
      // `min-row-height` (a floor, matching the model) takes precedence over
      // `row-height` (a fixed height, approximated the same way as the
      // .docx reader treats `w:hRule="exact"`).
      const height = props
        ? (odfLengthToCm(styleAttr(props, "min-row-height")) ?? odfLengthToCm(styleAttr(props, "row-height")))
        : undefined;
      if (name && height !== undefined) {
        heights.set(name, height);
      }
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return heights;
}

function collectTableStyles(contentDoc: Document, stylesDoc: Document | undefined): Map<string, ParsedTableStyle> {
  const styles = new Map<string, ParsedTableStyle>();
  const scan = (doc: Document | undefined) => {
    if (!doc) {
      return;
    }
    for (const el of Array.from(doc.getElementsByTagNameNS(STYLE_NAMESPACE, "style"))) {
      if (el.getAttributeNS(STYLE_NAMESPACE, "family") !== "table") {
        continue;
      }
      const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
      const props = firstChildElement(el, STYLE_NAMESPACE, "table-properties");
      if (!name || !props) {
        continue;
      }
      const relWidth = odfPercentage(styleAttr(props, "rel-width"));
      styles.set(
        name,
        relWidth !== undefined
          ? { widthMode: "percent", widthPercent: Math.min(100, Math.max(1, Math.round(relWidth))) }
          : { widthMode: "fixed" },
      );
    }
  };
  scan(contentDoc);
  scan(stylesDoc);
  return styles;
}

/** `table:align` on the `table:table` element itself (SPEC-023 AC-22, not a
 * style property): "margins" is ODF's own "left, indented like a paragraph"
 * default and reads the same as "left" — table RTL is out of scope. */
function parseTableAlignmentAttr(tableEl: Element): TableAlignment | undefined {
  const val = tableAttr(tableEl, "align");
  if (val === "center" || val === "right") {
    return val;
  }
  if (val === "left" || val === "margins") {
    return "left";
  }
  return undefined;
}

/** The leading tracked-change marker(s) SPEC-039 AC-16's writer prepends to
 * a table row/cell's first paragraph — see writer.ts's rowChangeMarkerXml/
 * cellChangeMarkerXml. A `table:change-scope="row"` marker is the row's
 * own (only ever found on the row's first real cell); an unscoped one is
 * this cell's own. Both can be present at once (the row's first cell also
 * independently cell-tracked). Read-only — the caller strips them via
 * stripLeadingChangeMarkers before parsing the paragraph's own content, so
 * they are never mistaken for it (or for the paragraph's own trailing
 * markInsert/markDelete marker, at the opposite end — see
 * paragraphMarkChange). */
interface LeadingCellMarkers {
  rowMark?: OdtChangedRegion;
  cellMark?: OdtChangedRegion;
}

function firstParagraphChild(el: Element): Element | undefined {
  return Array.from(el.children).find(
    (child) => child.namespaceURI === TEXT_NAMESPACE && (child.localName === "p" || child.localName === "h"),
  );
}

function leadingCellChangeMarkers(cellEl: Element): LeadingCellMarkers {
  const result: LeadingCellMarkers = {};
  const firstPara = firstParagraphChild(cellEl);
  if (!firstPara) {
    return result;
  }
  for (const child of Array.from(firstPara.children)) {
    if (child.namespaceURI !== TEXT_NAMESPACE || child.localName !== "change") {
      break;
    }
    const changeId = textAttr(child, "change-id");
    const region = changeId ? odtChangedRegions?.get(changeId) : undefined;
    if (!region || (region.kind !== "insert" && region.kind !== "delete")) {
      continue;
    }
    if (tableAttr(child, "change-scope") === "row") {
      result.rowMark = region;
    } else {
      result.cellMark = region;
    }
  }
  return result;
}

/** A clone of `cellEl` with its first paragraph's leading change markers
 * (see leadingCellChangeMarkers) removed — cheap (cells are small); a
 * no-op clone when there are none. */
function stripLeadingChangeMarkers(cellEl: Element): Element {
  const clone = cellEl.cloneNode(true) as Element;
  const firstPara = firstParagraphChild(clone);
  if (!firstPara) {
    return clone;
  }
  const toRemove: Element[] = [];
  for (const child of Array.from(firstPara.children)) {
    if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "change") {
      toRemove.push(child);
      continue;
    }
    break;
  }
  for (const el of toRemove) {
    firstPara.removeChild(el);
  }
  return clone;
}

function parseTableCellElement(
  cellEl: Element,
  catalog: OdfStyleCatalog,
  cellStyles: Map<string, ParsedCellStyle>,
): TableCell | null {
  const blocks: Block[] = [];
  const strayOpaque: OpaqueBlock[] = [];
  // The cell is its own numbering context (SPEC-022), so its lists get
  // their own instance plan (SPEC-032).
  const listPlan = planListInstances(cellEl, catalog);
  for (const child of Array.from(cellEl.children)) {
    if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "p") {
      blocks.push(parseTextBlock(child, catalog, false));
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "h") {
      blocks.push(parseTextBlock(child, catalog, true));
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "list") {
      parseListElement(
        child,
        catalog,
        1,
        undefined,
        blocks,
        strayOpaque,
        definedInstanceFor(child, listPlan, catalog),
      );
    } else {
      // Nested tables, frames, … — outside the model's cell shape.
      return null;
    }
  }
  if (strayOpaque.length > 0) {
    // A cell whose list carried unmodelled content has no body level to
    // anchor it to; preserve the whole table verbatim instead.
    return null;
  }
  const paragraphs: Paragraph[] = [];
  for (const block of blocks) {
    if (block.type !== "paragraph") {
      return null;
    }
    paragraphs.push(block);
  }
  if (paragraphs.length === 0) {
    paragraphs.push({ type: "paragraph", runs: [], formatting: {} });
  }
  const cell: TableCell = { blocks: paragraphs };
  // Merged cells (SPEC-023 AC-28): the origin carries its own span; the
  // grid slots it covers are separate `table:covered-table-cell` elements
  // (handled by the caller, not here).
  const colSpan = Number(tableAttr(cellEl, "number-columns-spanned") ?? "1");
  if (Number.isInteger(colSpan) && colSpan > 1) {
    cell.colSpan = colSpan;
  }
  const rowSpan = Number(tableAttr(cellEl, "number-rows-spanned") ?? "1");
  if (Number.isInteger(rowSpan) && rowSpan > 1) {
    cell.rowSpan = rowSpan;
  }
  const styleName = tableAttr(cellEl, "style-name");
  if (styleName !== null) {
    cell.odtStyleName = styleName;
    const parsedStyle = cellStyles.get(styleName);
    if (parsedStyle?.borders) {
      cell.borders = parsedStyle.borders;
    }
    if (parsedStyle?.shadingFill !== undefined) {
      cell.shadingFill = parsedStyle.shadingFill;
    }
    if (parsedStyle?.paddingCm !== undefined) {
      cell.paddingCm = parsedStyle.paddingCm;
    }
    if (parsedStyle?.verticalAlign !== undefined) {
      cell.verticalAlign = parsedStyle.verticalAlign;
    }
  }
  return cell;
}

/** Parses a `table:table` into a modelled Table (SPEC-022; merged cells and
 * formatting SPEC-023), or null when it uses features outside the model's
 * regular grid (header rows, nested tables, sub-tables, …) — the caller
 * then preserves the element verbatim as an opaque block (SPEC-008).
 *
 * A merged cell's covered neighbours are explicit `table:covered-table-cell`
 * elements in the source — unlike `.docx`'s inferred `w:gridSpan`/`w:vMerge`
 * chains, ODF already lists one element per grid slot, so no second pass is
 * needed: each becomes a `covered: true` placeholder directly. */
/** Parses one `table:table-row` (repeated `number-rows-repeated` times) into
 * its `TableRow`s, or null on a shape the model cannot express — shared by
 * the plain row branch and the `table:table-header-rows` wrapper branch of
 * `parseTableElement`, which differ only in whether the resulting rows are
 * tagged `isHeader` (SPEC-043). */
function parseTableRowElement(
  rowEl: Element,
  isHeaderRow: boolean,
  catalog: OdfStyleCatalog,
  styleTables: TableStyleTables,
): TableRow[] | null {
  const repeat = repeatCount(rowEl, "number-rows-repeated");
  if (repeat === null) {
    return null;
  }
  const cells: TableCell[] = [];
  // SPEC-039 AC-16: a pending row/cell insert/delete's own marker — see
  // writer.ts's rowChangeMarkerXml/cellChangeMarkerXml — leads the affected
  // cell's first paragraph; the row's own marker always lands on the row's
  // first real cell, so at most one is ever found per row.
  let rowChange: TrackedChangeInfo | undefined;
  let rowChangeKind: "insert" | "delete" | undefined;
  for (const cellEl of Array.from(rowEl.children)) {
    if (cellEl.namespaceURI !== TABLE_NAMESPACE) {
      return null;
    }
    const cellRepeat = repeatCount(cellEl, "number-columns-repeated");
    if (cellRepeat === null) {
      return null;
    }
    let cell: TableCell | null;
    if (cellEl.localName === "table-cell") {
      const markers = leadingCellChangeMarkers(cellEl);
      const parseTarget = markers.rowMark || markers.cellMark ? stripLeadingChangeMarkers(cellEl) : cellEl;
      cell = parseTableCellElement(parseTarget, catalog, styleTables.cellStyles);
      if (cell && markers.cellMark) {
        cell.changes = { ...cell.changes, [markers.cellMark.kind]: markers.cellMark };
      }
      if (markers.rowMark) {
        rowChange = markers.rowMark;
        rowChangeKind = markers.rowMark.kind as "insert" | "delete";
      }
    } else if (cellEl.localName === "covered-table-cell") {
      cell = { blocks: [], covered: true };
    } else {
      return null;
    }
    if (!cell) {
      return null;
    }
    cells.push(cell);
    for (let i = 1; i < cellRepeat; i++) {
      cells.push(structuredClone(cell));
    }
  }
  if (cells.length === 0) {
    return null;
  }
  const rowStyleName = tableAttr(rowEl, "style-name");
  const minHeightCm = rowStyleName ? styleTables.rowMinHeights.get(rowStyleName) : undefined;
  const rows: TableRow[] = [];
  for (let i = 0; i < repeat; i++) {
    const row: TableRow = { cells: i === 0 ? cells : structuredClone(cells) };
    if (rowStyleName !== null) {
      row.odtStyleName = rowStyleName;
    }
    if (minHeightCm !== undefined) {
      row.minHeightCm = minHeightCm;
    }
    if (isHeaderRow) {
      row.isHeader = true;
    }
    if (rowChange && rowChangeKind) {
      row.changes = { ...row.changes, [rowChangeKind]: rowChange };
    }
    rows.push(row);
  }
  return rows;
}

function parseTableElement(tableEl: Element, catalog: OdfStyleCatalog, styleTables: TableStyleTables): Table | null {
  const columnStyleNames: Array<string | undefined> = [];
  const rows: TableRow[] = [];
  for (const child of Array.from(tableEl.children)) {
    if (child.namespaceURI !== TABLE_NAMESPACE) {
      return null;
    }
    if (child.localName === "table-column") {
      const repeat = repeatCount(child, "number-columns-repeated");
      if (repeat === null) {
        return null;
      }
      const styleName = tableAttr(child, "style-name") ?? undefined;
      for (let i = 0; i < repeat; i++) {
        columnStyleNames.push(styleName);
      }
    } else if (child.localName === "table-row") {
      const parsedRows = parseTableRowElement(child, false, catalog, styleTables);
      if (!parsedRows) {
        return null;
      }
      rows.push(...parsedRows);
    } else if (child.localName === "table-header-rows") {
      // The leading run of header rows (SPEC-043): a structural wrapper
      // around plain table-row elements, not a per-row attribute — each row
      // inside it is otherwise parsed exactly like an ordinary table-row.
      for (const headerChild of Array.from(child.children)) {
        if (headerChild.namespaceURI !== TABLE_NAMESPACE || headerChild.localName !== "table-row") {
          return null;
        }
        const parsedRows = parseTableRowElement(headerChild, true, catalog, styleTables);
        if (!parsedRows) {
          return null;
        }
        rows.push(...parsedRows);
      }
    } else {
      // sub-tables, title/desc, … — the model cannot express these; keep
      // the table verbatim.
      return null;
    }
  }
  if (rows.length === 0 || columnStyleNames.length === 0) {
    return null;
  }
  if (rows.some((row) => row.cells.length !== columnStyleNames.length)) {
    return null;
  }
  // Column widths from the referenced column styles; when any is missing,
  // fall back to an equal split of the text area (or of the declared
  // widths' own total when only some are missing there is no better
  // anchor).
  const fallback = createDefaultStyles();
  const pageSize = catalog.pageSize ?? fallback.pageSize;
  const margins = catalog.margins ?? fallback.margins;
  const textAreaCm = Math.max(1, pageSize.widthCm - margins.leftCm - margins.rightCm);
  const declared = columnStyleNames.map((name) => (name ? styleTables.columnWidths.get(name) : undefined));
  const columnWidthsCm = declared.every((width) => width !== undefined && width > 0)
    ? (declared as number[])
    : columnStyleNames.map(() => textAreaCm / columnStyleNames.length);

  const table: Table = { type: "table", rows, columnWidthsCm };
  const name = tableAttr(tableEl, "name");
  if (name !== null) {
    table.odtName = name;
  }
  const styleName = tableAttr(tableEl, "style-name");
  if (styleName !== null) {
    table.odtStyleName = styleName;
    const parsedStyle = styleTables.tableStyles.get(styleName);
    if (parsedStyle?.widthMode !== undefined) {
      table.widthMode = parsedStyle.widthMode;
    }
    if (parsedStyle?.widthPercent !== undefined) {
      table.widthPercent = parsedStyle.widthPercent;
    }
  }
  const alignment = parseTableAlignmentAttr(tableEl);
  if (alignment !== undefined) {
    table.alignment = alignment;
  }
  if (columnStyleNames.some((columnName) => columnName !== undefined)) {
    table.odtColumnStyleNames = columnStyleNames;
  }
  return table;
}

// --- package ---------------------------------------------------------------

// --- headers, footers and page numbers (SPEC-029) ---------------------------

/** One `<style:header>`/`<style:header-left>`/`<style:footer>`/
 * `<style:footer-left>` element's paragraphs — reusing parseTextBlock the
 * same way body content does. Tables and images inside a master page's
 * header/footer are out of this pass's scope (unlike .docx, ODF has no
 * per-part opaque-block anchoring to fall back to here) and are silently
 * skipped rather than dropping the whole header/footer; undefined for a
 * master page with no such element at all (distinct from an element with
 * no paragraphs, which still yields one empty paragraph — a real, if
 * blank, header/footer). */
function parseHeaderFooterBlocks(el: Element | undefined, catalog: OdfStyleCatalog): Block[] | undefined {
  if (!el) {
    return undefined;
  }
  const blocks: Block[] = [];
  for (const child of Array.from(el.children)) {
    if (child.namespaceURI === TEXT_NAMESPACE && (child.localName === "p" || child.localName === "h")) {
      blocks.push(parseTextBlock(child, catalog, child.localName === "h"));
    }
  }
  return blocks.length > 0 ? blocks : [{ type: "paragraph", runs: [], formatting: {} }];
}

/** Resolves a master page's header/footer content onto a section (SPEC-029
 * AC-9, AC-12, AC-26): `-left` variants are ODF's own representation of
 * "different odd and even pages" — their mere presence turns it on, there
 * is no separate flag to read. `catalog` must be the styles-xml-scoped one
 * (see parseOdt) since header/footer paragraphs can only reference
 * automatic styles declared in styles.xml itself, a separate scope from
 * content.xml's. */
function applyHeaderFooterLayout(
  section: Section,
  layout: OdfMasterPageLayout | undefined,
  stylesScopedCatalog: OdfStyleCatalog,
): void {
  if (!layout) {
    return;
  }
  const defaultHeader = parseHeaderFooterBlocks(layout.headerEl, stylesScopedCatalog);
  const evenHeader = parseHeaderFooterBlocks(layout.headerLeftEl, stylesScopedCatalog);
  const defaultFooter = parseHeaderFooterBlocks(layout.footerEl, stylesScopedCatalog);
  const evenFooter = parseHeaderFooterBlocks(layout.footerLeftEl, stylesScopedCatalog);
  if (defaultHeader) {
    section.headers = { ...section.headers, default: { blocks: defaultHeader } };
  }
  if (evenHeader) {
    section.headers = { ...section.headers, even: { blocks: evenHeader } };
  }
  if (defaultFooter) {
    section.footers = { ...section.footers, default: { blocks: defaultFooter } };
  }
  if (evenFooter) {
    section.footers = { ...section.footers, even: { blocks: evenFooter } };
  }
  if (evenHeader || evenFooter) {
    section.evenAndOddHeaders = true;
  }
}

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
// The model note settings a styles.xml `text:notes-configuration` carries
// (SPEC-031): num-format 1/i/a (ODF has no symbol format), start-value
// (the counter's initial value, one below the first displayed number) and
// start-numbering-at (page, or chapter as the closest ODF notion of a
// section restart).
function parseOdtNoteSettings(
  stylesDoc: Document | undefined,
  cls: NoteKind,
): NoteSettings | undefined {
  const officeStyles = stylesDoc
    ? firstChildElement(stylesDoc.documentElement, OFFICE_NAMESPACE, "styles")
    : undefined;
  if (!officeStyles) {
    return undefined;
  }
  for (const el of Array.from(officeStyles.children)) {
    if (el.namespaceURI !== TEXT_NAMESPACE || el.localName !== "notes-configuration") {
      continue;
    }
    if ((textAttr(el, "note-class") ?? "footnote") !== cls) {
      continue;
    }
    const settings: NoteSettings = {};
    const format = styleAttr(el, "num-format");
    if (format === "i" || format === "I") {
      settings.numberFormat = "lowerRoman";
    } else if (format === "a" || format === "A") {
      settings.numberFormat = "lowerAlpha";
    } else if (format === "1") {
      settings.numberFormat = "decimal";
    }
    const startValue = Number(textAttr(el, "start-value") ?? "");
    if (Number.isInteger(startValue) && startValue >= 1) {
      settings.startAt = startValue + 1;
    }
    const restart = textAttr(el, "start-numbering-at");
    if (restart === "page") {
      settings.restart = "page";
    } else if (restart === "chapter") {
      settings.restart = "section";
    }
    return Object.keys(settings).length > 0 ? settings : undefined;
  }
  return undefined;
}

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
  // SPEC-032: the list-instance mint and definition registry the body's
  // (and each cell's) instance plans draw from.
  installListReadState(catalog);
  // Header/footer paragraphs (SPEC-029) live in styles.xml and can only
  // reference automatic styles styles.xml itself declares — a scope
  // parseOdfStyleCatalog never folds in for the main `catalog` above (which
  // is deliberately content.xml-scoped, see its own docs). Building a
  // second catalog with styles.xml standing in as "content" gets those
  // automatic styles into `.automatic` for free, no separate collector
  // needed. Only ever used for header/footer content, never body content.
  const stylesScopedCatalog = stylesDoc ? parseOdfStyleCatalog(stylesDoc, stylesDoc) : catalog;
  const tableColumnWidths = collectTableColumnWidths(contentDoc, stylesDoc);
  const tableStyleTables: TableStyleTables = {
    columnWidths: tableColumnWidths,
    cellStyles: collectTableCellStyles(contentDoc, stylesDoc),
    rowMinHeights: collectTableRowMinHeights(contentDoc, stylesDoc),
    tableStyles: collectTableStyles(contentDoc, stylesDoc),
  };

  const body = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "body");
  const textRoot = body ? firstChildElement(body, OFFICE_NAMESPACE, "text") : undefined;
  if (!textRoot) {
    throw new OdtOpenError(INVALID_ODT_MESSAGE);
  }

  const drawingCtx: OdfDrawingContext = {
    parts,
    graphicWrapStyles: collectGraphicWrapStyles(contentDoc, stylesDoc),
    contentDoc,
    stylesDoc,
  };
  const mediaById = new Map<string, MediaPart>();
  const sections: Section[] = [];
  let blocks: Block[] = [];
  let opaqueBlocks: OpaqueBlock[] = [];
  // The geometry of the master page the CURRENT section's first paragraph
  // switched to (SPEC-028 AC-33); undefined for the first section, whose
  // master is the document default folded into defaultStyles below.
  let sectionGeometry: OdfMasterPageLayout | undefined;
  // The master page named by the document's very FIRST body paragraph. The
  // reader never opens an empty leading section for such a reference (see
  // the guard below) — but that master, not the first-listed one, is the one
  // the body actually starts on, so it must supply the document defaults and
  // the first section's geometry (SPEC-028 AC-17/18/19/33). Undefined when
  // the first paragraph names no master: the first-listed master stays the
  // default, exactly as before.
  let leadingMasterName: string | undefined;
  // "Start numbering at" (SPEC-029 AC-22): style:page-number on the section-
  // starting paragraph's automatic style, alongside its master-page-name.
  let sectionPageNumberStart: number | undefined;
  let leadingPageNumberStart: number | undefined;
  const finishSection = () => {
    const section: Section = { blocks };
    if (opaqueBlocks.length > 0) {
      section.opaqueBlocks = opaqueBlocks;
    }
    // Explicit geometry on every later section — the master page pins it,
    // so it must not float with a later defaultStyles edit. `start` stays
    // absent: ODF's master-page switch is always a page break (nextPage).
    if (sectionGeometry?.pageSize) {
      section.pageSize = { ...sectionGeometry.pageSize };
    }
    if (sectionGeometry?.margins) {
      section.margins = { ...sectionGeometry.margins };
    }
    if (sectionGeometry?.columns) {
      section.columns = { ...sectionGeometry.columns };
    }
    applyHeaderFooterLayout(section, sectionGeometry, stylesScopedCatalog);
    if (sectionPageNumberStart !== undefined) {
      section.pageNumberStart = sectionPageNumberStart;
    }
    sections.push(section);
  };
  // SPEC-032: which top-level lists parse as defined instances, planned
  // before the walk (continuations need lookahead).
  const bodyListPlan = planListInstances(textRoot, catalog);
  // SPEC-031: notes anywhere in the body walk collect here (see odtNoteSink).
  const notes: NoteContent[] = [];
  odtNoteSink = notes;
  odtNoteCounter = 0;
  // SPEC-038: comments anywhere in the body walk collect here (see
  // odtCommentSink).
  const comments: CommentThread[] = [];
  odtCommentSink = comments;
  odtCommentCounter = 0;
  // SPEC-039: the tracked-change definitions text:change-start/-end/-change
  // marks anywhere in the body walk reference (see odtChangedRegions).
  odtChangedRegions = parseTrackedChangesRegions(textRoot, catalog);
  odtPendingChangeStarts = new Map();
  for (const child of Array.from(textRoot.children)) {
    if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "tracked-changes") {
      continue;
    }
    if (child.namespaceURI === TEXT_NAMESPACE && (child.localName === "p" || child.localName === "h")) {
      const signals = paragraphAutoStyleSignals(catalog, textAttr(child, "style-name") ?? undefined);
      // SPEC-029 AC-22: a style:page-number on the document's very first
      // paragraph restarts the first section's own numbering — independent
      // of whether that paragraph also switches master page (it usually
      // does not; section 0 has no master to switch to).
      if (sections.length === 0 && blocks.length === 0 && signals.pageNumberStart !== undefined) {
        leadingPageNumberStart = signals.pageNumberStart;
      }
      // A style:master-page-name on the paragraph's automatic style starts
      // a new section here (SPEC-028 AC-33) — except on the document's very
      // first block, where opening an empty leading section would be wrong.
      // But "first master page" is then the one this paragraph actually
      // names — NOT whichever is listed first in master-styles (a foreign
      // doc can list Standard first and have the body start on A5Master) —
      // so record it and let it provide the document defaults below.
      if (signals.masterPageName !== undefined) {
        if (sections.length === 0 && blocks.length === 0) {
          leadingMasterName = signals.masterPageName;
        } else {
          finishSection();
          blocks = [];
          opaqueBlocks = [];
          sectionGeometry = catalog.masterPageLayouts.get(signals.masterPageName);
          sectionPageNumberStart = signals.pageNumberStart;
        }
      }
      // fo:break-before="page" on the AUTOMATIC style materialises a manual
      // BreakBlock before the paragraph, and the flag is dropped from the
      // paragraph's own formatting (SPEC-028 AC-26/32). ODF has one channel
      // for both the manual break and the pageBreakBefore formatting flag,
      // so .odt import canonicalises the ambiguity to the break block —
      // what the writer's fo:break-before actually stood for; only a NAMED
      // style's break-before stays pure formatting (see
      // paragraphAutoStyleSignals).
      if (signals.pageBreakBefore) {
        blocks.push({ type: "break", kind: "page" });
      } else if (signals.columnBreakBefore) {
        // fo:break-before="column" (one attribute — mutually exclusive with
        // "page") likewise materialises a column BreakBlock (AC-10/11); no
        // formatting flag to strip afterwards, since the model has no
        // column counterpart of pageBreakBefore (SPEC-030).
        blocks.push({ type: "break", kind: "column" });
      }
      const { blocks: parsedBlocks, media } = parseBodyTextBlockOrSplit(
        child,
        catalog,
        child.localName === "h",
        drawingCtx,
      );
      if (signals.pageBreakBefore) {
        for (const parsed of parsedBlocks) {
          if (parsed.type === "paragraph" && parsed.formatting) {
            delete parsed.formatting.pageBreakBefore;
          }
        }
      }
      blocks.push(...parsedBlocks);
      for (const part of media) {
        mediaById.set(part.mediaId, part);
      }
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "list") {
      // Bulleted/numbered lists (SPEC-018 AC-21; instances SPEC-032).
      parseListElement(
        child,
        catalog,
        1,
        undefined,
        blocks,
        opaqueBlocks,
        definedInstanceFor(child, bodyListPlan, catalog),
      );
    } else if (child.namespaceURI === TABLE_NAMESPACE && child.localName === "table") {
      // Regular-grid tables are modelled (SPEC-022); anything the model
      // cannot express falls through to the opaque path below.
      const table = parseTableElement(child, catalog, tableStyleTables);
      if (table) {
        blocks.push(table);
      } else {
        opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
      }
    } else if (child.namespaceURI === TEXT_NAMESPACE && child.localName === "table-of-content") {
      // SPEC-033 AC-18..20: a generated table of contents.
      blocks.push(parseTableOfContent(child));
    } else {
      // Sequence declarations, frames, ... — unmodelled body children are
      // preserved verbatim, anchored to their position among the
      // paragraphs (SPEC-008).
      opaqueBlocks.push({ xml: serializeElement(child), blockIndex: blocks.length });
    }
  }
  finishSection();
  odtNoteSink = null;
  odtCommentSink = null;
  odtChangedRegions = null;
  odtPendingChangeStarts = null;

  const fallback = createDefaultStyles();
  // The document default page geometry is the master the body actually
  // starts on: the master named by the first body paragraph when there is
  // one, else the first-listed master folded into catalog.pageSize/margins.
  // parsePageLayout folds the first-listed master for the no-lead case; when
  // a lead exists it wins (SPEC-028 AC-17/18/19/33 — a foreign doc must read
  // back the geometry it actually starts on, not whichever master happens to
  // be listed first).
  const leadingGeometry = leadingMasterName
    ? catalog.masterPageLayouts.get(leadingMasterName)
    : undefined;
  // The leading section's geometry is the master the very first paragraph
  // named, so it carries explicit pageSize/margins like every later section
  // that switches masters — reported by AC-17/18/19/33 and rendered/round-
  // tripped correctly rather than silently inheriting the first-LISTED
  // master. A document whose first paragraph names no master keeps the
  // first-listed master folded into defaults exactly as before.
  const leadingSection = sections[0];
  if (leadingSection && leadingGeometry?.pageSize) {
    leadingSection.pageSize = { ...leadingGeometry.pageSize };
  }
  if (leadingSection && leadingGeometry?.margins) {
    leadingSection.margins = { ...leadingGeometry.margins };
  }
  // The leading section's columns come from the master it names, else the
  // first-listed master folded into the catalog (the single-section case).
  const leadingColumns = leadingGeometry?.columns ?? catalog.columns;
  if (leadingSection && leadingColumns) {
    leadingSection.columns = { ...leadingColumns };
  }
  // Header/footer content (SPEC-029): the master the first paragraph
  // explicitly named, or — the common case — whichever master is first in
  // styles.xml (catalog.firstMasterName), the same resolution pageSize/
  // margins above already follow.
  if (leadingSection) {
    const leadingLayout =
      leadingGeometry ??
      (catalog.firstMasterName ? catalog.masterPageLayouts.get(catalog.firstMasterName) : undefined);
    applyHeaderFooterLayout(leadingSection, leadingLayout, stylesScopedCatalog);
    if (leadingPageNumberStart !== undefined) {
      leadingSection.pageNumberStart = leadingPageNumberStart;
    }
  }
  const defaultStyles: DefaultStyles = {
    pageSize: leadingGeometry?.pageSize ?? catalog.pageSize ?? fallback.pageSize,
    margins: leadingGeometry?.margins ?? catalog.margins ?? fallback.margins,
    fontFamily: catalog.defaultRunFormatting.fontFamily ?? fallback.fontFamily,
    fontSizePt: catalog.defaultRunFormatting.fontSizePt ?? fallback.fontSizePt,
    language: catalog.defaultLanguage ?? fallback.language,
  };

  // content.xml and meta.xml are rebuilt on save (body/title/timestamp);
  // everything else — mimetype, styles.xml, the manifest, settings,
  // embedded objects — is retained verbatim for the lifetime of the
  // document (SPEC-008). Media parts a modelled Image references (SPEC-024)
  // are excluded here — they travel in `media` instead (see docx/reader.ts's
  // parseDocx for why: the writer regenerates their manifest/relationship
  // entries from `media` on every save, same-format or not).
  const opaque: OpaquePart[] = [];
  for (const [name, data] of parts) {
    if (name !== "content.xml" && name !== "meta.xml" && !mediaById.has(name)) {
      opaque.push({ name, data: bytesToBase64(data) });
    }
  }

  const sourceXml: SourceXml = {};
  const fontFaceDecls = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "font-face-decls");
  if (fontFaceDecls) {
    sourceXml.odtFontFaceDecls = serializeElement(fontFaceDecls);
  }
  const automaticStyles = firstChildElement(contentDoc.documentElement, OFFICE_NAMESPACE, "automatic-styles");
  if (automaticStyles) {
    // Paragraph/text-family styles are resolved into the model and
    // regenerated flat on save; every other automatic style (graphic,
    // table, list, … families) may be referenced by preserved opaque nodes
    // and is kept verbatim. Families never share a name scope in ODF, so
    // the regenerated and preserved definitions cannot collide.
    const kept: string[] = [];
    for (const def of Array.from(automaticStyles.children)) {
      const family =
        def.namespaceURI === STYLE_NAMESPACE && def.localName === "style"
          ? def.getAttributeNS(STYLE_NAMESPACE, "family")
          : null;
      if (family !== "paragraph" && family !== "text") {
        kept.push(serializeElement(def));
      }
    }
    if (kept.length > 0) {
      sourceXml.odtAutomaticStyles = kept;
    }
  }

  // SPEC-039: settings.xml's RecordChanges config-item — the odt counterpart
  // of docx settings.xml's w:trackChanges, read the same regex-on-raw-text
  // way (settings.xml itself is otherwise unmodelled, travelling through
  // `opaque` above).
  const settingsBytes = parts.get("settings.xml");
  const settingsText = settingsBytes ? decodeUtf8(settingsBytes) : undefined;
  const trackChangesRecording =
    !!settingsText &&
    /<[\w:]*config-item\b[^>]*\bconfig:name="RecordChanges"[^>]*>\s*true\s*<\/[\w:]*config-item>/.test(settingsText);

  // Document properties and metadata (SPEC-041).
  const metaBytes = parts.get("meta.xml");
  const metaText = metaBytes ? decodeUtf8(metaBytes) : undefined;
  const properties: DocumentProperties = metaText ? parseOdfMeta(metaText) : {};
  if (metaText) {
    const { customProperties, unsupportedCustomProperties } = parseOdfCustomProperties(metaText);
    if (customProperties.length > 0) {
      properties.customProperties = customProperties;
    }
    if (unsupportedCustomProperties.length > 0) {
      properties.unsupportedCustomProperties = unsupportedCustomProperties;
    }
  }

  const result: ParsedOdtDocument = {
    sections,
    defaultStyles,
    styles: catalogToNamedStyles(catalog),
    opaque,
    sourceXml,
    media: Array.from(mediaById.values()),
    listDefinitions: catalog.listRead?.definitions ?? [],
    notes,
    comments,
    trackChangesRecording,
    properties,
  };
  const footnoteSettings = parseOdtNoteSettings(stylesDoc, "footnote");
  if (footnoteSettings) {
    result.footnoteSettings = footnoteSettings;
  }
  const endnoteSettings = parseOdtNoteSettings(stylesDoc, "endnote");
  if (endnoteSettings) {
    result.endnoteSettings = endnoteSettings;
  }
  return result;
}
