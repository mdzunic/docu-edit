// Orchestrates the .docx -> DocumentModel pipeline: unzip the OPC package,
// walk the package/document relationships to find the real part names (per
// AC-3, never assume "word/document.xml" is where it is), map w:p/w:r onto
// the document model, resolve w:pStyle against styles.xml, and apply the
// section's page setup.
import type {
  Block,
  BookmarkAnchor,
  BorderLineStyle,
  BorderSide,
  BreakKind,
  CellBorders,
  CellVerticalAlign,
  CommentAnchor,
  CommentMessage,
  CommentThread,
  DefaultStyles,
  FieldKind,
  HeaderFooterKind,
  HeaderFooterPart,
  ListDefinition,
  MediaPart,
  NamedStyle,
  NoteContent,
  NoteKind,
  NoteNumberFormat,
  NoteRestart,
  NoteSettings,
  OpaqueBlock,
  OpaqueInline,
  OpaquePart,
  PageNumberFormat,
  Paragraph,
  Run,
  RunFormatting,
  Section,
  SectionColumns,
  SectionStart,
  SourceXml,
  Table,
  TableAlignment,
  TableCell,
  TableRow,
  TableWidthMode,
  TableOfContents,
  TocEntry,
  TocOptions,
  TrackedChangeInfo,
} from "../documentModel";
import {
  createDefaultStyles,
  createFieldRun,
  createNoteRun,
  DEFAULT_COLUMN_GAP_CM,
  MAX_COLUMN_COUNT,
} from "../documentModel";
import { sanitizeHref } from "../links";
import { createMessageId, parseReplyParagraphText } from "../comments";
import { bytesToBase64 } from "../opaque";
import { blockText } from "../textEditing";
import { ZipError, unzip } from "./zip";
import {
  childElements,
  firstChildElement,
  parseRelationships,
  parseXml,
  relsPathFor,
  resolvePartPath,
  serializeElement,
  wordAttr,
  WORD_NAMESPACE,
  XmlParseError,
  type Relationship,
} from "./xml";
import { parseDrawing, type ParsedDrawing } from "./drawing";
import { catalogToNamedStyles, parseStylesCatalog, resolveStyle, type StyleCatalog } from "./styles";
import {
  EMPTY_NUMBERING_CATALOG,
  NUMBERING_RELATIONSHIP,
  parseNumberingCatalog,
  type NumberingCatalog,
} from "./numbering";
import {
  mergeParagraphFormatting,
  mergeRunFormatting,
  parseOutlineLevel,
  parseParagraphProperties,
  parseParagraphPropertiesChange,
  parseRunProperties,
  parseRunPropertiesChange,
} from "./properties";
import { eighthPointsToPt, ooxmlColorToCss, twipsToCm } from "./units";

export class DocxOpenError extends Error {}
export class LegacyDocError extends Error {}

export const INVALID_DOCX_MESSAGE = "This file isn't a Word document we can open";
export const LEGACY_DOC_MESSAGE = "Old .doc files aren't supported — save as .docx first";

const OFFICE_DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STYLES_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const HYPERLINK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
// The `r:` attribute namespace (w:hyperlink/@r:id) — same URI drawing.ts
// uses for a:blip/@r:embed.
const R_ATTRIBUTE_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Resolves a `w:hyperlink`'s r:id to its external URL (SPEC-025): the main
 * part's hyperlink relationships, sanitized — a dangerous scheme (AC-17)
 * resolves to undefined, which strips the link on open but keeps its text. */
export interface LinkReadContext {
  hrefForRelId(relId: string): string | undefined;
}

// Transient markup Word itself regenerates on every open — spell/grammar
// error ranges and rendered pagination markers. Preserving these as opaque
// nodes would freeze stale derived data into the file, so they are the one
// deliberate exception to SPEC-008's "carry everything through" rule.
const TRANSIENT_PARAGRAPH_CHILDREN = new Set(["proofErr"]);
const MODELLED_RUN_CHILDREN = new Set([
  "rPr",
  "t",
  "delText",
  "tab",
  "br",
  "cr",
  "noBreakHyphen",
  "softHyphen",
  "lastRenderedPageBreak",
]);

/** True when a run carries content the model cannot express (w:drawing,
 * w:object, w:pict, field chars, …) — such a run is preserved whole as an
 * opaque inline node rather than flattened to text. */
function runHasUnmodelledContent(rEl: Element): boolean {
  return Array.from(rEl.children).some(
    (child) => child.namespaceURI !== WORD_NAMESPACE || !MODELLED_RUN_CHILDREN.has(child.localName),
  );
}

/** The stable model note id for a source note's numeric id (SPEC-031) —
 * shared by the reference runs and the parsed footnotes.xml/endnotes.xml
 * contents so they link up. */
function docxNoteId(kind: NoteKind, sourceId: string): string {
  return `${kind === "footnote" ? "fn" : "en"}${sourceId}`;
}

/** A run's footnote/endnote reference (SPEC-031): the `w:footnoteReference`/
 * `w:endnoteReference` child's kind and target id, or null for an ordinary
 * run. Custom-mark references (`w:customMarkFollows`) are left to the
 * opaque-inline path — the model only numbers automatically. */
function noteReferenceOf(rEl: Element): { kind: NoteKind; noteId: string } | null {
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    const isFootnote = child.localName === "footnoteReference";
    const isEndnote = child.localName === "endnoteReference";
    if (!isFootnote && !isEndnote) {
      continue;
    }
    const sourceId = wordAttr(child, "id");
    if (!sourceId || wordAttr(child, "customMarkFollows") === "1") {
      return null;
    }
    const kind: NoteKind = isFootnote ? "footnote" : "endnote";
    return { kind, noteId: docxNoteId(kind, sourceId) };
  }
  return null;
}

/** The stable model thread id for a source comment's numeric range id
 * (SPEC-038) — shared by the body's commentRangeStart/End anchors and the
 * parsed word/comments.xml entries so they link up, mirroring docxNoteId.
 * The range id and the `<w:comment w:id>` it belongs to are always written
 * as the same number (see writer.ts's commentIdsFor), so no separate
 * commentReference lookup is needed to find which comment a range names. */
function docxCommentId(sourceId: string): string {
  return `wc${sourceId}`;
}

/** The model change id for a source `w:ins`/`w:del`/`w:rPrChange`/
 * `w:pPrChange`/`w:trPr` ins-or-del's `w:id`/its own identity (SPEC-039) —
 * same "derive, don't randomise" convention as docxCommentId, so re-parsing
 * the same file gives the same id and a same-format round trip compares
 * structurally equal. `kind` keeps the numeric ids each element type
 * happens to reuse from colliding (an insertion and a format change can
 * both be `w:id="1"` in the same file). */
function docxChangeId(kind: string, sourceId: string): string {
  return `wtc-${kind}-${sourceId}`;
}

const UNKNOWN_TRACK_CHANGE_AUTHOR = "Unknown";

/** A tracked-change range's author/date (SPEC-039), from the `w:author`/
 * `w:date` attributes every `w:ins`/`w:del`/`w:rPrChange`/`w:pPrChange`
 * carries — same fallback convention as word/comments.xml's author/date. */
function trackedChangeInfo(el: Element, kind: string): TrackedChangeInfo {
  const id = wordAttr(el, "id") ?? `${Math.random()}`;
  return {
    changeId: docxChangeId(kind, id),
    author: wordAttr(el, "author") || UNKNOWN_TRACK_CHANGE_AUTHOR,
    date: wordAttr(el, "date") || new Date(0).toISOString(),
  };
}

/** Whether a run's only content (besides `w:rPr`) is a `w:commentReference`
 * (SPEC-038) — dropped on read like a footnoteRef/endnoteRef marker: the
 * model derives which comment a range belongs to from the range's own id
 * (docxCommentId), so the reference run itself carries nothing the model
 * needs, and the writer regenerates a fresh one on save. */
function isCommentReferenceRun(rEl: Element): boolean {
  return Array.from(rEl.children).some(
    (child) => child.namespaceURI === WORD_NAMESPACE && child.localName === "commentReference",
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
  /** The paragraph styles of styles.xml as the model's style catalogue
   * (SPEC-019); empty when the package has no styles part. */
  styles: NamedStyle[];
  /** Every package part the reader did not model (SPEC-008), in package
   * order — written back byte-identical on a save to .docx. */
  opaque: OpaquePart[];
  /** Verbatim XML captured from modelled parts (the body-level w:sectPr). */
  sourceXml: SourceXml;
  /** Media parts modelled Image blocks reference (SPEC-024), one per
   * distinct source media part — deduplicated, since the same picture used
   * twice in the document shares one relationship target. */
  media: MediaPart[];
  /** Multilevel list definitions (SPEC-032) reconstructed from composite
   * abstractNums the body actually references; empty when there are none. */
  listDefinitions: ListDefinition[];
  /** Footnote/endnote contents (SPEC-031) the body's reference marks link
   * to; empty when the package has neither part. */
  notes: NoteContent[];
  /** Footnote numbering settings from the body sectPr's w:footnotePr. */
  footnoteSettings?: NoteSettings;
  /** Endnote numbering settings from the body sectPr's w:endnotePr. */
  endnoteSettings?: NoteSettings;
  /** Comment threads (SPEC-038) from word/comments.xml; empty when the
   * package has no comments part. */
  comments: CommentThread[];
  /** Whether settings.xml declared `w:trackChanges` (SPEC-039) — round-trips
   * onto DocumentModel.trackChangesRecording unchanged. */
  trackChangesRecording: boolean;
}

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

/** The literal text a modelled run child contributes to the run's `text` —
 * "\t" for w:tab, "\n" for w:br/w:cr, the hyphen characters for their
 * elements — or "" for the non-text children (w:rPr,
 * w:lastRenderedPageBreak). Shared by parseRun and the typed-break run
 * split in parseBodyParagraphBlocks (SPEC-028). */
function runChildText(child: Element): string {
  switch (child.localName) {
    case "t":
    case "delText":
      // w:delText (SPEC-039): a tracked-deletion run's text — Word's
      // substitute for w:t inside w:del, same content model otherwise.
      return child.textContent ?? "";
    case "tab":
      return "\t";
    case "br":
    case "cr":
      return "\n";
    case "noBreakHyphen":
      return "‑";
    case "softHyphen":
      return "­";
    default:
      return "";
  }
}

/** The block-level kind of a `w:br` (SPEC-028 AC-30): a page or column break
 * splits the enclosing paragraph into blocks; every other w:br (no w:type,
 * or "textWrapping") stays an inline "\n". Null for non-br elements. */
function typedBreakKind(el: Element): BreakKind | null {
  if (el.namespaceURI !== WORD_NAMESPACE || el.localName !== "br") {
    return null;
  }
  const type = wordAttr(el, "type");
  return type === "page" || type === "column" ? type : null;
}

function runHasTypedBreak(rEl: Element): boolean {
  return Array.from(rEl.children).some((child) => typedBreakKind(child) !== null);
}

/** Note: any `w:br` folds to "\n" here — parseRun serves the contexts where
 * a block-level break cannot exist (table cells, hyperlink children), so a
 * typed break degrades to a line break there. Body-level paragraphs go
 * through parseBodyParagraphBlocks, which splits around typed breaks
 * instead (SPEC-028 AC-30). */
function parseRun(rEl: Element, inheritedFormatting: RunFormatting): Run | undefined {
  const rPr = firstChildElement(rEl, WORD_NAMESPACE, "rPr");
  const formatting = mergeRunFormatting(inheritedFormatting, parseRunProperties(rPr));
  const formatChange = parseRunPropertiesChange(rPr);

  let text = "";
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI === WORD_NAMESPACE) {
      text += runChildText(child);
    }
  }

  if (text.length === 0) {
    return undefined;
  }
  return formatChange ? { text, formatting, changes: { formatChange } } : { text, formatting };
}

/** The run formatting a `w:hyperlink`'s child runs inherit: the wrapper's
 * target folded in as linkHref (external, resolved through the rels part)
 * or linkAnchor (w:anchor). An unresolvable or dangerous target leaves the
 * formatting unlinked — the text still reaches the document (AC-17). */
function hyperlinkFormatting(
  hyperlinkEl: Element,
  inherited: RunFormatting,
  links: LinkReadContext | undefined,
): RunFormatting {
  const relId = hyperlinkEl.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "id");
  if (relId && links) {
    const href = links.hrefForRelId(relId);
    if (href !== undefined) {
      return { ...inherited, linkHref: href };
    }
  }
  const anchor = wordAttr(hyperlinkEl, "anchor");
  if (anchor) {
    return { ...inherited, linkAnchor: anchor };
  }
  return inherited;
}

function textLengthOf(runs: Run[]): number {
  return runs.reduce((sum, run) => sum + run.text.length, 0);
}

/** Records a `w:bookmarkStart` as a model bookmark anchored at the current
 * text offset (SPEC-025). Word's transient `_GoBack` (regenerated on every
 * open, like proofErr) and a malformed nameless start are dropped.
 * `w:bookmarkEnd` is always dropped: the model holds point bookmarks — the
 * writer re-emits start+end adjacent — so a ranged source bookmark
 * collapses to its start position. */
function collectBookmark(child: Element, runs: Run[], bookmarks: BookmarkAnchor[]): void {
  if (child.localName !== "bookmarkStart") {
    return;
  }
  const name = wordAttr(child, "name");
  if (!name || name === "_GoBack") {
    return;
  }
  bookmarks.push({ name, offset: textLengthOf(runs) });
}

/** Records a `w:commentRangeStart`/`w:commentRangeEnd` as a model comment
 * anchor point at the current text offset (SPEC-038) — the boundary
 * counterpart of collectBookmark. A malformed id-less marker is dropped. */
function collectCommentAnchor(child: Element, runs: Run[], commentAnchors: CommentAnchor[]): void {
  const isStart = child.localName === "commentRangeStart";
  const isEnd = child.localName === "commentRangeEnd";
  if (!isStart && !isEnd) {
    return;
  }
  const id = wordAttr(child, "id");
  if (!id) {
    return;
  }
  commentAnchors.push({
    threadId: docxCommentId(id),
    boundary: isStart ? "start" : "end",
    offset: textLengthOf(runs),
  });
}

/** A tracked insertion/deletion wrapper's runs (SPEC-039): recurse with the
 * given `collect` callback (the caller's own collectRuns-shaped recursion,
 * since `parseBodyParagraphBlocks` and `collectRuns` each walk children
 * their own way), then stamp every run it produced with the wrapper's
 * author/date. w:del's runs use w:delText instead of w:t (handled by
 * runChildText) but are otherwise ordinary runs. */
function collectTrackedRun(child: Element, runs: Run[], collect: () => void): void {
  const kind = child.localName === "ins" ? "insert" : "delete";
  const mark = trackedChangeInfo(child, child.localName);
  const startIndex = runs.length;
  collect();
  for (let i = startIndex; i < runs.length; i++) {
    runs[i] = { ...runs[i], changes: { ...runs[i].changes, [kind]: mark } };
  }
}

function collectRuns(
  container: Element,
  inheritedFormatting: RunFormatting,
  runs: Run[],
  opaqueInlines: OpaqueInline[],
  bookmarks: BookmarkAnchor[],
  links?: LinkReadContext,
  commentAnchors: CommentAnchor[] = [],
): void {
  for (const child of Array.from(container.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "pPr") {
      continue; // Modelled separately by parseParagraph.
    }
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "r") {
      // A footnote/endnote reference run (SPEC-031) becomes a note marker
      // run linked to the parsed footnotes.xml/endnotes.xml content.
      const noteRef = noteReferenceOf(child);
      if (noteRef) {
        const rPr = firstChildElement(child, WORD_NAMESPACE, "rPr");
        const formatting = mergeRunFormatting(inheritedFormatting, parseRunProperties(rPr));
        runs.push(createNoteRun(noteRef.noteId, noteRef.kind, formatting));
        continue;
      }
      if (isCommentReferenceRun(child)) {
        // Dropped (SPEC-038): the range's own id already names its comment.
        continue;
      }
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
      // Hyperlinks (SPEC-025): the wrapper's target becomes link formatting
      // on its child runs; the wrapper element itself is not preserved.
      collectRuns(
        child,
        hyperlinkFormatting(child, inheritedFormatting, links),
        runs,
        opaqueInlines,
        bookmarks,
        links,
        commentAnchors,
      );
    } else if (
      child.namespaceURI === WORD_NAMESPACE &&
      (child.localName === "bookmarkStart" || child.localName === "bookmarkEnd")
    ) {
      collectBookmark(child, runs, bookmarks);
    } else if (
      child.namespaceURI === WORD_NAMESPACE &&
      (child.localName === "commentRangeStart" || child.localName === "commentRangeEnd")
    ) {
      collectCommentAnchor(child, runs, commentAnchors);
    } else if (child.namespaceURI === WORD_NAMESPACE && (child.localName === "ins" || child.localName === "del")) {
      collectTrackedRun(child, runs, () =>
        collectRuns(child, inheritedFormatting, runs, opaqueInlines, bookmarks, links, commentAnchors),
      );
    } else if (
      child.namespaceURI === WORD_NAMESPACE &&
      TRANSIENT_PARAGRAPH_CHILDREN.has(child.localName)
    ) {
      // Regenerated by Word on open; deliberately not preserved.
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "fldSimple") {
      // A PAGE/NUMPAGES field (SPEC-029 AC-25) becomes a live field run;
      // any other field (DATE, AUTHOR, TOC, …) is out of this spec's scope
      // and falls through to the opaque-inline path below, unchanged from
      // before this spec existed.
      const field = parseFieldSimple(child);
      if (field) {
        runs.push(field);
      } else {
        opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
      }
    } else {
      // Anything else (w:sdt, other namespaces) is an unmodelled inline
      // element: preserved verbatim, anchored to its position between the
      // runs (SPEC-008).
      opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
    }
  }
}

/** A paragraph's block-level metadata resolved from its `w:pPr` — the parts
 * shared by parseParagraph and parseBodyParagraphBlocks (which needs the
 * same resolution applied to every piece a mid-paragraph image splits off,
 * SPEC-024). */
interface ParagraphMeta {
  resolved: ReturnType<typeof resolveStyle>;
  formatting: Paragraph["formatting"];
  headingLevel: number | undefined;
  styleId: string | undefined;
  list: Paragraph["list"];
  changes: Paragraph["changes"];
}

/** A paragraph's own tracked-change state (SPEC-039): `w:pPrChange` for a
 * pending formatting change, and `w:pPr/w:rPr/w:ins`|`w:del` for the
 * paragraph mark itself (the pilcrow) being tracked-inserted/deleted —
 * Word nests these inside the mark's own run-properties element, distinct
 * from the paragraph's other `w:pPr` children. */
function parseParagraphChanges(pPr: Element | undefined): Paragraph["changes"] {
  if (!pPr) {
    return undefined;
  }
  const formatChange = parseParagraphPropertiesChange(pPr);
  const markRPr = firstChildElement(pPr, WORD_NAMESPACE, "rPr");
  const markInsertEl = markRPr ? firstChildElement(markRPr, WORD_NAMESPACE, "ins") : undefined;
  const markDeleteEl = markRPr ? firstChildElement(markRPr, WORD_NAMESPACE, "del") : undefined;
  if (!formatChange && !markInsertEl && !markDeleteEl) {
    return undefined;
  }
  const changes: Paragraph["changes"] = {};
  if (formatChange) {
    changes.formatChange = formatChange;
  }
  if (markInsertEl) {
    changes.markInsert = trackedChangeInfo(markInsertEl, "ins");
  }
  if (markDeleteEl) {
    changes.markDelete = trackedChangeInfo(markDeleteEl, "del");
  }
  return changes;
}

function parseParagraphMeta(
  pEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
): ParagraphMeta {
  const pPr = firstChildElement(pEl, WORD_NAMESPACE, "pPr");
  const pStyleEl = pPr ? firstChildElement(pPr, WORD_NAMESPACE, "pStyle") : undefined;
  const styleId = pStyleEl ? (wordAttr(pStyleEl, "val") ?? undefined) : undefined;

  const resolved = resolveStyle(styleCatalog, styleId);
  const formatting = mergeParagraphFormatting(resolved.paragraphFormatting, parseParagraphProperties(pPr));
  // "ltr" exists only to let a direct w:bidi="0" override a style's RTL
  // during the merge above; the model itself holds "rtl" or nothing.
  if (formatting.direction === "ltr") {
    delete formatting.direction;
  }
  // Same normalisation for the SPEC-028 pagination toggles: the parse keeps
  // the non-canonical values only so direct formatting can override a style
  // chain; the model stores keeps/pageBreakBefore only when on and
  // widowControl only when explicitly off (absent means on).
  if (formatting.keepNext === false) {
    delete formatting.keepNext;
  }
  if (formatting.keepLines === false) {
    delete formatting.keepLines;
  }
  if (formatting.pageBreakBefore === false) {
    delete formatting.pageBreakBefore;
  }
  if (formatting.widowControl === true) {
    delete formatting.widowControl;
  }
  const headingLevel = parseOutlineLevel(pPr) ?? resolved.headingLevel;

  // w:numPr resolves against the numbering catalog (SPEC-018 AC-20): the
  // referenced level definition decides type, marker, start and indent.
  let list: Paragraph["list"];
  const numPr = pPr ? firstChildElement(pPr, WORD_NAMESPACE, "numPr") : undefined;
  if (numPr) {
    const numIdEl = firstChildElement(numPr, WORD_NAMESPACE, "numId");
    const ilvlEl = firstChildElement(numPr, WORD_NAMESPACE, "ilvl");
    const numId = numIdEl ? wordAttr(numIdEl, "val") : null;
    const ilvl = Number((ilvlEl ? wordAttr(ilvlEl, "val") : null) ?? "0");
    if (numId !== null && Number.isInteger(ilvl) && ilvl >= 0) {
      list = numberingCatalog.listFormattingFor(numId, ilvl) ?? undefined;
    }
  }

  const changes = parseParagraphChanges(pPr);
  return { resolved, formatting, headingLevel, styleId, list, changes };
}

function assembleParagraph(
  meta: ParagraphMeta,
  runs: Run[],
  opaqueInlines: OpaqueInline[],
  includeList: boolean,
  bookmarks: BookmarkAnchor[] = [],
  commentAnchors: CommentAnchor[] = [],
): Paragraph {
  const paragraph: Paragraph = { type: "paragraph", runs, formatting: { ...meta.formatting } };
  if (includeList && meta.list) {
    paragraph.list = meta.list;
  }
  if (meta.styleId) {
    paragraph.styleId = meta.styleId;
  }
  if (meta.headingLevel !== undefined) {
    paragraph.headingLevel = meta.headingLevel;
  }
  if (opaqueInlines.length > 0) {
    paragraph.opaqueInlines = opaqueInlines;
  }
  if (bookmarks.length > 0) {
    paragraph.bookmarks = bookmarks;
  }
  if (commentAnchors.length > 0) {
    paragraph.commentAnchors = commentAnchors;
  }
  if (meta.changes) {
    paragraph.changes = meta.changes;
  }
  return paragraph;
}

function parseParagraph(
  pEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  links?: LinkReadContext,
): Paragraph {
  const meta = parseParagraphMeta(pEl, styleCatalog, numberingCatalog);
  const runs: Run[] = [];
  const opaqueInlines: OpaqueInline[] = [];
  const bookmarks: BookmarkAnchor[] = [];
  const commentAnchors: CommentAnchor[] = [];
  collectRuns(pEl, meta.resolved.runFormatting, runs, opaqueInlines, bookmarks, links, commentAnchors);
  return assembleParagraph(meta, runs, opaqueInlines, true, bookmarks, commentAnchors);
}

/** Context a body-level paragraph parse needs to resolve a `w:drawing`'s
 * media relationship into actual bytes (SPEC-024) — threaded down from
 * parseDocx, which already unzipped the package and read the main part's
 * relationships for the numbering/styles lookups. */
export interface DrawingContext {
  mainRels: Map<string, Relationship>;
  mainPartPath: string;
  parts: Map<string, Uint8Array>;
}

/** A `w:r` whose only content (besides `w:rPr`) is one `w:drawing` that
 * resolves to a modelled picture — the mid-paragraph split point
 * parseBodyParagraphBlocks looks for. Null for every other run, including
 * one holding a drawing this spec doesn't model (a chart, shape, ...), which
 * keeps falling through to the ordinary opaque-inline path. */
function imageRunDrawing(rEl: Element, ctx: DrawingContext): ParsedDrawing | null {
  const contentChildren = Array.from(rEl.children).filter(
    (child) => child.namespaceURI === WORD_NAMESPACE && child.localName !== "rPr",
  );
  if (contentChildren.length !== 1 || contentChildren[0].localName !== "drawing") {
    return null;
  }
  return parseDrawing(contentChildren[0], ctx.mainRels, ctx.mainPartPath, ctx.parts);
}

// --- table of contents (SPEC-033) -------------------------------------------

/** The `w:fldChar` type of a run, or null for a run that carries none
 * (SPEC-033's TOC field detection). */
function runFldCharType(rEl: Element): string | null {
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "fldChar") {
      return wordAttr(child, "fldCharType");
    }
  }
  return null;
}

function runInstrText(rEl: Element): string | null {
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "instrText") {
      return child.textContent ?? "";
    }
  }
  return null;
}

/** The TOC field options a paragraph's complex-field begin/instrText
 * sequence describes (SPEC-033 AC-19), or null when this paragraph does not
 * begin a TOC field. `showPageNumbers`/`hyperlinksOnly` are inferred from
 * whether entries actually carry a page-number trailer (see
 * tocEntryFromParagraph) rather than from the instruction text, since a
 * hand-authored Word field's switches vary more than this reader needs to
 * distinguish — what matters is what the cached result actually shows. */
function tocFieldStartOptions(pEl: Element): TocOptions | null {
  let sawBegin = false;
  let instr = "";
  for (const rEl of childElements(pEl, WORD_NAMESPACE, "r")) {
    const fldType = runFldCharType(rEl);
    if (fldType === "begin") {
      sawBegin = true;
      continue;
    }
    if (fldType === "separate" || fldType === "end") {
      break;
    }
    if (sawBegin) {
      instr += runInstrText(rEl) ?? "";
    }
  }
  if (!sawBegin || !/^\s*TOC\b/i.test(instr)) {
    return null;
  }
  const rangeMatch = /\\o\s*"(\d+)\s*-\s*(\d+)"/.exec(instr);
  const min = rangeMatch ? Math.max(1, Number(rangeMatch[1])) : 1;
  const max = rangeMatch ? Math.min(6, Number(rangeMatch[2])) : 3;
  const levels: number[] = [];
  for (let level = Math.min(min, max); level <= Math.max(min, max); level += 1) {
    levels.push(level);
  }
  return { levels: levels.length > 0 ? levels : [1], showPageNumbers: true, hyperlinksOnly: false };
}

/** Walks a paragraph's runs in document order — its own direct runs and, in
 * place, each hyperlink's runs — tracking complex-field nesting depth
 * (SPEC-033 AC-19, AC-20). A real Word TOC entry commonly caches its page
 * number as a NESTED `PAGEREF` complex field inside the entry's own
 * hyperlink; that nested field's `end` must not be mistaken for the outer
 * TOC field's own `end`, so `depth` only reaches 0 again — `endedHere` true
 * — when every nested field opened has also closed. `depth` carries across
 * paragraphs (the caller threads it through the whole field's paragraph
 * span), starting at 0 for the field's very first paragraph. */
function tocFieldDepthWalk(pEl: Element, depth: number): { depth: number; endedHere: boolean } {
  let d = depth;
  let endedHere = false;
  for (const child of Array.from(pEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    const runs =
      child.localName === "r"
        ? [child]
        : child.localName === "hyperlink"
          ? childElements(child, WORD_NAMESPACE, "r")
          : [];
    for (const rEl of runs) {
      const fldType = runFldCharType(rEl);
      if (fldType === "begin") {
        d += 1;
      } else if (fldType === "end") {
        d -= 1;
        if (d === 0) {
          endedHere = true;
        }
      }
    }
  }
  return { depth: d, endedHere };
}

/** The heading anchor, text and cached page number a TOC entry paragraph's
 * `w:hyperlink` carries (SPEC-033 AC-19): text before its `w:tab`, the page
 * number after — absent a tab, the entry is a hyperlink-only line (no page
 * number was cached for it). Returns null for a paragraph with no
 * hyperlink at all (the placeholder paragraph of an empty TOC). */
function tocEntryFromParagraph(pEl: Element): TocEntry | null {
  const hyperlink = firstChildElement(pEl, WORD_NAMESPACE, "hyperlink");
  if (!hyperlink) {
    return null;
  }
  const anchor = wordAttr(hyperlink, "anchor") ?? "";
  let text = "";
  let pageNumberText = "";
  let sawTab = false;
  for (const rEl of childElements(hyperlink, WORD_NAMESPACE, "r")) {
    for (const child of Array.from(rEl.children)) {
      if (child.namespaceURI !== WORD_NAMESPACE) {
        continue;
      }
      if (child.localName === "tab") {
        sawTab = true;
      } else if (child.localName === "t") {
        if (sawTab) {
          pageNumberText += child.textContent ?? "";
        } else {
          text += child.textContent ?? "";
        }
      }
    }
  }
  const pPr = firstChildElement(pEl, WORD_NAMESPACE, "pPr");
  const styleVal = pPr ? wordAttr(firstChildElement(pPr, WORD_NAMESPACE, "pStyle") ?? pPr, "val") : null;
  const levelMatch = styleVal ? /^TOC([1-6])$/.exec(styleVal) : null;
  const pageNumber = /^\d+$/.test(pageNumberText) ? Number(pageNumberText) : null;
  return {
    text,
    level: levelMatch ? Number(levelMatch[1]) : 1,
    anchor,
    pageNumber,
  };
}

/** Parses a TOC field's full paragraph run (SPEC-033 AC-19, AC-20): starting
 * at `startIndex` (already confirmed to begin a TOC field by
 * tocFieldStartOptions), consumes body paragraphs up to and including the
 * one whose field closes, folding their cached entries into one
 * TableOfContents block. Returns the block and the index of the last
 * paragraph it consumed, so the caller's loop can skip past it. A field
 * that never finds its own end (malformed input) consumes the rest of the
 * body — the same "better than losing the content" tradeoff the opaque-node
 * fallback elsewhere in this reader makes. */
function parseTocField(
  bodyChildren: Element[],
  startIndex: number,
  options: TocOptions,
): { toc: TableOfContents; endIndex: number } {
  const entries: TocEntry[] = [];
  let endIndex = startIndex;
  let depth = 0;
  for (let i = startIndex; i < bodyChildren.length; i += 1) {
    const pEl = bodyChildren[i];
    if (pEl.namespaceURI !== WORD_NAMESPACE || pEl.localName !== "p") {
      continue;
    }
    const entry = tocEntryFromParagraph(pEl);
    if (entry) {
      entries.push(entry);
    }
    endIndex = i;
    const walked = tocFieldDepthWalk(pEl, depth);
    depth = walked.depth;
    if (walked.endedHere) {
      break;
    }
  }
  const hasPageNumbers = entries.some((entry) => entry.pageNumber !== null);
  const toc: TableOfContents = {
    type: "toc",
    options: { ...options, showPageNumbers: hasPageNumbers, hyperlinksOnly: !hasPageNumbers },
    entries,
  };
  return { toc, endIndex };
}

/** Parses one body-level `<w:p>` into one or more blocks (SPEC-024, breaks
 * SPEC-028): a paragraph containing a modelled inline/floating picture
 * splits around each image — and one containing a typed `w:br` (page/column)
 * splits around the break the same way (AC-30) — since the document model
 * addresses both as their own top-level blocks (see documentModel.ts's
 * Image and BreakBlock doc comments) — mirroring
 * how a body-level `w:tbl` already breaks the block sequence, just with the
 * split happening here instead of being structural in the source XML. Every
 * resulting paragraph piece keeps the source paragraph's own formatting/
 * style/heading level; list membership (SPEC-018) is kept only on the first
 * piece, so an image inside a list item does not fork into two items.
 * Paragraphs with no modelled image parse exactly as parseParagraph always
 * did — this function is a pure superset, never a behaviour change for the
 * common case. Used only for section-level body content; table cells still
 * go through parseTableCell -> parseParagraph, since images are not modelled
 * inside cells. */
interface BodyParagraphResult {
  blocks: Block[];
  /** Media parts the split-off Image blocks reference, in document order —
   * the caller folds these into the document's overall media list. */
  media: MediaPart[];
}

function parseBodyParagraphBlocks(
  pEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  ctx: DrawingContext,
  links?: LinkReadContext,
): BodyParagraphResult {
  const runElements = Array.from(pEl.children).filter(
    (child) => child.namespaceURI === WORD_NAMESPACE && child.localName === "r",
  );
  const hasImageRun = runElements.some((r) => imageRunDrawing(r, ctx) !== null);
  // A typed w:br in an otherwise-modelled run splits the paragraph the same
  // way an image does (SPEC-028 AC-30); one inside an opaque run stays part
  // of that run's verbatim XML — splitting an unmodelled run could not
  // round-trip.
  const hasBreakRun = runElements.some(
    (r) => !runHasUnmodelledContent(r) && runHasTypedBreak(r),
  );
  if (!hasImageRun && !hasBreakRun) {
    return { blocks: [parseParagraph(pEl, styleCatalog, numberingCatalog, links)], media: [] };
  }

  const meta = parseParagraphMeta(pEl, styleCatalog, numberingCatalog);
  const blocks: Block[] = [];
  const media: MediaPart[] = [];
  let runs: Run[] = [];
  let opaqueInlines: OpaqueInline[] = [];
  let bookmarks: BookmarkAnchor[] = [];
  let commentAnchors: CommentAnchor[] = [];
  const flush = () => {
    if (runs.length === 0 && opaqueInlines.length === 0 && bookmarks.length === 0 && commentAnchors.length === 0) {
      return;
    }
    blocks.push(
      assembleParagraph(meta, runs, opaqueInlines, blocks.length === 0, bookmarks, commentAnchors),
    );
    runs = [];
    opaqueInlines = [];
    bookmarks = [];
    commentAnchors = [];
  };
  for (const child of Array.from(pEl.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "pPr") {
      continue;
    }
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "r") {
      const parsed = imageRunDrawing(child, ctx);
      if (parsed) {
        flush();
        blocks.push(parsed.image);
        media.push(parsed.media);
        continue;
      }
      if (isCommentReferenceRun(child)) {
        continue;
      }
      if (runHasUnmodelledContent(child)) {
        opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
        continue;
      }
      if (runHasTypedBreak(child)) {
        // Split the run around each typed break (SPEC-028 AC-30): the text
        // before it ends the current paragraph piece, the break becomes a
        // BreakBlock, the text after starts the next piece. A paragraph
        // whose only content is the break yields just the BreakBlock — the
        // writer's own break shape reads back as a fixed point.
        const rPrEl = firstChildElement(child, WORD_NAMESPACE, "rPr");
        const breakRunFormatting = mergeRunFormatting(
          meta.resolved.runFormatting,
          parseRunProperties(rPrEl),
        );
        let text = "";
        const pushText = () => {
          if (text.length > 0) {
            runs.push({ text, formatting: breakRunFormatting });
            text = "";
          }
        };
        for (const runChild of Array.from(child.children)) {
          if (runChild.namespaceURI !== WORD_NAMESPACE) {
            continue;
          }
          const kind = typedBreakKind(runChild);
          if (kind !== null) {
            pushText();
            flush();
            blocks.push({ type: "break", kind });
          } else {
            text += runChildText(runChild);
          }
        }
        pushText();
        continue;
      }
      const run = parseRun(child, meta.resolved.runFormatting);
      if (run) {
        runs.push(run);
      }
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "hyperlink") {
      collectRuns(
        child,
        hyperlinkFormatting(child, meta.resolved.runFormatting, links),
        runs,
        opaqueInlines,
        bookmarks,
        links,
        commentAnchors,
      );
    } else if (
      child.namespaceURI === WORD_NAMESPACE &&
      (child.localName === "bookmarkStart" || child.localName === "bookmarkEnd")
    ) {
      collectBookmark(child, runs, bookmarks);
    } else if (
      child.namespaceURI === WORD_NAMESPACE &&
      (child.localName === "commentRangeStart" || child.localName === "commentRangeEnd")
    ) {
      collectCommentAnchor(child, runs, commentAnchors);
    } else if (child.namespaceURI === WORD_NAMESPACE && (child.localName === "ins" || child.localName === "del")) {
      collectTrackedRun(child, runs, () =>
        collectRuns(child, meta.resolved.runFormatting, runs, opaqueInlines, bookmarks, links, commentAnchors),
      );
    } else if (child.namespaceURI === WORD_NAMESPACE && TRANSIENT_PARAGRAPH_CHILDREN.has(child.localName)) {
      // Regenerated by Word on open; deliberately not preserved.
    } else {
      opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
    }
  }
  flush();
  return { blocks, media };
}

/** OOXML `w:val`s for `w:tcBorders`' children collapse to the model's small
 * enum (SPEC-023): the handful of dashed variants all read as "dashed",
 * "nil"/"none" as "none", and anything else (including "thick", which the
 * model has no separate weight class for) as a plain single line. */
function mapBorderLineStyle(val: string): BorderLineStyle {
  const v = val.toLowerCase();
  if (v === "nil" || v === "none") {
    return "none";
  }
  if (v === "double" || v === "triple") {
    return "double";
  }
  if (v === "dotted") {
    return "dotted";
  }
  if (v.startsWith("dash") || v.includes("dotdash") || v.includes("dotdotdash")) {
    return "dashed";
  }
  return "single";
}

function parseBorderSide(el: Element | undefined): BorderSide | undefined {
  if (!el) {
    return undefined;
  }
  const val = wordAttr(el, "val");
  if (!val) {
    return undefined;
  }
  const style = mapBorderLineStyle(val);
  const sz = Number(wordAttr(el, "sz") ?? "");
  const widthPt = style !== "none" && Number.isFinite(sz) && sz > 0 ? eighthPointsToPt(sz) : 0;
  const color = ooxmlColorToCss(wordAttr(el, "color")) ?? "auto";
  return { style, widthPt, color };
}

/** A cell's `w:tcBorders` (SPEC-023 AC-10..12, AC-24). OOXML's logical
 * `w:start`/`w:end` are read (falling back to the legacy `w:left`/`w:right`
 * some producers still emit) and stored as the model's physical left/right —
 * table RTL is out of scope, so the two coincide. */
function parseCellBorders(tcPr: Element): CellBorders | undefined {
  const bordersEl = firstChildElement(tcPr, WORD_NAMESPACE, "tcBorders");
  if (!bordersEl) {
    return undefined;
  }
  const top = parseBorderSide(firstChildElement(bordersEl, WORD_NAMESPACE, "top"));
  const bottom = parseBorderSide(firstChildElement(bordersEl, WORD_NAMESPACE, "bottom"));
  const left = parseBorderSide(
    firstChildElement(bordersEl, WORD_NAMESPACE, "start") ??
      firstChildElement(bordersEl, WORD_NAMESPACE, "left"),
  );
  const right = parseBorderSide(
    firstChildElement(bordersEl, WORD_NAMESPACE, "end") ?? firstChildElement(bordersEl, WORD_NAMESPACE, "right"),
  );
  if (!top && !bottom && !left && !right) {
    return undefined;
  }
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
  return borders;
}

/** A cell's `w:shd` fill as CSS (SPEC-023 AC-15, AC-26): `w:fill` is the
 * background colour; `w:color` is the pattern colour on top of it, which the
 * model has no slot for and drops (the same simplification run shading
 * makes, SPEC-013). */
function parseCellShading(tcPr: Element): string | undefined {
  const shd = firstChildElement(tcPr, WORD_NAMESPACE, "shd");
  return shd ? ooxmlColorToCss(wordAttr(shd, "fill")) : undefined;
}

/** A cell's `w:tcMar` (SPEC-023 AC-16): the model holds one uniform padding,
 * so the top margin stands in for all four — same-format round trips are
 * exact because the writer re-emits all four equal to it. */
function parseCellPadding(tcPr: Element): number | undefined {
  const mar = firstChildElement(tcPr, WORD_NAMESPACE, "tcMar");
  const top = mar ? firstChildElement(mar, WORD_NAMESPACE, "top") : undefined;
  const w = top ? Number(wordAttr(top, "w") ?? "") : NaN;
  return Number.isFinite(w) && w >= 0 ? twipsToCm(w) : undefined;
}

function parseCellVerticalAlign(tcPr: Element): CellVerticalAlign | undefined {
  const el = firstChildElement(tcPr, WORD_NAMESPACE, "vAlign");
  const val = el ? wordAttr(el, "val") : null;
  return val === "top" || val === "center" || val === "bottom" ? val : undefined;
}

function parseGridSpan(tcPr: Element): number | undefined {
  const el = firstChildElement(tcPr, WORD_NAMESPACE, "gridSpan");
  const val = el ? Number(wordAttr(el, "val")) : NaN;
  return Number.isInteger(val) && val > 1 ? val : undefined;
}

/** A cell's `w:vMerge`: absent (not vertically merged), "restart" (the
 * origin of a new vertical merge — carries its own content) or "continue"
 * (covered by the origin above; an explicit `w:val` is required for
 * "restart", its absence defaults to "continue" per the schema). */
type VMergeKind = "restart" | "continue" | null;

function parseVMerge(tcPr: Element): VMergeKind {
  const el = firstChildElement(tcPr, WORD_NAMESPACE, "vMerge");
  if (!el) {
    return null;
  }
  return wordAttr(el, "val") === "restart" ? "restart" : "continue";
}

/** One `w:tc`'s parse result: the cell itself plus the two attributes
 * (SPEC-023) that decide how many grid slots it and its followers occupy —
 * kept out of the TableCell union so pass 2 below can resolve vMerge chains
 * before the row/table shape is finalised. */
interface ParsedCellSlot {
  cell: TableCell;
  colSpan: number;
  vMerge: VMergeKind;
}

function parseTableCell(
  tcEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  links?: LinkReadContext,
): ParsedCellSlot | null {
  const blocks: Paragraph[] = [];
  let tcPrEl: Element | undefined;
  let tcPrXml: string | undefined;
  for (const child of Array.from(tcEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      return null;
    }
    if (child.localName === "tcPr") {
      tcPrEl = child;
      tcPrXml = serializeElement(child);
    } else if (child.localName === "p") {
      blocks.push(parseParagraph(child, styleCatalog, numberingCatalog, links));
    } else {
      // Nested tables, block-level sdt, … — out of the model's table shape.
      return null;
    }
  }
  if (blocks.length === 0) {
    // OOXML requires a w:p per cell; tolerate its absence with an empty one
    // so every cell can hold a caret.
    blocks.push({ type: "paragraph", runs: [], formatting: {} });
  }
  const cell: TableCell = { blocks };
  const colSpan = tcPrEl ? parseGridSpan(tcPrEl) : undefined;
  if (colSpan !== undefined) {
    cell.colSpan = colSpan;
  }
  const borders = tcPrEl ? parseCellBorders(tcPrEl) : undefined;
  if (borders) {
    cell.borders = borders;
  }
  const shadingFill = tcPrEl ? parseCellShading(tcPrEl) : undefined;
  if (shadingFill !== undefined) {
    cell.shadingFill = shadingFill;
  }
  const paddingCm = tcPrEl ? parseCellPadding(tcPrEl) : undefined;
  if (paddingCm !== undefined) {
    cell.paddingCm = paddingCm;
  }
  const verticalAlign = tcPrEl ? parseCellVerticalAlign(tcPrEl) : undefined;
  if (verticalAlign !== undefined) {
    cell.verticalAlign = verticalAlign;
  }
  if (tcPrXml !== undefined) {
    cell.docxTcPr = tcPrXml;
  }
  return { cell, colSpan: colSpan ?? 1, vMerge: tcPrEl ? parseVMerge(tcPrEl) : null };
}

/** A `w:tc` occupies one grid column normally, or `colSpan` of them when it
 * carries `w:gridSpan` — the extra columns have no `w:tc` of their own, so
 * they become covered placeholder slots here to keep the model's per-row
 * cell count equal to the table's grid column count (SPEC-023). */
type RowSlot = { kind: "cell"; parsed: ParsedCellSlot } | { kind: "covered" };

function parseTableRowSlots(
  trEl: Element,
  gridColumns: number,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  links?: LinkReadContext,
): { slots: RowSlot[]; trPrEl?: Element; trPr?: string } | null {
  const slots: RowSlot[] = [];
  let trPrEl: Element | undefined;
  let trPr: string | undefined;
  for (const child of Array.from(trEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      return null;
    }
    if (child.localName === "trPr") {
      trPrEl = child;
      trPr = serializeElement(child);
    } else if (child.localName === "tc") {
      const parsed = parseTableCell(child, styleCatalog, numberingCatalog, links);
      if (!parsed) {
        return null;
      }
      slots.push({ kind: "cell", parsed });
      for (let i = 1; i < parsed.colSpan; i++) {
        slots.push({ kind: "covered" });
      }
    } else {
      // tblPrEx, bookmarks between cells, … — keep the table verbatim.
      return null;
    }
  }
  if (slots.length === 0 || slots.length !== gridColumns) {
    return null;
  }
  return { slots, trPrEl, trPr };
}

/** A row's minimum height from `w:trHeight` (SPEC-023 AC-4, AC-5): both
 * `hRule="atLeast"` (an explicit minimum) and its absence (Word's own
 * default meaning) map to the model's `minHeightCm`, which is always a
 * floor. `hRule="exact"` (a fixed height, not modelled) is read the same
 * way — an approximation, but a closer one than dropping the height. */
function parseRowMinHeight(trPr: Element | undefined): number | undefined {
  const heightEl = trPr ? firstChildElement(trPr, WORD_NAMESPACE, "trHeight") : undefined;
  const val = heightEl ? Number(wordAttr(heightEl, "val") ?? "") : NaN;
  return Number.isFinite(val) && val > 0 ? twipsToCm(val) : undefined;
}

/** A table's `w:tblW` (SPEC-023 AC-23, AC-27): `pct` is fiftieths of a
 * percent (5000 = 100%), `auto` sizes to content, anything else (including
 * no `w:tblW` at all) is a fixed width — the `columnWidthsCm` sum. */
function parseTableWidth(tblPr: Element | undefined): { widthMode?: TableWidthMode; widthPercent?: number } {
  const tblW = tblPr ? firstChildElement(tblPr, WORD_NAMESPACE, "tblW") : undefined;
  if (!tblW) {
    return {};
  }
  const type = wordAttr(tblW, "type");
  if (type === "pct") {
    const w = Number(wordAttr(tblW, "w") ?? "");
    if (Number.isFinite(w) && w > 0) {
      return { widthMode: "percent", widthPercent: Math.min(100, Math.max(1, Math.round(w / 50))) };
    }
  }
  if (type === "auto") {
    return { widthMode: "auto" };
  }
  return { widthMode: "fixed" };
}

/** A table's `w:jc` (SPEC-023 AC-22): "start" is the logical form of
 * "left" — table RTL is out of scope, so both map to the model's "left". */
function parseTableAlignment(tblPr: Element | undefined): TableAlignment | undefined {
  const jc = tblPr ? firstChildElement(tblPr, WORD_NAMESPACE, "jc") : undefined;
  const val = jc ? wordAttr(jc, "val") : null;
  if (val === "center" || val === "right") {
    return val;
  }
  if (val === "left" || val === "start") {
    return "left";
  }
  return undefined;
}

/** Parses a `w:tbl` into a modelled Table (SPEC-022; merged cells and
 * formatting SPEC-023), or null when the table uses features outside the
 * model's regular grid — nested tables, a grid that doesn't match the rows —
 * in which case the caller preserves the element verbatim as an opaque
 * block instead (SPEC-008: fidelity over feature count).
 *
 * Merged cells keep the grid regular for indexing (every row still has
 * `columnWidthsCm.length` cells): a `w:gridSpan` cell's extra columns and a
 * `w:vMerge="continue"` cell both become `covered: true` placeholders, and
 * the real cell they are covered by carries `colSpan`/`rowSpan`. Resolving
 * `vMerge` needs a second pass over the whole table (a "continue" cell's
 * `rowSpan` accumulates on whichever earlier row most recently started a
 * chain at that column), so rows are parsed into slots first and only
 * assembled into the model's `TableRow`s afterwards. */
function parseTable(
  tblEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  links?: LinkReadContext,
): Table | null {
  const columnWidthsCm: number[] = [];
  let tblPrEl: Element | undefined;
  let tblPrXml: string | undefined;
  let sawGrid = false;
  const rowSlots: Array<{ slots: RowSlot[]; trPrEl?: Element; trPr?: string }> = [];
  for (const child of Array.from(tblEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      return null;
    }
    if (child.localName === "tblPr") {
      tblPrEl = child;
      tblPrXml = serializeElement(child);
    } else if (child.localName === "tblGrid") {
      sawGrid = true;
      for (const col of Array.from(child.children)) {
        if (col.namespaceURI !== WORD_NAMESPACE || col.localName !== "gridCol") {
          return null;
        }
        const w = Number(wordAttr(col, "w") ?? "");
        if (!Number.isFinite(w) || w <= 0) {
          return null;
        }
        columnWidthsCm.push(twipsToCm(w));
      }
    } else if (child.localName === "tr") {
      if (!sawGrid) {
        return null;
      }
      const parsedRow = parseTableRowSlots(child, columnWidthsCm.length, styleCatalog, numberingCatalog, links);
      if (!parsedRow) {
        return null;
      }
      rowSlots.push(parsedRow);
    } else {
      return null;
    }
  }
  if (!sawGrid || rowSlots.length === 0 || columnWidthsCm.length === 0) {
    return null;
  }

  const rows: TableRow[] = rowSlots.map(({ slots, trPrEl, trPr }) => {
    const cells: TableCell[] = slots.map((slot) =>
      slot.kind === "cell" ? slot.parsed.cell : { blocks: [], covered: true },
    );
    const row: TableRow = { cells };
    if (trPr !== undefined) {
      row.docxTrPr = trPr;
    }
    const minHeightCm = parseRowMinHeight(trPrEl);
    if (minHeightCm !== undefined) {
      row.minHeightCm = minHeightCm;
    }
    // Tracked row insert/delete (SPEC-039 AC-15): w:trPr/w:ins|w:del.
    const insEl = trPrEl ? firstChildElement(trPrEl, WORD_NAMESPACE, "ins") : undefined;
    const delEl = trPrEl ? firstChildElement(trPrEl, WORD_NAMESPACE, "del") : undefined;
    if (insEl || delEl) {
      row.changes = {};
      if (insEl) {
        row.changes.insert = trackedChangeInfo(insEl, "ins");
      }
      if (delEl) {
        row.changes.delete = trackedChangeInfo(delEl, "del");
      }
    }
    return row;
  });

  // Pass 2: a "continue" cell's rowSpan accumulates onto whichever row most
  // recently started (or continued) a merge at that column.
  const openOrigin: Array<number | null> = Array.from({ length: columnWidthsCm.length }, () => null);
  for (let r = 0; r < rowSlots.length; r++) {
    const slots = rowSlots[r].slots;
    for (let col = 0; col < slots.length; col++) {
      const slot = slots[col];
      if (slot.kind !== "cell") {
        continue;
      }
      if (slot.parsed.vMerge === "continue" && openOrigin[col] !== null) {
        const originRow = openOrigin[col] as number;
        const originCell = rows[originRow].cells[col];
        originCell.rowSpan = (originCell.rowSpan ?? 1) + 1;
        rows[r].cells[col] = { blocks: [], covered: true };
      } else {
        openOrigin[col] = r;
      }
    }
  }

  const table: Table = { type: "table", rows, columnWidthsCm };
  if (tblPrXml !== undefined) {
    table.docxTblPr = tblPrXml;
  }
  const width = parseTableWidth(tblPrEl);
  if (width.widthMode !== undefined) {
    table.widthMode = width.widthMode;
  }
  if (width.widthPercent !== undefined) {
    table.widthPercent = width.widthPercent;
  }
  const alignment = parseTableAlignment(tblPrEl);
  if (alignment !== undefined) {
    table.alignment = alignment;
  }
  return table;
}

/** A sectPr's `w:type` as the model's SectionStart (SPEC-028 AC-31): only
 * "continuous" is stored — an absent `start` already means "nextPage", and
 * the odd/even-page variants collapse to it (unmodelled distinctions). */
function parseSectionStart(sectPr: Element | undefined): SectionStart | undefined {
  const typeEl = sectPr ? firstChildElement(sectPr, WORD_NAMESPACE, "type") : undefined;
  return typeEl && wordAttr(typeEl, "val") === "continuous" ? "continuous" : undefined;
}

function boolAttr(value: string | null | undefined): boolean {
  return value === "1" || value === "true" || value === "on";
}

/** A sectPr's `w:cols` as the model's multi-column layout (SPEC-030
 * AC-16/AC-16a): `w:num`/`w:space`/`w:sep`, plus explicit `w:col` children
 * (each `w:w` a column width, `w:space` its trailing gutter) for unequal
 * columns. A single column (num ≤ 1) or an absent `w:cols` is left unmodelled
 * (an ordinary single-column section). */
function parseColumns(sectPr: Element | undefined): SectionColumns | undefined {
  const cols = sectPr ? firstChildElement(sectPr, WORD_NAMESPACE, "cols") : undefined;
  if (!cols) {
    return undefined;
  }
  const num = Number(wordAttr(cols, "num") ?? "1");
  const count = Number.isFinite(num) ? Math.max(1, Math.min(MAX_COLUMN_COUNT, Math.round(num))) : 1;
  if (count <= 1) {
    return undefined;
  }
  const spaceTwips = Number(wordAttr(cols, "space") ?? "");
  const gapCm = Number.isFinite(spaceTwips) ? twipsToCm(spaceTwips) : DEFAULT_COLUMN_GAP_CM;
  const separator = boolAttr(wordAttr(cols, "sep"));
  const colChildren = childElements(cols, WORD_NAMESPACE, "col");
  if (colChildren.length === count) {
    const widthsCm = colChildren.map((col) => {
      const w = Number(wordAttr(col, "w") ?? "");
      return Number.isFinite(w) ? twipsToCm(w) : 0;
    });
    if (widthsCm.every((w) => w > 0)) {
      const firstGap = Number(wordAttr(colChildren[0], "space") ?? "");
      const gutter = Number.isFinite(firstGap) ? twipsToCm(firstGap) : gapCm;
      return { count, equalWidth: false, gapCm: gutter, separator, widthsCm };
    }
  }
  return { count, equalWidth: true, gapCm, separator };
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

// --- headers, footers and page numbers (SPEC-029) ---------------------------

const SETTINGS_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";

const HEADER_FOOTER_KIND_BY_TYPE: Record<string, HeaderFooterKind> = {
  default: "default",
  first: "first",
  even: "even",
};

/** A PAGE/NUMPAGES field instruction's kind and numbering format (SPEC-029
 * AC-16..18, AC-25) — `\* roman`/`\* alphabetic` switches map to the
 * dialog's lowercase formats (case folded: this app never emits the
 * uppercase variants, and reads either case back the same way); anything
 * else (no switch, `\* ARABIC`, `\* MERGEFORMAT`, …) is "decimal". Null for
 * an instruction that is not a page-number field — a field this spec is
 * explicitly not scoped to (DATE, AUTHOR, TOC, …), which falls through to
 * the ordinary opaque-inline path exactly as before this spec existed. */
function fieldInstrParts(instr: string): { kind: FieldKind; format: PageNumberFormat } | null {
  const match = /^\s*(PAGE|NUMPAGES)\b([\s\S]*)$/i.exec(instr);
  if (!match) {
    return null;
  }
  const kind: FieldKind = match[1].toUpperCase() === "PAGE" ? "page" : "numpages";
  const switchMatch = /\\\*\s*(roman|alphabetic|arabic)/i.exec(match[2]);
  let format: PageNumberFormat = "decimal";
  if (switchMatch) {
    const sw = switchMatch[1].toLowerCase();
    if (sw === "roman") {
      format = "lowerRoman";
    } else if (sw === "alphabetic") {
      format = "lowerAlpha";
    }
  }
  return { kind, format };
}

/** A `w:fldSimple` as a page-number/page-count field run (SPEC-029), or
 * null when its instruction is something else — the caller falls back to
 * preserving it as an opaque inline the same way it always has. */
function parseFieldSimple(el: Element): Run | null {
  const instr = wordAttr(el, "instr");
  if (!instr) {
    return null;
  }
  const parsed = fieldInstrParts(instr);
  return parsed ? createFieldRun(parsed.kind, parsed.format) : null;
}

interface HeaderFooterReference {
  part: HeaderFooterPart;
  kind: HeaderFooterKind;
  relId: string;
}

/** A section's `w:headerReference`/`w:footerReference` elements (SPEC-029
 * AC-24): each names which variant (`w:type`) and the relationship id its
 * content part resolves through. A reference with no `r:id` (malformed) is
 * skipped rather than crashing the open. */
function parseHeaderFooterReferences(sectPr: Element | undefined): HeaderFooterReference[] {
  if (!sectPr) {
    return [];
  }
  const refs: HeaderFooterReference[] = [];
  for (const child of Array.from(sectPr.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    const isHeader = child.localName === "headerReference";
    const isFooter = child.localName === "footerReference";
    if (!isHeader && !isFooter) {
      continue;
    }
    const relId = child.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "id");
    if (!relId) {
      continue;
    }
    const kind = HEADER_FOOTER_KIND_BY_TYPE[wordAttr(child, "type") ?? "default"] ?? "default";
    refs.push({ part: isHeader ? "header" : "footer", kind, relId });
  }
  return refs;
}

/** "Different first page" (SPEC-029 AC-11): `w:titlePg`'s mere presence. */
function parseTitlePg(sectPr: Element | undefined): boolean {
  return !!sectPr && !!firstChildElement(sectPr, WORD_NAMESPACE, "titlePg");
}

/** "Start numbering at" (SPEC-029 AC-22): `w:pgNumType/@w:start`, when it is
 * a valid non-negative integer. */
function parsePageNumberStart(sectPr: Element | undefined): number | undefined {
  const pgNumType = sectPr ? firstChildElement(sectPr, WORD_NAMESPACE, "pgNumType") : undefined;
  const start = pgNumType ? wordAttr(pgNumType, "start") : null;
  if (start === null) {
    return undefined;
  }
  const parsed = Number(start);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

/** One header/footer part's content (SPEC-029 AC-3..6, AC-24): its own
 * paragraphs/tables, parsed the same way body content is — reusing
 * parseBodyParagraphBlocks/parseTable rather than a second copy of run/table
 * parsing — resolved against the PART'S OWN relationships (a header's
 * `r:embed`/hyperlink targets are relative to `word/_rels/header1.xml.rels`,
 * never the main document's). Manual breaks are dropped: a header/footer
 * never paginates, so a stray `w:br type="page"` from a hand-edited file
 * would otherwise become a dead, unremovable block. */
function parseHeaderFooterContent(
  partPath: string,
  parts: Map<string, Uint8Array>,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
): { blocks: Block[]; media: MediaPart[] } {
  const raw = parts.get(partPath);
  const fallback = { blocks: [{ type: "paragraph" as const, runs: [] }], media: [] };
  if (!raw) {
    return fallback;
  }
  let root: Element;
  try {
    root = parseXml(decodeUtf8(raw)).documentElement;
  } catch (error) {
    if (error instanceof XmlParseError) {
      return fallback;
    }
    throw error;
  }
  const relsBytes = parts.get(relsPathFor(partPath));
  const ownRels = relsBytes ? parseRelationships(decodeUtf8(relsBytes)) : new Map<string, Relationship>();
  const ctx: DrawingContext = { mainRels: ownRels, mainPartPath: partPath, parts };
  const linkCtx: LinkReadContext = {
    hrefForRelId: (relId) => {
      const rel = ownRels.get(relId);
      if (!rel || rel.type !== HYPERLINK_RELATIONSHIP) {
        return undefined;
      }
      return sanitizeHref(rel.target) ?? undefined;
    },
  };
  const blocks: Block[] = [];
  const media: MediaPart[] = [];
  for (const child of Array.from(root.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "p") {
      const result = parseBodyParagraphBlocks(child, styleCatalog, numberingCatalog, ctx, linkCtx);
      for (const block of result.blocks) {
        if (block.type !== "break") {
          blocks.push(block);
        }
      }
      media.push(...result.media);
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "tbl") {
      const table = parseTable(child, styleCatalog, numberingCatalog, linkCtx);
      if (table) {
        blocks.push(table);
      }
    }
  }
  return blocks.length > 0 ? { blocks, media } : fallback;
}

// --- footnotes and endnotes (SPEC-031) ---------------------------------------

const FOOTNOTES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const ENDNOTES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";

const NOTE_NUM_FMT_BY_VAL: Record<string, NoteNumberFormat> = {
  decimal: "decimal",
  lowerRoman: "lowerRoman",
  lowerLetter: "lowerAlpha",
  chicago: "symbol",
};

const NOTE_RESTART_BY_VAL: Record<string, NoteRestart> = {
  eachPage: "page",
  eachSect: "section",
};

/** A sectPr's `w:footnotePr`/`w:endnotePr` as model note settings (SPEC-031
 * AC-17..20), or undefined when absent or carrying nothing the model
 * expresses. `continuous` restart maps to absent — the model default. */
function parseNoteSettings(
  sectPr: Element | undefined,
  tag: "footnotePr" | "endnotePr",
): NoteSettings | undefined {
  const pr = sectPr ? firstChildElement(sectPr, WORD_NAMESPACE, tag) : undefined;
  if (!pr) {
    return undefined;
  }
  const settings: NoteSettings = {};
  const numFmt = firstChildElement(pr, WORD_NAMESPACE, "numFmt");
  const format = numFmt ? NOTE_NUM_FMT_BY_VAL[wordAttr(numFmt, "val") ?? ""] : undefined;
  if (format !== undefined) {
    settings.numberFormat = format;
  }
  const numStart = firstChildElement(pr, WORD_NAMESPACE, "numStart");
  const start = numStart ? Number(wordAttr(numStart, "val")) : NaN;
  if (Number.isInteger(start) && start >= 1) {
    settings.startAt = start;
  }
  const numRestart = firstChildElement(pr, WORD_NAMESPACE, "numRestart");
  const restart = numRestart ? NOTE_RESTART_BY_VAL[wordAttr(numRestart, "val") ?? ""] : undefined;
  if (restart !== undefined) {
    settings.restart = restart;
  }
  return Object.keys(settings).length > 0 ? settings : undefined;
}

/** One footnotes.xml/endnotes.xml part's real notes (SPEC-031 AC-21/22):
 * each `w:footnote`/`w:endnote` with no `w:type` (separator and
 * continuation entries carry one) becomes a NoteContent whose blocks parse
 * exactly like body content — through the part's own relationships, the
 * parseHeaderFooterContent pattern. The `w:footnoteRef`/`w:endnoteRef`
 * marker elements inside the note text are Word's rendering of the note's
 * own number and are dropped: the model derives numbers from mark order,
 * and the writer re-emits a fresh marker run on save. */
function parseNotesPart(
  partPath: string,
  parts: Map<string, Uint8Array>,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  kind: NoteKind,
): { notes: NoteContent[]; media: MediaPart[] } {
  const raw = parts.get(partPath);
  if (!raw) {
    return { notes: [], media: [] };
  }
  let root: Element;
  try {
    root = parseXml(decodeUtf8(raw)).documentElement;
  } catch (error) {
    if (error instanceof XmlParseError) {
      return { notes: [], media: [] };
    }
    throw error;
  }
  const relsBytes = parts.get(relsPathFor(partPath));
  const ownRels = relsBytes
    ? parseRelationships(decodeUtf8(relsBytes))
    : new Map<string, Relationship>();
  const ctx: DrawingContext = { mainRels: ownRels, mainPartPath: partPath, parts };
  const linkCtx: LinkReadContext = {
    hrefForRelId: (relId) => {
      const rel = ownRels.get(relId);
      if (!rel || rel.type !== HYPERLINK_RELATIONSHIP) {
        return undefined;
      }
      return sanitizeHref(rel.target) ?? undefined;
    },
  };
  const tag = kind === "footnote" ? "footnote" : "endnote";
  const refTag = kind === "footnote" ? "footnoteRef" : "endnoteRef";
  const notes: NoteContent[] = [];
  const media: MediaPart[] = [];
  for (const noteEl of childElements(root, WORD_NAMESPACE, tag)) {
    if (wordAttr(noteEl, "type") !== null) {
      continue; // separator / continuationSeparator / continuationNotice
    }
    const sourceId = wordAttr(noteEl, "id");
    if (sourceId === null) {
      continue;
    }
    for (const marker of Array.from(noteEl.getElementsByTagNameNS(WORD_NAMESPACE, refTag))) {
      marker.parentElement?.removeChild(marker);
    }
    const blocks: Block[] = [];
    for (const child of Array.from(noteEl.children)) {
      if (child.namespaceURI === WORD_NAMESPACE && child.localName === "p") {
        const result = parseBodyParagraphBlocks(child, styleCatalog, numberingCatalog, ctx, linkCtx);
        for (const block of result.blocks) {
          if (block.type !== "break") {
            blocks.push(block);
          }
        }
        media.push(...result.media);
      } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "tbl") {
        const table = parseTable(child, styleCatalog, numberingCatalog, linkCtx);
        if (table) {
          blocks.push(table);
        }
      }
    }
    notes.push({
      noteId: docxNoteId(kind, sourceId),
      kind,
      blocks: blocks.length > 0 ? blocks : [{ type: "paragraph", runs: [] }],
    });
  }
  return { notes, media };
}

// --- comments (SPEC-038) -----------------------------------------------------

const COMMENTS_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const UNKNOWN_COMMENT_AUTHOR = "Unknown";

interface ParsedComment {
  threadId: string;
  messages: CommentMessage[];
}

/** word/comments.xml's comments (SPEC-038, "Opening a Word document with
 * comments shows every comment, its author and its date"): each
 * `w:comment` becomes one thread, keyed by its numeric id — matched to the
 * body's commentRangeStart/End via docxCommentId. Reverses writer.ts's
 * commentsPartXml encoding: the first paragraph is the opening message's
 * plain text, later paragraphs are replies encoded as
 * "{author} ({createdAt}): {text}". A comment with no paragraphs still
 * becomes a thread with one empty-text message, so an empty comment
 * round-trips rather than vanishing. */
function parseCommentsPart(
  partPath: string,
  parts: Map<string, Uint8Array>,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
): ParsedComment[] {
  const raw = parts.get(partPath);
  if (!raw) {
    return [];
  }
  let root: Element;
  try {
    root = parseXml(decodeUtf8(raw)).documentElement;
  } catch (error) {
    if (error instanceof XmlParseError) {
      return [];
    }
    throw error;
  }
  const comments: ParsedComment[] = [];
  for (const commentEl of childElements(root, WORD_NAMESPACE, "comment")) {
    const sourceId = wordAttr(commentEl, "id");
    if (sourceId === null) {
      continue;
    }
    const author = wordAttr(commentEl, "author") || UNKNOWN_COMMENT_AUTHOR;
    const createdAt = wordAttr(commentEl, "date") || new Date(0).toISOString();
    const texts = childElements(commentEl, WORD_NAMESPACE, "p").map((pEl) =>
      blockText(parseParagraph(pEl, styleCatalog, numberingCatalog)),
    );
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
    comments.push({ threadId: docxCommentId(sourceId), messages });
  }
  return comments;
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
  let numberingCatalog: NumberingCatalog = EMPTY_NUMBERING_CATALOG;
  const mainRelsText = readPart(relsPathFor(mainPartPath));
  const mainRels: Map<string, Relationship> = mainRelsText ? parseRelationships(mainRelsText) : new Map();
  if (mainRelsText) {
    const stylesRel = Array.from(mainRels.values()).find((rel) => rel.type === STYLES_RELATIONSHIP);
    if (stylesRel) {
      const stylesText = readPart(resolvePartPath(mainPartPath, stylesRel.target));
      if (stylesText) {
        stylesDoc = parseXml(stylesText);
      }
    }
    const numberingRel = Array.from(mainRels.values()).find(
      (rel) => rel.type === NUMBERING_RELATIONSHIP,
    );
    if (numberingRel) {
      const numberingText = readPart(resolvePartPath(mainPartPath, numberingRel.target));
      if (numberingText) {
        numberingCatalog = parseNumberingCatalog(parseXml(numberingText));
      }
    }
  }
  const styleCatalog = parseStylesCatalog(stylesDoc);
  const drawingCtx: DrawingContext = { mainRels, mainPartPath, parts };
  const linkContext: LinkReadContext = {
    hrefForRelId: (relId) => {
      const rel = mainRels.get(relId);
      if (!rel || rel.type !== HYPERLINK_RELATIONSHIP) {
        return undefined;
      }
      return sanitizeHref(rel.target) ?? undefined;
    },
  };

  // Sections (SPEC-028 AC-31): a w:sectPr inside a paragraph's w:pPr ends a
  // section — that paragraph is the section's last block and the sectPr
  // describes the section just closed; subsequent body children accumulate
  // into the next one. Opaque-block anchors are per section, so each
  // accumulator carries its own list. The body-level sectPr (last body
  // child) describes the final section, exactly as before SPEC-028.
  interface RawSection {
    blocks: Block[];
    opaqueBlocks: OpaqueBlock[];
    sectPrEl?: Element;
  }
  const rawSections: RawSection[] = [{ blocks: [], opaqueBlocks: [] }];
  let current = rawSections[0];
  const mediaById = new Map<string, MediaPart>();
  let sectPr: Element | undefined;
  const bodyChildren = Array.from(body.children);
  for (let bodyIndex = 0; bodyIndex < bodyChildren.length; bodyIndex += 1) {
    const child = bodyChildren[bodyIndex];
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "p") {
      // A TOC field (SPEC-033 AC-19, AC-20): its own paragraphs are the
      // field's cached result, so the whole span becomes one TOC block
      // rather than the ordinary per-paragraph parse below.
      const tocOptions = tocFieldStartOptions(child);
      if (tocOptions) {
        const { toc, endIndex } = parseTocField(bodyChildren, bodyIndex, tocOptions);
        current.blocks.push(toc);
        bodyIndex = endIndex;
        continue;
      }
      const { blocks: parsedBlocks, media } = parseBodyParagraphBlocks(
        child,
        styleCatalog,
        numberingCatalog,
        drawingCtx,
        linkContext,
      );
      current.blocks.push(...parsedBlocks);
      for (const part of media) {
        mediaById.set(part.mediaId, part);
      }
      const pPr = firstChildElement(child, WORD_NAMESPACE, "pPr");
      const paragraphSectPr = pPr ? firstChildElement(pPr, WORD_NAMESPACE, "sectPr") : undefined;
      if (paragraphSectPr) {
        current.sectPrEl = paragraphSectPr;
        current = { blocks: [], opaqueBlocks: [] };
        rawSections.push(current);
      }
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "tbl") {
      // Regular-grid tables are modelled (SPEC-022); a table the model
      // cannot express (merged cells, nested tables, …) falls through to
      // the opaque path below.
      const table = parseTable(child, styleCatalog, numberingCatalog, linkContext);
      if (table) {
        current.blocks.push(table);
      } else {
        current.opaqueBlocks.push({ xml: serializeElement(child), blockIndex: current.blocks.length });
      }
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "sectPr") {
      sectPr = child;
    } else {
      // Block-level content controls, bookmarks, anything from another
      // namespace: unmodelled body children are preserved verbatim,
      // anchored to their position among the paragraphs (SPEC-008).
      current.opaqueBlocks.push({ xml: serializeElement(child), blockIndex: current.blocks.length });
    }
  }

  const fallbackStyles = createDefaultStyles();
  const defaultStyles = applySectionProperties(sectPr, {
    ...fallbackStyles,
    fontFamily: styleCatalog.docDefaultRunFormatting.fontFamily ?? fallbackStyles.fontFamily,
    fontSizePt: styleCatalog.docDefaultRunFormatting.fontSizePt ?? fallbackStyles.fontSizePt,
    language: styleCatalog.docDefaultLanguage ?? fallbackStyles.language,
  });

  const sections: Section[] = rawSections.map((raw) => {
    const section: Section = { blocks: raw.blocks };
    if (raw.opaqueBlocks.length > 0) {
      section.opaqueBlocks = raw.opaqueBlocks;
    }
    return section;
  });
  // A single-section document keeps its geometry on defaultStyles alone —
  // the pre-SPEC-028 model shape, unchanged. A multi-section one puts
  // explicit geometry on every non-final section (its own sectPr's values,
  // falling back to the body-adjusted defaults for anything it omits, the
  // same rule applySectionProperties applies to its base), leaving the
  // final section on defaultStyles as before.
  if (sections.length > 1) {
    rawSections.forEach((raw, index) => {
      const section = sections[index];
      if (raw.sectPrEl) {
        const resolved = applySectionProperties(raw.sectPrEl, defaultStyles);
        section.pageSize = { ...resolved.pageSize };
        section.margins = { ...resolved.margins };
        const start = parseSectionStart(raw.sectPrEl);
        if (start !== undefined) {
          section.start = start;
        }
        // Preserved verbatim so unmodelled section properties survive a
        // same-format save (SPEC-008) — the per-section counterpart of
        // sourceXml.docxSectPr below.
        section.docxSectPr = serializeElement(raw.sectPrEl);
      } else {
        const start = parseSectionStart(sectPr);
        if (start !== undefined) {
          section.start = start;
        }
      }
    });
  }

  // SPEC-030: column layouts, read for every document (single- and
  // multi-section). A non-final section reads its own sectPr; the final
  // section reads the body-level sectPr — the same source split the geometry
  // above uses.
  sections.forEach((section, index) => {
    const isFinal = index === sections.length - 1;
    const sectPrEl = rawSections[index]?.sectPrEl ?? (isFinal ? sectPr : undefined);
    const columns = parseColumns(sectPrEl);
    if (columns) {
      section.columns = columns;
    }
  });

  const sourceXml: SourceXml = {};
  if (sectPr) {
    sourceXml.docxSectPr = serializeElement(sectPr);
  }

  // Headers, footers and page numbering (SPEC-029): resolved for every
  // section regardless of section count — unlike page geometry above, which
  // stays implicit on defaultStyles for a single-section document, headers/
  // footers/titlePg/pgNumType can appear even then. Header/footer PARTS are
  // always regenerated from the model on save (see writer.ts) rather than
  // preserved-then-patched the way styles.xml/numbering.xml are, since their
  // content is now fully modelled — so a section with any reference has its
  // preserved sectPr dropped here: reusing it verbatim would keep pointing
  // at relationship ids the writer's freshly minted ones do not reproduce.
  // Footnotes and endnotes (SPEC-031): the note parts are consumed and
  // rebuilt from the model on save — like headers/footers — so they are
  // excluded from `opaque` below; the reference runs already became note
  // marker runs in collectRuns.
  const consumedHeaderFooterParts = new Set<string>();
  const notes: NoteContent[] = [];
  for (const [relType, kind] of [
    [FOOTNOTES_RELATIONSHIP, "footnote"],
    [ENDNOTES_RELATIONSHIP, "endnote"],
  ] as Array<[string, NoteKind]>) {
    const rel = Array.from(mainRels.values()).find((entry) => entry.type === relType);
    if (!rel) {
      continue;
    }
    const partPath = resolvePartPath(mainPartPath, rel.target);
    const parsed = parseNotesPart(partPath, parts, styleCatalog, numberingCatalog, kind);
    notes.push(...parsed.notes);
    for (const part of parsed.media) {
      mediaById.set(part.mediaId, part);
    }
    consumedHeaderFooterParts.add(partPath);
    consumedHeaderFooterParts.add(relsPathFor(partPath));
  }
  const footnoteSettings = parseNoteSettings(sectPr, "footnotePr");
  const endnoteSettings = parseNoteSettings(sectPr, "endnotePr");

  // Comments (SPEC-038): word/comments.xml is consumed and rebuilt from the
  // model on save, like the note parts above — the range/reference elements
  // already became commentAnchors side-lists in collectRuns. A thread whose
  // range anchors turn out not to be present anywhere in the body (a
  // malformed source) is still kept here rather than dropped: it simply
  // stays invisible everywhere, the same "derive, don't discard" contract
  // NoteContent uses (see lib/comments.ts's visibleCommentThreads).
  const comments: CommentThread[] = [];
  const commentsRel = Array.from(mainRels.values()).find((entry) => entry.type === COMMENTS_RELATIONSHIP);
  if (commentsRel) {
    const partPath = resolvePartPath(mainPartPath, commentsRel.target);
    for (const parsed of parseCommentsPart(partPath, parts, styleCatalog, numberingCatalog)) {
      comments.push({ threadId: parsed.threadId, resolved: false, messages: parsed.messages });
    }
    consumedHeaderFooterParts.add(partPath);
    consumedHeaderFooterParts.add(relsPathFor(partPath));
  }

  const settingsRel = Array.from(mainRels.values()).find((rel) => rel.type === SETTINGS_RELATIONSHIP);
  const settingsPartPath = settingsRel ? resolvePartPath(mainPartPath, settingsRel.target) : "word/settings.xml";
  const settingsText = readPart(settingsPartPath);
  const evenAndOddHeaders = !!settingsText && /<[\w:]*evenAndOddHeaders\b/.test(settingsText);
  const trackChangesRecording = !!settingsText && /<[\w:]*trackChanges\b/.test(settingsText);
  const headerFooterContentCache = new Map<string, { blocks: Block[]; media: MediaPart[] }>();
  rawSections.forEach((raw, index) => {
    const isFinal = index === rawSections.length - 1;
    const ownSectPr = isFinal ? sectPr : raw.sectPrEl;
    const refs = parseHeaderFooterReferences(ownSectPr);
    const titlePage = parseTitlePg(ownSectPr);
    const pageNumberStart = parsePageNumberStart(ownSectPr);
    if (refs.length === 0 && !titlePage && pageNumberStart === undefined && !evenAndOddHeaders) {
      return;
    }
    const section = sections[index];
    if (titlePage) {
      section.titlePage = true;
    }
    if (pageNumberStart !== undefined) {
      section.pageNumberStart = pageNumberStart;
    }
    if (evenAndOddHeaders) {
      section.evenAndOddHeaders = true;
    }
    for (const ref of refs) {
      const rel = mainRels.get(ref.relId);
      if (!rel) {
        continue;
      }
      const partPath = resolvePartPath(mainPartPath, rel.target);
      if (!headerFooterContentCache.has(partPath)) {
        headerFooterContentCache.set(
          partPath,
          parseHeaderFooterContent(partPath, parts, styleCatalog, numberingCatalog),
        );
        consumedHeaderFooterParts.add(partPath);
        consumedHeaderFooterParts.add(relsPathFor(partPath));
      }
      const content = headerFooterContentCache.get(partPath)!;
      for (const part of content.media) {
        mediaById.set(part.mediaId, part);
      }
      const key = ref.part === "header" ? "headers" : "footers";
      section[key] = { ...section[key], [ref.kind]: { blocks: content.blocks } };
    }
    if (refs.length > 0) {
      delete section.docxSectPr;
      if (isFinal) {
        delete sourceXml.docxSectPr;
      }
    }
  });

  // Everything the pipeline above did not consume-and-rebuild is retained
  // verbatim for the lifetime of the document (SPEC-008 AC): styles.xml and
  // the relationship/content-type parts included, since a same-format save
  // reuses them as-is. Media parts a modelled Image references (SPEC-024)
  // are excluded here — they travel in `media` instead, not duplicated into
  // `opaque` (the writer regenerates their relationships/content types from
  // `media` on every save, same-format or not) — header/footer parts
  // (SPEC-029) are excluded the same way, for the same reason.
  const opaque: OpaquePart[] = [];
  for (const [name, data] of parts) {
    if (
      name !== mainPartPath &&
      name !== corePartPath &&
      !mediaById.has(name) &&
      !consumedHeaderFooterParts.has(name)
    ) {
      opaque.push({ name, data: bytesToBase64(data) });
    }
  }

  const result: ParsedDocxDocument = {
    sections,
    defaultStyles,
    styles: catalogToNamedStyles(styleCatalog),
    opaque,
    sourceXml,
    media: Array.from(mediaById.values()),
    // Multilevel definitions (SPEC-032) the resolved w:numPr references
    // actually used — reconstructed from their composite abstractNums.
    listDefinitions: numberingCatalog.definitionsUsed(),
    notes,
    comments,
    trackChangesRecording,
  };
  if (footnoteSettings) {
    result.footnoteSettings = footnoteSettings;
  }
  if (endnoteSettings) {
    result.endnoteSettings = endnoteSettings;
  }
  return result;
}
