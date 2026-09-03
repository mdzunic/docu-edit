// Orchestrates the .docx -> DocumentModel pipeline: unzip the OPC package,
// walk the package/document relationships to find the real part names (per
// AC-3, never assume "word/document.xml" is where it is), map w:p/w:r onto
// the document model, resolve w:pStyle against styles.xml, and apply the
// section's page setup.
import type {
  Block,
  BookmarkAnchor,
  BreakKind,
  CellPropertyChange,
  Chart,
  CommentAnchor,
  CommentMessage,
  CommentThread,
  DefaultStyles,
  DocTheme,
  DocumentProperties,
  FieldKind,
  HeaderFooterKind,
  HeaderFooterPart,
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
  RowPropertyChange,
  Run,
  RunFormatting,
  Section,
  SectionColumns,
  SectionStart,
  Shape,
  SourceXml,
  Table,
  TableCell,
  TablePropertyChange,
  TableRow,
  TableOfContents,
  TocEntry,
  TocOptions,
  TableOfFigures,
  TofEntry,
  RefFormat,
  TrackedChangeInfo,
} from "../documentModel";
import type { PageBorders } from "../documentModel";
import type { ListDefinition } from "../core/lists";
import type { MediaPart } from "../core/media";
import type {
  BorderSide,
  CellBorders,
  CellVerticalAlign,
  TableAlignment,
  TableWidthMode,
} from "../core/tables";
import { parseBorderElement } from "./borders";
import {
  createDefaultStyles,
  createEquationRun,
  createFieldSpecRun,
  createNoteRun,
  DEFAULT_COLUMN_GAP_CM,
  MAX_COLUMN_COUNT,
  mergeParagraphFormatting,
  mergeRunFormatting,
  type FieldSpec,
} from "../documentModel";
import { parseChipText } from "../mergeChips";
import { dropRedundantPaintOverride } from "../paragraphBorders";
import { NORMAL_STYLE_ID } from "../styles";
import { canonicalDateTimeFormatFromDocx } from "../fields";
import { sanitizeHref } from "../links";
import { createMessageId, parseReplyParagraphText } from "../comments";
import { yieldToMainThread } from "../core/idleScheduler";
import { bytesToBase64 } from "../opaque";
import { blockText } from "../textEditing";
import { ZipError, unzip } from "../core/zip";
import { wordAttr, WORD_NAMESPACE, MATH_NAMESPACE } from "./xml";
import { parseOmmlEquation } from "./omml";
import {
  childElements,
  firstChildElement,
  MARKUP_COMPATIBILITY_NAMESPACE,
  parseRelationships,
  parseXml,
  relsPathFor,
  resolvePartPath,
  selectAlternateContent,
  serializeElement,
  XmlParseError,
  type Relationship,
} from "../core/xml";
import { isFloatingDrawing } from "../drawingAnchors";
import { blipMediaResolver, parseDrawing, type ParsedDrawing } from "./drawing";
import { ensureMediaEncoded, pendingMediaByteLength } from "../core/media";
import { parseShapeDrawing } from "./shape";
import { parseGroupDrawing, type ParsedGroupDrawing } from "./group";
import { parseVmlPicture } from "./vml";
import { parseConnectorDrawing, resolveDocxConnectorAttachments, stampConnectorTargetId } from "./connector";
import { parseChartDrawing } from "./chart";
import { catalogToNamedStyles, parseStylesCatalog, resolveStyle, type StyleCatalog } from "./styles";
import {
  EMPTY_NUMBERING_CATALOG,
  NUMBERING_RELATIONSHIP,
  parseNumberingCatalog,
  type NumberingCatalog,
} from "./numbering";
import {
  parseOutlineLevel,
  parseParagraphProperties,
  parseParagraphPropertiesChange,
  parseRunProperties,
  parseRunPropertiesChange,
} from "./properties";
import { ooxmlColorToCss } from "../core/colors";
import type { ModelColor } from "../core/theme";
import { readDocxTheme, THEME_RELATIONSHIP, themeRefFromAttrs } from "./theme";
import { twipsToCm } from "../core/units";
import { parseCoreProperties, parseCustomProperties } from "../core/opcProperties";
import { OFFICE_DOCUMENT_RELATIONSHIP } from "../core/packageDetect";

export class DocxOpenError extends Error {}
export class LegacyDocError extends Error {}

export const INVALID_DOCX_MESSAGE = "This file isn't a Word document we can open";
export const LEGACY_DOC_MESSAGE = "Old .doc files aren't supported — save as .docx first";

const STYLES_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const CUSTOM_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
const HYPERLINK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
// The `r:` attribute namespace (w:hyperlink/@r:id) — same URI drawing.ts
// uses for a:blip/@r:embed.
const R_ATTRIBUTE_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// `wp:docPr` lives here, and `de:group` — this editor's own group ordinal
// (shape.ts's DE_NAMESPACE) — is an attribute on it. Both are needed only to
// seed the imported-group ordinal counter below.
const WP_DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const DE_NAMESPACE = "urn:docu-edit:shape";

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
  /** The document's theme (SPEC-077 AC-25), parsed from the theme part the
   * main part's rels name; absent when the package has none. */
  theme?: DocTheme;
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
  /** Document properties and metadata (SPEC-041) read from docProps/core.xml
   * and docProps/custom.xml; `{}` when the package had neither. */
  properties: DocumentProperties;
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
/** A run's character-style reference (SPEC-042 `w:rStyle`) resolved against
 * the style catalog and merged UNDER the run's own direct `rPr` (OOXML's
 * character-style-overrides-paragraph-style, direct-overrides-both
 * precedence — `inheritedFormatting` is already the paragraph's resolution). */
function parseRun(
  rEl: Element,
  inheritedFormatting: RunFormatting,
  styleCatalog: StyleCatalog,
): Run | undefined {
  const rPr = firstChildElement(rEl, WORD_NAMESPACE, "rPr");
  const rStyleEl = rPr ? firstChildElement(rPr, WORD_NAMESPACE, "rStyle") : undefined;
  const styleId = rStyleEl ? (wordAttr(rStyleEl, "val") ?? undefined) : undefined;
  const base = styleId
    ? mergeRunFormatting(inheritedFormatting, resolveStyle(styleCatalog, styleId).runFormatting)
    : inheritedFormatting;
  const formatting = mergeRunFormatting(base, parseRunProperties(rPr));
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
  const run: Run = { text, formatting };
  if (styleId) {
    run.styleId = styleId;
  }
  if (formatChange) {
    run.changes = { formatChange };
  }
  return run;
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
  styleCatalog: StyleCatalog,
  runs: Run[],
  opaqueInlines: OpaqueInline[],
  bookmarks: BookmarkAnchor[],
  links?: LinkReadContext,
  commentAnchors: CommentAnchor[] = [],
): void {
  const children = Array.from(container.children);
  for (let ci = 0; ci < children.length; ci++) {
    const child = children[ci];
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "pPr") {
      continue; // Modelled separately by parseParagraph.
    }
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "r") {
      // A run-split complex field (SPEC-053 AC-47): a `w:fldChar begin` opens
      // an `instrText` / `separate` / result / `end` sequence that reads into
      // one field run, its cached value the result runs' text. A TOC field is
      // handled at block level, so it falls through here unchanged.
      if (runFldCharType(child) === "begin") {
        const consumed = consumeComplexField(children, ci);
        if (consumed) {
          runs.push(consumed.run);
          ci = consumed.next - 1;
          continue;
        }
      }
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
      const run = parseRun(child, inheritedFormatting, styleCatalog);
      if (run) {
        runs.push(run);
      }
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "hyperlink") {
      // Hyperlinks (SPEC-025): the wrapper's target becomes link formatting
      // on its child runs; the wrapper element itself is not preserved.
      collectRuns(
        child,
        hyperlinkFormatting(child, inheritedFormatting, links),
        styleCatalog,
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
        collectRuns(child, inheritedFormatting, styleCatalog, runs, opaqueInlines, bookmarks, links, commentAnchors),
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
    } else if (
      child.namespaceURI === MATH_NAMESPACE &&
      (child.localName === "oMath" || child.localName === "oMathPara")
    ) {
      // An inline (`m:oMath`) or display (`m:oMathPara`) equation (SPEC-056):
      // parsed into an equation-marker run; a construct outside the closed set
      // is kept verbatim inside it as an `opaqueMath` node.
      runs.push(createEquationRun(parseOmmlEquation(child), inheritedFormatting));
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
  // Same idea for the SPEC-074 AC-17 paint overrides: an all-nil w:pBdr and
  // a fill-less w:shd exist to switch a style's box or shading back off, and
  // Word writes them freely on paragraphs whose style has neither, where
  // they say nothing the absent field did not. The comparison resolves
  // through Normal for a paragraph carrying no w:pStyle — it still has a
  // style, and its box comes from there (styles.inheritedParagraphPaint's
  // rule, which is what will paint it).
  dropRedundantPaintOverride(
    formatting,
    styleId === undefined
      ? resolveStyle(styleCatalog, NORMAL_STYLE_ID).paragraphFormatting
      : resolved.paragraphFormatting,
  );
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
  collectRuns(pEl, meta.resolved.runFormatting, styleCatalog, runs, opaqueInlines, bookmarks, links, commentAnchors);
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
  /** SPEC-078: the running anchor-ordinal counter for this parse. A source
   * paragraph containing floating shape/chart drawings mints the next value
   * onto its first split-off piece and those drawings — containment IS the
   * wire pairing, and minting in document order is what keeps
   * read → write → read structurally equal. Absent on contexts that never
   * reach parseBodyParagraphBlocks (headers/footers, the comments part). */
  anchorOrdinals?: { next: number };
  /** The document-wide counter imported `wpg:wgp` groups take their
   * `groupOrdinal` from (SPEC-051's flat grouping — see group.ts). Seeded
   * above every `de:group` already in the part so a group Word added to a
   * document this editor had saved cannot land on an ordinal that document
   * already uses. Absent on contexts that never reach
   * parseBodyParagraphBlocks (headers/footers, the comments part), where a
   * drawing is preserved opaquely anyway. */
  groupOrdinals?: { next: number };
}

/** One more than the largest `de:group` ordinal the part already carries —
 * the seed for ordinals minted for imported `wpg:wgp` groups (group.ts).
 *
 * Both kinds of group can coexist in one file: this editor saves its groups
 * as `de:group` on sibling drawings, and a user who then groups two more
 * shapes in Word gets a real `wpg:wgp` next to them. Minting from 1 would put
 * the imported group on an ordinal the document already uses, silently
 * merging two unrelated groups. */
function firstFreeGroupOrdinal(root: Element): number {
  let max = 0;
  const docPrs = root.getElementsByTagNameNS(WP_DRAWING_NAMESPACE, "docPr");
  for (let i = 0; i < docPrs.length; i++) {
    const value = Number(docPrs[i].getAttributeNS(DE_NAMESPACE, "group"));
    if (Number.isFinite(value)) {
      max = Math.max(max, value);
    }
  }
  return max + 1;
}

/** The DrawingML extension namespaces whose markup the readers below can
 * actually parse — the answer to an `mc:Choice`'s `Requires`. `wp14` adds
 * relative sizing/positioning attributes to a `wp:anchor` we already ignore;
 * `wps` is SPEC-051's shape/SPEC-070's connector namespace; `wpg` is the
 * group namespace group.ts flattens.
 *
 * Deliberately absent: `wpc` (a drawing canvas), which no reader models — its
 * Choice is skipped, the VML `mc:Fallback` is read instead, nothing parses out
 * of it, and the run stays preserved verbatim. Add a namespace here only once
 * something can read what it gates: selecting a Choice we then fail to parse
 * still preserves the run, but only because every caller falls back to the
 * original element rather than the branch. */
const UNDERSTOOD_DRAWING_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
]);

/** The single `w:drawing` a run's content resolves to, or null when the run
 * holds anything else (text, a field, several drawings, a `w:pict`).
 *
 * Word writes a drawing bare only when its markup predates the 2010
 * extensions; a shape, a text box, or a picture with relative sizing is
 * wrapped in `mc:AlternateContent`, so this looks through that wrapper via
 * the MCE rules (core/xml's selectAlternateContent) rather than requiring a
 * direct child. Looking through it is a read-side lens only: when nothing
 * downstream models what is inside, every caller falls back to the
 * opaque-inline path, which serialises the *original* run — wrapper, chosen
 * branch and VML fallback intact — so preservation is unchanged. */
function runDrawingElement(rEl: Element): Element | null {
  let drawing: Element | null = null;
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI === WORD_NAMESPACE) {
      if (child.localName === "rPr") {
        continue;
      }
      if (child.localName !== "drawing" || drawing) {
        return null;
      }
      drawing = child;
      continue;
    }
    if (child.namespaceURI === MARKUP_COMPATIBILITY_NAMESPACE && child.localName === "AlternateContent") {
      const selected = selectAlternateContent(child, UNDERSTOOD_DRAWING_NAMESPACES);
      if (selected.length !== 1 || selected[0].namespaceURI !== WORD_NAMESPACE) {
        return null;
      }
      if (selected[0].localName !== "drawing" || drawing) {
        return null;
      }
      drawing = selected[0];
      continue;
    }
    // Any other foreign-namespaced child is skipped rather than disqualifying
    // the run, matching the `namespaceURI === WORD_NAMESPACE` filter this
    // replaced — widening the unwrap must not narrow what already reads.
  }
  return drawing;
}

/** A `w:r` whose only content (besides `w:rPr`) is one `w:drawing` that
 * resolves to a modelled picture — the mid-paragraph split point
 * parseBodyParagraphBlocks looks for. Null for every other run, including
 * one holding a drawing this spec doesn't model (a chart, shape, ...), which
 * keeps falling through to the ordinary opaque-inline path. */
function imageRunDrawing(rEl: Element, ctx: DrawingContext): ParsedDrawing | null {
  const drawing = runDrawingElement(rEl);
  if (!drawing) {
    return null;
  }
  return parseDrawing(drawing, ctx.mainRels, ctx.mainPartPath, ctx.parts);
}

/** A `w:r` whose only content is one `w:drawing` holding a `wps:wsp` this
 * spec models (SPEC-051) — the shape-flavoured sibling of imageRunDrawing.
 * Null for every other run, including one holding a drawing neither this
 * spec nor SPEC-024 models (a chart, an unsupported wrap...), which keeps
 * falling through to the opaque-inline path. */
function shapeRunDrawing(
  rEl: Element,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  links: LinkReadContext | undefined,
  resolveBlip?: (embedId: string) => string | null,
): Shape | null {
  const drawing = runDrawingElement(rEl);
  if (!drawing) {
    return null;
  }
  const parsed = parseShapeDrawing(drawing, {
    parseParagraph: (pEl) => parseParagraph(pEl, styleCatalog, numberingCatalog, links),
    resolveBlip,
  });
  return parsed?.shape ?? null;
}

/** A `w:r` whose only content is one legacy `w:pict` holding an embedded VML
 * picture (vml.ts) — the Word 97-2003 sibling of imageRunDrawing. Null for
 * every other run, including a `w:pict` holding a VML shape or a horizontal
 * rule, which keeps its verbatim preservation.
 *
 * Deliberately separate from runDrawingElement: a `w:pict` is not a
 * `w:drawing`, and the two never appear as alternatives to each other in a
 * run — Word writes VML markup instead of DrawingML, or as the `mc:Fallback`
 * of an AlternateContent whose Choice this reader takes first. */
function hasVmlPict(rEl: Element): boolean {
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "pict") {
      return true;
    }
  }
  return false;
}

function vmlRunPicture(rEl: Element, ctx: DrawingContext): ParsedDrawing | null {
  let pict: Element | null = null;
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI !== WORD_NAMESPACE) {
      continue;
    }
    if (child.localName === "rPr") {
      continue;
    }
    if (child.localName !== "pict" || pict) {
      return null;
    }
    pict = child;
  }
  return pict ? parseVmlPicture(pict, ctx.mainRels, ctx.mainPartPath, ctx.parts) : null;
}

/** A `w:r` whose only content is one `w:drawing` holding a `wpg:wgp` group,
 * flattened into its member blocks — the many-blocks sibling of
 * imageRunDrawing/shapeRunDrawing. Null for every other run, and for a group
 * holding a member no reader models, which keeps the whole run on the
 * opaque-inline path.
 *
 * Deliberately free of side effects (it mints no ordinal): the caller runs it
 * once to decide whether the paragraph splits and again to build the blocks,
 * then stamps the shared `groupOrdinal` itself. */
function groupRunDrawing(
  rEl: Element,
  ctx: DrawingContext,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  links: LinkReadContext | undefined,
): ParsedGroupDrawing | null {
  const drawing = runDrawingElement(rEl);
  if (!drawing) {
    return null;
  }
  return parseGroupDrawing(drawing, {
    mainRels: ctx.mainRels,
    mainPartPath: ctx.mainPartPath,
    parts: ctx.parts,
    parseParagraph: (pEl) => parseParagraph(pEl, styleCatalog, numberingCatalog, links),
  });
}

/** The single `w:drawing` a run resolves to (the caller has already
 * confirmed it has one via imageRunDrawing/shapeRunDrawing/etc.) — used to
 * stamp the drawing's `wp:docPr/@id` onto its parsed block for connector
 * resolution. Resolved through runDrawingElement so a drawing inside an
 * `mc:AlternateContent` is stamped from the same element that was parsed. */
function drawingElementOf(rEl: Element): Element {
  return runDrawingElement(rEl) as Element;
}

/** A `w:r` whose only content is one `w:drawing` holding a `wps:cxnSp`
 * (SPEC-070) — the connector-flavoured sibling of shapeRunDrawing. Null for
 * every other run. A `wps:cxnSp` reads as a Connector, never a Shape. */
function connectorRunDrawing(rEl: Element): { connector: import("../documentModel").Connector; drawing: Element } | null {
  const drawing = runDrawingElement(rEl);
  if (!drawing) {
    return null;
  }
  const parsed = parseConnectorDrawing(drawing);
  return parsed ? { connector: parsed.connector, drawing } : null;
}

/** A `w:r` whose only content is one `w:drawing` holding a `c:chart`
 * reference (SPEC-052) — the chart-flavoured sibling of imageRunDrawing/
 * shapeRunDrawing. Null for every other run, which keeps falling through to
 * the opaque-inline path. Unlike the other two, this never returns null just
 * because the referenced chart is outside "Chart catalogue" — an
 * out-of-catalogue chart still becomes a (placeholder) Chart block, per
 * SPEC-051's own `unsupportedGeometryXml` convention for a shape, so it gets
 * the anchored-drawing selection/paint/preserve treatment rather than
 * falling back to an inert opaque-inline run. */
function chartRunDrawing(rEl: Element, ctx: DrawingContext): Chart | null {
  const drawing = runDrawingElement(rEl);
  if (!drawing) {
    return null;
  }
  const parsed = parseChartDrawing(drawing, ctx.mainRels, ctx.mainPartPath, ctx.parts);
  if (!parsed) {
    return null;
  }
  const chart: Chart = {
    type: "chart",
    chartType: parsed.chartType,
    categories: parsed.categories,
    series: parsed.series,
    options: parsed.options,
    sourcePart: parsed.sourcePart,
    widthCm: parsed.widthCm,
    heightCm: parsed.heightCm,
    wrap: parsed.wrap,
  };
  if (parsed.unsupportedChartXml !== undefined) {
    chart.unsupportedChartXml = parsed.unsupportedChartXml;
  }
  if (parsed.altText) {
    chart.altText = parsed.altText;
  }
  if (parsed.groupOrdinal !== undefined) {
    chart.groupOrdinal = parsed.groupOrdinal;
  }
  if (parsed.offsetXCm !== undefined) {
    chart.offsetXCm = parsed.offsetXCm;
  }
  if (parsed.offsetYCm !== undefined) {
    chart.offsetYCm = parsed.offsetYCm;
  }
  if (parsed.relativeToH !== undefined) {
    chart.relativeToH = parsed.relativeToH;
  }
  if (parsed.relativeToV !== undefined) {
    chart.relativeToV = parsed.relativeToV;
  }
  if (parsed.zOrder !== undefined) {
    chart.zOrder = parsed.zOrder;
  }
  return chart;
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

/** The caption label a `TOC \c "Figure"` field names (SPEC-055), or null when
 * the paragraph does not begin a table-of-figures field. A `\c` switch is what
 * distinguishes a table of figures from a table of contents. */
function tofFieldStartLabel(pEl: Element): string | null {
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
  const match = /\\c\s*"([^"]*)"/.exec(instr);
  return match ? match[1] : null;
}

/** One table-of-figures entry from a cached field paragraph (SPEC-055): the
 * caption text as the hyperlink text, its anchor, and the trailing page
 * number after the tab. Mirrors tocEntryFromParagraph. */
function tofEntryFromParagraph(pEl: Element): TofEntry | null {
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
  const page = /^\d+$/.test(pageNumberText) ? Number(pageNumberText) : null;
  return { text, page, targetBookmark: anchor };
}

/** Parses a `TOC \c` field's full paragraph run into one TableOfFigures block
 * (SPEC-055), mirroring parseTocField. */
function parseTofField(
  bodyChildren: Element[],
  startIndex: number,
  label: string,
): { tof: TableOfFigures; endIndex: number } {
  const entries: TofEntry[] = [];
  let endIndex = startIndex;
  let depth = 0;
  for (let i = startIndex; i < bodyChildren.length; i += 1) {
    const pEl = bodyChildren[i];
    if (pEl.namespaceURI !== WORD_NAMESPACE || pEl.localName !== "p") {
      continue;
    }
    const entry = tofEntryFromParagraph(pEl);
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
  const hasPageNumbers = entries.some((entry) => entry.page !== null);
  const tof: TableOfFigures = {
    type: "tof",
    label,
    options: { pageNumbers: hasPageNumbers, dotLeader: true, hyperlinks: hasPageNumbers },
    entries,
  };
  return { tof, endIndex };
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
  // Every drawing precheck below starts by resolving the run's `w:drawing`,
  // so ask that once first: a paragraph of plain text — the overwhelming
  // majority — then pays one cheap scan per run instead of one per drawing
  // kind, which is what keeps adding a kind from costing anything on open.
  const hasDrawingRun = runElements.some((r) => runDrawingElement(r) !== null);
  const hasImageRun = hasDrawingRun && runElements.some((r) => imageRunDrawing(r, ctx) !== null);
  const hasShapeRun =
    hasDrawingRun && runElements.some((r) => shapeRunDrawing(r, styleCatalog, numberingCatalog, links) !== null);
  const hasChartRun = hasDrawingRun && runElements.some((r) => chartRunDrawing(r, ctx) !== null);
  const hasConnectorRun = hasDrawingRun && runElements.some((r) => connectorRunDrawing(r) !== null);
  const hasGroupRun =
    hasDrawingRun && runElements.some((r) => groupRunDrawing(r, ctx, styleCatalog, numberingCatalog, links) !== null);
  // Legacy VML has no `w:drawing` to resolve, so it is asked separately —
  // and only of runs that actually hold a `w:pict`, which is why the cheap
  // element check comes before the parse.
  const hasVmlRun = runElements.some(
    (r) => hasVmlPict(r) && vmlRunPicture(r, ctx) !== null,
  );
  // A typed w:br in an otherwise-modelled run splits the paragraph the same
  // way an image does (SPEC-028 AC-30); one inside an opaque run stays part
  // of that run's verbatim XML — splitting an unmodelled run could not
  // round-trip.
  const hasBreakRun = runElements.some(
    (r) => !runHasUnmodelledContent(r) && runHasTypedBreak(r),
  );
  if (
    !hasImageRun &&
    !hasShapeRun &&
    !hasChartRun &&
    !hasConnectorRun &&
    !hasGroupRun &&
    !hasVmlRun &&
    !hasBreakRun
  ) {
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
        stampConnectorTargetId(drawingElementOf(child), parsed.image);
        media.push(parsed.media);
        continue;
      }
      const shape = shapeRunDrawing(child, styleCatalog, numberingCatalog, links, blipMediaResolver(ctx, media));
      if (shape) {
        flush();
        blocks.push(shape);
        stampConnectorTargetId(drawingElementOf(child), shape);
        continue;
      }
      // SPEC-070: a wps:cxnSp reads as a Connector, before the chart/opaque
      // fallbacks (its own graphicData uri is the same as a shape's).
      const connectorRun = connectorRunDrawing(child);
      if (connectorRun) {
        flush();
        blocks.push(connectorRun.connector);
        stampConnectorTargetId(connectorRun.drawing, connectorRun.connector);
        continue;
      }
      const chart = chartRunDrawing(child, ctx);
      if (chart) {
        flush();
        blocks.push(chart);
        stampConnectorTargetId(drawingElementOf(child), chart);
        continue;
      }
      // A `wpg:wgp` group flattens into several blocks at once (group.ts),
      // all sharing one freshly minted ordinal. Not stamped for connector
      // targeting: a member has no `wp:docPr` of its own to be named by.
      const group = groupRunDrawing(child, ctx, styleCatalog, numberingCatalog, links);
      if (group) {
        flush();
        const ordinal = ctx.groupOrdinals ? ctx.groupOrdinals.next++ : undefined;
        for (const member of group.blocks) {
          if (ordinal !== undefined) {
            member.groupOrdinal = ordinal;
          }
          blocks.push(member);
        }
        media.push(...group.media);
        continue;
      }
      // A legacy `w:pict` picture (SPEC-024's markup predecessor, vml.ts):
      // last of the drawing kinds, so a `w:pict` that is a VML shape rather
      // than a picture still falls through to the opaque-inline path below.
      if (hasVmlPict(child)) {
        const vml = vmlRunPicture(child, ctx);
        if (vml) {
          flush();
          blocks.push(vml.image);
          media.push(vml.media);
          continue;
        }
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
      const run = parseRun(child, meta.resolved.runFormatting, styleCatalog);
      if (run) {
        runs.push(run);
      }
    } else if (child.namespaceURI === WORD_NAMESPACE && child.localName === "hyperlink") {
      collectRuns(
        child,
        hyperlinkFormatting(child, meta.resolved.runFormatting, links),
        styleCatalog,
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
        collectRuns(child, meta.resolved.runFormatting, styleCatalog, runs, opaqueInlines, bookmarks, links, commentAnchors),
      );
    } else if (child.namespaceURI === WORD_NAMESPACE && TRANSIENT_PARAGRAPH_CHILDREN.has(child.localName)) {
      // Regenerated by Word on open; deliberately not preserved.
    } else if (
      child.namespaceURI === MATH_NAMESPACE &&
      (child.localName === "oMath" || child.localName === "oMathPara")
    ) {
      // An inline or display equation (SPEC-056), mirroring collectRuns.
      runs.push(createEquationRun(parseOmmlEquation(child), meta.resolved.runFormatting));
    } else {
      opaqueInlines.push({ xml: serializeElement(child), runIndex: runs.length });
    }
  }
  flush();
  // SPEC-078: floating shapes/charts found INSIDE this source paragraph are
  // anchored to it — containment is the wire pairing. Mint one canonical
  // ordinal onto the paragraph's first split-off piece (the piece that
  // keeps the paragraph's top edge) and every such drawing. A drawing-only
  // paragraph — every pre-SPEC-078 save — yields no paragraph piece and
  // stays unpaired, resolved by the adjacency fallback exactly as before.
  // Image never pairs (no anchor-relative offsets — SPEC-078 out of scope).
  if (ctx.anchorOrdinals) {
    const firstPiece = blocks.find((block) => block.type === "paragraph");
    const floating = blocks.filter(isFloatingDrawing);
    if (firstPiece?.type === "paragraph" && floating.length > 0) {
      const ordinal = ctx.anchorOrdinals.next++;
      firstPiece.anchorOrdinal = ordinal;
      for (const drawing of floating) {
        drawing.anchorOrdinal = ordinal;
      }
    }
  }
  return { blocks, media };
}

/** One `CT_Border` element as a `BorderSide` — the shared docx/borders.ts
 * mapping (SPEC-074 AC-1..3 widened the enum, so cell, paragraph and page
 * borders all read the same `w:val` table). */
function parseBorderSide(el: Element | undefined): BorderSide | undefined {
  return parseBorderElement(el, wordAttr);
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
function parseCellShading(tcPr: Element): ModelColor | undefined {
  const shd = firstChildElement(tcPr, WORD_NAMESPACE, "shd");
  if (!shd) {
    return undefined;
  }
  // @w:themeFill becomes a reference (SPEC-077 AC-27), same as run shading.
  const ref = themeRefFromAttrs(
    wordAttr(shd, "themeFill"),
    wordAttr(shd, "themeFillTint"),
    wordAttr(shd, "themeFillShade"),
  );
  return ref ?? ooxmlColorToCss(wordAttr(shd, "fill")) ?? undefined;
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
  // Tracked cell insert/delete/property-change (SPEC-039 AC-15): w:tcPr/
  // w:cellIns|w:cellDel|w:tcPrChange, the cell-level analogue of a run's
  // w:ins/w:del/w:rPrChange.
  const cellInsEl = tcPrEl ? firstChildElement(tcPrEl, WORD_NAMESPACE, "cellIns") : undefined;
  const cellDelEl = tcPrEl ? firstChildElement(tcPrEl, WORD_NAMESPACE, "cellDel") : undefined;
  if (cellInsEl || cellDelEl) {
    cell.changes = { ...cell.changes };
    if (cellInsEl) {
      cell.changes.insert = trackedChangeInfo(cellInsEl, "cellIns");
    }
    if (cellDelEl) {
      cell.changes.delete = trackedChangeInfo(cellDelEl, "cellDel");
    }
  }
  const tcPrChangeEl = tcPrEl ? firstChildElement(tcPrEl, WORD_NAMESPACE, "tcPrChange") : undefined;
  if (tcPrChangeEl) {
    const info = trackedChangeInfo(tcPrChangeEl, "tcPrChange");
    const prevTcPr = firstChildElement(tcPrChangeEl, WORD_NAMESPACE, "tcPr");
    const previous: CellPropertyChange["previous"] = {};
    if (prevTcPr) {
      const prevColSpan = parseGridSpan(prevTcPr);
      if (prevColSpan !== undefined) {
        previous.colSpan = prevColSpan;
      }
      const prevBorders = parseCellBorders(prevTcPr);
      if (prevBorders) {
        previous.borders = prevBorders;
      }
      const prevShading = parseCellShading(prevTcPr);
      if (prevShading !== undefined) {
        previous.shadingFill = prevShading;
      }
      const prevPadding = parseCellPadding(prevTcPr);
      if (prevPadding !== undefined) {
        previous.paddingCm = prevPadding;
      }
      const prevValign = parseCellVerticalAlign(prevTcPr);
      if (prevValign !== undefined) {
        previous.verticalAlign = prevValign;
      }
    }
    const prevRowSpan = Number(wordAttr(tcPrChangeEl, "prevRowSpan") ?? "");
    if (Number.isInteger(prevRowSpan) && prevRowSpan > 0) {
      previous.rowSpan = prevRowSpan;
    }
    cell.changes = { ...cell.changes, propertyChange: { ...info, previous } };
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

/** Whether a row repeats as the table's header row: `w:trPr/w:tblHeader`
 * (SPEC-043's "table without a header row" check). */
function parseRowIsHeader(trPr: Element | undefined): boolean {
  return trPr !== undefined && firstChildElement(trPr, WORD_NAMESPACE, "tblHeader") !== undefined;
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
    if (parseRowIsHeader(trPrEl)) {
      row.isHeader = true;
    }
    // Tracked row insert/delete (SPEC-039 AC-15): w:trPr/w:ins|w:del.
    const insEl = trPrEl ? firstChildElement(trPrEl, WORD_NAMESPACE, "ins") : undefined;
    const delEl = trPrEl ? firstChildElement(trPrEl, WORD_NAMESPACE, "del") : undefined;
    if (insEl || delEl) {
      row.changes = { ...row.changes };
      if (insEl) {
        row.changes.insert = trackedChangeInfo(insEl, "ins");
      }
      if (delEl) {
        row.changes.delete = trackedChangeInfo(delEl, "del");
      }
    }
    // Tracked row property change (SPEC-039 AC-15): w:trPr/w:trPrChange.
    const trPrChangeEl = trPrEl ? firstChildElement(trPrEl, WORD_NAMESPACE, "trPrChange") : undefined;
    if (trPrChangeEl) {
      const info = trackedChangeInfo(trPrChangeEl, "trPrChange");
      const prevTrPr = firstChildElement(trPrChangeEl, WORD_NAMESPACE, "trPr");
      const previous: RowPropertyChange["previous"] = {};
      const prevMinHeight = parseRowMinHeight(prevTrPr);
      if (prevMinHeight !== undefined) {
        previous.minHeightCm = prevMinHeight;
      }
      row.changes = { ...row.changes, propertyChange: { ...info, previous } };
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
  // Tracked table property change (SPEC-039 AC-15): w:tblPr/w:tblPrChange,
  // with an optional nested w:tblGrid for a column-widths edit (see
  // writer.ts's tablePropertyChangeXml).
  const tblPrChangeEl = tblPrEl ? firstChildElement(tblPrEl, WORD_NAMESPACE, "tblPrChange") : undefined;
  if (tblPrChangeEl) {
    const info = trackedChangeInfo(tblPrChangeEl, "tblPrChange");
    const prevTblPr = firstChildElement(tblPrChangeEl, WORD_NAMESPACE, "tblPr");
    const previous: TablePropertyChange["previous"] = {};
    if (prevTblPr) {
      const prevWidth = parseTableWidth(prevTblPr);
      if (prevWidth.widthMode !== undefined) {
        previous.widthMode = prevWidth.widthMode;
      }
      if (prevWidth.widthPercent !== undefined) {
        previous.widthPercent = prevWidth.widthPercent;
      }
      const prevAlignment = parseTableAlignment(prevTblPr);
      if (prevAlignment !== undefined) {
        previous.alignment = prevAlignment;
      }
    }
    const prevGridEl = firstChildElement(tblPrChangeEl, WORD_NAMESPACE, "tblGrid");
    if (prevGridEl) {
      const prevColumnWidthsCm: number[] = [];
      for (const col of Array.from(prevGridEl.children)) {
        const w = Number(wordAttr(col, "w") ?? "");
        if (Number.isFinite(w) && w > 0) {
          prevColumnWidthsCm.push(twipsToCm(w));
        }
      }
      if (prevColumnWidthsCm.length > 0) {
        previous.columnWidthsCm = prevColumnWidthsCm;
      }
    }
    table.changes = { propertyChange: { ...info, previous } };
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

/** A sectPr's `w:pgBorders` (SPEC-074 AC-30): the four edges, the reference
 * edge `@w:offsetFrom` measures from (AC-21) and which pages take it.
 * `@w:space` is the same on every edge in practice, so the first edge that
 * carries one supplies the model's single `offsetPt`. Art borders
 * (`@w:id`, out of scope per SPEC-074) read as the plain line they sit on
 * and are preserved verbatim by the sectPr passthrough. */
function parsePageBorders(sectPr: Element | undefined): PageBorders | undefined {
  const el = sectPr ? firstChildElement(sectPr, WORD_NAMESPACE, "pgBorders") : undefined;
  if (!el) {
    return undefined;
  }
  const borders: PageBorders = {};
  let offsetPt: number | undefined;
  for (const tag of ["top", "bottom", "left", "right"] as const) {
    const edgeEl = firstChildElement(el, WORD_NAMESPACE, tag);
    const side = parseBorderSide(edgeEl);
    if (!side) {
      continue;
    }
    borders[tag] = side;
    if (offsetPt === undefined && edgeEl) {
      const space = Number(wordAttr(edgeEl, "space") ?? "");
      if (Number.isFinite(space) && space >= 0) {
        offsetPt = space;
      }
    }
  }
  if (!borders.top && !borders.bottom && !borders.left && !borders.right) {
    return undefined;
  }
  if (offsetPt !== undefined) {
    borders.offsetPt = offsetPt;
  }
  borders.measureFrom = wordAttr(el, "offsetFrom") === "text" ? "text" : "page";
  const display = wordAttr(el, "display");
  if (display === "firstPage") {
    borders.appliesTo = "firstPage";
  } else if (display === "notFirstPage") {
    borders.appliesTo = "exceptFirstPage";
  }
  return borders;
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

/** A `MERGEFIELD "..."` instruction's canonical text, or null when the
 * instruction isn't a (quoted) MERGEFIELD — always quoted per SPEC-046
 * "Chip serialisation", so this is the one pattern the reader needs. */
function mergeFieldInstrText(instr: string): string | null {
  const match = /^\s*MERGEFIELD\s+"([^"]*)"\s*$/i.exec(instr);
  return match ? match[1] : null;
}

/** Maps a `.docx` field instruction and its cached result text to a
 * `FieldSpec` (SPEC-053). Page/count and merge chips keep their existing
 * shapes; every catalogue type maps to its kind; anything outside the
 * catalogue becomes an `unsupported` field carrying the instruction and its
 * result byte-for-byte (SPEC-008 at run level). Returns null only for a
 * `MERGEFIELD` whose quoted text isn't valid chip syntax — the caller then
 * preserves the whole element as an opaque inline, unchanged. */
/** The number format a SEQ/field `\* <switch>` names (SPEC-055): roman →
 * lowerRoman, alphabetic → lowerAlpha, everything else (ARABIC / absent) →
 * decimal — the mirror of the writer's SEQ_FORMAT_SWITCH. */
/** Whether a SEQ instruction's switches (everything after the label) can be
 * fully represented by a `seq` FieldSpec's {label, format}. True only when the
 * sole switch is the numeric-format switch (`\* arabic|roman|alphabetic`) or
 * there are none; any residual switch (chapter `\s`, restart `\r`, repeat `\c`,
 * `\n`, `\h`, an explicit bookmark, `\* MERGEFORMAT`, …) means the field must be
 * preserved verbatim instead (SPEC-056 fidelity guard). */
function seqSwitchesModellable(switches: string): boolean {
  return switches.replace(/\\\*\s*(arabic|roman|alphabetic)\b/i, " ").trim() === "";
}

function numberFormatFromSwitch(switches: string): "decimal" | "lowerRoman" | "lowerAlpha" {
  if (/\\\*\s*roman\b/i.test(switches)) {
    return "lowerRoman";
  }
  if (/\\\*\s*alphabetic\b/i.test(switches)) {
    return "lowerAlpha";
  }
  return "decimal";
}

/** The `show` a REF field's switches name (SPEC-053 + SPEC-055): the caption
 * markers first, then `\p` position, else plain text — the mirror of the
 * writer's CAPTION_REF_SWITCH. */
function refShowFromSwitches(switches: string): RefFormat {
  if (/\\\*\s*fullcaption\b/i.test(switches)) {
    return "full";
  }
  if (/\\\*\s*labelnumber\b/i.test(switches)) {
    return "labelAndNumber";
  }
  if (/\\\*\s*captionnumber\b/i.test(switches)) {
    return "numberOnly";
  }
  if (/\\\*\s*captiontext\b/i.test(switches)) {
    return "textOnly";
  }
  return /\\p\b/i.test(switches) ? "position" : "text";
}

function fieldSpecFromInstr(instr: string, cached: string): FieldSpec | null {
  const trimmed = instr.trim();
  const mergeText = mergeFieldInstrText(instr);
  if (mergeText !== null) {
    const chip = parseChipText(mergeText);
    return chip ?? null;
  }
  const pageParts = fieldInstrParts(instr);
  if (pageParts) {
    return { kind: pageParts.kind, format: pageParts.format };
  }
  let m: RegExpExecArray | null;
  if ((m = /^(DATE|TIME)\b(?:[\s\S]*?\\@\s*"([^"]*)")?/i.exec(trimmed))) {
    const kind = m[1].toUpperCase() === "DATE" ? "date" : "time";
    const picture = m[2] ?? (kind === "date" ? "d MMMM yyyy" : "HH:mm");
    return { kind, format: canonicalDateTimeFormatFromDocx(picture), cached };
  }
  if (/^AUTHOR\b/i.test(trimmed)) {
    return { kind: "author", cached };
  }
  if (/^TITLE\b/i.test(trimmed)) {
    return { kind: "docTitle", cached };
  }
  if (/^SUBJECT\b/i.test(trimmed)) {
    return { kind: "docSubject", cached };
  }
  if (/^KEYWORDS\b/i.test(trimmed)) {
    return { kind: "docKeywords", cached };
  }
  if (/^NUMWORDS\b/i.test(trimmed)) {
    return { kind: "wordCount", cached };
  }
  if (/^FILENAME\b/i.test(trimmed)) {
    return /\\p\b/i.test(trimmed) ? { kind: "fileName", path: true, cached } : { kind: "fileName", cached };
  }
  if ((m = /^DOCPROPERTY\s+"([^"]*)"/i.exec(trimmed))) {
    return { kind: "docProperty", name: m[1], cached };
  }
  if ((m = /^PAGEREF\s+(\S+)/i.exec(trimmed))) {
    return { kind: "pageRef", target: m[1], show: "page", cached };
  }
  if ((m = /^SEQ\s+(\S+)([\s\S]*)$/i.exec(trimmed))) {
    // A `seq` FieldSpec can only carry {label, format}; a SEQ whose switches go
    // beyond the numeric-format switch — Word's chapter-relative `\s`, restart
    // `\r`, repeat `\c`, `\n`/`\h`, an explicit bookmark — cannot be modelled
    // without dropping them. Preserve those verbatim as an unsupported field
    // (SPEC-053/SPEC-056 round-trip fidelity) rather than reparsing lossily.
    if (seqSwitchesModellable(m[2])) {
      return { kind: "seq", label: m[1], format: numberFormatFromSwitch(m[2]), cached };
    }
    return { kind: "unsupported", code: instr, cached };
  }
  if ((m = /^REF\s+(\S+)([\s\S]*)$/i.exec(trimmed))) {
    return { kind: "ref", target: m[1], show: refShowFromSwitches(m[2]), cached };
  }
  // Outside the catalogue: preserved verbatim (SPEC-053 unsupported case).
  return { kind: "unsupported", code: instr, cached };
}

/** A `w:fldSimple` as a field run (SPEC-053, widening SPEC-029/046), or null
 * only when its instruction is a `MERGEFIELD` whose text isn't valid chip
 * syntax (the caller then preserves it as an opaque inline). The cached
 * result is the element's own text content — the value the file was saved
 * with, which unsupported fields display and every field round-trips. */
function parseFieldSimple(el: Element): Run | null {
  const instr = wordAttr(el, "instr");
  if (!instr) {
    return null;
  }
  const spec = fieldSpecFromInstr(instr, fieldResultText(el));
  if (!spec) {
    return null;
  }
  return createFieldSpecRun(applyFieldLock(spec, isFldLocked(wordAttr(el, "fldLock"))));
}

/** Whether a `w:fldLock` attribute value marks the field locked (SPEC-053). */
function isFldLocked(value: string | null): boolean {
  return value === "1" || value === "true" || value === "on";
}

/** Adds `locked: true` to an evaluated field when its source marked it locked. */
function applyFieldLock(spec: FieldSpec, locked: boolean): FieldSpec {
  if (!locked || !("cached" in spec) || spec.kind === "unsupported") {
    return spec;
  }
  return { ...spec, locked: true };
}

/** The concatenated visible text of a field's result runs — every `w:t`
 * descendant of a `w:fldSimple` (or, for a complex field, its result runs). */
function fieldResultText(el: Element): string {
  let text = "";
  for (const t of Array.from(el.getElementsByTagNameNS(WORD_NAMESPACE, "t"))) {
    text += t.textContent ?? "";
  }
  return text;
}

/** Consumes a run-split complex field starting at `children[start]` (a
 * `w:fldChar begin` run), returning the field run it reads into and the index
 * just past its `end` run (SPEC-053 AC-47). The instruction is the text of the
 * `instrText` runs before `separate`; the cached result is the text of the
 * result runs between `separate` and the matching `end`. Nested fields (a
 * PAGEREF cached inside a REF result, say) only deepen the nesting count so
 * the outer `end` is matched correctly; their text still contributes to the
 * visible result. Returns null for a TOC field (handled at block level) so the
 * caller falls back to its ordinary run handling. */
function consumeComplexField(children: Element[], start: number): { run: Run; next: number } | null {
  let depth = 0;
  let sawSeparate = false;
  let endFound = false;
  let locked = false;
  let instr = "";
  let cached = "";
  let i = start;
  for (; i < children.length; i++) {
    const el = children[i];
    const isRun = el.namespaceURI === WORD_NAMESPACE && el.localName === "r";
    if (!isRun) {
      if (depth >= 1 && sawSeparate) {
        cached += fieldResultText(el);
      }
      continue;
    }
    const type = runFldCharType(el);
    if (type === "begin") {
      if (depth === 0) {
        locked = isFldLocked(runFldCharLock(el));
      }
      depth += 1;
      continue;
    }
    if (type === "end") {
      depth -= 1;
      if (depth === 0) {
        endFound = true;
        i += 1;
        break;
      }
      continue;
    }
    if (type === "separate") {
      if (depth === 1) {
        sawSeparate = true;
      }
      continue;
    }
    if (depth === 1 && !sawSeparate) {
      instr += runInstrText(el) ?? "";
    } else if (depth >= 1 && sawSeparate) {
      cached += fieldResultText(el);
    }
  }
  // A degenerate `begin` with no matching `end` (a malformed or unmodelled
  // run) is not a field — leave it to the ordinary opaque-inline path. A TOC
  // field is handled at block level, so fall through there too.
  if (!endFound || /^\s*TOC\b/i.test(instr)) {
    return null;
  }
  const spec = fieldSpecFromInstr(instr, cached);
  return spec ? { run: createFieldSpecRun(applyFieldLock(spec, locked)), next: i } : null;
}

/** The `w:fldLock` attribute of a run's `w:fldChar`, or null. */
function runFldCharLock(rEl: Element): string | null {
  for (const child of Array.from(rEl.children)) {
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "fldChar") {
      return wordAttr(child, "fldLock");
    }
  }
  return null;
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
async function parseNotesPart(
  partPath: string,
  parts: Map<string, Uint8Array>,
  styleCatalog: StyleCatalog,
  numberingCatalog: NumberingCatalog,
  kind: NoteKind,
): Promise<{ notes: NoteContent[]; media: MediaPart[] }> {
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
  // SPEC-047 AC-16's ≤150ms main-thread-block-during-open ceiling: unlike
  // the body walk above, this loop had no `await` of its own — a fixture
  // with hundreds of footnotes/endnotes (each parsed exactly like a body
  // paragraph) read as one long synchronous task on top of the chunked body
  // walk's own budget.
  let lastNoteYield = performance.now();
  for (const noteEl of childElements(root, WORD_NAMESPACE, tag)) {
    if (performance.now() - lastNoteYield > 4) {
      await yieldToMainThread();
      lastNoteYield = performance.now();
    }
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
export interface ParseDocxOptions {
  /** SPEC-047: when true, media parts are left base64-unencoded (their raw
   * bytes held in core/media's pending registry) so the ~1s encode of a
   * large document's images does not sit on the open path between choosing
   * the file and first paint. Every consumer force-encodes on demand via
   * `ensureMediaEncoded`/`mediaDataUri`, and the interactive open path runs a
   * chunked background pass after first paint. Default false: direct callers
   * (tests, headless generate, template import) get fully-encoded media as
   * before. */
  deferMedia?: boolean;
}

export async function parseDocx(
  data: ArrayBuffer | Uint8Array,
  preunzipped?: Map<string, Uint8Array>,
  options: ParseDocxOptions = {},
): Promise<ParsedDocxDocument> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

  if (isLegacyDoc(bytes)) {
    throw new LegacyDocError(LEGACY_DOC_MESSAGE);
  }

  // The shared open funnel (openDocument.ts) already inflated the package to
  // classify it; reuse those parts so a large `.docx` is decompressed once,
  // not twice. Direct callers pass only `data` and unzip here as before.
  let parts: Map<string, Uint8Array>;
  if (preunzipped) {
    parts = preunzipped;
  } else {
    try {
      parts = await unzip(bytes);
    } catch (error) {
      if (error instanceof ZipError) {
        throw new DocxOpenError(INVALID_DOCX_MESSAGE);
      }
      throw error;
    }
  }

  const readPart = (path: string): string | undefined => {
    const raw = parts.get(path);
    return raw ? decodeUtf8(raw) : undefined;
  };

  let mainPartPath = "word/document.xml";
  let corePartPath: string | undefined;
  let customPartPath: string | undefined;
  const rootRelsText = readPart("_rels/.rels");
  if (rootRelsText) {
    const rootRels = parseRelationships(rootRelsText);
    const officeDocRel = Array.from(rootRels.values()).find(
      (rel) => rel.type === OFFICE_DOCUMENT_RELATIONSHIP,
    );
    if (officeDocRel) {
      mainPartPath = resolvePartPath("", officeDocRel.target);
    }
    // Core/custom properties are rebuilt on save (SPEC-041), so these are
    // the only parts besides the main document that are not retained
    // verbatim.
    const coreRel = Array.from(rootRels.values()).find(
      (rel) => rel.type === CORE_PROPERTIES_RELATIONSHIP,
    );
    if (coreRel) {
      corePartPath = resolvePartPath("", coreRel.target);
    }
    const customRel = Array.from(rootRels.values()).find(
      (rel) => rel.type === CUSTOM_PROPERTIES_RELATIONSHIP,
    );
    if (customRel) {
      customPartPath = resolvePartPath("", customRel.target);
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
  // SPEC-074 AC-30: the document-wide page background, applied to every
  // section below (the model holds it per section).
  const backgroundEl = firstChildElement(documentXml.documentElement, WORD_NAMESPACE, "background");
  const pageBackground = backgroundEl ? (ooxmlColorToCss(wordAttr(backgroundEl, "color")) ?? undefined) : undefined;

  let stylesDoc: Document | undefined;
  let numberingCatalog: NumberingCatalog = EMPTY_NUMBERING_CATALOG;
  let docTheme: DocTheme | undefined;
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
    // The theme part (SPEC-077 AC-25): word/theme/themeN.xml into
    // model.theme, its part path recorded so the writer can splice or
    // preserve exactly that part. `dirty` stays unset on open (AC-18).
    const themeRel = Array.from(mainRels.values()).find((rel) => rel.type === THEME_RELATIONSHIP);
    if (themeRel) {
      const themePartPath = resolvePartPath(mainPartPath, themeRel.target);
      const themeText = readPart(themePartPath);
      if (themeText) {
        docTheme = readDocxTheme(themeText, themePartPath);
      }
    }
  }
  const styleCatalog = parseStylesCatalog(stylesDoc);
  const drawingCtx: DrawingContext = {
    mainRels,
    mainPartPath,
    parts,
    anchorOrdinals: { next: 1 },
    groupOrdinals: { next: firstFreeGroupOrdinal(documentXml.documentElement) },
  };
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
  // SPEC-047 AC-16's ≤150ms main-thread-block-during-open ceiling: a large
  // document's body walk (thousands of paragraphs/tables/images, each doing
  // real DOM/XML work) has no `await` of its own otherwise, so it reads as
  // one long task. Time-budgeted rather than every-N-elements because a
  // table or a paragraph with a `w:drawing` costs far more per element than
  // a plain paragraph does.
  let lastBodyWalkYield = performance.now();
  for (let bodyIndex = 0; bodyIndex < bodyChildren.length; bodyIndex += 1) {
    if (performance.now() - lastBodyWalkYield > 4) {
      await yieldToMainThread();
      lastBodyWalkYield = performance.now();
    }
    const child = bodyChildren[bodyIndex];
    if (child.namespaceURI === WORD_NAMESPACE && child.localName === "p") {
      // A TOC field (SPEC-033 AC-19, AC-20): its own paragraphs are the
      // field's cached result, so the whole span becomes one TOC block
      // rather than the ordinary per-paragraph parse below.
      // A `TOC \c "Figure"` field is a table of figures (SPEC-055), not a
      // table of contents — detected first since both begin `TOC`.
      const tofLabel = tofFieldStartLabel(child);
      if (tofLabel !== null) {
        const { tof, endIndex } = parseTofField(bodyChildren, bodyIndex, tofLabel);
        current.blocks.push(tof);
        bodyIndex = endIndex;
        continue;
      }
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
    // SPEC-074 AC-30: page borders, read from the same sectPr the columns
    // come from.
    const pageBorders = parsePageBorders(sectPrEl);
    if (pageBorders) {
      section.pageBorders = pageBorders;
    }
    // The page background is document-wide in OOXML (w:background) but per
    // section in the model, so every section reads the same value — which is
    // also what keeps a same-format round trip stable (the writer emits the
    // first section's, they all agree, and nothing is reported dropped).
    if (pageBackground !== undefined) {
      section.pageBackground = pageBackground;
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
    const parsed = await parseNotesPart(partPath, parts, styleCatalog, numberingCatalog, kind);
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
  // The document's default tab interval (SPEC-054): w:defaultTabStop in
  // settings.xml. Absent leaves defaultStyles.defaultTabStopCm undefined
  // (the model's 1.27 cm default applies).
  const defaultTabMatch = settingsText
    ? /<[\w:]*defaultTabStop\b[^>]*\bw:val="(-?\d+)"/.exec(settingsText)
    : null;
  if (defaultTabMatch) {
    const twips = Number(defaultTabMatch[1]);
    if (Number.isFinite(twips) && twips > 0) {
      defaultStyles.defaultTabStopCm = twipsToCm(twips);
    }
  }
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
  // Each part that still needs base64: the raw byte length (to bound each
  // encode chunk below) and the encode itself, run in the package's part
  // order so `opaque` stays deterministic.
  interface PendingEncode {
    byteLength: number;
    encode: () => void;
  }
  const pendingEncodes: PendingEncode[] = [];
  const opaque: OpaquePart[] = [];
  for (const [name, data] of parts) {
    if (
      name !== mainPartPath &&
      name !== corePartPath &&
      name !== customPartPath &&
      !mediaById.has(name) &&
      !consumedHeaderFooterParts.has(name)
    ) {
      // Placeholder now, bytes encoded in the chunked pass below.
      const part: OpaquePart = { name, data: "" };
      opaque.push(part);
      pendingEncodes.push({
        byteLength: data.length,
        encode: () => {
          part.data = bytesToBase64(data);
        },
      });
    }
  }
  // SPEC-047: with deferMedia, the media parts are left in the pending
  // registry (data still ""), encoded lazily per part when first read and in
  // a chunked background pass after first paint — see core/media. Opaque
  // parts above are still encoded here: they are small XML, never rendered,
  // and read only on save, so deferring them buys nothing.
  if (!options.deferMedia) {
    for (const part of mediaById.values()) {
      pendingEncodes.push({
        byteLength: pendingMediaByteLength(part),
        encode: () => ensureMediaEncoded(part),
      });
    }
  }

  // The one deferred base64 pass for every preserved part and image any walk
  // above collected — see parseDrawing/ensureMediaEncoded: encoding per
  // drawing put ~0.6s of the benchmark fixture's open inside the body walk.
  // Encoded SYNCHRONOUSLY here rather than through FileReader.readAsDataURL:
  // a large `.docx` carries hundreds of small media parts, and one FileReader
  // round trip per part costs more in scheduling than `btoa` spends on the
  // bytes — on the benchmark fixture the round-trip pass measured ~2x the
  // synchronous one, and its per-part latency is exactly what a loaded gate
  // host inflates (SPEC-047 AC-3). AC-16's 150ms main-thread-block budget is
  // held by running the encodes in small byte-bounded chunks with one real
  // yield between them: `btoa` costs ~18ms per source MB on the factory
  // container, so CHUNK_BYTES = 2 MB keeps each chunk's block to ~40ms — well
  // under the budget with room for a loaded host — while the whole pass still
  // costs about half what the per-part round-trip pass did.
  const CHUNK_BYTES = 2 * 1024 * 1024;
  let chunkBytes = 0;
  for (let i = 0; i < pendingEncodes.length; i++) {
    pendingEncodes[i].encode();
    chunkBytes += pendingEncodes[i].byteLength;
    if (chunkBytes >= CHUNK_BYTES && i < pendingEncodes.length - 1) {
      chunkBytes = 0;
      await yieldToMainThread();
    }
  }

  // Document properties and metadata (SPEC-041): core.xml's fields plus
  // custom.xml's typed properties. `docProps/app.xml` is not parsed here —
  // see docProps.ts's header comment — so it stays a plain OpaquePart in the
  // loop above.
  const properties: DocumentProperties = corePartPath ? parseCoreProperties(readPart(corePartPath) ?? "") : {};
  if (customPartPath) {
    const { customProperties, unsupportedCustomProperties } = parseCustomProperties(readPart(customPartPath) ?? "");
    if (customProperties.length > 0) {
      properties.customProperties = customProperties;
    }
    if (unsupportedCustomProperties.length > 0) {
      properties.unsupportedCustomProperties = unsupportedCustomProperties;
    }
  }

  // SPEC-070: resolve every connector's a:stCxn/a:endCxn id references to the
  // target's block index now that every section's blocks (and their docPr
  // ids) are known.
  resolveDocxConnectorAttachments({ sections });

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
    properties,
  };
  if (docTheme) {
    result.theme = docTheme;
  }
  if (footnoteSettings) {
    result.footnoteSettings = footnoteSettings;
  }
  if (endnoteSettings) {
    result.endnoteSettings = endnoteSettings;
  }
  return result;
}
