// The .xlsx wire mapping of cell annotations (SPEC-135 "xlsx wire mapping",
// "The thread placeholder", "Writer rules" — normative there): legacy notes
// (`xl/commentsN.xml` + the VML part's note shapes), Excel 365 threads
// (`xl/threadedComments/threadedCommentN.xml`, mirrored as placeholders),
// the workbook's persons (`xl/persons/person.xml`) and the worksheet's
// `<hyperlinks>`. Framework-free: the small parts are read with
// xlsx/elementTree.ts and the scanner, the worksheet's tail with
// drawingXml.ts's worksheetTail — never a DOM.
//
// Reading (readXlsxSheetAnnotations; ids derived from the wire, so
// re-reading a written file twice yields identical ids):
// - notes: the worksheet's `comments` relationship → `<authors>` and each
//   `<comment ref authorId>` (its `<text>` runs read by SPEC-117's run
//   reader, `_xHHHH_` decoded); the VML part `<legacyDrawing r:id>` names →
//   the `x:ClientData ObjectType="Note"` shape at the same `x:Row`/
//   `x:Column`: `<x:Visible/>` → `visible`, the style's `width`/`height` pt
//   → `box` (absent when it is the default 144 × 79 px). `noteId` is
//   `n-${A1}`; `source` keeps the comment element and the shape verbatim. A
//   comment whose author is `tc={threadId}` is a thread's placeholder, never
//   a note.
// - threads: the `threadedComment` relationship → `threadedComment`s grouped
//   by `parentId` under their root (`threadId` = the root's `id`), messages
//   in document order (`dT` → `createdAt`, `personId` → `author` through the
//   persons, `<mentions>` kept), the root's `done` → `resolved`.
// - the subset: two comments on one cell, an unresolvable author (comment
//   `authorId` or message `personId`), a reply to no thread or unparseable
//   VML leave the sheet with NO notes and threads, and the writer then keeps
//   every one of those parts (and `<legacyDrawing>`) byte-verbatim —
//   annotations never half-parse. Likewise a `<hyperlink>` without a
//   readable `ref` or target leaves `hyperlinks` unread and the element
//   verbatim.
// - hyperlinks: `<hyperlink ref r:id location display tooltip xr:uid>`; an
//   `r:id` → its relationship's `Target` as `{ kind: "url", href,
//   subaddress?: location }`, a `location` alone → `{ kind: "place" }` in
//   canonical reference text; `linkId` = `xr:uid`, else `h-${A1}`.
//
// Writing (writeXlsxSheetAnnotations, for a sheet that is `dirty` or
// `annotationsDirty`; xlsx/writer.ts wires it into both save paths):
// - `<hyperlinks>` re-serialised from the model (removed when empty) and
//   `<legacyDrawing r:id>` present iff the sheet keeps a VML part, both
//   spliced into the worksheet XML at their CT_Worksheet places — nothing
//   else in the part changes. A url link reuses its relationship while that
//   still targets its href; a new one is appended after the highest `rId`;
//   an unreferenced hyperlink relationship stays (the SPEC-097 rule).
// - the comments part is rebuilt: the source authors verbatim, new ones
//   appended (append-only); one comment per note — its source element
//   verbatim while unchanged (its `ref` patched when it moved, its
//   `authorId` re-pointed when it no longer names the author) — and one
//   placeholder per thread, regenerated every time (threadPlaceholder.ts;
//   none on a cell whose note holds the legacy slot), in (row, col) order.
// - the VML part is SPLICED: a note or placeholder shape still on its cell
//   stays byte-verbatim, one that moved is regenerated in place from
//   vmlNoteGolden.ts (its box kept, its custom position not), a new one is
//   appended after the last note shape with the next free shape id, a
//   deleted annotation's shape goes, and every non-note shape (a form
//   control, an orphan note shape no comment names) stays verbatim. The
//   part survives while it holds any shape.
// - the threaded part is rebuilt: each unchanged message's element
//   verbatim (its `ref` patched when the thread moved), others generated
//   (`dT` = `createdAt`, `<mentions>` while carried).
// - parts are minted at the next free number (with their relationships,
//   Overrides and the `vml` Default) when the first note or thread arrives,
//   and their relationships dropped when the last goes — xlsx/writer.ts's
//   reachability rule then drops the part and its Override.
// - persons (writeXlsxPersons): the source part verbatim with every person
//   it lacks appended (the model's, then any a message names), or a new
//   part related from the workbook.
import { escapeXmlAttribute, relsPathFor } from "../../core/xml";
import { scanXml, XmlScanError, type XmlScanAttribute } from "../../core/xmlScanner";
import { formatA1, formatRange, parseA1, parseRange, type CellRef } from "../address";
import { canonicalPlaceLocation, compareRefs, linkIdAt, noteIdAt } from "../annotations";
import { buildGridMetrics } from "../geometry";
import type { CellHyperlink, CellNote, CellThread, Sheet, SpreadsheetModel, TextRun, ThreadMessage, ThreadPerson } from "../model";
import { applySplices, attributeValue, escapeAttributeValue, escapeElementText, setAttribute, startTagAt, type Splice } from "../xmlSplice";
import { worksheetTail } from "./drawingXml";
import { elementTree, nodeAttr, type XmlNode } from "./elementTree";
import { nextFreePartName, nextRelationshipId, partTextOf, readRels, relationshipKind, relationshipsOf, relativeTarget, spliceRels, targetPart, workbookPathOf, type PartLookup, type RelEntry } from "./package";
import { decodeXmlEscapes, encodeXlsxAttribute, encodeXlsxText, xlsxTextElement } from "./sharedStrings";
import { StringItemReader } from "./sheetXml";
import { isSpreadsheetMl, xlsxFontProperties } from "./styles";
import { isPlaceholderAuthor, placeholderAuthor, placeholderCommentXml, placeholderThreadId } from "./threadPlaceholder";
import { VML_NOTE_SHAPETYPE, vmlDrawingXml, vmlNoteShapeFor } from "./vmlNoteGolden";

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL_NS = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const REL_NAMESPACES: ReadonlySet<string> = new Set([REL_NS, STRICT_REL_NS]);
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const REVISION_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
export const THREADED_COMMENTS_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const VML_NS = "urn:schemas-microsoft-com:vml";
const VML_OFFICE_NS = "urn:schemas-microsoft-com:office:office";
const VML_EXCEL_NS = "urn:schemas-microsoft-com:office:excel";

export const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
export const THREADED_COMMENTS_CONTENT_TYPE = "application/vnd.ms-excel.threadedcomments+xml";
export const PERSON_CONTENT_TYPE = "application/vnd.ms-excel.person+xml";
export const VML_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing";
export const THREADED_COMMENT_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
export const PERSON_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2017/10/relationships/person";

/** What a sheet's annotations read as (each absent when it has none). */
export interface XlsxSheetAnnotations {
  notes?: CellNote[];
  threads?: CellThread[];
  hyperlinks?: CellHyperlink[];
}

// --- the parts, parsed ------------------------------------------------------------------

/** One `<comment>` of a comments part. */
interface SourceComment {
  ref: CellRef;
  authorId: number;
  text: string;
  runs?: TextRun[];
  start: number;
  end: number;
}

/** A comments part: its authors and comments with their offsets. */
interface CommentsPart {
  xml: string;
  prefix: string;
  rootTag: { start: number; end: number };
  rootClose: number;
  authors: string[];
  authorsElement?: { start: number; end: number; tagEnd: number; selfClosing: boolean; close: number };
  listElement?: { start: number; end: number; tagEnd: number; selfClosing: boolean };
  comments: SourceComment[];
  /** A `<comment>` the subset does not cover (no parseable cell `ref`). */
  broken: boolean;
}

const prefixOf = (name: string): string => (name.includes(":") ? `${name.slice(0, name.indexOf(":"))}:` : "");

/** A comments part read with the scanner (undefined: not well-formed). */
function parseCommentsPart(xml: string, indexedColors: readonly string[] | undefined): CommentsPart | undefined {
  const part: CommentsPart = { xml, prefix: "", rootTag: { start: 0, end: 0 }, rootClose: xml.length, authors: [], comments: [], broken: false };
  let depth = 0;
  let section: "authors" | "list" | null = null;
  let author: string | null = null;
  let comment: { attrs: XmlScanAttribute[]; start: number; text?: { text: string; runs?: TextRun[] } } | null = null;
  let reader: StringItemReader | null = null;
  const finishComment = (end: number): void => {
    const refText = comment!.attrs.find((attr) => attr.local === "ref" && attr.ns === undefined)?.value ?? "";
    const ref = parseA1(refText.trim());
    const authorId = Number(comment!.attrs.find((attr) => attr.local === "authorId" && attr.ns === undefined)?.value ?? "NaN");
    if (ref === null || !Number.isInteger(authorId)) {
      part.broken = true;
    } else {
      const read: SourceComment = { ref: { row: ref.row, col: ref.col }, authorId, text: comment!.text?.text ?? "", start: comment!.start, end };
      if (comment!.text?.runs !== undefined) {
        read.runs = comment!.text.runs;
      }
      part.comments.push(read);
    }
    comment = null;
  };
  try {
    scanXml(xml, (event) => {
      if (event.type === "open") {
        depth += 1;
        const main = isSpreadsheetMl(event.ns);
        if (depth === 1) {
          part.prefix = prefixOf(event.name);
          part.rootTag = { start: event.start, end: event.end };
        } else if (depth === 2 && main && event.local === "authors") {
          section = "authors";
          part.authorsElement = { start: event.start, end: event.end, tagEnd: event.end, selfClosing: event.selfClosing, close: event.start };
        } else if (depth === 2 && main && event.local === "commentList") {
          section = "list";
          part.listElement = { start: event.start, end: event.end, tagEnd: event.end, selfClosing: event.selfClosing };
        } else if (depth === 3 && section === "authors" && main && event.local === "author") {
          if (event.selfClosing) {
            part.authors.push("");
          } else {
            author = "";
          }
        } else if (depth === 3 && section === "list" && main && event.local === "comment") {
          comment = { attrs: event.attrs, start: event.start };
          if (event.selfClosing) {
            finishComment(event.end);
          }
        } else if (depth === 4 && comment !== null && main && event.local === "text") {
          if (event.selfClosing) {
            comment.text = { text: "" };
          } else {
            reader = new StringItemReader(indexedColors);
          }
        } else if (depth > 4 && reader !== null) {
          reader.feed(event);
        }
        if (event.selfClosing) {
          depth -= 1;
          if (depth === 1) {
            section = null;
          }
        }
        return;
      }
      if (event.type === "close") {
        if (depth > 4 && reader !== null) {
          reader.feed(event);
        } else if (depth === 4 && reader !== null && comment !== null) {
          comment.text = reader.finish();
          reader = null;
        } else if (depth === 3 && author !== null) {
          part.authors.push(decodeXmlEscapes(author));
          author = null;
        } else if (depth === 3 && comment !== null) {
          finishComment(event.end);
        } else if (depth === 2) {
          if (section === "authors" && part.authorsElement) {
            part.authorsElement.end = event.end;
            part.authorsElement.close = event.start;
          } else if (section === "list" && part.listElement) {
            part.listElement.end = event.end;
          }
          section = null;
        } else if (depth === 1) {
          part.rootClose = event.start;
        }
        depth -= 1;
        return;
      }
      if (author !== null) {
        author += event.text;
      } else if (reader !== null) {
        reader.feed(event);
      }
    });
  } catch (error) {
    if (error instanceof XmlScanError) {
      return undefined;
    }
    throw error;
  }
  return part;
}

/** One top-level `v:shape` of a VML part. */
interface VmlShape {
  start: number;
  end: number;
  xml: string;
  /** The shape's `_x0000_s<n>` numbers (`id`, `o:spid`). */
  numbers: number[];
  /** A note shape's cell, visibility and size. */
  note?: { ref: CellRef; visible: boolean; widthPt?: number; heightPt?: number };
}

interface VmlPart {
  xml: string;
  idmap: number | undefined;
  shapes: VmlShape[];
  /** Where a note shape may be inserted: after the note shapetype, else after the shape layout. */
  shapetypeEnd: number | undefined;
  layoutEnd: number | undefined;
  rootTagEnd: number;
  rootClose: number;
}

const SHAPE_NUMBER = /_x0000_s(\d+)/;

function vmlChild(node: XmlNode, ns: string, local: string): XmlNode | undefined {
  return node.children.find((child) => child.ns === ns && child.local === local);
}

/** Whether `node`'s subtree holds a Note ClientData. */
function holdsNoteData(node: XmlNode): boolean {
  return node.children.some((child) => (child.ns === VML_EXCEL_NS && child.local === "ClientData" && nodeAttr(child, "ObjectType") === "Note") || holdsNoteData(child));
}

/** The pt value of `name` in a VML style, when written in pt. */
function stylePt(style: string, name: string): number | undefined {
  const match = new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*(-?[\\d.]+)pt`, "i").exec(style);
  return match ? Number(match[1]) : undefined;
}

/** A VML part read as a tree (undefined: unparseable, or a note shape
 * somewhere the subset does not reach — nested, or without its cell). */
function parseVmlPart(xml: string): VmlPart | undefined {
  const tree = elementTree(xml);
  const root = tree?.find((node) => node.local === "xml");
  if (!root) {
    return undefined;
  }
  const part: VmlPart = { xml, idmap: undefined, shapes: [], shapetypeEnd: undefined, layoutEnd: undefined, rootTagEnd: root.tagEnd, rootClose: root.end - `</${root.name}>`.length };
  for (const child of root.children) {
    if (child.ns === VML_OFFICE_NS && child.local === "shapelayout") {
      part.layoutEnd = child.end;
      const idmap = vmlChild(child, VML_OFFICE_NS, "idmap");
      const first = Number((idmap === undefined ? undefined : nodeAttr(idmap, "data"))?.split(",")[0]);
      if (Number.isInteger(first) && first > 0) {
        part.idmap = first;
      }
    } else if (child.ns === VML_NS && child.local === "shapetype" && nodeAttr(child, "id") === "_x0000_t202") {
      part.shapetypeEnd = child.end;
    } else if (child.ns === VML_NS && child.local === "shape") {
      const shape: VmlShape = { start: child.start, end: child.end, xml: xml.slice(child.start, child.end), numbers: [] };
      const spid = child.attrs.find((attr) => attr.local === "spid" && attr.ns === VML_OFFICE_NS)?.value;
      for (const id of [nodeAttr(child, "id"), spid]) {
        const match = id === undefined ? null : SHAPE_NUMBER.exec(id);
        if (match) {
          shape.numbers.push(Number(match[1]));
        }
      }
      const data = vmlChild(child, VML_EXCEL_NS, "ClientData");
      if (data !== undefined && nodeAttr(data, "ObjectType") === "Note") {
        const row = Number(vmlChild(data, VML_EXCEL_NS, "Row")?.text.trim() ?? "NaN");
        const col = Number(vmlChild(data, VML_EXCEL_NS, "Column")?.text.trim() ?? "NaN");
        if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) {
          return undefined;
        }
        const style = nodeAttr(child, "style") ?? "";
        shape.note = { ref: { row, col }, visible: vmlChild(data, VML_EXCEL_NS, "Visible") !== undefined };
        const widthPt = stylePt(style, "width");
        const heightPt = stylePt(style, "height");
        if (widthPt !== undefined && heightPt !== undefined) {
          shape.note.widthPt = widthPt;
          shape.note.heightPt = heightPt;
        }
      } else if (holdsNoteData(child)) {
        return undefined;
      }
      part.shapes.push(shape);
    } else if (holdsNoteData(child)) {
      return undefined;
    }
  }
  return part;
}

/** One `threadedComment` of a threaded-comments part. */
interface SourceMessage {
  ref: CellRef;
  id: string;
  parentId?: string;
  personId: string;
  dT?: string;
  done: boolean;
  text: string;
  mentions?: string;
  start: number;
  end: number;
}

interface ThreadedPart {
  xml: string;
  prefix: string;
  rootTagEnd: number;
  rootClose: number;
  messages: SourceMessage[];
}

/** A threaded-comments part (undefined: unparseable, or a message without
 * its `id`, `ref` or `personId`). */
function parseThreadedPart(xml: string): ThreadedPart | undefined {
  const root = elementTree(xml)?.[0];
  if (!root) {
    return undefined;
  }
  const part: ThreadedPart = { xml, prefix: prefixOf(root.name), rootTagEnd: root.tagEnd, rootClose: root.end - `</${root.name}>`.length, messages: [] };
  if (root.selfClosing) {
    part.rootClose = root.end;
  }
  for (const child of root.children) {
    if (child.local !== "threadedComment") {
      continue;
    }
    const ref = parseA1((nodeAttr(child, "ref") ?? "").trim());
    const id = nodeAttr(child, "id");
    const personId = nodeAttr(child, "personId");
    if (ref === null || id === undefined || personId === undefined) {
      return undefined;
    }
    const done = nodeAttr(child, "done");
    const message: SourceMessage = {
      ref: { row: ref.row, col: ref.col },
      id,
      personId,
      done: done === "1" || done === "true",
      text: child.children.find((each) => each.local === "text")?.text ?? "",
      start: child.start,
      end: child.end,
    };
    const parentId = nodeAttr(child, "parentId");
    if (parentId !== undefined) {
      message.parentId = parentId;
    }
    const dT = nodeAttr(child, "dT");
    if (dT !== undefined) {
      message.dT = dT;
    }
    const mentions = child.children.find((each) => each.local === "mentions");
    if (mentions !== undefined) {
      message.mentions = xml.slice(mentions.start, mentions.end);
    }
    part.messages.push(message);
  }
  return part;
}

/** A persons part's people, each with its element verbatim (undefined: unparseable). */
function parsePersonsPart(xml: string): { persons: ThreadPerson[]; rootClose: number; prefix: string; selfClosing: boolean; rootTagEnd: number } | undefined {
  const root = elementTree(xml)?.[0];
  if (!root) {
    return undefined;
  }
  const persons: ThreadPerson[] = [];
  for (const child of root.children) {
    const personId = nodeAttr(child, "id");
    if (child.local !== "person" || personId === undefined) {
      continue;
    }
    const person: ThreadPerson = { personId, displayName: nodeAttr(child, "displayName") ?? "", source: xml.slice(child.start, child.end) };
    const userId = nodeAttr(child, "userId");
    if (userId !== undefined) {
      person.userId = userId;
    }
    const providerId = nodeAttr(child, "providerId");
    if (providerId !== undefined) {
      person.providerId = providerId;
    }
    persons.push(person);
  }
  return {
    persons,
    rootClose: root.selfClosing ? root.end : root.end - `</${root.name}>`.length,
    prefix: prefixOf(root.name),
    selfClosing: root.selfClosing,
    rootTagEnd: root.tagEnd,
  };
}

// --- locating a sheet's parts -----------------------------------------------------------------

/** The part one of a sheet's relationships targets. */
interface RelatedPart {
  rel: RelEntry;
  partName: string;
  xml: string;
}

/** The relationship id an element's `r:id` holds. */
function relIdOf(attrs: readonly XmlScanAttribute[]): string | undefined {
  return attrs.find((attr) => attr.local === "id" && attr.ns !== undefined && REL_NAMESPACES.has(attr.ns))?.value;
}

/** A sheet's annotation parts as the package has them. */
interface SheetParts {
  rels: RelEntry[];
  comments?: RelatedPart;
  vml?: RelatedPart;
  threaded?: RelatedPart;
}

function relatedPart(parts: PartLookup, sheetPart: string, rel: RelEntry | undefined): RelatedPart | undefined {
  const partName = rel === undefined ? undefined : targetPart(sheetPart, rel);
  const xml = partName === undefined ? undefined : partTextOf(parts, partName);
  return rel !== undefined && partName !== undefined && xml !== undefined ? { rel, partName, xml } : undefined;
}

function sheetParts(parts: PartLookup, sheetPart: string, rels: RelEntry[], worksheetXml: string | undefined): SheetParts {
  const tail = worksheetXml === undefined ? undefined : worksheetTail(worksheetXml);
  const legacy = tail?.children.find((child) => child.main && child.local === "legacyDrawing");
  const vmlId = legacy === undefined ? undefined : relIdOf(legacy.attrs);
  const found: SheetParts = { rels };
  const comments = relatedPart(parts, sheetPart, rels.find((rel) => relationshipKind(rel.type) === "comments"));
  if (comments) {
    found.comments = comments;
  }
  const vml = relatedPart(parts, sheetPart, rels.find((rel) => rel.id === vmlId));
  if (vml) {
    found.vml = vml;
  }
  const threaded = relatedPart(parts, sheetPart, rels.find((rel) => relationshipKind(rel.type) === "threadedComment"));
  if (threaded) {
    found.threaded = threaded;
  }
  return found;
}

// --- reading -----------------------------------------------------------------------------------

/** A sheet's notes and threads read from its parts, with what the writer
 * needs to splice them; `modelled` false when they break the subset. */
interface ReadNotes {
  modelled: boolean;
  notes: CellNote[];
  threads: CellThread[];
  comments?: CommentsPart;
  vml?: VmlPart;
  threadedPart?: ThreadedPart;
  /** The source comment of each thread's placeholder, by thread id. */
  placeholders: Map<string, SourceComment>;
  /** The source comments by A1. */
  commentAt: Map<string, SourceComment>;
}

const UNMODELLED: ReadNotes = { modelled: false, notes: [], threads: [], placeholders: new Map(), commentAt: new Map() };

const PX_PER_PT = 4 / 3;

function readNotes(found: SheetParts, persons: readonly ThreadPerson[] | undefined, indexedColors: readonly string[] | undefined): ReadNotes {
  const out: ReadNotes = { modelled: true, notes: [], threads: [], placeholders: new Map(), commentAt: new Map() };
  if (found.comments) {
    const comments = parseCommentsPart(found.comments.xml, indexedColors);
    if (comments === undefined || comments.broken) {
      return UNMODELLED;
    }
    out.comments = comments;
  }
  if (found.vml) {
    const vml = parseVmlPart(found.vml.xml);
    if (vml === undefined) {
      return UNMODELLED;
    }
    out.vml = vml;
  }
  if (found.threaded) {
    const threaded = parseThreadedPart(found.threaded.xml);
    if (threaded === undefined) {
      return UNMODELLED;
    }
    out.threadedPart = threaded;
  }
  // The note shapes by cell: two on one cell break the subset.
  const shapeAt = new Map<string, VmlShape>();
  for (const shape of out.vml?.shapes ?? []) {
    if (shape.note) {
      const key = formatA1(shape.note.ref);
      if (shapeAt.has(key)) {
        return UNMODELLED;
      }
      shapeAt.set(key, shape);
    }
  }
  for (const comment of out.comments?.comments ?? []) {
    const key = formatA1(comment.ref);
    const author = out.comments!.authors[comment.authorId];
    if (out.commentAt.has(key) || author === undefined) {
      return UNMODELLED;
    }
    out.commentAt.set(key, comment);
    const threadId = placeholderThreadId(author);
    if (threadId !== undefined) {
      out.placeholders.set(threadId, comment);
      continue;
    }
    const note: CellNote = { noteId: noteIdAt(comment.ref), ref: comment.ref, author, text: comment.text, source: { xlsxComment: out.comments!.xml.slice(comment.start, comment.end) } };
    if (comment.runs !== undefined) {
      note.runs = comment.runs;
    }
    const shape = shapeAt.get(key);
    if (shape?.note) {
      note.source!.vmlShape = shape.xml;
      if (shape.note.visible) {
        note.visible = true;
      }
      if (shape.note.widthPt !== undefined && shape.note.heightPt !== undefined) {
        const widthPx = Math.round(shape.note.widthPt * PX_PER_PT * 100) / 100;
        const heightPx = Math.round(shape.note.heightPt * PX_PER_PT * 100) / 100;
        if (widthPx !== 144 || heightPx !== 79) {
          note.box = { widthPx, heightPx };
        }
      }
    }
    out.notes.push(note);
  }
  if (out.threadedPart) {
    const people = new Map((persons ?? []).map((person) => [person.personId, person] as const));
    const byId = new Map<string, CellThread>();
    for (const message of out.threadedPart.messages) {
      const person = people.get(message.personId);
      if (person === undefined) {
        return UNMODELLED;
      }
      const read: ThreadMessage = {
        messageId: message.id,
        author: person.displayName,
        text: message.text,
        createdAt: message.dT ?? "",
        personId: message.personId,
        source: out.threadedPart.xml.slice(message.start, message.end),
      };
      if (message.mentions !== undefined) {
        read.mentions = message.mentions;
      }
      if (message.parentId === undefined) {
        if (byId.has(message.id) || out.threads.some((thread) => compareRefs(thread.ref, message.ref) === 0)) {
          return UNMODELLED;
        }
        const thread: CellThread = { threadId: message.id, ref: message.ref, messages: [read], resolved: message.done };
        byId.set(message.id, thread);
        out.threads.push(thread);
      } else {
        const thread = byId.get(message.parentId);
        if (thread === undefined) {
          return UNMODELLED;
        }
        thread.messages.push(read);
      }
    }
  }
  out.notes.sort((a, b) => compareRefs(a.ref, b.ref));
  out.threads.sort((a, b) => compareRefs(a.ref, b.ref));
  return out;
}

/** A worksheet's `<hyperlinks>`: its element offsets and the links read
 * (`links` undefined when one breaks the subset). */
interface ReadLinks {
  element?: { start: number; end: number };
  links: CellHyperlink[] | undefined;
}

function readLinks(worksheetXml: string, rels: readonly RelEntry[]): ReadLinks {
  const tail = worksheetTail(worksheetXml);
  const child = tail.children.find((each) => each.main && each.local === "hyperlinks");
  if (child === undefined) {
    return { links: [] };
  }
  const element = { start: child.start, end: child.end };
  const node = elementTree(worksheetXml.slice(child.start, child.end), [tail.rootTag])?.[0];
  if (node === undefined) {
    return { element, links: undefined };
  }
  const byId = new Map(rels.map((rel) => [rel.id, rel] as const));
  const links: CellHyperlink[] = [];
  const taken = new Set<string>();
  for (const each of node.children) {
    if (each.local !== "hyperlink" || !isSpreadsheetMl(each.ns)) {
      continue;
    }
    const range = parseRange((nodeAttr(each, "ref") ?? "").trim());
    if (range === null) {
      return { element, links: undefined };
    }
    const relId = relIdOf(each.attrs);
    const location = nodeAttr(each, "location");
    const rel = relId === undefined ? undefined : byId.get(relId);
    if ((relId !== undefined && rel === undefined) || (relId === undefined && location === undefined)) {
      return { element, links: undefined };
    }
    const uid = each.attrs.find((attr) => attr.local === "uid" && attr.ns === REVISION_NS)?.value;
    let linkId = uid ?? linkIdAt(range);
    for (let n = 2; taken.has(linkId); n += 1) {
      linkId = `${uid ?? linkIdAt(range)}-${n}`;
    }
    taken.add(linkId);
    const link: CellHyperlink = {
      linkId,
      range,
      target:
        rel !== undefined
          ? { kind: "url", href: rel.target, ...(location !== undefined && location !== "" ? { subaddress: decodeXmlEscapes(location) } : {}) }
          : { kind: "place", location: canonicalPlaceLocation(decodeXmlEscapes(location!)) },
    };
    const tooltip = nodeAttr(each, "tooltip");
    if (tooltip !== undefined) {
      link.tooltip = decodeXmlEscapes(tooltip);
    }
    const display = nodeAttr(each, "display");
    if (display !== undefined) {
      link.display = decodeXmlEscapes(display);
    }
    if (relId !== undefined) {
      link.source = { relId };
    }
    links.push(link);
  }
  links.sort((a, b) => compareRefs({ row: a.range.top, col: a.range.left }, { row: b.range.top, col: b.range.left }));
  return { element, links };
}

/** The workbook's persons (xl/persons/person.xml through the workbook's
 * `person` relationship), each with its element as `source`; undefined
 * when the package has none (or it is unreadable). */
export function readXlsxPersons(parts: PartLookup): ThreadPerson[] | undefined {
  const workbookPath = workbookPathOf(parts);
  if (workbookPath === undefined) {
    return undefined;
  }
  const rel = relationshipsOf(parts, workbookPath).find((entry) => relationshipKind(entry.type) === "person");
  const partName = rel === undefined ? undefined : targetPart(workbookPath, rel);
  const xml = partName === undefined ? undefined : partTextOf(parts, partName);
  const parsed = xml === undefined ? undefined : parsePersonsPart(xml);
  return parsed === undefined ? undefined : parsed.persons;
}

/** A worksheet's notes, threads and hyperlinks (see the module header):
 * `sheetPart` the worksheet's part name, `worksheetXml` its text (a dirty
 * sheet's retained text, for a reader upgrade), `parts` the package's parts
 * (the same lookup the reader uses). `persons` are the workbook's
 * (readXlsxPersons — read from `parts` when not given); `indexedColors` the
 * workbook's palette override, for run colours. Reading sets nothing dirty. */
export function readXlsxSheetAnnotations(
  parts: PartLookup,
  sheetPart: string,
  worksheetXml: string,
  options: { persons?: readonly ThreadPerson[]; indexedColors?: readonly string[] } = {},
): XlsxSheetAnnotations {
  const rels = relationshipsOf(parts, sheetPart);
  const out: XlsxSheetAnnotations = {};
  const links = readLinks(worksheetXml, rels).links;
  if (links !== undefined && links.length > 0) {
    out.hyperlinks = links;
  }
  const found = sheetParts(parts, sheetPart, rels, worksheetXml);
  if (found.comments === undefined && found.vml === undefined && found.threaded === undefined) {
    return out;
  }
  const persons = options.persons ?? (found.threaded ? readXlsxPersons(parts) : undefined);
  const read = readNotes(found, persons, options.indexedColors);
  if (read.notes.length > 0) {
    out.notes = read.notes;
  }
  if (read.threads.length > 0) {
    out.threads = read.threads;
  }
  return out;
}

// --- writing: helpers ----------------------------------------------------------------------------

/** The prefix (with its colon) a start tag binds to `ns`, if any. */
function boundPrefix(tag: string, ns: string): string | undefined {
  for (const match of tag.matchAll(/\sxmlns:([\w.-]+)\s*=\s*(["'])(.*?)\2/g)) {
    if (match[3] === ns) {
      return `${match[1]}:`;
    }
  }
  return undefined;
}

/** A relationship element with an external target. */
function externalRelationshipXml(id: string, type: string, target: string): string {
  return `<Relationship Id="${escapeAttributeValue(id)}" Type="${escapeAttributeValue(type)}" Target="${escapeXmlAttribute(target)}" TargetMode="External"/>`;
}

function internalRelationshipXml(id: string, type: string, target: string): string {
  return `<Relationship Id="${escapeAttributeValue(id)}" Type="${escapeAttributeValue(type)}" Target="${escapeAttributeValue(target)}"/>`;
}

/** The officeDocument relationship base a `.rels` part uses. */
function relationshipBase(rels: readonly RelEntry[]): string {
  return rels.some((rel) => rel.type.startsWith(`${STRICT_REL_NS}/`)) ? STRICT_REL_NS : REL_NS;
}

const sameCell = (a: CellRef, b: CellRef): boolean => a.row === b.row && a.col === b.col;

/** A comment's `<text>` from plain text or runs (fonts through SPEC-117's writer). */
function commentTextXml(note: Pick<CellNote, "text" | "runs">, p: string): string {
  if (note.runs !== undefined && note.runs.length > 0) {
    return note.runs
      .map((run) => {
        const font = run.font !== undefined && Object.keys(run.font).length > 0 ? `<${p}rPr>${xlsxFontProperties(run.font, undefined, "rFont", p)}</${p}rPr>` : "";
        return `<${p}r>${font}${xlsxTextElement(run.text, p)}</${p}r>`;
      })
      .join("");
  }
  return xlsxTextElement(note.text, p);
}

/** A source element with one attribute of its start tag set. */
function withTagAttribute(element: string, name: string, value: string): string {
  const tag = startTagAt(element, 0);
  const next = setAttribute(tag, name, value);
  return next === tag ? element : next + element.slice(tag.length);
}

const ROW_IN_SHAPE = /<(?:[\w.-]+:)?Row>\s*(\d+)\s*<\/(?:[\w.-]+:)?Row>/;
const COLUMN_IN_SHAPE = /<(?:[\w.-]+:)?Column>\s*(\d+)\s*<\/(?:[\w.-]+:)?Column>/;

/** The cell a note shape's text names. */
function shapeCell(xml: string): CellRef | undefined {
  const row = ROW_IN_SHAPE.exec(xml);
  const col = COLUMN_IN_SHAPE.exec(xml);
  return row && col ? { row: Number(row[1]), col: Number(col[1]) } : undefined;
}

// --- writing: one sheet ---------------------------------------------------------------------------

/** What the save around one sheet's annotation write provides. */
export interface AnnotationWriteContext {
  /** The retained parts (the comment, VML and threaded parts' source). */
  parts: PartLookup;
  /** Part names in use; minted names are added. */
  taken: Set<string>;
  /** The next free VML `o:idmap` block. */
  nextIdmap(): number;
  /** The workbook's persons (`model.persons`): a source message whose
   * person is not among them broke the subset at read, and still does. */
  persons: readonly ThreadPerson[] | undefined;
}

/** What writing one sheet's annotations changes in its package. */
export interface SheetAnnotationWrite {
  /** Part texts written: comments, VML and threaded parts, the sheet `.rels`. */
  texts: Map<string, string>;
  /** Parts this write mints, each with its Override (none for VML, whose `vml` Default declares it). */
  added: { partName: string; contentType?: string }[];
  /** A VML part was minted. */
  vml: boolean;
  /** The sheet part's internal relationship targets after the write. */
  edges: string[];
  /** The worksheet XML with `<hyperlinks>` and `<legacyDrawing>` spliced. */
  worksheet(xml: string): string;
}

/** One sheet's annotations written (see the module header). `relsXml` is
 * the sheet's `.rels` as the save left it so far (the drawing and table
 * writes'), `retainedXml` the sheet's retained worksheet text (undefined for
 * a new sheet). Returns undefined when nothing changes. */
export function writeXlsxSheetAnnotations(
  sheet: Sheet,
  sheetPart: string,
  relsXml: string | undefined,
  retainedXml: string | undefined,
  ctx: AnnotationWriteContext,
): SheetAnnotationWrite | undefined {
  const rels = relsXml === undefined ? [] : readRels(relsXml);
  const found = sheetParts(ctx.parts, sheetPart, rels, retainedXml);
  const sourceLinks = retainedXml === undefined ? { links: [] as CellHyperlink[] } : readLinks(retainedXml, rels);
  const notes = sheet.notes ?? [];
  const threads = sheet.threads ?? [];
  const links = sheet.hyperlinks ?? [];
  const nothing = found.comments === undefined && found.vml === undefined && found.threaded === undefined && notes.length === 0 && threads.length === 0;
  if (nothing && links.length === 0 && (sourceLinks.links?.length ?? 0) === 0 && sourceLinks.element === undefined) {
    return undefined;
  }
  const texts = new Map<string, string>();
  const added: SheetAnnotationWrite["added"] = [];
  const removeRels = new Set<string>();
  const appendRels: string[] = [];
  const relIds = new Set(rels.map((rel) => rel.id));
  const newRelId = (): string => {
    const id = nextRelationshipId(relIds);
    relIds.add(id);
    return id;
  };
  const base = relationshipBase(rels);
  let mintedVml = false;

  // Hyperlinks: each url link's relationship (reused while it still
  // targets the href; else appended), unless the source broke the subset.
  let linkRelIds: Map<CellHyperlink, string> | undefined;
  if (sourceLinks.links !== undefined) {
    linkRelIds = new Map();
    for (const link of links) {
      if (link.target.kind !== "url") {
        continue;
      }
      const href = link.target.href;
      const known = link.source?.relId === undefined ? undefined : rels.find((rel) => rel.id === link.source!.relId);
      if (known !== undefined && known.target === href && relationshipKind(known.type) === "hyperlink") {
        linkRelIds.set(link, known.id);
      } else {
        const id = newRelId();
        appendRels.push(externalRelationshipXml(id, `${base}/hyperlink`, href));
        linkRelIds.set(link, id);
      }
    }
  }

  // Notes and threads, unless the source parts broke the subset.
  let legacyDrawing: string | null | undefined;
  const read = nothing ? undefined : readNotes(found, ctx.persons, undefined);
  if (read?.modelled) {
    const placeholderThreads = threads.filter((thread) => !notes.some((note) => sameCell(note.ref, thread.ref)));
    // The comments part.
    if (notes.length === 0 && placeholderThreads.length === 0) {
      if (found.comments) {
        removeRels.add(found.comments.rel.id);
      }
    } else {
      const partName = found.comments?.partName ?? mint(ctx.taken, "xl/comments1.xml");
      texts.set(partName, commentsPartXml(read.comments, notes, placeholderThreads));
      if (!found.comments) {
        added.push({ partName, contentType: COMMENTS_CONTENT_TYPE });
        appendRels.push(internalRelationshipXml(newRelId(), `${base}/comments`, relativeTarget(sheetPart, partName)));
      }
    }
    // The VML part.
    const vml = vmlPartXml(read, notes, placeholderThreads, sheet, ctx);
    if (vml === null) {
      if (found.vml) {
        removeRels.add(found.vml.rel.id);
      }
      legacyDrawing = null;
    } else {
      const partName = found.vml?.partName ?? mint(ctx.taken, "xl/drawings/vmlDrawing1.vml");
      if (vml !== found.vml?.xml) {
        texts.set(partName, vml);
      }
      if (found.vml) {
        legacyDrawing = found.vml.rel.id;
      } else {
        mintedVml = true;
        added.push({ partName });
        legacyDrawing = newRelId();
        appendRels.push(internalRelationshipXml(legacyDrawing, `${base}/vmlDrawing`, relativeTarget(sheetPart, partName)));
      }
    }
    // The threaded part.
    if (threads.length === 0) {
      if (found.threaded) {
        removeRels.add(found.threaded.rel.id);
      }
    } else {
      const partName = found.threaded?.partName ?? mint(ctx.taken, "xl/threadedComments/threadedComment1.xml");
      texts.set(partName, threadedPartXml(read.threadedPart, threads));
      if (!found.threaded) {
        added.push({ partName, contentType: THREADED_COMMENTS_CONTENT_TYPE });
        appendRels.push(internalRelationshipXml(newRelId(), THREADED_COMMENT_RELATIONSHIP_TYPE, relativeTarget(sheetPart, partName)));
      }
    }
  }

  // The sheet's relationships.
  if (removeRels.size > 0 || appendRels.length > 0) {
    const relsPath = relsPathFor(sheetPart);
    texts.set(relsPath, relsXml === undefined ? newRelsXml(appendRels) : spliceRels(relsXml, removeRels, appendRels));
  }
  const finalRels = texts.has(relsPathFor(sheetPart)) ? readRels(texts.get(relsPathFor(sheetPart))!) : rels;
  const edges = finalRels.flatMap((rel) => targetPart(sheetPart, rel) ?? []);
  const linkSplice = linkRelIds === undefined ? undefined : links.map((link) => ({ link, relId: linkRelIds!.get(link) }));
  return {
    texts,
    added,
    vml: mintedVml,
    edges,
    worksheet: (xml) => spliceWorksheetAnnotations(xml, linkSplice, legacyDrawing),
  };
}

function newRelsXml(relationships: readonly string[]): string {
  return `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>`;
}

function mint(taken: Set<string>, like: string): string {
  const name = taken.has(like) ? nextFreePartName(taken, like) : like;
  taken.add(name);
  return name;
}

/** The comments part: the source's authors (new ones appended) and every
 * note's comment and every thread's placeholder in (row, col) order. */
function commentsPartXml(source: CommentsPart | undefined, notes: readonly CellNote[], threads: readonly CellThread[]): string {
  const p = source?.prefix ?? "";
  const authors = [...(source?.authors ?? [])];
  const appended: string[] = [];
  const authorIndex = (name: string): number => {
    let at = authors.indexOf(name);
    if (at < 0) {
      authors.push(name);
      appended.push(name);
      at = authors.length - 1;
    }
    return at;
  };
  let rootTag = source === undefined ? "" : source.xml.slice(source.rootTag.start, source.rootTag.end);
  let xr = boundPrefix(rootTag, REVISION_NS);
  const entries = [...notes.map((note) => ({ ref: note.ref, note })), ...threads.map((thread) => ({ ref: thread.ref, thread }))].toSorted((a, b) => compareRefs(a.ref, b.ref));
  if (source !== undefined && xr === undefined && threads.length > 0) {
    // A placeholder's xr:uid needs the revision namespace (ignorable).
    const ignorable = attributeValue(rootTag, "mc:Ignorable");
    rootTag = setAttribute(rootTag, "xmlns:xr", REVISION_NS);
    rootTag = ignorable === undefined ? setAttribute(setAttribute(rootTag, "xmlns:mc", MC_NS), "mc:Ignorable", "xr") : setAttribute(rootTag, "mc:Ignorable", `${ignorable} xr`);
    xr = "xr:";
  }
  const comments = entries.map((entry) => {
    if ("thread" in entry && entry.thread !== undefined) {
      return placeholderCommentXml(entry.thread, formatA1(entry.ref), authorIndex(placeholderAuthor(entry.thread.threadId)), p, (xr ?? "xr:").slice(0, -1));
    }
    const note = (entry as { note: CellNote }).note;
    const a1 = formatA1(note.ref);
    if (note.source?.xlsxComment !== undefined) {
      let element = note.source.xlsxComment;
      const tag = startTagAt(element, 0);
      if (attributeValue(tag, "ref") !== a1) {
        element = withTagAttribute(element, "ref", a1);
      }
      const authorId = Number(attributeValue(tag, "authorId"));
      if (authors[authorId] !== note.author) {
        element = withTagAttribute(element, "authorId", String(authorIndex(note.author)));
      }
      return element;
    }
    return `<${p}comment ref="${a1}" authorId="${authorIndex(note.author)}" shapeId="0"><${p}text>${commentTextXml(note, p)}</${p}text></${p}comment>`;
  });
  const authorXml = (name: string): string => `<${p}author>${encodeXlsxText(name)}</${p}author>`;
  const list = `<${p}commentList>${comments.join("")}</${p}commentList>`;
  if (source === undefined) {
    return (
      `${XML_DECLARATION}<comments xmlns="${MAIN_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="xr" xmlns:xr="${REVISION_NS}">` +
      `<authors>${authors.map(authorXml).join("")}</authors>${list}</comments>`
    );
  }
  const splices: Splice[] = [];
  const originalTag = source.xml.slice(source.rootTag.start, source.rootTag.end);
  if (rootTag !== originalTag) {
    splices.push({ start: source.rootTag.start, end: source.rootTag.end, text: rootTag });
  }
  const authorsElement = source.authorsElement;
  if (authorsElement === undefined) {
    splices.push({ start: source.rootTag.end, end: source.rootTag.end, text: `<${p}authors>${appended.map(authorXml).join("")}</${p}authors>` });
  } else if (appended.length > 0) {
    if (authorsElement.selfClosing) {
      splices.push({ start: authorsElement.start, end: authorsElement.end, text: `<${p}authors>${appended.map(authorXml).join("")}</${p}authors>` });
    } else {
      splices.push({ start: authorsElement.close, end: authorsElement.close, text: appended.map(authorXml).join("") });
    }
  }
  const listElement = source.listElement;
  if (listElement === undefined) {
    const at = authorsElement?.end ?? source.rootTag.end;
    splices.push({ start: at, end: at, text: list });
  } else {
    const open = listElement.selfClosing ? `<${p}commentList>` : source.xml.slice(listElement.start, listElement.tagEnd);
    const text = `${open}${comments.join("")}</${p}commentList>`;
    if (source.xml.slice(listElement.start, listElement.end) !== text) {
      splices.push({ start: listElement.start, end: listElement.end, text });
    }
  }
  return applySplices(source.xml, splices);
}

/** The VML part after the write (see the module header) — the source's
 * bytes when nothing changed, a new part when the sheet had none, null when
 * no shape is left. */
function vmlPartXml(read: ReadNotes, notes: readonly CellNote[], threads: readonly CellThread[], sheet: Sheet, ctx: AnnotationWriteContext): string | null {
  const vml = read.vml;
  const shapes = vml?.shapes ?? [];
  const noteShapes = shapes.filter((shape) => shape.note !== undefined);
  // A source note shape some source comment names belongs to that comment
  // (and goes with it); one no comment names is an orphan, kept as is.
  const owned = (shape: VmlShape): boolean => read.commentAt.has(formatA1(shape.note!.ref));
  const claimed = new Map<VmlShape, { keep: true } | { regenerate: (n: number) => string }>();
  const fresh: ((n: number) => string)[] = [];
  const idmap = vml?.idmap ?? (notes.length > 0 || threads.length > 0 ? ctx.nextIdmap() : 1);
  let metrics: ReturnType<typeof buildGridMetrics> | undefined;
  const gridMetrics = (): ReturnType<typeof buildGridMetrics> => (metrics ??= buildGridMetrics(sheet));
  const noteShape =
    (ref: CellRef, visible: boolean, box?: CellNote["box"]) =>
    (n: number): string =>
      vmlNoteShapeFor(ref, { idmap, n, visible, ...(box ? { box } : {}), metrics: gridMetrics() });
  for (const note of notes) {
    const sourceShape = note.source?.vmlShape;
    let match: VmlShape | undefined;
    if (sourceShape !== undefined) {
      match = noteShapes.find((shape) => !claimed.has(shape) && owned(shape) && shape.xml === sourceShape);
      const cell = match === undefined ? shapeCell(sourceShape) : undefined;
      if (cell !== undefined) {
        match = noteShapes.find(
          (shape) => !claimed.has(shape) && owned(shape) && sameCell(shape.note!.ref, cell) && !isPlaceholderAuthorAt(read, shape.note!.ref),
        );
      }
    }
    const generate = noteShape(note.ref, note.visible === true, note.box);
    if (match === undefined) {
      fresh.push(generate);
    } else {
      claimed.set(match, sameCell(match.note!.ref, note.ref) ? { keep: true } : { regenerate: generate });
    }
  }
  for (const thread of threads) {
    const placeholder = read.placeholders.get(thread.threadId);
    const match = placeholder === undefined ? undefined : noteShapes.find((shape) => !claimed.has(shape) && sameCell(shape.note!.ref, placeholder.ref));
    const generate = noteShape(thread.ref, false);
    if (match === undefined) {
      fresh.push(generate);
    } else {
      claimed.set(match, sameCell(match.note!.ref, thread.ref) ? { keep: true } : { regenerate: generate });
    }
  }
  const removed = new Set(noteShapes.filter((shape) => !claimed.has(shape) && owned(shape)));
  const kept = shapes.filter((shape) => !removed.has(shape));
  if (kept.length === 0 && fresh.length === 0) {
    return null;
  }
  // Shape numbers: a shape regenerated in place keeps its own; a new one
  // takes the next above every number in use.
  let next = Math.max(idmap * 1024, ...kept.flatMap((shape) => shape.numbers));
  const nextN = (): number => {
    next += 1;
    return next - idmap * 1024;
  };
  if (vml === undefined) {
    return vmlDrawingXml(
      idmap,
      fresh.map((generate) => generate(nextN())),
    );
  }
  const splices: Splice[] = [];
  for (const shape of noteShapes) {
    const claim = claimed.get(shape);
    if (claim === undefined) {
      if (removed.has(shape)) {
        splices.push({ start: shape.start, end: shape.end, text: "" });
      }
    } else if ("regenerate" in claim) {
      const own = shape.numbers[0];
      const n = own !== undefined && own > idmap * 1024 ? own - idmap * 1024 : nextN();
      splices.push({ start: shape.start, end: shape.end, text: claim.regenerate(n) });
    }
  }
  if (fresh.length > 0) {
    const lastNote = noteShapes.at(-1);
    const at = lastNote?.end ?? vml.shapetypeEnd ?? vml.layoutEnd ?? vml.rootTagEnd;
    const shapetype = vml.shapetypeEnd === undefined ? VML_NOTE_SHAPETYPE : "";
    splices.push({ start: at, end: at, text: shapetype + fresh.map((generate) => generate(nextN())).join("") });
  }
  return applySplices(vml.xml, splices);
}

/** Whether the source comment on `ref` is a thread's placeholder. */
function isPlaceholderAuthorAt(read: ReadNotes, ref: CellRef): boolean {
  const comment = read.commentAt.get(formatA1(ref));
  const author = comment === undefined ? undefined : read.comments?.authors[comment.authorId];
  return author !== undefined && isPlaceholderAuthor(author);
}

/** The threaded part: every thread's messages in (row, col) order, an
 * unchanged message's element verbatim (its `ref`, `parentId` and the
 * root's `done` patched when they differ), others generated. */
function threadedPartXml(source: ThreadedPart | undefined, threads: readonly CellThread[]): string {
  const p = source?.prefix ?? "";
  const body = threads
    .flatMap((thread) =>
      thread.messages.map((message, index) => {
        const a1 = formatA1(thread.ref);
        const id = index === 0 ? thread.threadId : message.messageId;
        if (message.source !== undefined) {
          let element = message.source;
          const tag = startTagAt(element, 0);
          if (attributeValue(tag, "ref") !== a1) {
            element = withTagAttribute(element, "ref", a1);
          }
          if (index > 0 && attributeValue(tag, "parentId") !== thread.threadId) {
            element = withTagAttribute(element, "parentId", thread.threadId);
          }
          const done = attributeValue(tag, "done");
          if (index === 0 && (done === "1" || done === "true") !== thread.resolved) {
            const nextTag = setAttribute(startTagAt(element, 0), "done", thread.resolved ? "1" : null);
            element = nextTag + element.slice(startTagAt(element, 0).length);
          }
          return element;
        }
        return (
          `<${p}threadedComment ref="${a1}"${message.createdAt !== "" ? ` dT="${escapeAttributeValue(message.createdAt)}"` : ""} personId="${escapeAttributeValue(message.personId)}" id="${escapeAttributeValue(id)}"` +
          (index > 0 ? ` parentId="${escapeAttributeValue(thread.threadId)}"` : thread.resolved ? ` done="1"` : "") +
          `><${p}text>${escapeElementText(message.text)}</${p}text>${message.mentions ?? ""}</${p}threadedComment>`
        );
      }),
    )
    .join("");
  if (source === undefined) {
    return `${XML_DECLARATION}<ThreadedComments xmlns="${THREADED_COMMENTS_NS}" xmlns:x="${MAIN_NS}">${body}</ThreadedComments>`;
  }
  const first = source.messages[0];
  const last = source.messages.at(-1);
  if (first === undefined || last === undefined) {
    return source.xml.slice(0, source.rootClose) + body + source.xml.slice(source.rootClose);
  }
  return applySplices(source.xml, [{ start: first.start, end: last.end, text: body }]);
}

// --- the worksheet ------------------------------------------------------------------------------

/** A worksheet with its `<hyperlinks>` rewritten from `links` (each url
 * link with its relationship id; [] removes the element; undefined leaves it
 * as written) and its `<legacyDrawing r:id>` set (null removes it,
 * undefined leaves it), each at its CT_Worksheet place; every other byte
 * verbatim, the cells never scanned. */
export function spliceWorksheetAnnotations(
  xml: string,
  links: readonly { link: CellHyperlink; relId?: string | undefined }[] | undefined,
  legacyDrawing: string | null | undefined,
): string {
  if (links === undefined && legacyDrawing === undefined) {
    return xml;
  }
  const tail = worksheetTail(xml);
  const p = tail.prefix;
  const relPrefix = boundPrefix(tail.rootTag, REL_NS) ?? boundPrefix(tail.rootTag, STRICT_REL_NS);
  const splices: Splice[] = [];
  const place = (local: "hyperlinks" | "legacyDrawing", text: string): void => {
    const existing = tail.children.find((child) => child.main && child.local === local);
    if (existing !== undefined && local === "legacyDrawing" && legacyDrawing !== null && relIdOf(existing.attrs) === legacyDrawing) {
      return; // already names the part, however written
    }
    if (existing !== undefined) {
      if (xml.slice(existing.start, existing.end) !== text) {
        splices.push({ start: existing.start, end: existing.end, text });
      }
      return;
    }
    if (text === "") {
      return;
    }
    const rank = tailRank(local);
    const at = tail.children.find((child) => child.rank > rank)?.start ?? tail.rootClose;
    splices.push({ start: at, end: at, text });
  };
  if (links !== undefined) {
    place("hyperlinks", hyperlinksXml(links, p, relPrefix, boundPrefix(tail.rootTag, REVISION_NS)));
  }
  if (legacyDrawing !== undefined) {
    place(
      "legacyDrawing",
      legacyDrawing === null ? "" : `<${p}legacyDrawing${relPrefix === undefined ? ` xmlns:r="${REL_NS}"` : ""} ${relPrefix ?? "r:"}id="${escapeAttributeValue(legacyDrawing)}"/>`,
    );
  }
  return applySplices(xml, splices);
}

const tailRanks = new Map<string, number>();

/** An element's rank in drawingXml.ts's worksheet-tail order (what
 * worksheetTail reports for the children it finds). */
function tailRank(local: "hyperlinks" | "legacyDrawing"): number {
  let rank = tailRanks.get(local);
  if (rank === undefined) {
    rank = worksheetTail(`<worksheet xmlns="${MAIN_NS}"><sheetData/><${local}/></worksheet>`).children.find((child) => child.local === local)!.rank;
    tailRanks.set(local, rank);
  }
  return rank;
}

/** Whether a link id is a GUID (written as the link's `xr:uid`). */
function isGuid(id: string): boolean {
  return /^\{[0-9A-Fa-f-]{36}\}$/.test(id);
}

/** `<hyperlinks>` from the model (CT_Hyperlink attribute order: ref, r:id,
 * location, tooltip, display, then xr:uid when the link id is a GUID). */
function hyperlinksXml(links: readonly { link: CellHyperlink; relId?: string | undefined }[], p: string, relPrefix: string | undefined, xrPrefix: string | undefined): string {
  if (links.length === 0) {
    return "";
  }
  const declareR = relPrefix === undefined && links.some((each) => each.relId !== undefined) ? ` xmlns:r="${REL_NS}"` : "";
  const declareXr = xrPrefix === undefined && links.some((each) => isGuid(each.link.linkId)) ? ` xmlns:xr="${REVISION_NS}"` : "";
  const items = links.map(({ link, relId }) => {
    const target = link.target;
    let attrs = ` ref="${formatRange(link.range)}"`;
    if (relId !== undefined) {
      attrs += ` ${relPrefix ?? "r:"}id="${escapeAttributeValue(relId)}"`;
    }
    const location = target.kind === "place" ? target.location : target.subaddress;
    if (location !== undefined) {
      attrs += ` location="${encodeXlsxAttribute(location)}"`;
    }
    if (link.tooltip !== undefined) {
      attrs += ` tooltip="${encodeXlsxAttribute(link.tooltip)}"`;
    }
    if (link.display !== undefined) {
      attrs += ` display="${encodeXlsxAttribute(link.display)}"`;
    }
    if (isGuid(link.linkId)) {
      attrs += ` ${xrPrefix ?? "xr:"}uid="${link.linkId}"`;
    }
    return `<${p}hyperlink${attrs}/>`;
  });
  return `<${p}hyperlinks${declareR}${declareXr}>${items.join("")}</${p}hyperlinks>`;
}

// --- persons ------------------------------------------------------------------------------------

/** What writing the workbook's persons changes: the part's text and
 * whether it is new (it then needs the workbook relationship and Override). */
export interface PersonsWrite {
  partName: string;
  xml: string;
  minted: boolean;
}

/** The persons part after a save (see the module header), or undefined
 * when it already holds everyone. `parts` are the retained parts (empty for
 * a from-scratch package); `taken` the part names in use. */
export function writeXlsxPersons(model: Pick<SpreadsheetModel, "persons" | "sheets">, parts: PartLookup, taken: Set<string>): PersonsWrite | undefined {
  const wanted = new Map<string, ThreadPerson>();
  for (const person of model.persons ?? []) {
    if (!wanted.has(person.personId)) {
      wanted.set(person.personId, person);
    }
  }
  for (const sheet of model.sheets) {
    for (const thread of sheet.threads ?? []) {
      for (const message of thread.messages) {
        if (!wanted.has(message.personId)) {
          wanted.set(message.personId, { personId: message.personId, displayName: message.author, userId: message.author, providerId: "None" });
        }
      }
    }
  }
  if (wanted.size === 0) {
    return undefined;
  }
  const workbookPath = workbookPathOf(parts);
  const rel = workbookPath === undefined ? undefined : relationshipsOf(parts, workbookPath).find((entry) => relationshipKind(entry.type) === "person");
  const partName = rel === undefined || workbookPath === undefined ? undefined : targetPart(workbookPath, rel);
  const xml = partName === undefined ? undefined : partTextOf(parts, partName);
  const source = xml === undefined ? undefined : parsePersonsPart(xml);
  if (xml !== undefined && source === undefined) {
    return undefined; // unreadable: kept verbatim
  }
  const present = new Set(source?.persons.map((person) => person.personId) ?? []);
  const missing = [...wanted.values()].filter((person) => !present.has(person.personId));
  if (missing.length === 0) {
    return undefined;
  }
  const p = source?.prefix ?? "";
  const elements = missing.map(
    (person) =>
      `<${p}person displayName="${encodeXlsxAttribute(person.displayName)}" id="${escapeAttributeValue(person.personId)}"` +
      (person.userId !== undefined ? ` userId="${encodeXlsxAttribute(person.userId)}"` : "") +
      (person.providerId !== undefined ? ` providerId="${encodeXlsxAttribute(person.providerId)}"` : "") +
      `/>`,
  );
  if (source === undefined || xml === undefined || partName === undefined) {
    const name = mint(taken, "xl/persons/person.xml");
    return { partName: name, xml: `${XML_DECLARATION}<personList xmlns="${THREADED_COMMENTS_NS}" xmlns:x="${MAIN_NS}">${elements.join("")}</personList>`, minted: true };
  }
  if (source.selfClosing) {
    const tag = xml.slice(0, source.rootTagEnd);
    const open = tag.replace(/\s*\/>$/, ">");
    const name = /<([^\s/>]+)/.exec(open.slice(open.lastIndexOf("<")))?.[1] ?? "personList";
    return { partName, xml: `${open}${elements.join("")}</${name}>${xml.slice(source.rootTagEnd)}`, minted: false };
  }
  return { partName, xml: xml.slice(0, source.rootClose) + elements.join("") + xml.slice(source.rootClose), minted: false };
}

/** The highest VML `o:idmap` block any part of `parts` uses. */
export function maxVmlIdmap(parts: PartLookup): number {
  let max = 0;
  for (const name of parts.keys()) {
    if (!name.toLowerCase().endsWith(".vml")) {
      continue;
    }
    const xml = partTextOf(parts, name) ?? "";
    for (const match of xml.matchAll(/<(?:[\w.-]+:)?idmap\b[^>]*\bdata\s*=\s*["']([^"']*)["']/g)) {
      for (const value of match[1]!.split(",")) {
        const n = Number(value.trim());
        if (Number.isInteger(n)) {
          max = Math.max(max, n);
        }
      }
    }
  }
  return max;
}

/** Whether a sheet's annotations are written by this save. */
export function annotationsRebuilt(sheet: Pick<Sheet, "kind" | "dirty" | "annotationsDirty">): boolean {
  return sheet.kind === "grid" && (sheet.dirty === true || sheet.annotationsDirty === true);
}
