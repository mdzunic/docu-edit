// ODS cell notes and links (SPEC-135 "ods wire mapping") — the .ods half of
// the annotation model (model.ts CellNote / CellHyperlink; the commands are
// annotations.ts, the xlsx wire xlsx/annotations.ts). ODF has notes only:
// a sheet read from an .ods never holds threads.
//
// Reading (reader.ts's scan hands each cell's pieces to OdsAnnotationCollector):
// - an `office:annotation` child of a `table:table-cell` is a note —
//   `dc:creator` → `author` ("" when absent), `dc:date` → `createdAt`, the
//   `text:p`/`text:h` paragraphs (at any depth: a list's too) joined with
//   "\n" → `text` (`text:s`, `text:tab` and `text:line-break` expanded as a
//   cell's are), `office:display="true"` → `visible`, `svg:width` ×
//   `svg:height` (odfLengthToCm, 96 px to the inch, to 1/100 px) → `box`,
//   `office:name` → `noteId` (else `n-${A1}`, suffixed when taken). The
//   element itself is `source.odsAnnotation`, verbatim, and the cell it was
//   read on `source.odsRef`. It is LIFTED OUT of the cell's preservedXml
//   (SPEC-122): modelled once, a cell holding nothing else is not stored;
//   only a cell's first annotation is a note (a second stays preserved).
// - the FIRST `text:a` in a cell's paragraphs is a link over that one cell:
//   `xlink:href="#…"` is a `place` — an ODS address (`#Sheet2.A1`,
//   `#'My Sheet'.A1:B2`, `#$Sheet2.$A$1`) through
//   odsRangeAddressToCanonical (a range end without its `.` and a sheet
//   name holding a space written bare retried); anything else after the `#` (`#MyName`) is the location as
//   written (a defined name) — and every other href a `url`, a dangerous
//   scheme included (kept, never opened: openHyperlinkTarget refuses it);
//   `office:title` → `tooltip`; `linkId` is `h-${A1}`. A `text:a` whose
//   only attributes are `xlink:href`, `xlink:type` and `office:title` is
//   modelled inline content (isModelledOdsLink): it does not make the
//   paragraph SPEC-122's verbatim `preservedText`. A later link in the same
//   cell is not modelled — the sanctioned delta: a cell that carried two
//   links keeps the first once its table is rewritten.
// - a repeated cell or row gives every repetition the note and the link
//   (each counted against the workbook's cell budget by the reader; the
//   office:name only names the first). Reading sets nothing dirty.
//
// Writing (tableXml.ts's TableWriter, for a rebuilt or new table — a clean
// table is its sourceXml, byte-identical):
// - each note is the FIRST child of its `table:table-cell` (odsNoteXml): its
//   `source.odsAnnotation` verbatim while it has one (a command that changes
//   the note drops it). A note an adjuster or remap MOVED keeps its source,
//   but the element's caption position is the old cell's — Calc stores the
//   caption rectangle (`svg:x`/`svg:y`) and its tail (`draw:caption-point-x`
//   /`-y`) absolutely — so a note written on another cell than
//   `source.odsRef` drops those four attributes and keeps its size; Calc
//   then places the caption beside its cell itself (the ODS face of the
//   sanctioned delta "a moved note's box keeps its size but not its custom
//   position"). Otherwise it is generated: `<office:annotation
//   xmlns:dc="…" [office:display="true"] [office:name] [svg:width
//   svg:height]><dc:creator/><dc:date/><text:p/>…</office:annotation>` —
//   `xmlns:dc` declared locally as the text writer's annotationStartXml
//   does, `office:name` only for an id that is not a read-time `n-${A1}`,
//   the box in cm, the creator and date only when present, one paragraph
//   per line in the cell writer's paragraph form.
// - a note on a position nothing else materialises gets its own cell
//   element (a repeated empty run splits around it) and row, never
//   repeat-compressed with a neighbour; the extent covers every note.
// - a linked cell's paragraph content is wrapped, paragraph by paragraph,
//   in `<text:a xlink:type="simple" xlink:href=… [office:title=…]>`
//   (odsLinkHref: a `place` as `#` + canonicalToOdsRangeAddress, a name or
//   unconvertible location as `#` + itself; a `url` with its `#subaddress`).
//   A cell keeping SPEC-122's verbatim `preservedText` keeps it while its
//   first `text:a` still says what the model's link says (none and none
//   included); otherwise the text is regenerated from the value with the
//   model's link. A linked position with no text carries no link.
//
// The reader upgrade (open.ts's ReaderUpgrade calls upgradeOdsSheetAnnotations
// for a dirty sheet): a sheet stored before SPEC-135 holds its annotations
// in its cells' preservedXml (liftOdsSheetNotes: each anchored to the cell's
// CURRENT position) — or, stored before SPEC-122, nowhere but its retained
// sourceXml — and its links only in that sourceXml (retainedOdsAnnotations:
// an annotation or link at a source-time address is attached only where the
// cell there still holds the source's value; a value or structural edit
// since read makes the address unreliable, and nothing is resurrected).
import { DC_NAMESPACE, odfLengthToCm, OFFICE_NAMESPACE, SVG_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE } from "../../core/odfXml";
import { scanAttr, scanXml, type XmlScanAttribute } from "../../core/xmlScanner";
import { formatA1, getCell, type CellRef } from "../address";
import { isCellError, MAX_COLS, MAX_ROWS, toCellError, type Cell, type CellHyperlink, type CellNote, type CellValue, type HyperlinkTarget, type Row, type Sheet, type SpreadsheetModel } from "../model";
import { odsDateValueToSerial, odsTimeValueToSerial } from "../serial";
import { escapeAttributeText, escapeElementText } from "../xmlSplice";
import { canonicalToOdsRangeAddress, odsRangeAddressToCanonical } from "./rangeAddress";

/** LibreOffice's extension namespace (`calcext:value-type="error"`). */
const CALCEXT_NAMESPACE = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";

const PX_PER_CM = 96 / 2.54;

/** An element or attribute name test that resolves through the namespace
 * when the prefix is bound, and falls back to the qualified name a
 * fragment parsed without its declarations shows. */
function named(item: { ns: string | undefined; local: string; name: string }, ns: string, local: string, qname: string): boolean {
  return item.ns === undefined ? item.name === qname : item.ns === ns && item.local === local;
}

function attr(attrs: readonly XmlScanAttribute[], ns: string, local: string, qname: string): string | undefined {
  return attrs.find((each) => named(each, ns, local, qname))?.value;
}

/** The `xmlns` declarations on the content envelope prefix's root element. */
function prefixDeclarations(prefix: string): string {
  const root = /<[\w.-]+:document-content\b[^>]*>/.exec(prefix)?.[0] ?? "";
  return [...root.matchAll(/\s(xmlns(?::[\w.-]+)?)\s*=\s*("[^"]*"|'[^']*')/g)].map((match) => `${match[1]}=${match[2]}`).join(" ");
}

/** Rounds to 1/100 — a box written in cm (4 decimals) reads back exactly. */
const round2 = (n: number): number => Math.round(n * 100) / 100;

// --- reading ----------------------------------------------------------------------------------

/** A note's fields as an `office:annotation` gives them. */
export interface OdsAnnotationFields {
  name?: string;
  author: string;
  createdAt?: string;
  text: string;
  visible: boolean;
  box?: { widthPx: number; heightPx: number };
}

/** Parses one verbatim `office:annotation` element, under the document
 * element's namespace declarations (`declarations`, "" for none — the
 * qualified names then decide). Null when it does not parse. */
export function readOdsAnnotation(xml: string, declarations: string): OdsAnnotationFields | null {
  const wrapped = `<ods-annotation${declarations === "" ? "" : ` ${declarations}`}>${xml}</ods-annotation>`;
  const fields: OdsAnnotationFields = { author: "", text: "", visible: false };
  const paragraphs: string[] = [];
  let depth = 0;
  let found = false;
  /** What the open text belongs to: dc:creator, dc:date or a paragraph. */
  let field: "creator" | "date" | null = null;
  let fieldDepth = -1;
  let creator = "";
  let date = "";
  let paragraph: string | null = null;
  let paragraphDepth = -1;
  try {
    scanXml(wrapped, (event) => {
      if (event.type === "text") {
        if (paragraph !== null) {
          paragraph += event.text;
        } else if (field === "creator") {
          creator += event.text;
        } else if (field === "date") {
          date += event.text;
        }
        return undefined;
      }
      if (event.type === "close") {
        if (depth === paragraphDepth && paragraph !== null) {
          paragraphs.push(paragraph);
          paragraph = null;
        } else if (depth === fieldDepth) {
          field = null;
        }
        depth -= 1;
        return depth === 0 ? "stop" : undefined;
      }
      const level = depth + 1;
      if (level === 2) {
        if (found || !named(event, OFFICE_NAMESPACE, "annotation", "office:annotation")) {
          return "stop";
        }
        found = true;
        fields.visible = attr(event.attrs, OFFICE_NAMESPACE, "display", "office:display") === "true";
        const name = attr(event.attrs, OFFICE_NAMESPACE, "name", "office:name");
        if (name !== undefined && name !== "") {
          fields.name = name;
        }
        const width = odfLengthToCm(attr(event.attrs, SVG_NAMESPACE, "width", "svg:width"));
        const height = odfLengthToCm(attr(event.attrs, SVG_NAMESPACE, "height", "svg:height"));
        if (width !== undefined && height !== undefined && width > 0 && height > 0) {
          fields.box = { widthPx: round2(width * PX_PER_CM), heightPx: round2(height * PX_PER_CM) };
        }
      } else if (level === 3 && paragraph === null && named(event, DC_NAMESPACE, "creator", "dc:creator")) {
        field = "creator";
        fieldDepth = level;
      } else if (level === 3 && paragraph === null && named(event, DC_NAMESPACE, "date", "dc:date")) {
        field = "date";
        fieldDepth = level;
      } else if (paragraph === null && level > 2 && (named(event, TEXT_NAMESPACE, "p", "text:p") || named(event, TEXT_NAMESPACE, "h", "text:h"))) {
        if (event.selfClosing) {
          paragraphs.push("");
        } else {
          paragraph = "";
          paragraphDepth = level;
        }
      } else if (paragraph !== null) {
        paragraph += inlineText(event);
      }
      if (!event.selfClosing) {
        depth += 1;
      } else if (level === 2) {
        return "stop";
      }
      return undefined;
    });
  } catch {
    return null;
  }
  if (!found) {
    return null;
  }
  fields.author = creator;
  if (date.trim() !== "") {
    fields.createdAt = date.trim();
  }
  fields.text = paragraphs.join("\n");
  return fields;
}

/** The text an inline paragraph element stands for: `text:s` (its count,
 * capped as the cell reader caps it), `text:tab`, `text:line-break`. */
function inlineText(event: { ns: string | undefined; local: string; name: string; attrs: readonly XmlScanAttribute[] }): string {
  if (named(event, TEXT_NAMESPACE, "s", "text:s")) {
    const count = Number(attr(event.attrs, TEXT_NAMESPACE, "c", "text:c") ?? 1);
    return " ".repeat(Number.isInteger(count) && count > 0 ? Math.min(count, 10_000) : 1);
  }
  if (named(event, TEXT_NAMESPACE, "tab", "text:tab")) {
    return "\t";
  }
  if (named(event, TEXT_NAMESPACE, "line-break", "text:line-break")) {
    return "\n";
  }
  return "";
}

/** `n-${A1}` — the id of a note the wire does not name (annotations.ts noteIdAt). */
const noteIdAt = (ref: CellRef): string => `n-${formatA1(ref)}`;

/** Whether an id is a read-time `n-${A1}` one (or a command's suffixed
 * twin) — never written as `office:name`. */
const isPositionalNoteId = (noteId: string): boolean => /^n-[A-Z]{1,3}[1-9][0-9]*(?:-[0-9]+)?$/.test(noteId);

/** `base`, suffixed until it is not in `taken` (which gains it). */
function freeId(base: string, taken: Set<string>): string {
  let id = base;
  for (let n = 2; taken.has(id); n += 1) {
    id = `${base}-${n}`;
  }
  taken.add(id);
  return id;
}

/** The note an `office:annotation` (verbatim) makes at `ref`, or null when
 * it does not parse. `taken` holds the sheet's note ids so far (it gains
 * this one's); `named` false for a repetition (the name names the first). */
export function odsNote(xml: string, declarations: string, ref: CellRef, taken: Set<string>, useName = true): CellNote | null {
  const fields = readOdsAnnotation(xml, declarations);
  if (!fields) {
    return null;
  }
  const noteId = freeId(useName && fields.name !== undefined && !taken.has(fields.name) ? fields.name : noteIdAt(ref), taken);
  const note: CellNote = { noteId, ref: { row: ref.row, col: ref.col }, author: fields.author, text: fields.text };
  if (fields.createdAt !== undefined) {
    note.createdAt = fields.createdAt;
  }
  if (fields.visible) {
    note.visible = true;
  }
  if (fields.box) {
    note.box = fields.box;
  }
  note.source = { odsAnnotation: xml, odsRef: { row: ref.row, col: ref.col } };
  return note;
}

/** A cell's first link as the wire gives it. */
export interface OdsWireLink {
  href: string;
  title?: string;
}

/** The attributes a modelled `text:a` may carry. */
function isModelledLinkAttribute(each: XmlScanAttribute): boolean {
  return named(each, XLINK_NAMESPACE, "href", "xlink:href") || named(each, XLINK_NAMESPACE, "type", "xlink:type") || named(each, OFFICE_NAMESPACE, "title", "office:title");
}

/** Whether an element inside a cell's paragraph is a link the model holds
 * (see the module header): a `text:a` with an `xlink:href` and no other
 * attribute than `xlink:type` and `office:title`. */
export function isModelledOdsLink(ns: string | undefined, local: string, attrs: readonly XmlScanAttribute[]): boolean {
  return ns === TEXT_NAMESPACE && local === "a" && scanAttr(attrs, "href", XLINK_NAMESPACE) !== undefined && attrs.every(isModelledLinkAttribute);
}

/** A `text:a` start tag's link (null without an href). */
export function odsWireLink(attrs: readonly XmlScanAttribute[]): OdsWireLink | null {
  const href = attr(attrs, XLINK_NAMESPACE, "href", "xlink:href");
  if (href === undefined) {
    return null;
  }
  const title = attr(attrs, OFFICE_NAMESPACE, "title", "office:title");
  return title === undefined || title === "" ? { href } : { href, title };
}

/** A `text:a` href's target (see the module header). */
export function odsHyperlinkTarget(href: string): HyperlinkTarget {
  if (!href.startsWith("#")) {
    return { kind: "url", href };
  }
  const address = href.slice(1);
  let canonical = odsRangeAddressToCanonical(address);
  if (canonical === null) {
    // Two lenient spellings: a range's end without its `.` (`Sheet.A1:B2`),
    // and a sheet name holding a space written bare — each retried.
    let retry = address.replace(/:(?=\$?[A-Za-z]+\$?[0-9]+$)/, ":.");
    const bare = /^(\$?)([^'.$][^.']*)\.(.+)$/.exec(retry);
    if (bare && /\s/.test(bare[2]!)) {
      retry = `${bare[1]}'${bare[2]}'.${bare[3]}`;
    }
    canonical = retry === address ? null : odsRangeAddressToCanonical(retry);
  }
  return { kind: "place", location: canonical ?? address };
}

/** The hyperlink a cell's first `text:a` makes over `ref`. */
export function odsHyperlink(link: OdsWireLink, ref: CellRef): CellHyperlink {
  const hyperlink: CellHyperlink = {
    linkId: `h-${formatA1(ref)}`,
    range: { top: ref.row, left: ref.col, bottom: ref.row, right: ref.col },
    target: odsHyperlinkTarget(link.href),
  };
  if (link.title !== undefined) {
    hyperlink.tooltip = link.title;
  }
  return hyperlink;
}

const compareRefs = (a: CellRef, b: CellRef): number => a.row - b.row || a.col - b.col;

/** A table's notes and links as reader.ts's scan meets its cells. */
export class OdsAnnotationCollector {
  private readonly notes: CellNote[] = [];
  private readonly links: CellHyperlink[] = [];
  private readonly noteIds = new Set<string>();

  constructor(private readonly declarations: () => string) {}

  /** The positions a cell (run) at (row, col), repeated `rows` × `cols`,
   * covers inside the grid — what the reader counts against the budget
   * before calling `cell`. */
  static positions(row: number, rows: number, col: number, cols: number): number {
    return Math.max(0, Math.min(row + rows, MAX_ROWS) - row) * Math.max(0, Math.min(col + cols, MAX_COLS) - col);
  }

  /** A cell's annotation (verbatim) and first link, given to every
   * position of its repeated run inside the grid. */
  cell(row: number, rows: number, col: number, cols: number, annotation: string | undefined, link: OdsWireLink | null): void {
    for (let r = row; r < Math.min(row + rows, MAX_ROWS); r += 1) {
      for (let c = col; c < Math.min(col + cols, MAX_COLS); c += 1) {
        const ref = { row: r, col: c };
        if (annotation !== undefined) {
          const note = odsNote(annotation, this.declarations(), ref, this.noteIds, r === row && c === col);
          if (note) {
            this.notes.push(note);
          }
        }
        if (link) {
          this.links.push(odsHyperlink(link, ref));
        }
      }
    }
  }

  /** The table's annotation fields, each sorted by (row, col), absent when empty. */
  result(): Pick<Sheet, "notes" | "hyperlinks"> {
    const out: Pick<Sheet, "notes" | "hyperlinks"> = {};
    if (this.notes.length > 0) {
      out.notes = this.notes.toSorted((a, b) => compareRefs(a.ref, b.ref));
    }
    if (this.links.length > 0) {
      out.hyperlinks = this.links.toSorted((a, b) => a.range.top - b.range.top || a.range.left - b.range.left);
    }
    return out;
  }
}

// --- writing ----------------------------------------------------------------------------------

/** The caption position attributes of a Calc annotation (any prefix). */
const CAPTION_POSITION = /\s+[\w.-]+:(?:x|y|caption-point-x|caption-point-y)\s*=\s*(?:"[^"]*"|'[^']*')/g;

/** A note's `office:annotation` (see the module header); `paragraphs`
 * writes text as the cell writer writes a cell's paragraphs. */
export function odsNoteXml(note: CellNote, paragraphs: (text: string) => string): string {
  const source = note.source?.odsAnnotation;
  if (source !== undefined) {
    const at = note.source?.odsRef;
    if (at === undefined || (at.row === note.ref.row && at.col === note.ref.col)) {
      return source;
    }
    const end = source.indexOf(">");
    return end < 0 ? source : source.slice(0, end).replace(CAPTION_POSITION, "") + source.slice(end);
  }
  let start = `<office:annotation xmlns:dc="${DC_NAMESPACE}"`;
  if (note.visible) {
    start += ` office:display="true"`;
  }
  if (!isPositionalNoteId(note.noteId)) {
    start += ` office:name="${escapeAttributeText(note.noteId)}"`;
  }
  if (note.box) {
    const cm = (px: number): string => `${Number(((px * 2.54) / 96).toFixed(4))}cm`;
    start += ` svg:width="${cm(note.box.widthPx)}" svg:height="${cm(note.box.heightPx)}"`;
  }
  let body = note.author === "" ? "" : `<dc:creator>${escapeElementText(note.author)}</dc:creator>`;
  if (note.createdAt !== undefined && note.createdAt !== "") {
    body += `<dc:date>${escapeElementText(note.createdAt)}</dc:date>`;
  }
  return `${start}>${body}${paragraphs(note.text)}</office:annotation>`;
}

/** The href a link is written with (see the module header). */
export function odsLinkHref(target: HyperlinkTarget): string {
  if (target.kind === "url") {
    return target.subaddress ? `${target.href}#${target.subaddress}` : target.href;
  }
  return `#${canonicalToOdsRangeAddress(target.location) ?? target.location}`;
}

/** Generated paragraphs (`<text:p>…</text:p>`, `<text:p/>` left alone),
 * each one's content wrapped in the link's `text:a`. */
export function linkedParagraphsXml(paragraphs: string, link: Pick<CellHyperlink, "target" | "tooltip">): string {
  const title = link.tooltip === undefined || link.tooltip === "" ? "" : ` office:title="${escapeAttributeText(link.tooltip)}"`;
  const open = `<text:a xlink:type="simple" xlink:href="${escapeAttributeText(odsLinkHref(link.target))}"${title}>`;
  return paragraphs.replace(/<text:p>([\s\S]*?)<\/text:p>/g, (_whole, content: string) => `<text:p>${open}${content}</text:a></text:p>`);
}

/** The first `text:a` link of verbatim paragraphs (`preservedText`), or null. */
export function firstOdsLink(paragraphs: string): OdsWireLink | null {
  let found: OdsWireLink | null = null;
  try {
    scanXml(`<ods-paragraphs>${paragraphs}</ods-paragraphs>`, (event) => {
      if (event.type === "open" && named(event, TEXT_NAMESPACE, "a", "text:a")) {
        found = odsWireLink(event.attrs);
        if (found) {
          return "stop";
        }
      }
      return undefined;
    });
  } catch {
    return null;
  }
  return found;
}

function sameTarget(a: HyperlinkTarget, b: HyperlinkTarget): boolean {
  if (a.kind === "place" || b.kind === "place") {
    return a.kind === "place" && b.kind === "place" && a.location === b.location;
  }
  return a.href === b.href && (a.subaddress ?? "") === (b.subaddress ?? "");
}

/** Whether verbatim paragraphs still say what the model's link (or its
 * absence) says: the first `text:a`'s target and title against the link's. */
export function paragraphsMatchLink(paragraphs: string, link: Pick<CellHyperlink, "target" | "tooltip"> | undefined): boolean {
  const wire = firstOdsLink(paragraphs);
  if (!wire || !link) {
    return !wire && !link;
  }
  return sameTarget(odsHyperlinkTarget(wire.href), link.target) && (wire.title ?? "") === (link.tooltip ?? "");
}

// --- the reader upgrade -----------------------------------------------------------------------

/** What liftOdsSheetNotes returns: the sheet's rows without the lifted
 * annotations (copy-on-write — the same array when no cell held one; a cell
 * left holding nothing dropped, a row left holding nothing too) and the
 * notes they made, sorted by (row, col). */
export interface OdsNotesUpgrade {
  rows: Row[];
  notes: CellNote[];
}

/** The top-level `office:annotation` elements of a cell's preservedXml and
 * the rest, verbatim and in order. */
function splitAnnotations(xml: string): { annotations: string[]; rest: string } {
  const annotations: string[] = [];
  let rest = "";
  let cursor = 0;
  let depth = 0;
  let open = -1;
  const offset = "<w>".length;
  const take = (start: number, end: number): void => {
    rest += xml.slice(cursor, start);
    annotations.push(xml.slice(start, end));
    cursor = end;
  };
  try {
    scanXml(`<w>${xml}</w>`, (event) => {
      if (event.type === "text") {
        return undefined;
      }
      if (event.type === "open") {
        const annotation = depth === 1 && event.local === "annotation" && (event.ns === undefined ? /^[\w.-]*:?annotation$/.test(event.name) : event.ns === OFFICE_NAMESPACE);
        if (annotation && event.selfClosing) {
          take(event.start - offset, event.end - offset);
        } else if (annotation) {
          open = event.start - offset;
        }
        if (!event.selfClosing) {
          depth += 1;
        }
        return undefined;
      }
      depth -= 1;
      if (depth === 1 && open >= 0) {
        take(open, event.end - offset);
        open = -1;
      }
      return undefined;
    });
  } catch {
    return { annotations: [], rest: xml };
  }
  return { annotations, rest: rest + xml.slice(cursor) };
}

function cellHoldsSomething(cell: Cell): boolean {
  const { col: _col, value, ...rest } = cell;
  return value !== null || Object.keys(rest).length > 0;
}

function rowHoldsSomething(row: Row): boolean {
  const { index: _index, cells, ...rest } = row;
  return cells.length > 0 || Object.keys(rest).length > 0;
}

/** SPEC-135's reader-upgrade step for a dirty .ods sheet's notes: every
 * `office:annotation` its cells' preservedXml hold (SPEC-122 kept them
 * there) becomes a note anchored to that cell's CURRENT position (the cell
 * moved with every edit since) — the first per cell; another stays
 * preserved. `prefix` is the retained content envelope's prefix (its root
 * declares the namespaces). The sheet is not touched. */
export function liftOdsSheetNotes(input: { sheet: Pick<Sheet, "rows">; prefix: string }): OdsNotesUpgrade {
  const { sheet } = input;
  const declarations = prefixDeclarations(input.prefix);
  const taken = new Set<string>();
  const notes: CellNote[] = [];
  let rows: Row[] | null = null;
  sheet.rows.forEach((row, at) => {
    if (!row.cells.some((cell) => cell.preservedXml !== undefined && cell.preservedXml.includes("annotation"))) {
      return;
    }
    let changed = false;
    const cells: Cell[] = [];
    for (const cell of row.cells) {
      const split = cell.preservedXml === undefined ? null : splitAnnotations(cell.preservedXml);
      const [first, ...others] = split?.annotations ?? [];
      const note = first === undefined ? null : odsNote(first, declarations, { row: row.index, col: cell.col }, taken);
      if (!note) {
        cells.push(cell);
        continue;
      }
      notes.push(note);
      changed = true;
      const rest = others.join("") + split!.rest;
      const next: Cell = { ...cell, preservedXml: rest };
      if (rest === "") {
        delete next.preservedXml;
      }
      if (cellHoldsSomething(next)) {
        cells.push(next);
      }
    }
    if (changed) {
      rows ??= sheet.rows.slice();
      rows[at] = { ...row, cells };
    }
  });
  return { rows: rows === null ? sheet.rows : (rows as Row[]).filter(rowHoldsSomething), notes: notes.toSorted((a, b) => compareRefs(a.ref, b.ref)) };
}

/** A link a retained table text carries: the cell (source-time address),
 * the hyperlink it makes there and the value the source cell held. */
export interface RetainedOdsLink {
  ref: CellRef;
  hyperlink: CellHyperlink;
  sourceValue: CellValue;
}

/** An annotation a retained table text carries: the cell (source-time
 * address), its verbatim element and the value the source cell held. */
export interface RetainedOdsAnnotation {
  ref: CellRef;
  xml: string;
  sourceValue: CellValue;
}

/** The source value of a cell as the reader computes it (reader.ts
 * odsCellValue) — null where it cannot be read. */
function sourceCellValue(attrs: readonly XmlScanAttribute[], text: string, dates: Pick<SpreadsheetModel, "dateSystem" | "odsNullDate">): CellValue {
  const office = (local: string): string | undefined => attr(attrs, OFFICE_NAMESPACE, local, `office:${local}`);
  const typed = office("value-type") !== undefined || attr(attrs, CALCEXT_NAMESPACE, "value-type", "calcext:value-type") !== undefined;
  if (attr(attrs, TABLE_NAMESPACE, "formula", "table:formula") !== undefined && !typed) {
    return null;
  }
  const valueType = attr(attrs, CALCEXT_NAMESPACE, "value-type", "calcext:value-type") === "error" ? "error" : office("value-type");
  try {
    switch (valueType) {
      case "float":
      case "percentage":
      case "currency": {
        const n = Number(office("value"));
        return Number.isFinite(n) ? n : null;
      }
      case "date":
        return odsDateValueToSerial(office("date-value") ?? "", dates.dateSystem, dates.odsNullDate);
      case "time":
        return odsTimeValueToSerial(office("time-value") ?? "");
      case "boolean": {
        const raw = office("boolean-value");
        return raw === "true" || raw === "1" ? true : raw === "false" || raw === "0" ? false : null;
      }
      case "string":
        return office("string-value") ?? text;
      case "error":
        return toCellError(text);
      default:
        return text.length > 0 ? text : null;
    }
  } catch {
    return null;
  }
}

/** Every cell annotation and link a retained table text (`sourceXml`)
 * carries — each cell's first `office:annotation` (verbatim) and first
 * `text:a`, at its source-time address (every repetition of a repeated
 * run) — with the value the cell held. `prefix` is the retained content
 * envelope's prefix (its root declares the namespaces). */
export function retainedOdsAnnotations(
  sourceXml: string | undefined,
  prefix: string,
  dates: Pick<SpreadsheetModel, "dateSystem" | "odsNullDate">,
): { annotations: RetainedOdsAnnotation[]; links: RetainedOdsLink[] } {
  const annotations: RetainedOdsAnnotation[] = [];
  const links: RetainedOdsLink[] = [];
  if (sourceXml === undefined || !/<[\w.-]+:(?:a|annotation)[\s>/]/.test(sourceXml)) {
    return { annotations, links };
  }
  const declarations = prefixDeclarations(prefix);
  const head = `<ods-table ${declarations}>`;
  const is = (event: { ns: string | undefined; local: string; name: string }, ns: string, local: string, prefixName: string): boolean => named(event, ns, local, `${prefixName}:${local}`);
  const count = (attrs: readonly XmlScanAttribute[], local: string): number => {
    const value = Number(attr(attrs, TABLE_NAMESPACE, local, `table:${local}`));
    return Number.isInteger(value) && value > 1 ? value : 1;
  };
  let depth = 0;
  let row = 0;
  let rowRepeat = 1;
  let rowLevel = -1;
  let col = 0;
  // The open cell: its level, attributes, paragraphs, first annotation and link.
  let cellLevel = -1;
  let cellAttrs: readonly XmlScanAttribute[] = [];
  let paragraphs: string[] = [];
  let paragraph: string | null = null;
  let paragraphLevel = -1;
  let skip = 0;
  let annotationStart = -1;
  let annotation: string | undefined;
  let link: OdsWireLink | null = null;
  const finishCell = (): void => {
    const repeat = count(cellAttrs, "number-columns-repeated");
    if (link || annotation !== undefined) {
      const value = sourceCellValue(cellAttrs, paragraphs.join("\n"), dates);
      for (let r = row; r < Math.min(row + rowRepeat, MAX_ROWS); r += 1) {
        for (let c = col; c < Math.min(col + repeat, MAX_COLS); c += 1) {
          const ref = { row: r, col: c };
          if (annotation !== undefined) {
            annotations.push({ ref, xml: annotation, sourceValue: value });
          }
          if (link) {
            links.push({ ref, hyperlink: odsHyperlink(link, ref), sourceValue: value });
          }
        }
      }
    }
    col += repeat;
  };
  const wrapped = `${head}${sourceXml}</ods-table>`;
  try {
    scanXml(wrapped, (event) => {
      if (event.type === "text") {
        if (paragraph !== null && skip === 0) {
          paragraph += event.text;
        }
        return undefined;
      }
      if (event.type === "close") {
        if (depth === cellLevel) {
          finishCell();
          cellLevel = -1;
        } else if (depth === rowLevel) {
          row += rowRepeat;
          rowLevel = -1;
        } else if (skip > 0) {
          skip -= 1;
          if (skip === 0 && annotationStart >= 0) {
            annotation ??= wrapped.slice(annotationStart, event.end);
            annotationStart = -1;
          }
        } else if (depth === paragraphLevel && paragraph !== null) {
          paragraphs.push(paragraph);
          paragraph = null;
        }
        depth -= 1;
        return undefined;
      }
      const level = depth + 1;
      if (cellLevel < 0) {
        if (rowLevel < 0 && is(event, TABLE_NAMESPACE, "table-row", "table")) {
          rowRepeat = count(event.attrs, "number-rows-repeated");
          col = 0;
          if (event.selfClosing) {
            row += rowRepeat;
          } else {
            rowLevel = level;
          }
        } else if (rowLevel >= 0 && level === rowLevel + 1 && (is(event, TABLE_NAMESPACE, "table-cell", "table") || is(event, TABLE_NAMESPACE, "covered-table-cell", "table"))) {
          cellAttrs = event.attrs;
          paragraphs = [];
          paragraph = null;
          link = null;
          annotation = undefined;
          annotationStart = -1;
          skip = 0;
          if (event.selfClosing) {
            finishCell();
          } else {
            cellLevel = level;
          }
        }
      } else if (skip > 0) {
        if (!event.selfClosing) {
          skip += 1;
        }
      } else if (level === cellLevel + 1 && !(is(event, TEXT_NAMESPACE, "p", "text") || is(event, TEXT_NAMESPACE, "h", "text"))) {
        // An annotation, a drawing, anything but a paragraph: not the cell's text.
        const isAnnotation = is(event, OFFICE_NAMESPACE, "annotation", "office");
        if (event.selfClosing) {
          if (isAnnotation) {
            annotation ??= wrapped.slice(event.start, event.end);
          }
        } else {
          skip = 1;
          annotationStart = isAnnotation && annotation === undefined ? event.start : -1;
        }
      } else if (level === cellLevel + 1) {
        if (event.selfClosing) {
          paragraphs.push("");
        } else {
          paragraph = "";
          paragraphLevel = level;
        }
      } else if (paragraph !== null) {
        if (link === null && is(event, TEXT_NAMESPACE, "a", "text")) {
          link = odsWireLink(event.attrs);
        }
        paragraph += inlineText(event);
      }
      if (!event.selfClosing) {
        depth += 1;
      }
      return undefined;
    });
  } catch {
    return { annotations: [], links: [] };
  }
  return { annotations, links };
}

/** Whether two cell values are the same (errors by code). */
function sameValue(a: CellValue, b: CellValue): boolean {
  if (isCellError(a) || isCellError(b)) {
    return isCellError(a) && isCellError(b) && a.error === b.error;
  }
  return a === b;
}

/** Whether the sheet's cell at `ref` still holds `sourceValue` (a null
 * source value never matches: an empty cell's address says nothing). */
function stillHolds(sheet: Pick<Sheet, "rows">, ref: CellRef, sourceValue: CellValue): boolean {
  return sourceValue !== null && sameValue(getCell(sheet as Sheet, ref.row, ref.col)?.value ?? null, sourceValue);
}

/** What the upgrade's input is: the dirty stored sheet, the retained
 * content envelope's prefix, the workbook's date system, and an optional
 * remap of a `place` location's source-time sheet names onto the record's
 * (SPEC-121's sourceSheetRemap through remapSheetReferences; absent,
 * locations are kept as read). */
export interface OdsAnnotationsUpgradeInput {
  sheet: Pick<Sheet, "rows" | "sourceXml">;
  prefix: string;
  dates: Pick<SpreadsheetModel, "dateSystem" | "odsNullDate">;
  remapLocation?: (location: string) => string;
}

/** SPEC-135's reader-upgrade step for a dirty .ods sheet's links: each link
 * its retained `sourceXml` carries, attached only where the sheet's cell at
 * that address still holds the source cell's value (a value or structural
 * edit since read makes the address unreliable, and such a link is not
 * resurrected). Sorted by (row, col); empty when nothing survives. */
export function upgradeOdsSheetLinks(input: OdsAnnotationsUpgradeInput): CellHyperlink[] {
  return attachLinks(input, retainedOdsAnnotations(input.sheet.sourceXml, input.prefix, input.dates).links);
}

function attachLinks(input: OdsAnnotationsUpgradeInput, retained: readonly RetainedOdsLink[]): CellHyperlink[] {
  const { sheet, remapLocation } = input;
  const out: CellHyperlink[] = [];
  for (const { ref, hyperlink, sourceValue } of retained) {
    if (!stillHolds(sheet, ref, sourceValue)) {
      continue;
    }
    out.push(remapLocation && hyperlink.target.kind === "place" ? { ...hyperlink, target: { kind: "place", location: remapLocation(hyperlink.target.location) } } : hyperlink);
  }
  return out.toSorted((a, b) => a.range.top - b.range.top || a.range.left - b.range.left);
}

/** SPEC-135's whole reader-upgrade step for a dirty .ods sheet — the one
 * call open.ts's ReaderUpgrade makes:
 * - notes: every annotation its cells' preservedXml hold (records read by
 *   SPEC-122's to SPEC-133's readers), at each cell's current position
 *   (liftOdsSheetNotes); then each annotation its retained sourceXml
 *   carries that is not one of those (by its verbatim element) — the only
 *   source for a record older than SPEC-122's reader, whose step 5
 *   re-read with today's reader leaves annotations out of preservedXml —
 *   attached where no note is and the cell at its address still holds the
 *   source's value;
 * - links: upgradeOdsSheetLinks.
 * `rows` is the sheet's rows with the lifted annotations gone
 * (copy-on-write, see liftOdsSheetNotes); the arrays are sorted by (row,
 * col) and empty when there is nothing — the caller sets `notes` and
 * `hyperlinks` only when non-empty. */
export function upgradeOdsSheetAnnotations(input: OdsAnnotationsUpgradeInput): { rows: Row[]; notes: CellNote[]; hyperlinks: CellHyperlink[] } {
  const lifted = liftOdsSheetNotes({ sheet: input.sheet, prefix: input.prefix });
  const retained = retainedOdsAnnotations(input.sheet.sourceXml, input.prefix, input.dates);
  const declarations = prefixDeclarations(input.prefix);
  const notes = [...lifted.notes];
  const taken = new Set(notes.map((note) => note.noteId));
  const at = new Set(notes.map((note) => `${note.ref.row}:${note.ref.col}`));
  // Each lifted element accounts for one retained twin.
  const liftedXml = new Map<string, number>();
  for (const note of lifted.notes) {
    const xml = note.source!.odsAnnotation!;
    liftedXml.set(xml, (liftedXml.get(xml) ?? 0) + 1);
  }
  for (const { ref, xml, sourceValue } of retained.annotations) {
    const twins = liftedXml.get(xml) ?? 0;
    if (twins > 0) {
      liftedXml.set(xml, twins - 1);
      continue;
    }
    if (at.has(`${ref.row}:${ref.col}`) || !stillHolds({ rows: lifted.rows }, ref, sourceValue)) {
      continue;
    }
    const note = odsNote(xml, declarations, ref, taken);
    if (note) {
      notes.push(note);
      at.add(`${ref.row}:${ref.col}`);
    }
  }
  return {
    rows: lifted.rows,
    notes: notes.toSorted((a, b) => compareRefs(a.ref, b.ref)),
    hyperlinks: attachLinks({ ...input, sheet: { ...input.sheet, rows: lifted.rows } }, retained.links),
  };
}
