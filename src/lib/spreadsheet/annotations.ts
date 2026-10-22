// Notes, threads and hyperlinks on cells (SPEC-135) — THIS FILE is the
// normative source for the annotation commands, the rule for which link
// targets a sheet opens, and how annotations follow their cells. The model
// is model.ts's CellNote / CellThread / ThreadPerson / CellHyperlink; the
// wire mappings are xlsx/annotations.ts and ods/annotations.ts.
//
// Commands. Each replaces the arrays it touches (copy-on-write, like `rows`)
// and installs the sheet through markAnnotationsEdited, which sets
// `annotationsDirty` on an .xlsx workbook (the writer rebuilds the comment
// parts and splices `<hyperlinks>`/`<legacyDrawing>` only) and `dirty` on
// any other (ODS annotations live inside the table's cells, so the table is
// re-serialised — the SPEC-101 markCommentsEdited precedent). A command that
// changes a note's content builds a new note WITHOUT `source` (and `runs`:
// an edited note is plain text); an adjuster or remap that only moves an
// annotation keeps its `source`, and the writer patches the position fields
// (a moved note's box keeps its size but not its custom position). New
// GUIDs are `{${crypto.randomUUID().toUpperCase()}}`; a new author joins
// `model.persons` (replaced, never mutated) as `{ displayName: name,
// userId: name, providerId: "None" }`. Every text a command takes in (note
// and message text, authors, link targets and tooltips) passes core's
// sanitizeXmlText first: a pasted control character (Word's U+000B) would
// otherwise make the comments, threaded-comments or content.xml part the
// writer builds from it ill-formed.
//
// Following a link (openHyperlinkTarget, a pure decision — the view acts):
// `http:`, `https:` and `mailto:` open externally; any other scheme, a
// dangerous one included, and a relative path are refused with
// LINK_NOT_OPENABLE_MESSAGE; a `place` activates its sheet and range, and an
// unresolvable one (missing sheet, `#REF!`, unknown name) shows
// LINK_DESTINATION_MISSING_MESSAGE. A dangerous scheme read from a file is
// KEPT in the model and the file (round-trip fidelity) and simply never
// opened. A cell whose formula's top-level node is a `HYPERLINK(…)` call
// (SPEC-127) behaves as linked (hyperlinkFormulaCall); its target is computed
// at follow time (hyperlinkFormulaTarget) and it is never stored in
// `Sheet.hyperlinks`.
//
// Adjusters (registered by structure.ts, move.ts and sort.ts, which call
// these functions through arrows — the modules import each other):
// - "notes" (notes and threads), structural edit: on the edited sheet a ref
//   in an inserted band's wake shifts by `count`; a ref inside a deleted
//   band is removed.
// - "hyperlinks", structural edit: on the edited sheet ranges shift, a range
//   straddling a deletion shrinks, a range wholly deleted is removed; then,
//   on EVERY sheet, each `place` location goes through adjustFormula (host:
//   the link's sheet), a location left holding `#REF!` stays.
// - the same ids in MOVE_ADJUSTERS: a note, thread or hyperlink range
//   entirely inside the moved block moves by the offset onto the destination
//   sheet; one entirely inside the destination and not moving is deleted
//   (its cells were overwritten); a hyperlink range only partly overlapping
//   either stays; every `place` location passes through repointFormula.
// - "annotations" in ROW_PERMUTATION_ADJUSTERS: remapAnnotations with the
//   sort's permutation, so a sort carries notes, threads and links with
//   their records.
// Each marks a sheet whose annotations it changed through
// markAnnotationsEdited.
import { replyParagraphText } from "../core/comments";
import { urlScheme } from "../core/urls";
import { sanitizeXmlText } from "../core/xmlText";
import { formatA1, type CellRef, type GridRange } from "./address";
import { adjustFormula } from "./formula/adjust";
import type { FormulaNode } from "./formula/ast";
import { storedValueContext, evaluate } from "./formula/evaluate";
import { FormulaSyntaxError, parseFormula, parseFormulaAt } from "./formula/parse";
import { formatReference, formatSheetPrefix } from "./formula/print";
import { withReferenceHooks } from "./formula/referenceHooks";
import { remapSheetReferences, type SheetRemap } from "./formula/sheetRefs";
import { isArray, isEmpty, isError, isRange, type EvalValue } from "./formula/values";
import type { Cell, CellHyperlink, CellNote, CellThread, HyperlinkTarget, Sheet, SpreadsheetFormat, SpreadsheetModel, ThreadMessage, ThreadPerson } from "./model";
import { moveDestination, repointFormula, type RangeMove } from "./move";
import { adjustIndex, adjustRange, editAxis, type StructuralEdit } from "./structure";

/** How long the pointer rests on an annotated cell before its popover shows. */
export const NOTE_HOVER_DELAY_MS = 400;
/** Excel's default note box, 108 × 59.25 pt at 96 dpi. */
export const NOTE_DEFAULT_BOX: Readonly<{ widthPx: number; heightPx: number }> = { widthPx: 144, heightPx: 79 };
export const LINK_NOT_OPENABLE_MESSAGE = "This link can't be opened from the browser";
export const LINK_DESTINATION_MISSING_MESSAGE = "This link's destination doesn't exist";
/** The title "New comment" carries, disabled, in an .ods workbook. */
export const ODS_NOTES_ONLY_TITLE = "OpenDocument spreadsheets support notes only";
/** The titles the note and comment commands, and the link commands, carry
 * disabled on an .xlsx sheet whose annotation parts (or `<hyperlinks>`) the
 * reader could not model: those stay byte-verbatim, so nothing added there
 * could be saved (AC-11). */
export const NOTES_KEPT_AS_READ_TITLE = "This sheet's notes and comments are kept exactly as the file has them, so they can't be changed here";
export const LINKS_KEPT_AS_READ_TITLE = "This sheet's links are kept exactly as the file has them, so they can't be changed here";
/** The title Link… carries, disabled, in a CSV or TSV workbook (which keeps values only). */
export const CSV_NO_LINKS_TITLE = "CSV and TSV files can't keep links";
/** The schemes a sheet opens in a new browsing context. */
export const OPENABLE_SCHEMES: ReadonlySet<string> = new Set(["http", "https", "mailto"]);

export const NOTE_UNDO_LABELS = { add: "Add note", edit: "Edit note", delete: "Delete note", show: "Show note", hide: "Hide note" } as const;
export const COMMENT_UNDO_LABELS = {
  add: "Add comment",
  reply: "Reply to comment",
  edit: "Edit comment",
  resolve: "Resolve comment",
  reopen: "Reopen comment",
  delete: "Delete comment",
} as const;
export const LINK_UNDO_LABELS = { insert: "Insert link", edit: "Edit link", remove: "Remove link" } as const;

// --- queries -----------------------------------------------------------------------------------

export const sheetNotes = (sheet: Pick<Sheet, "notes">): readonly CellNote[] => sheet.notes ?? [];
export const sheetThreads = (sheet: Pick<Sheet, "threads">): readonly CellThread[] => sheet.threads ?? [];
export const sheetHyperlinks = (sheet: Pick<Sheet, "hyperlinks">): readonly CellHyperlink[] => sheet.hyperlinks ?? [];

/** Whether a sheet holds any annotation. */
export function hasAnnotations(sheet: Pick<Sheet, "notes" | "threads" | "hyperlinks">): boolean {
  return sheetNotes(sheet).length > 0 || sheetThreads(sheet).length > 0 || sheetHyperlinks(sheet).length > 0;
}

/** (row, col) order — the order every annotation array is kept in. */
export function compareRefs(a: CellRef, b: CellRef): number {
  return a.row - b.row || a.col - b.col;
}

const sameRef = (a: CellRef, b: CellRef): boolean => a.row === b.row && a.col === b.col;
const rangeStart = (range: GridRange): CellRef => ({ row: range.top, col: range.left });
const inRange = (range: GridRange, ref: CellRef): boolean => ref.row >= range.top && ref.row <= range.bottom && ref.col >= range.left && ref.col <= range.right;
const insideRange = (outer: GridRange, inner: GridRange): boolean =>
  inner.top >= outer.top && inner.bottom <= outer.bottom && inner.left >= outer.left && inner.right <= outer.right;
const intersects = (a: GridRange, b: GridRange): boolean => a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;

/** Notes or threads sorted by ref (a new array). */
export function sortedByRef<T extends { ref: CellRef }>(items: readonly T[]): T[] {
  return items.toSorted((a, b) => compareRefs(a.ref, b.ref));
}

/** Hyperlinks sorted by their range's top-left (a new array). */
export function sortedLinks(links: readonly CellHyperlink[]): CellHyperlink[] {
  return links.toSorted((a, b) => compareRefs(rangeStart(a.range), rangeStart(b.range)));
}

export function noteAt(sheet: Pick<Sheet, "notes">, ref: CellRef): CellNote | undefined {
  return sheetNotes(sheet).find((note) => sameRef(note.ref, ref));
}

export function threadAt(sheet: Pick<Sheet, "threads">, ref: CellRef): CellThread | undefined {
  return sheetThreads(sheet).find((thread) => sameRef(thread.ref, ref));
}

/** The hyperlink whose range covers `ref` (the first in model order). */
export function hyperlinkAt(sheet: Pick<Sheet, "hyperlinks">, ref: CellRef): CellHyperlink | undefined {
  return sheetHyperlinks(sheet).find((link) => inRange(link.range, ref));
}

/** A note's plain text as the wire's paragraphs. */
export function noteParagraphs(note: Pick<CellNote, "text">): string[] {
  return note.text.split("\n");
}

// --- ids ----------------------------------------------------------------------------------------

/** A new GUID in Excel's written form. */
export function newGuid(): string {
  return `{${crypto.randomUUID().toUpperCase()}}`;
}

/** `${prefix}-${A1}`, suffixed until it is free. */
function freeId(prefix: string, ref: CellRef, taken: ReadonlySet<string>): string {
  const base = `${prefix}-${formatA1(ref)}`;
  if (!taken.has(base)) {
    return base;
  }
  let n = 2;
  while (taken.has(`${base}-${n}`)) {
    n += 1;
  }
  return `${base}-${n}`;
}

/** The read-time id of a note at `ref`: `n-${A1}`. */
export function noteIdAt(ref: CellRef): string {
  return `n-${formatA1(ref)}`;
}

/** The read-time id of a hyperlink with no `xr:uid`: `h-${A1 of its start}`. */
export function linkIdAt(range: GridRange): string {
  return `h-${formatA1(rangeStart(range))}`;
}

/** The current time as an ISO 8601 stamp (the message `createdAt`). */
function nowIso(): string {
  return new Date().toISOString();
}

// --- marking ------------------------------------------------------------------------------------

/** Installs `sheet` (by id) marked as annotation-edited: `annotationsDirty`
 * on an .xlsx workbook, `dirty` on any other. Returns the installed sheet. */
export function markAnnotationsEdited(model: SpreadsheetModel, sheet: Sheet): Sheet {
  const marked: Sheet = model.format === "xlsx" ? (sheet.annotationsDirty ? sheet : { ...sheet, annotationsDirty: true }) : sheet.dirty ? sheet : { ...sheet, dirty: true };
  model.sheets = model.sheets.map((each) => (each.id === sheet.id ? marked : each));
  return marked;
}

function sheetById(model: SpreadsheetModel, sheetId: string): Sheet {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  if (!sheet) {
    throw new Error(`No sheet ${sheetId}`);
  }
  return sheet;
}

/** The sheet with `field` set to `items` — deleted when empty. */
function withField<K extends "notes" | "threads" | "hyperlinks">(sheet: Sheet, field: K, items: NonNullable<Sheet[K]>): Sheet {
  const next: Sheet = { ...sheet };
  if (items.length === 0) {
    delete next[field];
  } else {
    next[field] = items;
  }
  return next;
}

// --- note commands ------------------------------------------------------------------------------

/** Adds a note to `ref` (one per cell: a cell holding one has it edited).
 * Blank text adds nothing. Returns the note now there, or null. */
export function addNote(model: SpreadsheetModel, sheetId: string, ref: CellRef, text: string, author: string): CellNote | null {
  const sheet = sheetById(model, sheetId);
  const existing = noteAt(sheet, ref);
  if (existing) {
    return editNote(model, sheetId, existing.noteId, text);
  }
  if (text.trim() === "") {
    return null;
  }
  const notes = sheetNotes(sheet);
  const note: CellNote = {
    noteId: freeId("n", ref, new Set(notes.map((each) => each.noteId))),
    ref: { row: ref.row, col: ref.col },
    author: sanitizeXmlText(author),
    text: sanitizeXmlText(text),
  };
  markAnnotationsEdited(model, withField(sheet, "notes", sortedByRef([...notes, note])));
  return note;
}

/** Replaces a note's text (plain from now on: no runs, no source); emptied
 * text deletes it. Returns the note, or null when it went. */
export function editNote(model: SpreadsheetModel, sheetId: string, noteId: string, typed: string): CellNote | null {
  const text = sanitizeXmlText(typed);
  const sheet = sheetById(model, sheetId);
  const note = sheetNotes(sheet).find((each) => each.noteId === noteId);
  if (!note) {
    return null;
  }
  if (text.trim() === "") {
    deleteNote(model, sheetId, noteId);
    return null;
  }
  if (text === note.text && note.runs === undefined) {
    return note;
  }
  const { source: _source, runs: _runs, ...rest } = note;
  const next: CellNote = { ...rest, text };
  markAnnotationsEdited(model, withField(sheet, "notes", sheetNotes(sheet).map((each) => (each === note ? next : each))));
  return next;
}

export function deleteNote(model: SpreadsheetModel, sheetId: string, noteId: string): boolean {
  const sheet = sheetById(model, sheetId);
  const notes = sheetNotes(sheet);
  const next = notes.filter((each) => each.noteId !== noteId);
  if (next.length === notes.length) {
    return false;
  }
  markAnnotationsEdited(model, withField(sheet, "notes", next));
  return true;
}

/** Shows or hides a note's box ("Show note" / "Hide note"). Its shape is
 * regenerated (the VML and ODS display state are in the shape), its
 * comment text kept verbatim. */
export function setNoteVisible(model: SpreadsheetModel, sheetId: string, noteId: string, visible: boolean): boolean {
  const sheet = sheetById(model, sheetId);
  const note = sheetNotes(sheet).find((each) => each.noteId === noteId);
  if (!note || (note.visible === true) === visible) {
    return false;
  }
  const { visible: _visible, source, ...rest } = note;
  const next: CellNote = { ...rest };
  if (visible) {
    next.visible = true;
  }
  if (source?.xlsxComment !== undefined) {
    next.source = { xlsxComment: source.xlsxComment };
  }
  markAnnotationsEdited(model, withField(sheet, "notes", sheetNotes(sheet).map((each) => (each === note ? next : each))));
  return true;
}

// --- thread commands ----------------------------------------------------------------------------

/** The person a message by `author` is written as: the first whose
 * display name is the author's, else a new one appended (the array
 * replaced). */
export function ensurePerson(model: SpreadsheetModel, author: string): ThreadPerson {
  const persons = model.persons ?? [];
  const known = persons.find((person) => person.displayName === author);
  if (known) {
    return known;
  }
  const person: ThreadPerson = { personId: newGuid(), displayName: author, userId: author, providerId: "None" };
  model.persons = [...persons, person];
  return person;
}

function newMessage(model: SpreadsheetModel, typed: string, typedAuthor: string): ThreadMessage {
  const author = sanitizeXmlText(typedAuthor);
  return { messageId: newGuid(), author, text: sanitizeXmlText(typed), createdAt: nowIso(), personId: ensurePerson(model, author).personId };
}

function replaceThread(model: SpreadsheetModel, sheet: Sheet, thread: CellThread, next: CellThread | null): void {
  const threads = sheetThreads(sheet).flatMap((each) => (each === thread ? (next ? [next] : []) : [each]));
  markAnnotationsEdited(model, withField(sheet, "threads", threads));
}

function findThread(sheet: Sheet, threadId: string): CellThread | undefined {
  return sheetThreads(sheet).find((each) => each.threadId === threadId);
}

/** Starts a thread on `ref` (one per cell: a cell holding one gets a reply).
 * Blank text starts nothing. */
export function addThread(model: SpreadsheetModel, sheetId: string, ref: CellRef, text: string, author: string): CellThread | null {
  const sheet = sheetById(model, sheetId);
  const existing = threadAt(sheet, ref);
  if (existing) {
    return replyToThread(model, sheetId, existing.threadId, text, author);
  }
  if (text.trim() === "") {
    return null;
  }
  const message = newMessage(model, text, author);
  const thread: CellThread = { threadId: message.messageId, ref: { row: ref.row, col: ref.col }, messages: [message], resolved: false };
  markAnnotationsEdited(model, withField(sheet, "threads", sortedByRef([...sheetThreads(sheet), thread])));
  return thread;
}

export function replyToThread(model: SpreadsheetModel, sheetId: string, threadId: string, text: string, author: string): CellThread | null {
  const sheet = sheetById(model, sheetId);
  const thread = findThread(sheet, threadId);
  if (!thread || text.trim() === "") {
    return null;
  }
  const next: CellThread = { ...thread, messages: [...thread.messages, newMessage(model, text, author)] };
  replaceThread(model, sheetById(model, sheetId), thread, next);
  return next;
}

/** Replaces a message's text: it loses its source and its `<mentions>`
 * (the sanctioned delta — @-mention authoring is out of scope). Emptied
 * text deletes the message. */
export function editThreadMessage(model: SpreadsheetModel, sheetId: string, threadId: string, messageId: string, typed: string): CellThread | null {
  const text = sanitizeXmlText(typed);
  const sheet = sheetById(model, sheetId);
  const thread = findThread(sheet, threadId);
  const message = thread?.messages.find((each) => each.messageId === messageId);
  if (!thread || !message) {
    return null;
  }
  if (text.trim() === "") {
    return deleteThreadMessage(model, sheetId, threadId, messageId);
  }
  if (text === message.text) {
    return thread;
  }
  const { source: _source, mentions: _mentions, ...rest } = message;
  const next: CellThread = { ...thread, messages: thread.messages.map((each) => (each === message ? { ...rest, text } : each)) };
  replaceThread(model, sheet, thread, next);
  return next;
}

/** Deletes a message; deleting `messages[0]` deletes the thread. Returns the
 * thread left, or null. */
export function deleteThreadMessage(model: SpreadsheetModel, sheetId: string, threadId: string, messageId: string): CellThread | null {
  const sheet = sheetById(model, sheetId);
  const thread = findThread(sheet, threadId);
  if (!thread) {
    return null;
  }
  const at = thread.messages.findIndex((each) => each.messageId === messageId);
  if (at < 0) {
    return thread;
  }
  const next = at === 0 ? null : { ...thread, messages: thread.messages.filter((_, i) => i !== at) };
  replaceThread(model, sheet, thread, next);
  return next;
}

/** Resolves or reopens a thread (the root message's `done`). */
export function setThreadResolved(model: SpreadsheetModel, sheetId: string, threadId: string, resolved: boolean): boolean {
  const sheet = sheetById(model, sheetId);
  const thread = findThread(sheet, threadId);
  if (!thread || thread.resolved === resolved) {
    return false;
  }
  const [root, ...replies] = thread.messages;
  const { source: _source, ...bare } = root!;
  replaceThread(model, sheet, thread, { ...thread, resolved, messages: [bare, ...replies] });
  return true;
}

// --- hyperlink commands -------------------------------------------------------------------------

/** Sets one hyperlink over `range`: every link intersecting it goes; a link
 * over exactly that range keeps its id (Edit link). Returns the link. */
export function setHyperlink(model: SpreadsheetModel, sheetId: string, range: GridRange, target: HyperlinkTarget, tooltip?: string): CellHyperlink {
  const sheet = sheetById(model, sheetId);
  const links = sheetHyperlinks(sheet);
  const same = links.find((link) => link.range.top === range.top && link.range.left === range.left && link.range.bottom === range.bottom && link.range.right === range.right);
  const rest = links.filter((link) => !intersects(link.range, range));
  const link: CellHyperlink = {
    linkId: same?.linkId ?? freeId("h", rangeStart(range), new Set(rest.map((each) => each.linkId))),
    range: { top: range.top, left: range.left, bottom: range.bottom, right: range.right },
    target:
      target.kind === "url"
        ? { kind: "url", href: sanitizeXmlText(target.href), ...(target.subaddress ? { subaddress: sanitizeXmlText(target.subaddress) } : {}) }
        : { kind: "place", location: sanitizeXmlText(target.location) },
  };
  if (tooltip !== undefined && tooltip !== "") {
    link.tooltip = sanitizeXmlText(tooltip);
  }
  markAnnotationsEdited(model, withField(sheet, "hyperlinks", sortedLinks([...rest, link])));
  return link;
}

export function removeHyperlink(model: SpreadsheetModel, sheetId: string, linkId: string): boolean {
  const sheet = sheetById(model, sheetId);
  const links = sheetHyperlinks(sheet);
  const next = links.filter((link) => link.linkId !== linkId);
  if (next.length === links.length) {
    return false;
  }
  markAnnotationsEdited(model, withField(sheet, "hyperlinks", next));
  return true;
}

// --- following --------------------------------------------------------------------------------

/** What following `target` does (see the module header) — the view acts. */
export function openHyperlinkTarget(target: HyperlinkTarget): "external" | "place" | "refused" {
  if (target.kind === "place") {
    return "place";
  }
  const scheme = urlScheme(target.href);
  return scheme !== null && OPENABLE_SCHEMES.has(scheme) ? "external" : "refused";
}

/** The href an external target opens: the address with its `#subaddress`. */
export function externalHref(target: Extract<HyperlinkTarget, { kind: "url" }>): string {
  return target.subaddress ? `${target.href}#${target.subaddress}` : target.href;
}

/** A link target's one-line label (the chip's). */
export function hyperlinkLabel(target: HyperlinkTarget): string {
  return target.kind === "url" ? externalHref(target) : target.location;
}

/** A `place` location as canonical reference text (SPEC-119): a lone
 * reference, name or error that parses is written back in canonical form
 * (one token, so no whitespace of the user's is lost — not a print(parse())
 * of a formula); anything else (text in no dialect) is kept as written. */
export function canonicalPlaceLocation(text: string): string {
  const trimmed = text.trim().replace(/^=/, "");
  try {
    const node = parseFormula(trimmed);
    switch (node.kind) {
      case "ref":
        return formatReference(node.ref);
      case "name":
        return node.sheet === undefined ? node.name : formatSheetPrefix(node.sheet) + node.name;
      case "error":
        return node.text;
      default:
        return trimmed;
    }
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return trimmed;
    }
    throw error;
  }
}

/** The `HYPERLINK(…)` call that is a formula cell's top-level node (SPEC-127),
 * or null — read from SPEC-119's shared AST cache at the cell's position. */
export function hyperlinkFormulaCall(cell: Pick<Cell, "formula" | "formulaUnparsed">, host: CellRef): Extract<FormulaNode, { kind: "function" }> | null {
  if (cell.formula === undefined || cell.formulaUnparsed || !/hyperlink/i.test(cell.formula)) {
    return null;
  }
  const parsed = parseFormulaAt(cell.formula, host);
  if (parsed instanceof FormulaSyntaxError) {
    return null;
  }
  const node = parsed.node;
  return node.kind === "function" && node.name.toUpperCase() === "HYPERLINK" && node.args.length > 0 ? node : null;
}

/** The text an evaluation gives, or null (an error, an empty value). */
function evaluatedText(value: EvalValue): string | null {
  if (isError(value) || isEmpty(value) || isRange(value) || isArray(value)) {
    return null;
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return String(value);
}

/** A `HYPERLINK()` cell's target, computed now: its first argument
 * evaluated for that cell (SPEC-120) and coerced to text — `#…` is a place
 * (the rest canonicalised), anything else a url (a bare `Sheet!A1` too,
 * which the open rule then refuses as a relative path); null when it gives
 * no text. */
export function hyperlinkFormulaTarget(model: SpreadsheetModel, sheetId: string, ref: CellRef, cell: Pick<Cell, "formula" | "formulaUnparsed">): HyperlinkTarget | null {
  const call = hyperlinkFormulaCall(cell, ref);
  const parsed = call === null ? null : parseFormulaAt(cell.formula!, ref);
  if (call === null || parsed === null || parsed instanceof FormulaSyntaxError) {
    return null;
  }
  let text: string | null;
  try {
    const ctx = withReferenceHooks({ ...storedValueContext(model), host: { sheetId, row: ref.row, col: ref.col }, origin: parsed.origin, arrayMode: false });
    text = evaluatedText(evaluate(call.args[0]!, ctx));
  } catch {
    return null;
  }
  if (text === null || text.trim() === "") {
    return null;
  }
  const trimmed = text.trim();
  if (trimmed.startsWith("#")) {
    return { kind: "place", location: canonicalPlaceLocation(trimmed.slice(1)) };
  }
  return { kind: "url", href: trimmed };
}

// --- following cells ---------------------------------------------------------------------------

/** Every `place` location of a sheet's links through `rewrite` (given the
 * link's host sheet id); the same array when none changes. */
function mapPlaces(links: readonly CellHyperlink[], rewrite: (location: string) => string): readonly CellHyperlink[] {
  let changed = false;
  const next = links.map((link) => {
    if (link.target.kind !== "place") {
      return link;
    }
    const location = rewrite(link.target.location);
    if (location === link.target.location) {
      return link;
    }
    changed = true;
    return { ...link, target: { kind: "place" as const, location } };
  });
  return changed ? next : links;
}

/** Applies an arbitrary cell mapping to a sheet's annotations — null
 * deletes — keeping each moved annotation's source (the writer patches its
 * position). A hyperlink range moves when its corners map onto a range of
 * the same size; one whose start maps to null goes; any other stays. Marks
 * the sheet when something changed. */
export function remapAnnotations(model: SpreadsheetModel, sheetId: string, map: (ref: CellRef) => CellRef | null): void {
  const sheet = model.sheets.find((each) => each.id === sheetId);
  if (!sheet || !hasAnnotations(sheet)) {
    return;
  }
  const remapped = remapSheet(sheet, map);
  if (remapped !== sheet) {
    markAnnotationsEdited(model, remapped);
  }
}

function remapSheet(sheet: Sheet, map: (ref: CellRef) => CellRef | null): Sheet {
  let changed = false;
  const move = <T extends { ref: CellRef }>(items: readonly T[]): T[] =>
    items.flatMap((item) => {
      const to = map(item.ref);
      if (to === null) {
        changed = true;
        return [];
      }
      if (sameRef(to, item.ref)) {
        return [item];
      }
      changed = true;
      return [{ ...item, ref: { row: to.row, col: to.col } }];
    });
  const notes = move(sheetNotes(sheet));
  const threads = move(sheetThreads(sheet));
  const links = sheetHyperlinks(sheet).flatMap((link) => {
    const start = map(rangeStart(link.range));
    if (start === null) {
      changed = true;
      return [];
    }
    const end = map({ row: link.range.bottom, col: link.range.right });
    const rows = link.range.bottom - link.range.top;
    const cols = link.range.right - link.range.left;
    if (end === null || end.row - start.row !== rows || end.col - start.col !== cols) {
      return [link];
    }
    if (start.row === link.range.top && start.col === link.range.left) {
      return [link];
    }
    changed = true;
    return [{ ...link, range: { top: start.row, left: start.col, bottom: end.row, right: end.col } }];
  });
  if (!changed) {
    return sheet;
  }
  return withField(withField(withField(sheet, "notes", sortedByRef(notes)), "threads", sortedByRef(threads)), "hyperlinks", sortedLinks(links));
}

/** Sheet id → name, for adjustFormula and repointFormula. */
function sheetNames(model: SpreadsheetModel): (sheetId: string) => string {
  const names = new Map(model.sheets.map((sheet) => [sheet.id, sheet.name] as const));
  return (sheetId) => names.get(sheetId) ?? "";
}

/** Installs `next` (a sheet list parallel to the model's), every sheet it
 * changed marked. */
function installChanged(model: SpreadsheetModel, next: readonly Sheet[]): void {
  const before = model.sheets;
  if (next.every((sheet, i) => sheet === before[i])) {
    return;
  }
  model.sheets = [...next];
  next.forEach((sheet, i) => {
    if (sheet !== before[i]) {
      markAnnotationsEdited(model, sheet);
    }
  });
}

/** SPEC-124's "notes" adjuster (notes and threads; see the module header). */
export function adjustNotesForEdit(model: SpreadsheetModel, edit: StructuralEdit): void {
  const sheet = model.sheets.find((each) => each.id === edit.sheetId);
  if (!sheet || (sheetNotes(sheet).length === 0 && sheetThreads(sheet).length === 0)) {
    return;
  }
  const rows = editAxis(edit) === "row";
  const map = (ref: CellRef): CellRef | null => {
    const index = adjustIndex(rows ? ref.row : ref.col, edit);
    return index === null ? null : rows ? { row: index, col: ref.col } : { row: ref.row, col: index };
  };
  const remapped = remapSheet({ ...sheet, hyperlinks: [] }, map);
  if (remapped.notes === sheet.notes && remapped.threads === sheet.threads) {
    return;
  }
  const next: Sheet = withField(withField({ ...sheet }, "notes", remapped.notes ?? []), "threads", remapped.threads ?? []);
  installChanged(
    model,
    model.sheets.map((each) => (each.id === sheet.id ? next : each)),
  );
}

/** SPEC-124's "hyperlinks" adjuster (see the module header). */
export function adjustHyperlinksForEdit(model: SpreadsheetModel, edit: StructuralEdit): void {
  const nameOf = sheetNames(model);
  const next = model.sheets.map((sheet) => {
    let links = sheetHyperlinks(sheet);
    if (links.length === 0) {
      return sheet;
    }
    if (sheet.id === edit.sheetId) {
      let moved = false;
      const adjusted = links.flatMap((link) => {
        const range = adjustRange(link.range, edit);
        if (range === null) {
          moved = true;
          return [];
        }
        if (range.top === link.range.top && range.left === link.range.left && range.bottom === link.range.bottom && range.right === link.range.right) {
          return [link];
        }
        moved = true;
        return [{ ...link, range }];
      });
      if (moved) {
        links = adjusted;
      }
    }
    links = mapPlaces(links, (location) => adjustPlace(location, (text) => adjustFormula(text, sheet.id, edit, nameOf)));
    return links === sheet.hyperlinks ? sheet : withField(sheet, "hyperlinks", [...links]);
  });
  installChanged(model, next);
}

/** A place location through a formula rewrite — kept when the rewrite
 * cannot read it. */
function adjustPlace(location: string, rewrite: (text: string) => string): string {
  try {
    return rewrite(location);
  } catch {
    return location;
  }
}

/** SPEC-125's "notes" move adjuster (see the module header). */
export function moveNotesForMove(model: SpreadsheetModel, move: RangeMove): void {
  moveItems(model, move, "notes");
  moveItems(model, move, "threads");
}

function moveItems(model: SpreadsheetModel, move: RangeMove, field: "notes" | "threads"): void {
  const from = model.sheets.find((sheet) => sheet.id === move.fromSheetId);
  const to = model.sheets.find((sheet) => sheet.id === move.toSheetId);
  if (!from || !to) {
    return;
  }
  const dest = moveDestination(move);
  const dRow = move.toTop - move.from.top;
  const dCol = move.toLeft - move.from.left;
  const items = (sheet: Sheet): readonly (CellNote | CellThread)[] => (field === "notes" ? sheetNotes(sheet) : sheetThreads(sheet));
  const fromItems = items(from);
  const travelling = fromItems.filter((item) => inRange(move.from, item.ref));
  const toItems = items(to);
  const overwritten = toItems.filter((item) => inRange(dest, item.ref) && !(to === from && inRange(move.from, item.ref)));
  if (travelling.length === 0 && overwritten.length === 0) {
    return;
  }
  const arrived = travelling.map((item) => ({ ...item, ref: { row: item.ref.row + dRow, col: item.ref.col + dCol } }));
  if (from === to) {
    const kept = fromItems.filter((item) => !travelling.includes(item) && !overwritten.includes(item));
    installChanged(model, model.sheets.map((sheet) => (sheet === from ? withField(sheet, field, sortedByRef([...kept, ...arrived]) as never) : sheet)));
    return;
  }
  const fromNext = fromItems.filter((item) => !travelling.includes(item));
  const toNext = sortedByRef([...toItems.filter((item) => !overwritten.includes(item)), ...arrived]);
  installChanged(
    model,
    model.sheets.map((sheet) => (sheet === from ? withField(sheet, field, fromNext as never) : sheet === to ? withField(sheet, field, toNext as never) : sheet)),
  );
}

/** SPEC-125's "hyperlinks" move adjuster (see the module header). */
export function moveHyperlinksForMove(model: SpreadsheetModel, move: RangeMove): void {
  const nameOf = sheetNames(model);
  const from = model.sheets.find((sheet) => sheet.id === move.fromSheetId);
  const to = model.sheets.find((sheet) => sheet.id === move.toSheetId);
  if (!from || !to) {
    return;
  }
  const dest = moveDestination(move);
  const dRow = move.toTop - move.from.top;
  const dCol = move.toLeft - move.from.left;
  const travelling = new Set(sheetHyperlinks(from).filter((link) => insideRange(move.from, link.range)));
  const overwritten = new Set(sheetHyperlinks(to).filter((link) => insideRange(dest, link.range) && !travelling.has(link)));
  // The links that arrive, each re-pointed as a formula moved with them.
  const arrived = [...travelling].map((link) => {
    const range = { top: link.range.top + dRow, left: link.range.left + dCol, bottom: link.range.bottom + dRow, right: link.range.right + dCol };
    if (link.target.kind !== "place") {
      return { ...link, range };
    }
    const location = adjustPlace(link.target.location, (text) => repointFormula(text, move.fromSheetId, move, nameOf, move.toSheetId));
    return { ...link, range, target: { kind: "place" as const, location } };
  });
  const next = model.sheets.map((sheet) => {
    const links = sheetHyperlinks(sheet);
    const leaving = sheet === from && travelling.size > 0;
    const losing = sheet === to && overwritten.size > 0;
    const staying = leaving || losing ? links.filter((link) => !(sheet === from && travelling.has(link)) && !(sheet === to && overwritten.has(link))) : links;
    const repointed = mapPlaces(staying, (location) => adjustPlace(location, (text) => repointFormula(text, sheet.id, move, nameOf)));
    if (sheet === to && arrived.length > 0) {
      return withField(sheet, "hyperlinks", sortedLinks([...repointed, ...arrived]));
    }
    return repointed === links ? sheet : withField(sheet, "hyperlinks", [...repointed]);
  });
  installChanged(model, next);
}

/** SPEC-128's "annotations" row-permutation adjuster: a sort inside
 * `range` carries notes, threads and links with their records. */
export function permuteAnnotations(model: SpreadsheetModel, sheetId: string, range: GridRange, permutation: readonly { from: number; to: number }[]): void {
  const rows = new Map(permutation.map((entry) => [entry.from, entry.to] as const));
  remapAnnotations(model, sheetId, (ref) => {
    if (ref.col < range.left || ref.col > range.right) {
      return ref;
    }
    const to = rows.get(ref.row);
    return to === undefined ? ref : { row: to, col: ref.col };
  });
}

/** SPEC-121's rename and delete: every `place` location, on every sheet,
 * through remapSheetReferences (a deleted sheet's places become `#REF!`). */
export function remapPlaceSheets(model: SpreadsheetModel, remap: SheetRemap): void {
  const next = model.sheets.map((sheet) => {
    const links = sheetHyperlinks(sheet);
    const mapped = mapPlaces(links, (location) => adjustPlace(location, (text) => remapSheetReferences(text, remap)));
    return mapped === links ? sheet : withField(sheet, "hyperlinks", [...mapped]);
  });
  installChanged(model, next);
}

// --- crossing formats ---------------------------------------------------------------------------

/** A sheet's annotations as a cross-format copy carries them (SPEC-135
 * "Download a copy as…"; `target` the copy's format): notes and links field
 * for field without their sources (rich runs only to an .xlsx); threads kept
 * for an .xlsx, and for any other target each made a note — the opening
 * message followed by one replyParagraphText paragraph per reply, its
 * author and date the opening message's, `resolved` dropped — appended to
 * the cell's own note when it has one. Fields are absent when empty. */
export function crossAnnotations(sheet: Pick<Sheet, "notes" | "threads" | "hyperlinks">, target: SpreadsheetFormat): Pick<Sheet, "notes" | "threads" | "hyperlinks"> {
  const out: Pick<Sheet, "notes" | "threads" | "hyperlinks"> = {};
  let notes: CellNote[] = sheetNotes(sheet).map(({ source: _source, runs, ...note }) => (target === "xlsx" && runs !== undefined ? { ...note, runs } : note));
  const threads = sheetThreads(sheet).map((thread) => ({ ...thread, messages: thread.messages.map(({ source: _source, ...message }) => message) }));
  if (target === "xlsx") {
    if (threads.length > 0) {
      out.threads = threads;
    }
  } else if (threads.length > 0) {
    const taken = new Set(notes.map((note) => note.noteId));
    for (const thread of threads) {
      const [opening, ...replies] = thread.messages;
      if (!opening) {
        continue;
      }
      const paragraphs = [opening.text, ...replies.map(replyParagraphText)].join("\n");
      const own = notes.find((note) => sameRef(note.ref, thread.ref));
      if (own) {
        notes = notes.map((note) => (note === own ? { ...note, text: `${note.text}\n${paragraphs}` } : note));
        continue;
      }
      const noteId = freeId("n", thread.ref, taken);
      taken.add(noteId);
      notes.push({ noteId, ref: { ...thread.ref }, author: opening.author, createdAt: opening.createdAt, text: paragraphs });
    }
    notes = sortedByRef(notes);
  }
  if (notes.length > 0) {
    out.notes = notes;
  }
  const links = sheetHyperlinks(sheet).map(({ source: _source, ...link }) => link);
  if (links.length > 0) {
    out.hyperlinks = links;
  }
  return out;
}
