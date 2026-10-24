// Hyperlinks and bookmarks (SPEC-025): URL safety, bare-URL autolink
// detection, the link/bookmark commands, and internal link targets (heading
// anchors and bookmark names). Framework-free like every lib module, and —
// like lib/clipboard.ts — a command module OUTSIDE commands.ts, built on
// documentModel.ts's shared cloneContainers, so commands.undo.spec.ts's
// completeness guard stays scoped to commands.ts (see the comment there on
// the same split).
import {
  cloneContainers,
  forEachParagraph,
  type DocumentModel,
  type Paragraph,
  type Run,
  type RunFormatting,
} from "./documentModel";
import {
  blockText,
  clampCaret,
  deleteRange as deleteRangeInPlace,
  insertText as insertTextInPlace,
  normalizeRuns,
  paragraphAt,
  paragraphSlicesInRange,
  replaceParagraphAt,
  splitRuns,
  type CaretPosition,
  type ParagraphAddress,
} from "./textEditing";
import {
  caretSelection,
  isCollapsed,
  orderedRange,
  type CommandResult,
  type ModelSelection,
} from "./commands";
import { listCaptions } from "./captions";
import { detectBareUrl } from "./core/urls";

/** A link's destination: exactly one of `href` (external URL) or `anchor`
 * (a bookmark name or heading anchor in this document) is set. */
export interface LinkTarget {
  href?: string;
  anchor?: string;
}

/** Typed `null` seed for link-target refs in `.vue` scripts, which espree
 * bars from writing the `LinkTarget | null` annotation themselves (see
 * commands.ts's `noSelection` for the same constraint). */
export function noLinkTarget(): LinkTarget | null {
  return null;
}

/** One internal-target entry the link dialog lists (AC-19..21): the anchor
 * it links to, the display text, and — for headings — the outline level. */
export interface LinkTargetEntry {
  anchor: string;
  text: string;
  level?: number;
}

/** Typed pass-through for target entries arriving as a runtime-declared
 * Array prop (same espree constraint as styles.ts's asStyleEntries). */
export function asLinkTargetEntries(value: unknown[]): LinkTargetEntry[] {
  return value as LinkTargetEntry[];
}

// --- URL safety (AC-14..18) and autolink detection (AC-5..7) --------------
// Phase one of SPEC-097's two-phase move: the neutral URL helpers now live in
// src/lib/core/urls.ts (so lib/presentation can sanitise slide links) and are
// re-exported here, so every existing import of them from this module keeps
// working unchanged.
export {
  DANGEROUS_SCHEMES,
  urlScheme,
  isDangerousUrl,
  sanitizeHref,
  detectBareUrl,
  type AutolinkMatch,
} from "./core/urls";

// The only two schemes anything here will *fetch*. A stricter rule than
// sanitizeHref's: that one governs what a link may point at (the browser
// navigates it, and the user chose to click), this governs what the app
// will request on the reader's behalf, off a URL a document or a query
// parameter supplied. `data:`/`blob:` carry their own content type and
// would walk past a content-type gate; `file:` and the rest have no
// business being fetched by a document reader at all.
const FETCHABLE_PROTOCOLS = new Set(["http:", "https:"]);

// Nor will it fetch from this computer. The page's CSP admits plain-http
// loopback since SPEC-113 (a local model server listens there, and the model
// gateway pins its own origin), so the CSP no longer keeps a document's
// image or a link's URL off the user's own services — this does. Every
// spelling of "this machine" the URL parser leaves: localhost and its
// subdomains (RFC 6761, with or without the root dot), all of 127.0.0.0/8
// (the parser has already turned 0x7f.1 or 2130706433 into dotted form),
// ::1 and its IPv4-mapped form, and the unspecified addresses, which
// connect locally. Broader on purpose than the gateway's three accepted
// loopback hosts: that list says where a request MAY go, this one where a
// document's request may not.
const LOCAL_MACHINE_HOSTNAME =
  /^(?:(?:[^.]+\.)*localhost|127\.\d+\.\d+\.\d+|0\.0\.0\.0|\[::1?\]|\[::ffff:7f[0-9a-f]{2}:[0-9a-f]{1,4}\])$/;

/** Whether `url` parses as absolute and names this computer (see above). */
export function isLocalMachineUrl(url: string): boolean {
  try {
    return LOCAL_MACHINE_HOSTNAME.test(new URL(url).hostname.replace(/\.$/, ""));
  } catch {
    return false;
  }
}

/** Whether `url` parses as absolute, names a scheme we are willing to fetch
 * and is not on this computer — unless it is the app's own origin
 * (`appOrigin`, the page's by default), which a dev server on localhost is
 * and the CSP's 'self' always admitted. Callers that resolve a relative
 * reference against a base must check the RESOLVED url, not the raw
 * reference — "images/x.png" is not absolute and would fail here on its
 * own. */
export function isFetchableUrl(url: string, appOrigin: string | undefined = globalThis.location?.origin): boolean {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return false;
  }
  if (!FETCHABLE_PROTOCOLS.has(parsed.protocol)) {
    return false;
  }
  return parsed.origin === appOrigin || !isLocalMachineUrl(url);
}

// --- link reading -----------------------------------------------------------

/** The formatting's link target, or null for plain text. */
export function linkTargetOf(formatting: RunFormatting): LinkTarget | null {
  if (formatting.linkHref !== undefined) {
    return { href: formatting.linkHref };
  }
  if (formatting.linkAnchor !== undefined) {
    return { anchor: formatting.linkAnchor };
  }
  return null;
}

function sameLinkTarget(formatting: RunFormatting, target: LinkTarget): boolean {
  return formatting.linkHref === target.href && formatting.linkAnchor === target.anchor;
}

/** `formatting` with its link target replaced (or removed with null) —
 * every other formatting key untouched (AC-33, AC-34). */
export function withLinkTarget(
  formatting: RunFormatting,
  target: LinkTarget | null,
): RunFormatting {
  const rest = { ...formatting };
  delete rest.linkHref;
  delete rest.linkAnchor;
  if (target?.href !== undefined) {
    rest.linkHref = target.href;
  } else if (target?.anchor !== undefined) {
    rest.linkAnchor = target.anchor;
  }
  return rest;
}

/** One contiguous link in a paragraph: its address, covered span, target
 * and visible text — what the chip shows and what Edit/Remove act on. */
export interface LinkExtent {
  address: ParagraphAddress;
  from: number;
  to: number;
  target: LinkTarget;
  text: string;
}

/** Typed `null` seed for link-extent refs in `.vue` scripts (same espree
 * constraint as noLinkTarget). */
export function noLinkExtent(): LinkExtent | null {
  return null;
}

/** The link the caret sits in, expanded to the maximal contiguous span of
 * runs sharing its target (AC-8). At a run boundary the run *ending* there
 * wins, matching insertText's formatting-inheritance rule; a caret at the
 * paragraph start falls through to the first run. */
export function linkAtCaret(model: DocumentModel, caret: CaretPosition): LinkExtent | null {
  const c = clampCaret(model, caret);
  const paragraph = paragraphAt(model, c);
  if (!paragraph) {
    return null;
  }
  const spans: Array<{ formatting: RunFormatting; start: number; end: number }> = [];
  let pos = 0;
  for (const run of paragraph.runs) {
    spans.push({ formatting: run.formatting, start: pos, end: pos + run.text.length });
    pos += run.text.length;
  }
  let index = spans.findIndex((span) => span.start < c.offset && c.offset <= span.end);
  if (index < 0) {
    index = spans.findIndex((span) => span.start === c.offset);
  }
  if (index < 0) {
    return null;
  }
  const target = linkTargetOf(spans[index].formatting);
  if (!target) {
    return null;
  }
  let first = index;
  while (first > 0 && sameLinkTarget(spans[first - 1].formatting, target)) {
    first--;
  }
  let last = index;
  while (last + 1 < spans.length && sameLinkTarget(spans[last + 1].formatting, target)) {
    last++;
  }
  const from = spans[first].start;
  const to = spans[last].end;
  const address: ParagraphAddress = c.cell
    ? { section: c.section, block: c.block, cell: c.cell }
    : { section: c.section, block: c.block };
  return { address, from, to, target, text: blockText(paragraph).slice(from, to) };
}

// --- link commands ----------------------------------------------------------

function setLinkOnParagraphRange(
  model: DocumentModel,
  address: ParagraphAddress,
  from: number,
  to: number,
  target: LinkTarget | null,
): void {
  const paragraph = paragraphAt(model, address);
  if (!paragraph || from >= to) {
    return;
  }
  const [before, tail] = splitRuns(paragraph.runs, from);
  const [within, after] = splitRuns(tail, to - from);
  const changed = within.map((run) => ({
    text: run.text,
    formatting: withLinkTarget(run.formatting, target),
  }));
  replaceParagraphAt(model, address, {
    ...paragraph,
    runs: normalizeRuns([...before, ...changed, ...after]),
  });
}

/** Wraps the selected text in a hyperlink (AC-3). The selection is
 * preserved so the caller's chip/toolbar reflect the new link over the same
 * text. A collapsed selection is a no-op (use insertLinkText instead). */
export function applyLink(
  model: DocumentModel,
  selection: ModelSelection,
  target: LinkTarget,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  for (const { address, from, to } of paragraphSlicesInRange(next, start, end)) {
    setLinkOnParagraphRange(next, address, from, to, target);
  }
  return { model: next, selection: { anchor: start, focus: end } };
}

/** Replaces the selection (or inserts at the caret) with `text` wrapped in
 * a hyperlink — the dialog's save when the display text was typed or edited
 * (AC-2, AC-3). The caret ends up after the linked text. */
export function insertLinkText(
  model: DocumentModel,
  selection: ModelSelection,
  text: string,
  target: LinkTarget,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const base = isCollapsed(selection) ? start : deleteRangeInPlace(next, start, end);
  const after = insertTextInPlace(next, base, text);
  if (after.offset > base.offset) {
    const address: ParagraphAddress = after.cell
      ? { section: after.section, block: after.block, cell: after.cell }
      : { section: after.section, block: after.block };
    setLinkOnParagraphRange(next, address, base.offset, after.offset, target);
  }
  return { model: next, selection: caretSelection(after) };
}

/** Removes the link the caret sits in (its whole extent), or every link in
 * a non-collapsed selection — the text and its other formatting stay
 * exactly as they were (AC-11, AC-33, AC-34). */
export function removeLink(
  model: DocumentModel,
  selection: ModelSelection,
): CommandResult {
  const next = cloneContainers(model);
  if (isCollapsed(selection)) {
    const extent = linkAtCaret(next, selection.focus);
    if (extent) {
      setLinkOnParagraphRange(next, extent.address, extent.from, extent.to, null);
    }
    return { model: next, selection };
  }
  const { start, end } = orderedRange(next, selection);
  for (const { address, from, to } of paragraphSlicesInRange(next, start, end)) {
    setLinkOnParagraphRange(next, address, from, to, null);
  }
  return { model: next, selection };
}

/** Autolinks a bare URL ending at `endOffset` of the addressed paragraph
 * (AC-5, AC-6), or returns null when there is nothing to link. Returns the
 * new MODEL only — deliberately no caret: the caller keeps whatever caret
 * its own space/Enter edit produced (after the space, or at the start of
 * the paragraph Enter created), because the autolink never moves the caret.
 * Text already inside a link is never re-linked. */
export function autolinkAt(
  model: DocumentModel,
  address: ParagraphAddress,
  endOffset: number,
): DocumentModel | null {
  const paragraph = paragraphAt(model, address);
  if (!paragraph) {
    return null;
  }
  const match = detectBareUrl(blockText(paragraph), endOffset);
  if (!match) {
    return null;
  }
  const [, tail] = splitRuns(paragraph.runs, match.start);
  const [within] = splitRuns(tail, match.end - match.start);
  if (within.some((run) => linkTargetOf(run.formatting) !== null)) {
    return null;
  }
  const next = cloneContainers(model);
  setLinkOnParagraphRange(next, address, match.start, match.end, { href: match.href });
  return next;
}

// --- writer support (AC-29..32) ---------------------------------------------

/** A paragraph's inline content laid out for a file-format writer: runs
 * split at every bookmark AND comment-anchor offset (so each lands exactly
 * on a run boundary), the bookmark names per boundary, the comment thread
 * ids whose range starts/ends at each boundary (SPEC-038), and a remapping
 * of opaque-inline anchors (which index the ORIGINAL runs) onto the split
 * list. Shared by the .docx and .odt writers. */
export interface InlineWriteLayout {
  runs: Run[];
  bookmarksAtBoundary: Map<number, string[]>;
  /** Thread ids whose anchored range STARTS at this run boundary. */
  commentStartsAtBoundary: Map<number, string[]>;
  /** Thread ids whose anchored range ENDS at this run boundary. */
  commentEndsAtBoundary: Map<number, string[]>;
  boundaryFor(originalIndex: number): number;
}

export function inlineWriteLayout(paragraph: Paragraph): InlineWriteLayout {
  const bookmarks = paragraph.bookmarks ?? [];
  const commentAnchors = paragraph.commentAnchors ?? [];
  const cutOffsets = Array.from(
    new Set([...bookmarks.map((b) => b.offset), ...commentAnchors.map((a) => a.offset)]),
  ).toSorted((a, b) => a - b);
  const runs: Run[] = [];
  const indexMap: number[] = [0];
  let pos = 0;
  for (const run of paragraph.runs) {
    const end = pos + run.text.length;
    let sliceStart = 0;
    for (const offset of cutOffsets) {
      if (offset > pos && offset < end) {
        // Spread the run so a split text run keeps every property (field/note/
        // styleId/changes); cut offsets fall on run boundaries, never inside a
        // one-character field atom, so a field run is never split (SPEC-055 —
        // a caption carries both a hidden bookmark and a seq field run).
        runs.push({ ...run, text: run.text.slice(sliceStart, offset - pos) });
        sliceStart = offset - pos;
      }
    }
    runs.push({ ...run, text: run.text.slice(sliceStart) });
    indexMap.push(runs.length);
    pos = end;
  }
  const boundaryForOffset = (offset: number): number => {
    let accumulated = 0;
    for (let i = 0; i < runs.length; i++) {
      if (accumulated >= offset) {
        return i;
      }
      accumulated += runs[i].text.length;
    }
    return runs.length;
  };
  const bookmarksAtBoundary = new Map<number, string[]>();
  for (const bookmark of bookmarks) {
    const boundary = boundaryForOffset(bookmark.offset);
    const names = bookmarksAtBoundary.get(boundary) ?? [];
    names.push(bookmark.name);
    bookmarksAtBoundary.set(boundary, names);
  }
  const commentStartsAtBoundary = new Map<number, string[]>();
  const commentEndsAtBoundary = new Map<number, string[]>();
  for (const anchor of commentAnchors) {
    const boundary = boundaryForOffset(anchor.offset);
    const target = anchor.boundary === "start" ? commentStartsAtBoundary : commentEndsAtBoundary;
    const ids = target.get(boundary) ?? [];
    ids.push(anchor.threadId);
    target.set(boundary, ids);
  }
  return {
    runs,
    bookmarksAtBoundary,
    commentStartsAtBoundary,
    commentEndsAtBoundary,
    boundaryFor: (originalIndex: number) => indexMap[Math.min(originalIndex, indexMap.length - 1)],
  };
}

// --- bookmarks (AC-22..28) --------------------------------------------------

/** A bookmark with its resolved document position, in document order. */
export interface BookmarkLocation {
  name: string;
  address: ParagraphAddress;
  offset: number;
}

/** Every bookmark in the document, in document order — the manager's list
 * (AC-23). Hidden bookmarks (auto-created at headings for SPEC-053
 * cross-references) are omitted: the manager never shows them. */
export function listBookmarks(model: DocumentModel): BookmarkLocation[] {
  return listAllBookmarks(model).filter((bookmark) => !bookmark.hidden);
}

/** Every bookmark including hidden ones, in document order — the list
 * cross-reference resolution needs (SPEC-053). `hidden` mirrors the anchor's
 * flag so callers can tell the two apart. */
export function listAllBookmarks(model: DocumentModel): Array<BookmarkLocation & { hidden?: true }> {
  const found: Array<BookmarkLocation & { hidden?: true }> = [];
  forEachParagraph(model, (paragraph, address) => {
    for (const bookmark of paragraph.bookmarks ?? []) {
      found.push({ name: bookmark.name, address, offset: bookmark.offset, ...(bookmark.hidden ? { hidden: true } : {}) });
    }
  });
  return found;
}

/** Like `bookmarkCaret` but resolves hidden bookmarks too (SPEC-053
 * cross-reference targets). */
export function bookmarkCaretIncludingHidden(model: DocumentModel, name: string): CaretPosition | null {
  const location = listAllBookmarks(model).find((bookmark) => bookmark.name === name);
  if (!location) {
    return null;
  }
  return clampCaret(model, {
    section: location.address.section,
    block: location.address.block,
    offset: location.offset,
    cell: location.address.cell,
  });
}

/** "" when the name is valid, else the inline error to show (AC-26..28).
 * `existing` is the document's current bookmark names. */
export function validateBookmarkName(name: string, existing: string[]): string {
  if (name.trim().length === 0) {
    return "Bookmark name cannot be empty";
  }
  if (/\s/.test(name)) {
    return "Bookmark name cannot contain spaces";
  }
  if (existing.includes(name)) {
    return "A bookmark with this name already exists";
  }
  return "";
}

/** Names a location: adds a bookmark at the selection's start (AC-22). The
 * caller validates the name first; the selection passes through unchanged. */
export function insertBookmark(
  model: DocumentModel,
  selection: ModelSelection,
  name: string,
): CommandResult {
  const next = cloneContainers(model);
  const caret = orderedRange(next, selection).start;
  const paragraph = paragraphAt(next, caret);
  if (!paragraph) {
    return { model: next, selection };
  }
  const bookmarks = [...(paragraph.bookmarks ?? []), { name, offset: caret.offset }].toSorted(
    (a, b) => a.offset - b.offset,
  );
  replaceParagraphAt(next, caret, { ...paragraph, bookmarks });
  return { model: next, selection };
}

/** Ensures a heading paragraph carries a bookmark a SPEC-053 cross-reference
 * can target, returning the (possibly new) model and the bookmark's name. If
 * the heading already has a bookmark at its start, that name is reused; else a
 * hidden bookmark named after the heading's derived anchor is inserted (Word's
 * behaviour, and what lets `.odt`'s `text:bookmark-ref` carry the reference).
 * Hidden bookmarks are not shown by the bookmark manager (see listBookmarks). */
export function ensureHeadingBookmark(
  model: DocumentModel,
  address: ParagraphAddress,
  fallbackName: string,
): { model: DocumentModel; name: string } {
  const paragraph = paragraphAt(model, address);
  if (!paragraph) {
    return { model, name: fallbackName };
  }
  const existing = (paragraph.bookmarks ?? []).find((bookmark) => bookmark.offset === 0);
  if (existing) {
    return { model, name: existing.name };
  }
  const next = cloneContainers(model);
  const target = paragraphAt(next, address)!;
  const bookmarks = [...(target.bookmarks ?? []), { name: fallbackName, offset: 0, hidden: true as const }].toSorted(
    (a, b) => a.offset - b.offset,
  );
  replaceParagraphAt(next, address, { ...target, bookmarks });
  return { model: next, name: fallbackName };
}

/** Deletes the named bookmark (AC-25); the document text is untouched. */
export function deleteBookmark(
  model: DocumentModel,
  selection: ModelSelection,
  name: string,
): CommandResult {
  const next = cloneContainers(model);
  const location = listBookmarks(next).find((bookmark) => bookmark.name === name);
  if (!location) {
    return { model: next, selection };
  }
  const paragraph = paragraphAt(next, location.address);
  if (!paragraph) {
    return { model: next, selection };
  }
  const bookmarks = (paragraph.bookmarks ?? []).filter((bookmark) => bookmark.name !== name);
  const replacement: Paragraph = { ...paragraph };
  if (bookmarks.length > 0) {
    replacement.bookmarks = bookmarks;
  } else {
    delete replacement.bookmarks;
  }
  replaceParagraphAt(next, location.address, replacement);
  return { model: next, selection };
}

/** The caret position a bookmark names, or null for an unknown name —
 * the manager's jump target (AC-24). */
export function bookmarkCaret(model: DocumentModel, name: string): CaretPosition | null {
  const location = listBookmarks(model).find((bookmark) => bookmark.name === name);
  if (!location) {
    return null;
  }
  return clampCaret(model, {
    section: location.address.section,
    block: location.address.block,
    offset: location.offset,
    cell: location.address.cell,
  });
}

// --- heading link targets (AC-19..21) ----------------------------------------

/** A heading offered as an internal link target: its derived anchor, text,
 * level and address. */
export interface HeadingTarget {
  anchor: string;
  text: string;
  level: number;
  address: ParagraphAddress;
}

/** The derived anchor id of the `ordinal`-th heading. Deterministic in the
 * heading's text and document position, so a saved link's anchor resolves
 * to the same heading after a round trip — both sides re-derive it. The
 * charset is letters/digits/underscore only, valid as a `w:anchor` bookmark
 * name and an ODF `#ref`. */
export function headingAnchor(text: string, ordinal: number): string {
  const slug = text
    .replace(/[^A-Za-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 40);
  return `_h${ordinal}_${slug || "heading"}`;
}

/** Every heading in the document, in document order, with its derived
 * anchor — the dialog's heading target list (AC-19, AC-21). */
export function listHeadings(model: DocumentModel): HeadingTarget[] {
  const found: Array<Omit<HeadingTarget, "anchor">> = [];
  forEachParagraph(model, (paragraph, address) => {
    if (!paragraph.headingLevel) {
      return;
    }
    found.push({ text: blockText(paragraph), level: paragraph.headingLevel, address });
  });
  return found.map((heading, ordinal) => ({
    ...heading,
    anchor: headingAnchor(heading.text, ordinal),
  }));
}

/** Resolves an internal anchor to a caret: bookmark names first, then
 * derived heading anchors. Null when the target no longer exists. */
export function anchorCaret(model: DocumentModel, anchor: string): CaretPosition | null {
  const bookmark = bookmarkCaret(model, anchor);
  if (bookmark) {
    return bookmark;
  }
  const hiddenBookmark = listAllBookmarks(model).find((entry) => entry.name === anchor);
  if (hiddenBookmark) {
    return clampCaret(model, {
      section: hiddenBookmark.address.section,
      block: hiddenBookmark.address.block,
      offset: hiddenBookmark.offset,
      cell: hiddenBookmark.address.cell,
    });
  }
  const heading = listHeadings(model).find((entry) => entry.anchor === anchor);
  if (heading) {
    return clampCaret(model, {
      section: heading.address.section,
      block: heading.address.block,
      offset: 0,
      cell: heading.address.cell,
    });
  }
  // A table-of-figures entry / caption cross-reference derived anchor
  // (captions.ts's captionAnchor) — resolved so TOF page numbers and entry
  // clicks land on the caption even before any hidden bookmark is inserted.
  const caption = listCaptions(model).find((entry) => entry.anchor === anchor);
  if (!caption) {
    return null;
  }
  return clampCaret(model, {
    section: caption.address.section,
    block: caption.address.block,
    offset: 0,
    cell: caption.address.cell,
  });
}

/** What the chip displays for a target (AC-8): the URL itself, a heading's
 * own text, or the bookmark name. */
export function linkTargetLabel(model: DocumentModel, target: LinkTarget): string {
  if (target.href !== undefined) {
    return target.href;
  }
  const anchor = target.anchor ?? "";
  const heading = listHeadings(model).find((entry) => entry.anchor === anchor);
  return heading ? heading.text : anchor;
}
