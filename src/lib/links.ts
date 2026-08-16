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

// --- URL safety (AC-14..18) -------------------------------------------------

// The schemes that execute script or smuggle content when navigated to.
// Rejected on entry (the link dialog), stripped on every other ingress path
// (file open, paste, autolink) and re-checked before any navigation.
const DANGEROUS_SCHEMES = new Set(["javascript", "data", "vbscript"]);

/** The URL's scheme, lowercased — parsed the forgiving way browsers do:
 * ASCII control characters and whitespace are stripped first, so
 * "  java\tscript:alert(1)" is still recognised as `javascript`. */
export function urlScheme(url: string): string | null {
  let cleaned = "";
  for (const char of url) {
    if (char.charCodeAt(0) > 0x20) {
      cleaned += char;
    }
  }
  const match = /^([a-z][a-z0-9+.-]*):/.exec(cleaned.toLowerCase());
  return match ? match[1] : null;
}

/** Whether the URL's scheme is one of the dangerous set above. */
export function isDangerousUrl(url: string): boolean {
  const scheme = urlScheme(url);
  return scheme !== null && DANGEROUS_SCHEMES.has(scheme);
}

/** The URL a link may actually carry: trimmed, or null when it must be
 * rejected (dialog entry, AC-14..16) or stripped (file open / paste,
 * AC-17). */
export function sanitizeHref(url: string): string | null {
  const trimmed = url.trim();
  if (trimmed.length === 0 || isDangerousUrl(trimmed)) {
    return null;
  }
  return trimmed;
}

// --- autolink detection (AC-5..7) -------------------------------------------

/** A bare URL found in a paragraph's text: the covered character span plus
 * the (possibly scheme-prefixed) href to link it to. */
export interface AutolinkMatch {
  start: number;
  end: number;
  href: string;
}

// Punctuation a sentence hangs on a URL's end ("see http://x.com.") — never
// part of the autolinked target, matching what word processors do.
const TRAILING_PUNCTUATION = /[.,;:!?)\]}>"']+$/;
const LEADING_WRAPPERS = /^[([{<"']+/;

/** Finds a bare URL ending at `endOffset` in `text` — the word between the
 * previous whitespace and `endOffset`, with surrounding punctuation peeled
 * off. Recognises http(s) URLs and `www.` shorthand (prefixed with https://
 * in the href). Dangerous schemes never match (AC-17's paste path). */
export function detectBareUrl(text: string, endOffset: number): AutolinkMatch | null {
  const upto = text.slice(0, Math.max(0, Math.min(endOffset, text.length)));
  let start = upto.length;
  while (start > 0 && !/\s/.test(upto[start - 1])) {
    start--;
  }
  let word = upto.slice(start);
  const lead = LEADING_WRAPPERS.exec(word);
  if (lead) {
    start += lead[0].length;
    word = word.slice(lead[0].length);
  }
  let end = upto.length;
  const trail = TRAILING_PUNCTUATION.exec(word);
  if (trail) {
    word = word.slice(0, word.length - trail[0].length);
    end -= trail[0].length;
  }
  if (word.length === 0) {
    return null;
  }
  let href: string | null = null;
  if (/^https?:\/\/\S+$/i.test(word)) {
    href = word;
  } else if (/^www\.[^\s.]+\.\S+$/i.test(word)) {
    href = `https://${word}`;
  }
  if (href === null || isDangerousUrl(href)) {
    return null;
  }
  return { start, end, href };
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
  ).sort((a, b) => a - b);
  const runs: Run[] = [];
  const indexMap: number[] = [0];
  let pos = 0;
  for (const run of paragraph.runs) {
    const end = pos + run.text.length;
    let sliceStart = 0;
    for (const offset of cutOffsets) {
      if (offset > pos && offset < end) {
        runs.push({ text: run.text.slice(sliceStart, offset - pos), formatting: run.formatting });
        sliceStart = offset - pos;
      }
    }
    runs.push({ text: run.text.slice(sliceStart), formatting: run.formatting });
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
 * (AC-23). */
export function listBookmarks(model: DocumentModel): BookmarkLocation[] {
  const found: BookmarkLocation[] = [];
  forEachParagraph(model, (paragraph, address) => {
    for (const bookmark of paragraph.bookmarks ?? []) {
      found.push({ name: bookmark.name, address, offset: bookmark.offset });
    }
  });
  return found;
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
  const bookmarks = [...(paragraph.bookmarks ?? []), { name, offset: caret.offset }].sort(
    (a, b) => a.offset - b.offset,
  );
  replaceParagraphAt(next, caret, { ...paragraph, bookmarks });
  return { model: next, selection };
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
  const heading = listHeadings(model).find((entry) => entry.anchor === anchor);
  if (!heading) {
    return null;
  }
  return clampCaret(model, {
    section: heading.address.section,
    block: heading.address.block,
    offset: 0,
    cell: heading.address.cell,
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
