// Cut/copy/paste over the document model (SPEC-016). Framework-free: the
// fragment a copy captures, the three clipboard flavours it serialises to,
// the validating parser for the internal payload, and the paste commands
// that splice a fragment (or plain text) back into a model. DocumentView
// owns only the DOM-facing wiring (ClipboardEvent handlers, the paste
// options chip, the context menu); everything here is unit-testable.
//
// The paste commands follow commands.ts's immutability contract: containers
// are cloned, blocks are replaced (never mutated), untouched paragraphs are
// shared by reference with the input model.
import {
  cloneContainers,
  type DocumentModel,
  type Paragraph,
  type ParagraphAlignment,
  type ParagraphFormatting,
  type PendingFormat,
  type Run,
  type RunFormatting,
} from "./documentModel";
import {
  applyFormatPatch,
  blockLength,
  blockText,
  compareCaret,
  deleteRange as deleteRangeInPlace,
  formattingAtOffset,
  insertText as insertTextInPlace,
  normalizeRuns,
  paragraphAt,
  paragraphAtCaret,
  paragraphSlicesInRange,
  spliceParagraphs,
  splitParagraph as splitParagraphInPlace,
  splitRuns,
  type CaretPosition,
  type ParagraphAddress,
} from "./textEditing";
import { orderedRange, type CommandResult, type ModelSelection } from "./commands";

/** The private MIME type the internal payload travels under. Only
 * `ClipboardEvent.clipboardData.setData/getData` can carry it —
 * `navigator.clipboard.read()/write()` accept a fixed allowlist of types,
 * so the async API is never used for the internal flavour. */
export const DOC_MODEL_MIME = "application/x-docu-edit-docmodel+json";

/** The message shown when an image paste is gated (AC-46). SPEC-024 replaces
 * this path with a real image insert. */
export const IMAGES_UNSUPPORTED_MESSAGE = "Images aren't supported yet";

/** The paste options chip's three interpretations of one paste (SPEC-016):
 * `source` is the verbatim/imported paste, `match` re-resolves formatting to
 * the destination, `text` is the plain-text path (Ctrl/Cmd+Shift+V). */
export type PasteMode = "source" | "match" | "text";

/** Typed pass-through for a paste mode coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asPasteMode(value: string): PasteMode {
  return value as PasteMode;
}

/** What a paste keeps around so the chip can re-apply it: the model
 * fragment (internal payload or HTML import; null for a text-only paste)
 * and the plain-text form. */
export interface PastePayload {
  fragment: Paragraph[] | null;
  text: string;
}

/** Typed `null` seed for paste-payload variables in `.vue` scripts (espree
 * bars the annotation — see commands.ts's noSelection). */
export function noPastePayload(): PastePayload | null {
  return null;
}

/** What copy/cut remember locally: the flavours written plus the fragment
 * they came from. The context menu's paste — which reads the clipboard via
 * the async API and therefore can never see the private MIME type — matches
 * the clipboard's `text/html` against `html` to recognise its own copy and
 * re-use `fragment` for exact fidelity. */
export interface ClipboardSnapshot {
  text: string;
  html: string;
  fragment: Paragraph[];
}

/** Typed `null` seed for clipboard-snapshot variables in `.vue` scripts. */
export function noClipboardSnapshot(): ClipboardSnapshot | null {
  return null;
}

// --- fragment capture -------------------------------------------------------

/** A deep copy of a paragraph carrying only the fields a fragment keeps:
 * runs plus paragraph-level formatting/style/heading. Document-scoped data
 * and preservation side-lists (opaqueInlines) stay behind — their anchors
 * would be meaningless inside a spliced fragment. */
function copyParagraph(block: Paragraph, runs: Run[]): Paragraph {
  const copy: Paragraph = {
    type: "paragraph",
    runs: normalizeRuns(
      runs.map((run) => ({ text: run.text, formatting: { ...run.formatting } })),
    ),
  };
  if (block.formatting && Object.keys(block.formatting).length > 0) {
    copy.formatting = { ...block.formatting };
  }
  if (block.styleId !== undefined) {
    copy.styleId = block.styleId;
  }
  if (block.headingLevel !== undefined) {
    copy.headingLevel = block.headingLevel;
  }
  return copy;
}

/** The `Paragraph[]` spanning a selection: boundary runs split at the
 * selection edges, every run re-normalised so the fragment obeys the model
 * invariants. Empty for a collapsed selection. A selection that covers a
 * table contributes the table's cell paragraphs in reading order — the
 * fragment stays a flat paragraph list (SPEC-022: table structure is not a
 * clipboard flavour yet), so a cut never silently loses cell text. */
export function extractFragment(
  model: DocumentModel,
  selection: ModelSelection,
): Paragraph[] {
  const { start, end } = orderedRange(model, selection);
  if (compareCaret(start, end) === 0) {
    return [];
  }
  const fragment: Paragraph[] = [];
  for (const { address, from, to } of paragraphSlicesInRange(model, start, end)) {
    const paragraph = paragraphAt(model, address);
    if (!paragraph) {
      continue;
    }
    const [, tail] = splitRuns(paragraph.runs, from);
    const [within] = splitRuns(tail, to - from);
    fragment.push(copyParagraph(paragraph, within));
  }
  return fragment;
}

// --- clipboard flavours -----------------------------------------------------

/** The `text/plain` flavour: paragraphs joined with `\n`, no formatting. */
export function fragmentText(fragment: Paragraph[]): string {
  return fragment.map(blockText).join("\n");
}

function escapeHtmlText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeHtmlAttribute(value: string): string {
  return escapeHtmlText(value).replace(/"/g, "&quot;");
}

function runHtml(run: Run): string {
  let html = escapeHtmlText(run.text).replace(/\n/g, "<br>");
  const formatting = run.formatting;
  const styles: string[] = [];
  if (formatting.color) {
    styles.push(`color:${formatting.color}`);
  }
  if (formatting.fontFamily) {
    styles.push(`font-family:${formatting.fontFamily}`);
  }
  if (formatting.fontSizePt) {
    styles.push(`font-size:${formatting.fontSizePt}pt`);
  }
  if (styles.length > 0) {
    html = `<span style="${escapeHtmlAttribute(styles.join(";"))}">${html}</span>`;
  }
  if (formatting.vertAlign === "subscript") {
    html = `<sub>${html}</sub>`;
  } else if (formatting.vertAlign === "superscript") {
    html = `<sup>${html}</sup>`;
  }
  if (formatting.strikethrough) {
    html = `<s>${html}</s>`;
  }
  if (formatting.underline) {
    html = `<u>${html}</u>`;
  }
  if (formatting.italic) {
    html = `<i>${html}</i>`;
  }
  if (formatting.bold) {
    html = `<b>${html}</b>`;
  }
  return html;
}

/** The `text/html` flavour: a minimal self-generated fragment — `<p>`,
 * `<h1>`–`<h6>`, `<b>/<i>/<u>/<s>/<sub>/<sup>` and `<span>`s carrying only
 * colour/font-family/font-size, plus `text-align` on blocks. No classes, no
 * layout markup, nothing the import table cannot map back. */
export function fragmentHtml(fragment: Paragraph[]): string {
  return fragment
    .map((block) => {
      const level = block.headingLevel;
      const tag = level ? `h${Math.min(Math.max(level, 1), 6)}` : "p";
      const alignment = block.formatting?.alignment;
      const open = alignment
        ? `<${tag} style="text-align:${alignment}">`
        : `<${tag}>`;
      return `${open}${block.runs.map(runHtml).join("")}</${tag}>`;
    })
    .join("");
}

/** The internal flavour: the fragment labelled with the MIME type and a
 * version, exactly the format the paste-side parser validates. */
export function serializeFragment(fragment: Paragraph[]): string {
  return JSON.stringify({
    mimeType: DOC_MODEL_MIME,
    version: 1,
    blocks: fragment,
  });
}

/** All three flavours of one fragment, ready for `setData`. */
export function clipboardFlavours(fragment: Paragraph[]): {
  text: string;
  html: string;
  payload: string;
} {
  return {
    text: fragmentText(fragment),
    html: fragmentHtml(fragment),
    payload: serializeFragment(fragment),
  };
}

// --- internal payload parsing ----------------------------------------------
// The clipboard is an untrusted boundary even for our own MIME type —
// anything can write any string under any key — so the parser rebuilds the
// fragment field by field from a whitelist instead of splicing parsed JSON
// into the model.

const ALIGNMENTS = new Set(["left", "center", "right", "justify"]);
const SAFE_COLOR = /^#[0-9a-f]{3,8}$/i;
const SAFE_COLOR_NAME = /^[a-zA-Z]+$/;

function isSafeColor(value: unknown): value is string {
  return (
    typeof value === "string" &&
    (SAFE_COLOR.test(value) || SAFE_COLOR_NAME.test(value))
  );
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function sanitizeRunFormatting(value: unknown): RunFormatting {
  const formatting: RunFormatting = {};
  if (typeof value !== "object" || value === null) {
    return formatting;
  }
  const raw = value as Record<string, unknown>;
  for (const key of ["bold", "italic", "underline", "strikethrough"] as const) {
    if (raw[key] === true) {
      formatting[key] = true;
    }
  }
  if (typeof raw.fontFamily === "string" && raw.fontFamily.length > 0) {
    formatting.fontFamily = raw.fontFamily;
  }
  if (isFiniteNumber(raw.fontSizePt) && raw.fontSizePt > 0) {
    formatting.fontSizePt = raw.fontSizePt;
  }
  if (isSafeColor(raw.color)) {
    formatting.color = raw.color;
  }
  if (raw.highlight === false || isSafeColor(raw.highlight)) {
    formatting.highlight = raw.highlight;
  }
  if (isSafeColor(raw.shadingFill)) {
    formatting.shadingFill = raw.shadingFill;
  }
  if (raw.vertAlign === "subscript" || raw.vertAlign === "superscript") {
    formatting.vertAlign = raw.vertAlign;
  }
  return formatting;
}

const NUMERIC_PARAGRAPH_KEYS = [
  "indentStartCm",
  "indentEndCm",
  "indentFirstLineCm",
  "spaceBeforePt",
  "spaceAfterPt",
  "lineSpacing",
  "lineSpacingPt",
] as const;

function sanitizeParagraphFormatting(value: unknown): ParagraphFormatting | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  const raw = value as Record<string, unknown>;
  const formatting: ParagraphFormatting = {};
  if (typeof raw.alignment === "string" && ALIGNMENTS.has(raw.alignment)) {
    formatting.alignment = raw.alignment as ParagraphAlignment;
  }
  for (const key of NUMERIC_PARAGRAPH_KEYS) {
    if (isFiniteNumber(raw[key])) {
      formatting[key] = raw[key] as number;
    }
  }
  if (raw.lineSpacingRule === "exact") {
    formatting.lineSpacingRule = "exact";
  }
  if (raw.direction === "rtl") {
    formatting.direction = "rtl";
  }
  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

function sanitizeParagraph(value: unknown): Paragraph | null {
  if (typeof value !== "object" || value === null) {
    return null;
  }
  const raw = value as Record<string, unknown>;
  if (raw.type !== "paragraph" || !Array.isArray(raw.runs)) {
    return null;
  }
  const runs: Run[] = [];
  for (const entry of raw.runs) {
    if (typeof entry !== "object" || entry === null) {
      return null;
    }
    const rawRun = entry as Record<string, unknown>;
    if (typeof rawRun.text !== "string") {
      return null;
    }
    runs.push({ text: rawRun.text, formatting: sanitizeRunFormatting(rawRun.formatting) });
  }
  const paragraph: Paragraph = { type: "paragraph", runs: normalizeRuns(runs) };
  const formatting = sanitizeParagraphFormatting(raw.formatting);
  if (formatting) {
    paragraph.formatting = formatting;
  }
  if (typeof raw.styleId === "string" && raw.styleId.length > 0) {
    paragraph.styleId = raw.styleId;
  }
  if (
    isFiniteNumber(raw.headingLevel) &&
    Number.isInteger(raw.headingLevel) &&
    raw.headingLevel >= 1 &&
    raw.headingLevel <= 9
  ) {
    paragraph.headingLevel = raw.headingLevel;
  }
  return paragraph;
}

/** Parses (and validates) the internal clipboard payload. Detection is the
 * caller's non-empty `getData(DOC_MODEL_MIME)`; here `mimeType` and
 * `version` are checked as defence-in-depth and every block is rebuilt from
 * whitelisted fields. Null when anything doesn't hold — the caller falls
 * through to the `text/html` flavour. */
export function parseFragmentPayload(raw: string): Paragraph[] | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (typeof parsed !== "object" || parsed === null) {
    return null;
  }
  const payload = parsed as Record<string, unknown>;
  if (payload.mimeType !== DOC_MODEL_MIME || payload.version !== 1) {
    return null;
  }
  if (!Array.isArray(payload.blocks)) {
    return null;
  }
  const blocks: Paragraph[] = [];
  for (const entry of payload.blocks) {
    const paragraph = sanitizeParagraph(entry);
    if (!paragraph) {
      return null;
    }
    blocks.push(paragraph);
  }
  return blocks;
}

// --- paste commands ---------------------------------------------------------

function collapsedResult(
  model: DocumentModel,
  caret: { section: number; block: number; offset: number },
): CommandResult {
  return { model, selection: { anchor: caret, focus: caret } };
}

/** Splices a fragment into the model at the selection (replacing it when
 * non-collapsed — AC-34). A single-block fragment carries no paragraph
 * break: its runs are spliced into the destination paragraph, which keeps
 * its own paragraph formatting — unless the destination is empty, in which
 * case the pasted paragraph replaces it wholesale (so pasting a copied
 * heading onto an empty line restores the heading). A multi-block fragment
 * splits the destination paragraph at the caret (AC-35): the first block
 * merges into the text before the caret and the last into the text after
 * it (both keep the destination's paragraph formatting); an empty half is
 * instead replaced by the boundary block verbatim, so pasting whole
 * paragraphs restores them exactly (AC-10..13). Blocks between the
 * boundaries land verbatim. The caret ends at the end of the pasted
 * content. */
export function pasteBlocks(
  model: DocumentModel,
  selection: ModelSelection,
  fragment: Paragraph[],
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  const caret =
    compareCaret(start, end) === 0 ? start : deleteRangeInPlace(next, start, end);
  // Fresh copies so repeated pastes of one payload never share objects.
  const blocks = fragment.map((block) => copyParagraph(block, block.runs));
  if (blocks.length === 0) {
    return collapsedResult(next, caret);
  }
  const target = paragraphAtCaret(next, caret);
  if (!target) {
    return collapsedResult(next, caret);
  }
  // Splices land in the caret's own container — the cell's block list when
  // the caret sits in a table (SPEC-022), so a paste into a cell grows the
  // cell, never the section around the table.
  const base = caret.cell ? caret.cell.block : caret.block;
  const addressAt = (index: number): ParagraphAddress =>
    caret.cell
      ? { section: caret.section, block: caret.block, cell: { ...caret.cell, block: index } }
      : { section: caret.section, block: index };
  const caretAt = (index: number, offset: number): CaretPosition =>
    caret.cell
      ? { section: caret.section, block: caret.block, offset, cell: { ...caret.cell, block: index } }
      : { section: caret.section, block: index, offset };
  const [before, after] = splitRuns(target.runs, caret.offset);
  if (blocks.length === 1) {
    const only = blocks[0];
    const emptyTarget = blockLength(target) === 0;
    const merged = emptyTarget
      ? only
      : { ...target, runs: normalizeRuns([...before, ...only.runs, ...after]) };
    spliceParagraphs(next, addressAt(base), 1, merged);
    return collapsedResult(next, caretAt(base, (emptyTarget ? 0 : caret.offset) + blockLength(only)));
  }
  const first = blocks[0];
  const last = blocks[blocks.length - 1];
  const head =
    before.length === 0
      ? first
      : { ...target, runs: normalizeRuns([...before, ...first.runs]) };
  const tail =
    after.length === 0
      ? last
      : { ...target, runs: normalizeRuns([...last.runs, ...after]) };
  spliceParagraphs(next, addressAt(base), 1, head, ...blocks.slice(1, -1), tail);
  return collapsedResult(next, caretAt(base + blocks.length - 1, blockLength(last)));
}

/** The plain-text paste path (Ctrl/Cmd+Shift+V and the chip's "Keep text
 * only"): each newline becomes a paragraph break, and the inserted text
 * takes the formatting of the paragraph at the caret (AC-28) — insertText's
 * inheritance rule, with the caller's pending format still honoured the way
 * typing would honour it. Replaces a non-collapsed selection. */
export function pastePlainText(
  model: DocumentModel,
  selection: ModelSelection,
  text: string,
  formatOverride: PendingFormat | null = null,
): CommandResult {
  const next = cloneContainers(model);
  const { start, end } = orderedRange(next, selection);
  let caret =
    compareCaret(start, end) === 0 ? start : deleteRangeInPlace(next, start, end);
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  caret = insertTextInPlace(next, caret, lines[0], formatOverride);
  for (let i = 1; i < lines.length; i++) {
    caret = splitParagraphInPlace(next, caret);
    caret = insertTextInPlace(next, caret, lines[i], formatOverride);
  }
  return collapsedResult(next, caret);
}

/** "Match destination formatting": paragraph boundaries and run text are
 * kept, everything else is re-resolved to the destination — each block
 * takes the destination paragraph's formatting/style/heading, each run the
 * formatting a freshly typed character would have. */
export function matchDestination(
  fragment: Paragraph[],
  destination: Paragraph,
  typedFormatting: RunFormatting,
): Paragraph[] {
  return fragment.map((block) => {
    const matched: Paragraph = {
      type: "paragraph",
      runs: normalizeRuns(
        block.runs.map((run) => ({
          text: run.text,
          formatting: { ...typedFormatting },
        })),
      ),
    };
    if (destination.formatting && Object.keys(destination.formatting).length > 0) {
      matched.formatting = { ...destination.formatting };
    }
    if (destination.styleId !== undefined) {
      matched.styleId = destination.styleId;
    }
    if (destination.headingLevel !== undefined) {
      matched.headingLevel = destination.headingLevel;
    }
    return matched;
  });
}

/** One paste, interpreted under a chip mode — the initial paste and every
 * chip re-application go through here so they cannot drift apart. `pending`
 * is the caret's pending format at paste time (what a freshly typed
 * character would carry). A payload without a fragment (text-only clipboard)
 * always takes the plain-text path. */
export function executePaste(
  model: DocumentModel,
  selection: ModelSelection,
  payload: PastePayload,
  mode: PasteMode,
  pending: PendingFormat | null = null,
): CommandResult {
  if (!payload.fragment || mode === "text") {
    const text = payload.text || (payload.fragment ? fragmentText(payload.fragment) : "");
    return pastePlainText(model, selection, text, pending);
  }
  if (mode === "match") {
    const { start } = orderedRange(model, selection);
    const destination = paragraphAtCaret(model, start);
    if (!destination) {
      const text = payload.text || fragmentText(payload.fragment);
      return pastePlainText(model, selection, text, pending);
    }
    const typedFormatting = applyFormatPatch(
      formattingAtOffset(destination, start.offset),
      pending ?? {},
    );
    return pasteBlocks(
      model,
      selection,
      matchDestination(payload.fragment, destination, typedFormatting),
    );
  }
  return pasteBlocks(model, selection, payload.fragment);
}
