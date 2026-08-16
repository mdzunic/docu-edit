// Find and replace over the document model (SPEC-017). Framework-free like
// the rest of lib/: pure functions over the model, unit-testable without a
// DOM. The view (DocumentView.vue) owns the find bar wiring, the on-screen
// highlight overlay and the caret/scroll behaviour; everything that touches
// only the model — compiling a query, locating matches, and rewriting the
// model for a replace — lives here.
//
// Scope (AC "search covers body text"): every paragraph of every section is
// searched. Headers, footers and footnotes are not part of the model yet, so
// they are excluded for free — the find bar states this to the user.
import { cloneContainers, isBreak, isChart, isImage, isShape, isTable, isToc, type Block, type DocumentModel, type FieldSpec, type RunFormatting } from "./documentModel";
import { cachedFieldText } from "./fields";
import {
  clampCaret,
  compareCaret,
  formattingAtOffset,
  normalizeRuns,
  splitRuns,
  type CaretPosition,
} from "./textEditing";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";
import { overlapsDeletion } from "./trackChanges";

/** The three search options, each persisted for the session by the view. */
export interface FindOptions {
  matchCase: boolean;
  wholeWords: boolean;
  regex: boolean;
}

/** One located match: a block address plus the half-open character range
 * [start, end) into that block's concatenated run text. Matches never span a
 * paragraph boundary — search runs per block. */
export interface FindMatch {
  section: number;
  block: number;
  start: number;
  end: number;
  /** SPEC-053 AC-45/46: true when the matched text falls partly or wholly
   * inside a field's displayed value. Replace skips these — a field's value
   * is computed, not literal text a replace may edit — and the view reports
   * them instead of corrupting the field. */
  insideField?: boolean;
}

/** A successfully compiled query — a global RegExp findMatches iterates. */
export interface CompiledQuery {
  regExp: RegExp;
}

/** One highlight rectangle the view paints over a match, in coordinates
 * relative to the page element; `current` marks the distinctly-styled current
 * match (AC). */
export interface HighlightRect {
  left: number;
  top: number;
  width: number;
  height: number;
  current: boolean;
}

/** Typed empty seeds for the view's `shallowRef`s — `.vue` scripts can't
 * write the `FindMatch[]` / `HighlightRect[]` generic annotations themselves
 * (see textEditing.ts's `noCaret` for the same espree constraint). */
export function noMatches(): FindMatch[] {
  return [];
}

export function noHighlightRects(): HighlightRect[] {
  return [];
}

/** A query that could not be compiled (an invalid regular expression), so the
 * search must not run and the bar shows the message inline (AC). */
export interface QueryError {
  error: string;
}

export const INVALID_REGEX_MESSAGE = "Invalid regular expression";

/** Narrows a compileQuery result to the error case (a `.vue` template can't
 * write the `in` narrowing itself; exported for symmetry with the view). */
export function isQueryError(value: CompiledQuery | QueryError): value is QueryError {
  return (value as QueryError).error !== undefined;
}

/** Escapes a literal string so it matches verbatim when used as a regex
 * source (non-regex search). */
export function escapeRegExp(literal: string): string {
  return literal.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Compiles a query string under the current options into a global RegExp, or
 * an error for an invalid regular expression (AC: invalid regex shows an
 * inline error and does not run). Returns null for an empty query — there is
 * nothing to search for, which the view renders as "0 of 0" rather than an
 * error. A non-regex query is escaped, so only regex mode can fail to
 * compile. "Whole words only" wraps the source in word boundaries, combining
 * with every mode. */
export function compileQuery(
  query: string,
  options: FindOptions,
): CompiledQuery | QueryError | null {
  if (query === "") {
    return null;
  }
  let source = options.regex ? query : escapeRegExp(query);
  if (options.wholeWords) {
    source = `\\b(?:${source})\\b`;
  }
  const flags = options.matchCase ? "g" : "gi";
  try {
    return { regExp: new RegExp(source, flags) };
  } catch {
    return { error: INVALID_REGEX_MESSAGE };
  }
}

/** All matches for a compiled query, in document order. Zero-length matches
 * (e.g. the regex `a*` against text with no `a`) are skipped and never stall
 * the scan — lastIndex is advanced past them — so a pathological query yields
 * no matches rather than an infinite loop. */
export function findMatches(model: DocumentModel, compiled: CompiledQuery): FindMatch[] {
  const matches: FindMatch[] = [];
  const regExp = compiled.regExp;
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      // Table/image content is not searched (SPEC-022/024): a FindMatch
      // addresses a top-level block's character space, which neither has.
      // A TOC's generated text is not searched either (SPEC-033) — it is
      // not directly editable, so a replace could never land there anyway.
      if (isTable(block) || isImage(block) || isShape(block) || isChart(block) || isToc(block)) {
        return;
      }
      // SPEC-053 AC-44: Find matches the text a field DISPLAYS, not the
      // one-character placeholder in the model text. The display mapping lets
      // a match's display offsets convert back to model offsets (for a
      // replace's splitRuns) and flags matches that touch a field.
      const display = blockDisplayInfo(block);
      const text = display.text;
      regExp.lastIndex = 0;
      let match = regExp.exec(text);
      while (match !== null) {
        const length = match[0].length;
        if (length === 0) {
          // A zero-width match makes no progress; step past it so exec()
          // cannot loop forever, and don't record an empty highlight.
          regExp.lastIndex += 1;
          match = regExp.exec(text);
          continue;
        }
        const displayStart = match.index;
        const displayEnd = match.index + length;
        const modelStart = display.toModel(displayStart);
        const modelEnd = display.toModel(displayEnd);
        const insideField = display.touchesField(displayStart, displayEnd);
        // A match inside a tracked deletion is invisible text (SPEC-039
        // AC-17) — matching stays against the full text so offsets keep
        // lining up, but the match itself is dropped.
        if (block.type !== "paragraph" || !overlapsDeletion(block, modelStart, modelEnd)) {
          matches.push({
            section: sectionIndex,
            block: blockIndex,
            start: modelStart,
            end: modelEnd,
            ...(insideField ? { insideField: true } : {}),
          });
        }
        match = regExp.exec(text);
      }
    });
  });
  return matches;
}

/** How many of `matches` were skipped for landing inside a field's value
 * (SPEC-053 AC-46) — the view reports this after a replace. */
export function fieldSkippedCount(matches: FindMatch[]): number {
  return matches.filter((match) => match.insideField).length;
}

/** A block's displayed text (SPEC-053): each field run contributes its shown
 * value rather than the one-character placeholder, with a map back to model
 * offsets and a test for whether a display range touches a field. Plain runs
 * advance display and model offsets together; a field run spans its value's
 * length in display but exactly one character in the model. */
function blockDisplayInfo(block: Block): {
  text: string;
  toModel: (displayOffset: number) => number;
  touchesField: (start: number, end: number) => boolean;
} {
  const runs = "runs" in block ? block.runs : [];
  let text = "";
  // Per run: its display start/length and its model start, so an offset maps
  // back precisely; a field run records its own span for the touches test.
  const segments: Array<{ dStart: number; dLen: number; mStart: number; mLen: number; field: boolean }> = [];
  let mOffset = 0;
  for (const run of runs) {
    const shown = run.field ? cachedFieldText(run.field as FieldSpec) : run.text;
    segments.push({ dStart: text.length, dLen: shown.length, mStart: mOffset, mLen: run.text.length, field: !!run.field });
    text += shown;
    mOffset += run.text.length;
  }
  const toModel = (d: number): number => {
    for (const seg of segments) {
      if (d < seg.dStart + seg.dLen) {
        // Inside a field, snap to its model start; inside a plain run, advance
        // in step with the offset.
        return seg.field ? seg.mStart : seg.mStart + (d - seg.dStart);
      }
    }
    return mOffset;
  };
  const touchesField = (start: number, end: number): boolean =>
    segments.some((seg) => seg.field && seg.dLen > 0 && start < seg.dStart + seg.dLen && end > seg.dStart);
  return { text, toModel, touchesField };
}

/** A match as a model selection (its two ends as carets) — how the view
 * highlights and scrolls the current match. */
export function matchSelection(match: FindMatch): ModelSelection {
  return {
    anchor: { section: match.section, block: match.block, offset: match.start },
    focus: { section: match.section, block: match.block, offset: match.end },
  };
}

/** The formatting the replacement text takes: that of the first run of the
 * text being replaced (AC). `within` is the run slice covering the match. */
function replacementFormatting(within: { formatting: RunFormatting }[], fallback: RunFormatting): RunFormatting {
  return within.length > 0 ? within[0].formatting : fallback;
}

/** Rewrites the model so every match in `matches` becomes `replacement`. The
 * whole set is applied in one pass and returned as a single CommandResult, so
 * the view records it as one undo step (AC: "Replace all" is a single undo
 * step). Each replacement takes the formatting of the first run it replaced
 * (AC). `matches` need not be sorted; within each block they are applied
 * right-to-left so earlier offsets stay valid. The returned selection is a
 * caret just after the first (document-order) replacement — a stable,
 * well-defined position independent of how later matches shifted. */
export function replaceMatches(
  model: DocumentModel,
  matches: FindMatch[],
  replacement: string,
): CommandResult {
  const next = cloneContainers(model);
  // SPEC-053 AC-45: a match inside a field's value is never rewritten — its
  // value is computed, not literal text. The view reports the skipped count
  // (fieldSkippedCount); here they are simply filtered out.
  const editable = matches.filter((match) => !match.insideField);
  if (editable.length === 0) {
    return { model: next, selection: caretSelection(clampCaret(next, documentStartCaret())) };
  }
  matches = editable;
  // Group by block so a block is rewritten once even with several matches.
  const byBlock = new Map<string, FindMatch[]>();
  for (const match of matches) {
    const key = `${match.section}:${match.block}`;
    const list = byBlock.get(key);
    if (list) {
      list.push(match);
    } else {
      byBlock.set(key, [match]);
    }
  }
  for (const [key, blockMatches] of byBlock) {
    const [section, block] = key.split(":").map(Number);
    const targetBlock = next.sections[section].blocks[block];
    if (isTable(targetBlock) || isImage(targetBlock) || isShape(targetBlock) || isChart(targetBlock) || isBreak(targetBlock) || isToc(targetBlock)) {
      continue;
    }
    let runs = targetBlock.runs;
    // Right-to-left: a later replacement never shifts an earlier match's
    // offsets, so each split uses offsets that are still valid.
    const ordered = [...blockMatches].sort((a, b) => b.start - a.start);
    for (const match of ordered) {
      const [before, tail] = splitRuns(runs, match.start);
      const [within, after] = splitRuns(tail, match.end - match.start);
      const formatting = replacementFormatting(
        within,
        formattingAtOffset(targetBlock, match.start),
      );
      const inserted = replacement
        ? [{ text: replacement, formatting: { ...formatting } }]
        : [];
      runs = normalizeRuns([...before, ...inserted, ...after]);
    }
    next.sections[section].blocks[block] = { ...targetBlock, runs };
  }
  const first = firstInDocumentOrder(matches);
  const caret = clampCaret(next, {
    section: first.section,
    block: first.block,
    offset: first.start + replacement.length,
  });
  return { model: next, selection: caretSelection(caret) };
}

/** Replaces a single match (the "Replace" button): a one-element
 * replaceMatches, sharing its single-undo-step and formatting behaviour. */
export function replaceMatch(
  model: DocumentModel,
  match: FindMatch,
  replacement: string,
): CommandResult {
  return replaceMatches(model, [match], replacement);
}

function firstInDocumentOrder(matches: FindMatch[]): FindMatch {
  let first = matches[0];
  for (const match of matches) {
    if (
      match.section < first.section ||
      (match.section === first.section && match.block < first.block) ||
      (match.section === first.section && match.block === first.block && match.start < first.start)
    ) {
      first = match;
    }
  }
  return first;
}

function documentStartCaret(): CaretPosition {
  return { section: 0, block: 0, offset: 0 };
}

/** The index of the match the view should treat as current after a replace or
 * an edit: the first match at or after `anchor` in document order, wrapping to
 * the first match when none follow, or -1 when there are no matches. Keeps the
 * "current" position sensible as the match list shifts under the caret. */
export function indexAtOrAfter(matches: FindMatch[], anchor: CaretPosition): number {
  if (matches.length === 0) {
    return -1;
  }
  for (let i = 0; i < matches.length; i++) {
    if (compareCaret(matchSelection(matches[i]).anchor, anchor) >= 0) {
      return i;
    }
  }
  return 0;
}
