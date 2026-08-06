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
import { isImage, isTable, type DocumentModel, type RunFormatting } from "./documentModel";
import {
  blockText,
  clampCaret,
  compareCaret,
  formattingAtOffset,
  normalizeRuns,
  splitRuns,
  type CaretPosition,
} from "./textEditing";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";

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
      if (isTable(block) || isImage(block)) {
        return;
      }
      const text = blockText(block);
      regExp.lastIndex = 0;
      let match = regExp.exec(text);
      while (match !== null) {
        const length = match[0].length;
        if (length === 0) {
          // A zero-width match makes no progress; step past it so exec()
          // cannot loop forever, and don't record an empty highlight.
          regExp.lastIndex += 1;
        } else {
          matches.push({
            section: sectionIndex,
            block: blockIndex,
            start: match.index,
            end: match.index + length,
          });
        }
        match = regExp.exec(text);
      }
    });
  });
  return matches;
}

/** A match as a model selection (its two ends as carets) — how the view
 * highlights and scrolls the current match. */
export function matchSelection(match: FindMatch): ModelSelection {
  return {
    anchor: { section: match.section, block: match.block, offset: match.start },
    focus: { section: match.section, block: match.block, offset: match.end },
  };
}

// Clones the containers a replace rewrites (the model object, its sections
// array, each section and its blocks array); untouched blocks are shared by
// reference, exactly as commands.ts does, so a replace stays O(edit). Kept
// local rather than exported from commands.ts, whose completeness guard
// (commands.undo.spec.ts) requires every export there to be a command/query.
function cloneContainers(model: DocumentModel): DocumentModel {
  return {
    ...model,
    sections: model.sections.map((section) => ({
      ...section,
      blocks: [...section.blocks],
    })),
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
  if (matches.length === 0) {
    return { model: next, selection: caretSelection(clampCaret(next, documentStartCaret())) };
  }
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
    if (isTable(targetBlock) || isImage(targetBlock)) {
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
