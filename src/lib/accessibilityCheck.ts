// The accessibility checker's detection engine (SPEC-043): scans the
// document model for issues a screen-reader user or someone who cannot
// perceive colour would hit, and groups them by severity. Framework-free
// (no Vue imports) so it is unit-testable without mounting a component and
// reusable headlessly, matching every other lib/ module's own convention.
//
// The engine is split into a target list (`buildAccessibilityScanTargets`,
// a cheap document walk) and a per-target scan (`scanAccessibilityTarget`)
// so `useAccessibilityCheck.ts` can chunk the scan across `scheduleIdleWork`
// slots (AC "runs on a 200-page document... without blocking the UI") while
// `checkAccessibility` below runs the whole thing synchronously in one call
// for callers (tests, a fix's re-check) that don't need chunking.
import {
  isBreak,
  isChart,
  isConnector,
  isImage,
  isShape,
  isTable,
  documentThemeOf,
  isToc,
  isTof,
  type Chart,
  type DocumentModel,
  type Image,
  type Paragraph,
  type Run,
  type Shape,
  type Table,
} from "./documentModel";
import { blockText, compareCaret, type CaretPosition, type ParagraphAddress } from "./textEditing";
import {
  NORMAL_STYLE_ID,
  namedStyleIdFor,
  paintedParagraphFormatting,
  resolveNamedStyle,
  type ResolvedNamedStyle,
} from "./styles";
import { LINK_COLOR } from "./renderFormatting";
import { paintedShadingFill } from "./paragraphBorders";
import { HIGHLIGHT_HEX } from "./core/colors";
import { resolveColor } from "./core/theme";
import type { DocumentTheme, ModelColor } from "./core/theme";

// --- issue shape (AC-2, AC-3) ------------------------------------------

export type AccessibilitySeverity = "error" | "warning" | "tip";

export type AccessibilityIssueKind =
  | "missing-alt-text"
  | "table-no-header-row"
  | "table-merged-cells"
  | "heading-level-skipped"
  | "empty-heading"
  | "link-bare-url"
  | "link-click-here"
  | "low-contrast"
  | "colour-alone"
  | "no-heading-structure"
  | "visual-heading";

/** The three mechanical fixes SPEC-043 offers (AC-27); every other kind has
 * no fix action at all — the pane simply omits the button. */
export type AccessibilityFixKind = "add-alt-text" | "mark-header-row" | "fix-heading-level";

/** Every detection's severity (the "Severity" reference table): assigned by
 * the principle — Error when a screen-reader user cannot reach or is given
 * the wrong content, Warning when it is reachable but misleading, Tip when
 * it is correct but better structure is available — never by which row a
 * new detection superficially resembles. */
export const SEVERITY_BY_KIND: Record<AccessibilityIssueKind, AccessibilitySeverity> = {
  "missing-alt-text": "error",
  "table-no-header-row": "error",
  "empty-heading": "error",
  "low-contrast": "error",
  "heading-level-skipped": "warning",
  "link-bare-url": "warning",
  "link-click-here": "warning",
  "colour-alone": "warning",
  "table-merged-cells": "warning",
  "no-heading-structure": "tip",
  "visual-heading": "tip",
};

export interface AccessibilityIssue {
  /** Stable within one scan pass — a fresh scan (including a fix's re-run,
   * AC-28) produces a fresh array, so this only needs to be a good `:key`. */
  id: string;
  kind: AccessibilityIssueKind;
  severity: AccessibilitySeverity;
  /** One sentence explaining why it matters (AC-26). */
  explanation: string;
  /** Where selecting the issue scrolls to and highlights (AC-25): a caret
   * range from `at` to `endAt`, extended as a selection rather than just a
   * collapsed caret so the offending run/paragraph/table/image is visibly
   * highlighted, not merely pointed at. */
  at: CaretPosition;
  endAt: CaretPosition;
  fix?: AccessibilityFixKind;
  /** `fix-heading-level`'s target level — the next expected level after the
   * heading before this one in document order (H1 -> H3 fixes to H2). */
  fixHeadingLevel?: number;
}

// --- colour and contrast (AC-13..19) ------------------------------------

function hexToRgb(hex: string): { r: number; g: number; b: number } {
  const clean = hex.replace("#", "");
  return {
    r: Number.parseInt(clean.slice(0, 2), 16) || 0,
    g: Number.parseInt(clean.slice(2, 4), 16) || 0,
    b: Number.parseInt(clean.slice(4, 6), 16) || 0,
  };
}

function srgbChannelToLinear(channel: number): number {
  const c = channel / 255;
  return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
}

function relativeLuminance(hex: string): number {
  const { r, g, b } = hexToRgb(hex);
  return 0.2126 * srgbChannelToLinear(r) + 0.7152 * srgbChannelToLinear(g) + 0.0722 * srgbChannelToLinear(b);
}

/** The WCAG contrast ratio between two colours (AC-16), rounded to two
 * decimals before any comparison — so a value that computes as 4.4999999 is
 * not an Error while its neighbour at 4.5000001 passes. */
export function contrastRatio(hexA: string, hexB: string): number {
  const a = relativeLuminance(hexA);
  const b = relativeLuminance(hexB);
  const lighter = Math.max(a, b);
  const darker = Math.min(a, b);
  return Math.round(((lighter + 0.05) / (darker + 0.05)) * 100) / 100;
}

/** "Large text" (AC-15): an effective size of 18pt+, or 14pt+ when the run
 * is bold — bold is the run's own `formatting.bold`, never inherited, the
 * same convention "looks like a heading" (AC-22) uses. */
export function isLargeText(effectiveFontSizePt: number, bold: boolean): boolean {
  return effectiveFontSizePt >= 18 || (bold && effectiveFontSizePt >= 14);
}

/** Foreground resolution (AC-17): the run's own colour, else the resolved
 * named style's, else `#000000` — `DefaultStyles` has no colour field of
 * its own to fall back to, so that step of the documented chain is vacuous
 * on the current model. */
function effectiveForegroundColor(resolvedStyle: ResolvedNamedStyle, run: Run, theme: DocumentTheme): string {
  // Contrast is computed on the RESOLVED colour (SPEC-077 AC-47): a theme
  // change can turn a passing document into a failing one, and the checker
  // must see the very colour the renderer paints.
  return resolveColor(run.formatting.color ?? resolvedStyle.runFormatting.color ?? "#000000", theme);
}

/** Background resolution (AC-18, AC-19; SPEC-074 AC-27): the run's highlight
 * (mapped through the renderer's own `HIGHLIGHT_HEX` palette so the checker
 * never disagrees with what was painted) — an explicit `false` ("No colour")
 * falls through exactly like unset — then the run's own shading, then the
 * *paragraph's* shading, then the enclosing cell's, then the page.
 *
 * The paragraph step is SPEC-074's: text on a shaded paragraph is checked
 * against the shading, not against the page, because the shading is what it
 * is actually read on. */
function effectiveBackgroundColor(
  run: Run,
  paragraphShadingFill: ModelColor | undefined,
  cellShadingFill: ModelColor | undefined,
  theme: DocumentTheme,
): string {
  if (run.formatting.highlight) {
    const highlight = resolveColor(run.formatting.highlight, theme);
    return HIGHLIGHT_HEX[highlight] ?? highlight;
  }
  if (run.formatting.shadingFill !== undefined) {
    return resolveColor(run.formatting.shadingFill, theme);
  }
  if (paragraphShadingFill !== undefined) {
    return resolveColor(paragraphShadingFill, theme);
  }
  if (cellShadingFill !== undefined) {
    return resolveColor(cellShadingFill, theme);
  }
  return "#ffffff";
}

function effectiveFontSizePt(model: DocumentModel, resolvedStyle: ResolvedNamedStyle, run: Run): number {
  return run.formatting.fontSizePt ?? resolvedStyle.runFormatting.fontSizePt ?? model.defaultStyles.fontSizePt;
}

/** The colour a sighted user actually sees (used only by the "colour alone"
 * check, AC-23/AC-24) — unlike `effectiveForegroundColor`'s contrast chain,
 * this also falls back to the renderer's link-blue default (SPEC-025) so an
 * unstyled hyperlink is compared against the surrounding text's colour
 * rather than invisibly matching it. */
function paintedColor(resolvedStyle: ResolvedNamedStyle, run: Run, theme: DocumentTheme): string {
  if (run.formatting.color) {
    return resolveColor(run.formatting.color, theme);
  }
  if (run.formatting.linkHref !== undefined || run.formatting.linkAnchor !== undefined) {
    return LINK_COLOR;
  }
  return resolveColor(resolvedStyle.runFormatting.color ?? "#000000", theme);
}

// --- addressing helpers ---------------------------------------------------

type BlockAddress = { section: number; block: number };

function paragraphCaret(address: ParagraphAddress, offset: number): CaretPosition {
  return { section: address.section, block: address.block, offset, cell: address.cell };
}

function atomCaret(address: BlockAddress, offset: 0 | 1): CaretPosition {
  return { section: address.section, block: address.block, offset };
}

function addressKey(address: ParagraphAddress): string {
  return address.cell ? `${address.section}.${address.block}.${address.cell.row}-${address.cell.col}-${address.cell.block}` : `${address.section}.${address.block}`;
}

// --- links (AC-9, AC-10) --------------------------------------------------

interface ParagraphLink {
  text: string;
  startOffset: number;
}

function linkKeyOf(run: Run): string | undefined {
  if (run.formatting.linkHref !== undefined) {
    return `href:${run.formatting.linkHref}`;
  }
  if (run.formatting.linkAnchor !== undefined) {
    return `anchor:${run.formatting.linkAnchor}`;
  }
  return undefined;
}

/** Every hyperlink in a paragraph, as one entry per contiguous run of runs
 * sharing the same link target — a link split across several runs (e.g. a
 * bold word mid-link) is one link, matching how a user perceives it. */
function paragraphLinks(paragraph: Paragraph): ParagraphLink[] {
  const links: ParagraphLink[] = [];
  let offset = 0;
  let current: { text: string; startOffset: number; key: string } | undefined;
  for (const run of paragraph.runs) {
    const key = linkKeyOf(run);
    if (key !== undefined) {
      if (current && current.key === key) {
        current.text += run.text;
      } else {
        if (current) {
          links.push({ text: current.text, startOffset: current.startOffset });
        }
        current = { text: run.text, startOffset: offset, key };
      }
    } else if (current) {
      links.push({ text: current.text, startOffset: current.startOffset });
      current = undefined;
    }
    offset += run.text.length;
  }
  if (current) {
    links.push({ text: current.text, startOffset: current.startOffset });
  }
  return links;
}

/** A bare URL (AC-9): starts with a scheme or "www." and has no whitespace —
 * a link whose visible text is the address itself rather than a description. */
const BARE_URL_RE = /^(?:[a-z][a-z0-9+.-]*:\/\/|www\.)\S+$/i;

function isClickHereText(text: string): boolean {
  return text.trim().replace(/[.!?]+$/, "").trim().toLowerCase() === "click here";
}

// --- "colour alone" (AC-23, AC-24) ----------------------------------------

interface RunColourInfo {
  run: Run;
  startOffset: number;
  length: number;
  color: string;
  fontSizePt: number;
  styleId: string | undefined;
}

function paragraphRunColourInfos(
  model: DocumentModel,
  paragraph: Paragraph,
  resolvedStyle: ResolvedNamedStyle,
): RunColourInfo[] {
  const infos: RunColourInfo[] = [];
  let offset = 0;
  for (const run of paragraph.runs) {
    if (run.text.length > 0) {
      infos.push({
        run,
        startOffset: offset,
        length: run.text.length,
        color: paintedColor(resolvedStyle, run, documentThemeOf(model)),
        fontSizePt: effectiveFontSizePt(model, resolvedStyle, run),
        styleId: run.styleId,
      });
    }
    offset += run.text.length;
  }
  return infos;
}

/** The value covering the most characters — "the paragraph's dominant
 * colour" (AC-23) and, by the same logic, its dominant size/style, the
 * baseline a run is compared against for "no other distinguishing
 * formatting". Ties keep whichever value was seen first. */
function dominantValue<T>(infos: RunColourInfo[], key: (info: RunColourInfo) => T): T | undefined {
  const coverage = new Map<T, number>();
  for (const info of infos) {
    const value = key(info);
    coverage.set(value, (coverage.get(value) ?? 0) + info.length);
  }
  let best: T | undefined;
  let bestCoverage = -1;
  for (const [value, count] of coverage) {
    if (count > bestCoverage) {
      bestCoverage = count;
      best = value;
    }
  }
  return best;
}

// --- issue construction -----------------------------------------------

function makeIssue(partial: Omit<AccessibilityIssue, "severity">): AccessibilityIssue {
  return { ...partial, severity: SEVERITY_BY_KIND[partial.kind] };
}

// --- per-target detections -----------------------------------------------

function checkImageAltText(image: Image | Shape | Chart, address: BlockAddress, issues: AccessibilityIssue[]): void {
  if (image.altText !== undefined && image.altText.trim().length > 0) {
    return;
  }
  issues.push(
    makeIssue({
      id: `missing-alt-text:${address.section}.${address.block}`,
      kind: "missing-alt-text",
      explanation: "A screen reader has nothing to announce for this image without alternative text.",
      at: atomCaret(address, 0),
      endAt: atomCaret(address, 1),
      fix: "add-alt-text",
    }),
  );
}

function checkTable(table: Table, address: BlockAddress, issues: AccessibilityIssue[]): void {
  if (!table.rows.some((row) => row.isHeader)) {
    issues.push(
      makeIssue({
        id: `table-no-header-row:${address.section}.${address.block}`,
        kind: "table-no-header-row",
        explanation: "Without a header row a screen reader cannot tell a user which column or row a cell belongs to.",
        at: atomCaret(address, 0),
        endAt: atomCaret(address, 1),
        fix: "mark-header-row",
      }),
    );
  }
  const hasMergedCells = table.rows.some((row) =>
    row.cells.some((cell) => (cell.colSpan ?? 1) > 1 || (cell.rowSpan ?? 1) > 1),
  );
  if (hasMergedCells) {
    issues.push(
      makeIssue({
        id: `table-merged-cells:${address.section}.${address.block}`,
        kind: "table-merged-cells",
        explanation: "Merged cells can confuse the row/column relationships a screen reader announces while reading the table.",
        at: atomCaret(address, 0),
        endAt: atomCaret(address, 1),
      }),
    );
  }
}

function checkParagraphLinks(paragraph: Paragraph, address: ParagraphAddress, issues: AccessibilityIssue[]): void {
  for (const link of paragraphLinks(paragraph)) {
    const text = link.text.trim();
    if (text.length === 0) {
      continue;
    }
    const at = paragraphCaret(address, link.startOffset);
    const endAt = paragraphCaret(address, link.startOffset + link.text.length);
    if (isClickHereText(text)) {
      issues.push(
        makeIssue({
          id: `link-click-here:${addressKey(address)}:${link.startOffset}`,
          kind: "link-click-here",
          explanation: '"Click here" tells a screen-reader user nothing about where the link goes when read out of context.',
          at,
          endAt,
        }),
      );
    } else if (BARE_URL_RE.test(text)) {
      issues.push(
        makeIssue({
          id: `link-bare-url:${addressKey(address)}:${link.startOffset}`,
          kind: "link-bare-url",
          explanation: "A raw web address is read out character by character, which is tedious and uninformative for a screen-reader user.",
          at,
          endAt,
        }),
      );
    }
  }
}

function checkParagraphContrastAndColour(
  model: DocumentModel,
  paragraph: Paragraph,
  address: ParagraphAddress,
  cellShadingFill: ModelColor | undefined,
  issues: AccessibilityIssue[],
): void {
  const resolvedStyle = resolveNamedStyle(model, paragraph.styleId);
  const theme = documentThemeOf(model);
  // SPEC-074 AC-27: the paragraph's own shading, else the one its style sets
  // (AC-17's inheritance) — the surface this paragraph's text is read on.
  // Resolved through `paintedParagraphFormatting` rather than `resolvedStyle`
  // so the checker reads the very fill the renderer painted, including for a
  // paragraph that names no style and inherits Normal's. A paragraph that
  // explicitly overrides its style's fill with "no fill" is read on the page
  // again, so the fill goes through paintedShadingFill.
  const paragraphShadingFill = paintedShadingFill(paintedParagraphFormatting(model, paragraph));

  // Contrast (AC-11, AC-13..19): every run with visible text against its
  // own resolved background.
  let offset = 0;
  for (const run of paragraph.runs) {
    if (run.text.trim().length > 0) {
      const foreground = effectiveForegroundColor(resolvedStyle, run, theme);
      const background = effectiveBackgroundColor(run, paragraphShadingFill, cellShadingFill, theme);
      const ratio = contrastRatio(foreground, background);
      const fontSizePt = effectiveFontSizePt(model, resolvedStyle, run);
      const large = isLargeText(fontSizePt, run.formatting.bold === true);
      const threshold = large ? 3 : 4.5;
      if (ratio < threshold) {
        issues.push(
          makeIssue({
            id: `low-contrast:${addressKey(address)}:${offset}`,
            kind: "low-contrast",
            explanation: `This text's contrast ratio (${ratio.toFixed(2)}:1) is below the ${threshold}:1 WCAG AA threshold, making it hard to read for low-vision users.`,
            at: paragraphCaret(address, offset),
            endAt: paragraphCaret(address, offset + run.text.length),
          }),
        );
      }
    }
    offset += run.text.length;
  }

  // Colour alone (AC-12, AC-23, AC-24).
  const infos = paragraphRunColourInfos(model, paragraph, resolvedStyle);
  if (infos.length < 2) {
    return;
  }
  const dominantColor = dominantValue(infos, (info) => info.color);
  const dominantSize = dominantValue(infos, (info) => info.fontSizePt);
  const dominantStyleId = dominantValue(infos, (info) => info.styleId);
  for (const info of infos) {
    if (info.color === dominantColor) {
      continue;
    }
    const distinguished =
      info.run.formatting.bold === true ||
      info.run.formatting.italic === true ||
      info.run.formatting.underline === true ||
      info.fontSizePt !== dominantSize ||
      info.styleId !== dominantStyleId;
    if (distinguished) {
      continue;
    }
    issues.push(
      makeIssue({
        id: `colour-alone:${addressKey(address)}:${info.startOffset}`,
        kind: "colour-alone",
        explanation: "Colour is the only thing distinguishing this text, which conveys nothing to a reader who cannot perceive it.",
        at: paragraphCaret(address, info.startOffset),
        endAt: paragraphCaret(address, info.startOffset + info.length),
      }),
    );
  }
}

/** "Resolves to Normal" (AC-22's condition 1): the paragraph's own style, or
 * an unstyled paragraph, which both read as "Normal" (see styles.ts's own
 * `styleDisplayName` for the same undefined-means-Normal convention). */
function resolvesToNormalStyle(model: DocumentModel, paragraph: Paragraph): boolean {
  const namedId = namedStyleIdFor(model, paragraph.styleId);
  return namedId === undefined || namedId === NORMAL_STYLE_ID;
}

function checkVisualHeading(
  model: DocumentModel,
  paragraph: Paragraph,
  address: ParagraphAddress,
  nextParagraph: Paragraph | undefined,
  issues: AccessibilityIssue[],
): void {
  if (paragraph.headingLevel !== undefined) {
    return;
  }
  const text = blockText(paragraph);
  if (text.trim().length === 0 || text.length > 100) {
    return;
  }
  if (!resolvesToNormalStyle(model, paragraph)) {
    return;
  }
  const textRuns = paragraph.runs.filter((run) => run.text.length > 0);
  if (textRuns.length === 0) {
    return;
  }
  const everyRunBold = textRuns.every((run) => run.formatting.bold === true);
  const threshold = model.defaultStyles.fontSizePt * 1.2;
  const resolvedStyle = resolveNamedStyle(model, paragraph.styleId);
  const everyRunLarge = textRuns.every((run) => effectiveFontSizePt(model, resolvedStyle, run) >= threshold);
  if (!everyRunBold && !everyRunLarge) {
    return;
  }
  if (!nextParagraph || blockText(nextParagraph).trim().length === 0) {
    return;
  }
  issues.push(
    makeIssue({
      id: `visual-heading:${addressKey(address)}`,
      kind: "visual-heading",
      explanation: "This paragraph is styled to look like a heading but is not one, so it will not appear in navigation by heading.",
      at: paragraphCaret(address, 0),
      endAt: paragraphCaret(address, text.length),
    }),
  );
}

// --- scan targets and state (chunking support) ----------------------------

export type AccessibilityScanTarget =
  | { kind: "paragraph"; address: ParagraphAddress; paragraph: Paragraph; topLevelNext?: Paragraph; cellShadingFill?: ModelColor }
  | { kind: "table"; address: BlockAddress; table: Table }
  | { kind: "image"; address: BlockAddress; image: Image }
  | { kind: "shape"; address: BlockAddress; shape: Shape }
  | { kind: "chart"; address: BlockAddress; chart: Chart };

export interface AccessibilityScanState {
  previousHeadingLevel: number | null;
  anyHeadingFound: boolean;
  anyContent: boolean;
}

export function createAccessibilityScanState(): AccessibilityScanState {
  return { previousHeadingLevel: null, anyHeadingFound: false, anyContent: false };
}

/** Builds the flat, document-order list of things to check — a cheap walk
 * (no formatting resolution) so it can run up front, letting the actual
 * scan be chunked across it one target at a time (AC "chunked through
 * scheduleIdleWork"). */
export function buildAccessibilityScanTargets(model: DocumentModel): AccessibilityScanTarget[] {
  const targets: AccessibilityScanTarget[] = [];
  model.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (isTable(block)) {
        targets.push({ kind: "table", address: { section: sectionIndex, block: blockIndex }, table: block });
        block.rows.forEach((row, rowIndex) => {
          row.cells.forEach((cell, colIndex) => {
            cell.blocks.forEach((paragraph, cellBlockIndex) => {
              targets.push({
                kind: "paragraph",
                address: {
                  section: sectionIndex,
                  block: blockIndex,
                  cell: { row: rowIndex, col: colIndex, block: cellBlockIndex },
                },
                paragraph,
                cellShadingFill: cell.shadingFill,
              });
            });
          });
        });
      } else if (isImage(block)) {
        targets.push({ kind: "image", address: { section: sectionIndex, block: blockIndex }, image: block });
      } else if (isShape(block)) {
        targets.push({ kind: "shape", address: { section: sectionIndex, block: blockIndex }, shape: block });
      } else if (isChart(block)) {
        targets.push({ kind: "chart", address: { section: sectionIndex, block: blockIndex }, chart: block });
      } else if (isConnector(block) || isBreak(block) || isToc(block) || isTof(block)) {
        // Nothing to check (a connector carries alt text but is not part of
        // this spec's image/shape alt-text audit).
      } else {
        const next = section.blocks[blockIndex + 1];
        targets.push({
          kind: "paragraph",
          address: { section: sectionIndex, block: blockIndex },
          paragraph: block,
          topLevelNext: next && !isTable(next) && !isImage(next) && !isShape(next) && !isChart(next) && !isConnector(next) && !isBreak(next) && !isToc(next) && !isTof(next) ? next : undefined,
        });
      }
    });
  });
  return targets;
}

/** Scans one target, appending any issues it produces — the unit the
 * idle-scheduled composable advances one (or a budgeted few) of per slot. */
export function scanAccessibilityTarget(
  model: DocumentModel,
  target: AccessibilityScanTarget,
  state: AccessibilityScanState,
  issues: AccessibilityIssue[],
): void {
  if (target.kind === "image") {
    state.anyContent = true;
    checkImageAltText(target.image, target.address, issues);
    return;
  }
  if (target.kind === "shape") {
    state.anyContent = true;
    // Same "missing-alt-text" detection as an image (AC-48 — no change to
    // SPEC-043's check itself, Shape carries altText in the same field).
    checkImageAltText(target.shape, target.address, issues);
    return;
  }
  if (target.kind === "chart") {
    state.anyContent = true;
    // Same "missing-alt-text" detection as an image/shape — Chart carries
    // altText in the same DrawingPlacement field (SPEC-052).
    checkImageAltText(target.chart, target.address, issues);
    return;
  }
  if (target.kind === "table") {
    state.anyContent = true;
    checkTable(target.table, target.address, issues);
    return;
  }
  const { paragraph, address } = target;
  const text = blockText(paragraph);
  if (text.trim().length > 0) {
    state.anyContent = true;
  }
  checkParagraphLinks(paragraph, address, issues);
  checkParagraphContrastAndColour(model, paragraph, address, target.cellShadingFill, issues);

  // Heading structure (AC-7, AC-8, AC-20, AC-21) applies to top-level flow
  // only — table-cell headings do not participate in outline navigation.
  if (address.cell) {
    return;
  }
  if (paragraph.headingLevel !== undefined) {
    state.anyHeadingFound = true;
    if (blockText(paragraph).trim().length === 0) {
      issues.push(
        makeIssue({
          id: `empty-heading:${addressKey(address)}`,
          kind: "empty-heading",
          explanation: "An empty heading announces a navigation stop with no content, wasting a screen-reader user's time.",
          at: paragraphCaret(address, 0),
          endAt: paragraphCaret(address, 0),
        }),
      );
    } else if (state.previousHeadingLevel !== null && paragraph.headingLevel > state.previousHeadingLevel + 1) {
      const expectedLevel = state.previousHeadingLevel + 1;
      issues.push(
        makeIssue({
          id: `heading-level-skipped:${addressKey(address)}`,
          kind: "heading-level-skipped",
          explanation: `Jumping from heading level ${state.previousHeadingLevel} to ${paragraph.headingLevel} skips a level, which disorients a screen-reader user navigating by heading.`,
          at: paragraphCaret(address, 0),
          endAt: paragraphCaret(address, text.length),
          fix: "fix-heading-level",
          fixHeadingLevel: expectedLevel,
        }),
      );
    }
    state.previousHeadingLevel = paragraph.headingLevel;
  } else {
    checkVisualHeading(model, paragraph, address, target.topLevelNext, issues);
  }
}

/** Appends the one document-level detection (AC-20's "no heading structure
 * at all", a Tip) once the whole document has been scanned. Guarded on the
 * document having any real content so a fresh blank document — one empty
 * paragraph, nothing to structure — does not open with a Tip already
 * showing. */
export function finalizeAccessibilityScan(state: AccessibilityScanState, issues: AccessibilityIssue[]): void {
  if (state.anyHeadingFound || !state.anyContent) {
    return;
  }
  issues.push(
    makeIssue({
      id: "no-heading-structure",
      kind: "no-heading-structure",
      explanation: "This document has no headings at all, so a screen-reader user cannot jump between its sections.",
      at: { section: 0, block: 0, offset: 0 },
      endAt: { section: 0, block: 0, offset: 0 },
    }),
  );
}

/** Document order: section, then block, then cell (row/col/block), then
 * offset — what the results pane lists within each severity group. Exported
 * so the idle-chunked composable can sort its own accumulated array the
 * same way `checkAccessibility` sorts its synchronous result. */
export function sortAccessibilityIssues(issues: AccessibilityIssue[]): AccessibilityIssue[] {
  return issues.sort((a, b) => compareCaret(a.at, b.at));
}

/** The whole scan, synchronously — used directly by tests and by a fix's
 * re-check (AC-28), where the document is already known to be small enough
 * (one fixed issue) that chunking would add nothing. The idle-scheduled
 * composable instead drives `buildAccessibilityScanTargets`/
 * `scanAccessibilityTarget`/`finalizeAccessibilityScan` itself across
 * budgeted slots for the initial whole-document pass. */
export function checkAccessibility(model: DocumentModel): AccessibilityIssue[] {
  const targets = buildAccessibilityScanTargets(model);
  const state = createAccessibilityScanState();
  const issues: AccessibilityIssue[] = [];
  for (const target of targets) {
    scanAccessibilityTarget(model, target, state, issues);
  }
  finalizeAccessibilityScan(state, issues);
  return sortAccessibilityIssues(issues);
}
