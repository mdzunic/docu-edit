// Maps the document model's formatting onto inline CSS for the canvas.
// Kept framework-free (no Vue imports) so it's testable in isolation, and
// split out of DocumentView.vue because that file's <script setup> is
// linted by ESLint's plain espree parser, which rejects TypeScript type
// annotations (see AppShell.vue's comment on the same constraint).
import {
  HORIZONTAL_RULE_STYLE_ID,
  type Chart,
  type Image,
  type Paragraph,
  type ParagraphFormatting,
  type Run,
  type RunFormatting,
  type Shape,
  type Table,
  type TableCell,
  type TableRow,
  type TrackChangesDisplayMode,
} from "./documentModel";
import { type ListDefinition } from "./core/lists";
import { type BorderSide } from "./core/tables";
import { listIndentCm } from "./core/lists";
import { definedTextIndentCm } from "./core/listDefinitions";
import { authorColor } from "./trackChangesIdentity";
// The one highlight palette (OOXML name -> hex), shared with both file
// formats' readers/writers (SPEC-013) — the model stores the OOXML colour
// name as-is.
import { HIGHLIGHT_HEX } from "./core/colors";
// Metric-compatible substitution (SPEC-048): the model keeps the original
// family name; here it becomes the CSS stack that renders the bundled twin
// when the original is unavailable. Measurement (textMetrics) routes the same
// family through the same helper, so measure and paint agree.
import { fontStack } from "./fonts/substitution";

/** The HTML tag a paragraph should render as: h1-h6 for a resolved heading
 * level (clamped to h6), otherwise a plain paragraph. */
export function headingTag(level: number | undefined): string {
  if (!level) {
    return "p";
  }
  return `h${Math.min(Math.max(level, 1), 6)}`;
}

/** The CSS class a paragraph's render entry carries — the shared
 * "doc-paragraph" plus SPEC-050's horizontal-rule modifier. */
export function paragraphClassName(paragraph: Paragraph): string {
  return paragraph.styleId === HORIZONTAL_RULE_STYLE_ID ? "doc-paragraph doc-horizontal-rule" : "doc-paragraph";
}

/** The hyperlink colour (SPEC-025 AC-4) — the blue Word and Docs both use
 * for the built-in Hyperlink character style. */
export const LINK_COLOR = "#1a73e8";

/** The href a rendered link run carries: the external URL as-is, or
 * `#anchor` for an internal bookmark/heading target. Undefined for plain
 * text — the template picks `<a>` vs `<span>` on this. */
export function runLinkHref(formatting: RunFormatting): string | undefined {
  if (formatting.linkHref !== undefined) {
    return formatting.linkHref;
  }
  if (formatting.linkAnchor !== undefined) {
    return `#${formatting.linkAnchor}`;
  }
  return undefined;
}

/** Whether the run links outside the document — only those open in a new
 * tab and carry `target`/`rel` (AC-12, AC-13); internal links jump. */
export function isExternalLink(formatting: RunFormatting): boolean {
  return formatting.linkHref !== undefined;
}

export function runStyle(formatting: RunFormatting): Record<string, string> {
  const style: Record<string, string> = {};
  const linked = formatting.linkHref !== undefined || formatting.linkAnchor !== undefined;
  if (formatting.bold) {
    style.fontWeight = "bold";
  }
  if (formatting.italic) {
    style.fontStyle = "italic";
  }
  const decorations: string[] = [];
  // A hyperlink is always underlined in the link colour (SPEC-025 AC-4); an
  // explicit run colour still wins over the link blue, like Word's direct
  // formatting over the Hyperlink style.
  if (formatting.underline || linked) {
    decorations.push("underline");
  }
  if (formatting.strikethrough) {
    decorations.push("line-through");
  }
  if (decorations.length > 0) {
    style.textDecorationLine = decorations.join(" ");
  }
  if (formatting.fontFamily) {
    // Substitution is a rendering decision only (AC-5): the run keeps its
    // original family name in the model; the CSS gets the metric twin stack.
    style.fontFamily = fontStack(formatting.fontFamily);
  }
  if (formatting.fontSizePt) {
    style.fontSize = `${formatting.fontSizePt}pt`;
  }
  if (formatting.color) {
    style.color = formatting.color;
  } else if (linked) {
    style.color = LINK_COLOR;
  }
  // highlight false is an explicit "No colour": like undefined it paints
  // nothing, unlike undefined it survives saving (SPEC-013).
  if (formatting.highlight) {
    style.backgroundColor = HIGHLIGHT_HEX[formatting.highlight] ?? formatting.highlight;
  }
  if (formatting.vertAlign) {
    // Word renders sub/superscript at roughly two-thirds size on a shifted
    // baseline; em-based so it composes with the run's own font size.
    style.verticalAlign = formatting.vertAlign === "subscript" ? "sub" : "super";
    style.fontSize = formatting.fontSizePt
      ? `${formatting.fontSizePt * 0.65}pt`
      : "0.65em";
  }
  return style;
}

/** A paragraph's formatting as this display mode should render it (SPEC-039
 * AC-9): "Original" shows the pre-change formatting for a pending paragraph
 * formatChange (as if it were rejected), every other mode shows the current
 * (possibly already-changed) formatting, same as `paragraph.formatting`. */
export function displayParagraphFormatting(
  paragraph: Paragraph,
  displayMode: TrackChangesDisplayMode,
): ParagraphFormatting | undefined {
  const formatChange = paragraph.changes?.formatChange;
  if (displayMode === "original" && formatChange) {
    return formatChange.previous;
  }
  return paragraph.formatting;
}

export function paragraphStyle(
  paragraph: Paragraph,
  listDefinitions: readonly ListDefinition[] = [],
  displayMode: TrackChangesDisplayMode = "all",
): Record<string, string> {
  const formatting = displayParagraphFormatting(paragraph, displayMode);
  const style: Record<string, string> = { whiteSpace: "pre-wrap" };
  // A list item's text sits at the level's indent (SPEC-018 AC-5's per-level
  // indents); the marker itself hangs before it via the canvas's ::before
  // rule. A direct start indent on the paragraph wins, matching how both
  // file formats let direct indentation override the list definition's.
  // A defined item (SPEC-032) takes its level's text position from its
  // definition instead.
  if (paragraph.list && formatting?.indentStartCm === undefined) {
    const def =
      paragraph.list.defId !== undefined
        ? listDefinitions.find((entry) => entry.defId === paragraph.list?.defId)
        : undefined;
    style.marginInlineStart = def
      ? `${definedTextIndentCm(def, paragraph.list.level)}cm`
      : `${listIndentCm(paragraph.list)}cm`;
  }
  if (!formatting) {
    return style;
  }
  if (formatting.alignment) {
    style.textAlign = formatting.alignment;
  }
  // Logical start/end indents map to margin-inline-start/end, and the
  // paragraph's own `direction` decides which physical edge each paints on —
  // the browser mirrors RTL paragraphs for free (SPEC-014). text-indent is
  // direction-relative in CSS already, matching the model's start-relative
  // first-line indent.
  if (formatting.direction === "rtl") {
    style.direction = "rtl";
  }
  if (formatting.indentStartCm) {
    style.marginInlineStart = `${formatting.indentStartCm}cm`;
  }
  if (formatting.indentEndCm) {
    style.marginInlineEnd = `${formatting.indentEndCm}cm`;
  }
  if (formatting.indentFirstLineCm) {
    style.textIndent = `${formatting.indentFirstLineCm}cm`;
  }
  if (formatting.spaceBeforePt !== undefined) {
    style.marginTop = `${formatting.spaceBeforePt}pt`;
  }
  if (formatting.spaceAfterPt !== undefined) {
    style.marginBottom = `${formatting.spaceAfterPt}pt`;
  }
  if (formatting.lineSpacingPt !== undefined) {
    style.lineHeight = `${formatting.lineSpacingPt}pt`;
  } else if (formatting.lineSpacing !== undefined) {
    style.lineHeight = String(formatting.lineSpacing);
  }
  return style;
}

export function runKey(run: Run, index: number): string {
  return `${index}-${run.text.length}`;
}

// --- track changes (SPEC-039) ------------------------------------------

/** A run's tracked-change markup, layered on top of `runStyle` (SPEC-039
 * AC "inserted text renders underlined in the author's colour and deleted
 * text renders struck through"): "All markup" only, since "Simple markup"
 * shows plain text with the margin change bar doing the signalling, and
 * "No markup"/"Original" render as if one side of the change already won
 * (see `trackChangeRunHidden`, which decides whether this run's text shows
 * at all in those two modes). */
export function trackChangeRunStyle(run: Run, displayMode: TrackChangesDisplayMode): Record<string, string> {
  const changes = run.changes;
  if (!changes || displayMode !== "all") {
    return {};
  }
  const style: Record<string, string> = {};
  const decorations: string[] = [];
  if (changes.insert) {
    decorations.push("underline");
    style.color = authorColor(changes.insert.author);
  }
  if (changes.delete) {
    decorations.push("line-through");
    style.color = authorColor(changes.delete.author);
  }
  if (decorations.length > 0) {
    style.textDecorationLine = decorations.join(" ");
    style.textDecorationColor = style.color;
  }
  return style;
}

/** Whether a run's text should render as invisible (zero rendered
 * characters) under this display mode (SPEC-039 AC-9): "No markup" reads as
 * if every change were already accepted, so a pending deletion's text is
 * gone; "Original" reads as if every change were rejected, so a pending
 * insertion's text never happened. The run itself always stays in the DOM
 * (see DocBlockEntries.vue's `caretAtom` branch) — only its rendered text
 * collapses to "", with the full text kept as the caret-atom's model text
 * (lib/domCaret.ts) so caret/offset mapping stays exactly as if the text
 * were showing. */
export function trackChangeRunHidden(run: Run, displayMode: TrackChangesDisplayMode): boolean {
  const changes = run.changes;
  if (!changes) {
    return false;
  }
  if (displayMode === "noMarkup" && changes.delete) {
    return true;
  }
  if (displayMode === "original" && changes.insert) {
    return true;
  }
  return false;
}

/** A run's formatting as this display mode should render it (SPEC-039
 * AC-9): "Original" shows the pre-change formatting for a pending run
 * formatChange (as if it were rejected), every other mode shows the
 * current formatting. */
export function displayRunFormatting(run: Run, displayMode: TrackChangesDisplayMode): RunFormatting {
  const formatChange = run.changes?.formatChange;
  if (displayMode === "original" && formatChange) {
    return formatChange.previous;
  }
  return run.formatting;
}

/** Whether a paragraph carries any tracked change at all — text, its own
 * mark, or block-level (SPEC-039 AC-5's margin change bar, shown in every
 * display mode except "No markup"/"Original", which render as if the
 * document already reflects one side of every change). */
export function paragraphHasTrackedChange(paragraph: Paragraph): boolean {
  if (paragraph.changes?.markInsert || paragraph.changes?.markDelete || paragraph.changes?.formatChange) {
    return true;
  }
  return paragraph.runs.some((run) => run.changes?.insert || run.changes?.delete || run.changes?.formatChange);
}

/** Whether a whole tracked-inserted/deleted block (a table row or an image
 * — see `BlockTrackedChanges`) should be hidden entirely under this display
 * mode (SPEC-039 AC-9): "No markup" drops a pending deletion (as if
 * accepted), "Original" drops a pending insertion (as if rejected). Unlike
 * run-level hiding, a whole row/image is its own addressable block, so it
 * is safe to omit outright rather than collapse to a caret atom. */
export function trackChangeBlockHidden(
  changes: { insert?: unknown; delete?: unknown } | undefined,
  displayMode: TrackChangesDisplayMode,
): boolean {
  if (!changes) {
    return false;
  }
  if (displayMode === "noMarkup" && changes.delete) {
    return true;
  }
  if (displayMode === "original" && changes.insert) {
    return true;
  }
  return false;
}

// --- tables (SPEC-023) -------------------------------------------------

function borderSideCss(side: BorderSide): string {
  if (side.style === "none") {
    return "none";
  }
  const cssStyle = side.style === "single" ? "solid" : side.style;
  const color = side.color !== "auto" ? side.color : "#000000";
  return `${side.widthPt}pt ${cssStyle} ${color}`;
}

/** A cell's own border/shading/padding/vertical-alignment as inline CSS
 * (AC-10..17): only the properties the cell actually sets — an unset side
 * falls through to the canvas's default hairline gridline (see
 * DocTableBlock.vue's `.doc-table-cell` rule). */
export function tableCellStyle(cell: TableCell): Record<string, string> {
  const style: Record<string, string> = {};
  const borders = cell.borders;
  if (borders?.top) {
    style.borderTop = borderSideCss(borders.top);
  }
  if (borders?.bottom) {
    style.borderBottom = borderSideCss(borders.bottom);
  }
  if (borders?.left) {
    style.borderLeft = borderSideCss(borders.left);
  }
  if (borders?.right) {
    style.borderRight = borderSideCss(borders.right);
  }
  if (cell.shadingFill !== undefined) {
    style.backgroundColor = cell.shadingFill;
  }
  if (cell.paddingCm !== undefined) {
    style.padding = `${cell.paddingCm}cm`;
  }
  if (cell.verticalAlign !== undefined) {
    // CSS's table-cell vertical-align keyword for the model's "center" is
    // "middle" — "center" itself is not a valid value and is silently
    // dropped, leaving the canvas's default (top).
    style.verticalAlign = cell.verticalAlign === "center" ? "middle" : cell.verticalAlign;
  }
  return style;
}

/** A cell's inline style (AC-15) with any explicit shading dropped while it
 * is part of an active range selection — the canvas's `.cell-selected`
 * class highlight would otherwise lose to the cell's own `style` attribute,
 * which always wins over a class rule regardless of shading. */
export function tableCellDisplayStyle(cell: TableCell, selected: boolean): Record<string, string> {
  const style = tableCellStyle(cell);
  if (selected) {
    delete style.backgroundColor;
  }
  return style;
}

/** A row's minimum height (AC-4, AC-5): `height` on a `<tr>` is a floor in
 * table layout, not a fixed size — the row still grows for its content. */
export function tableRowStyle(row: TableRow): Record<string, string> {
  return row.minHeightCm !== undefined ? { height: `${row.minHeightCm}cm` } : {};
}

/** A table's own width and page alignment (AC-22, AC-23): "auto" hands
 * sizing to the browser's own content-based table layout; "percent" scales
 * the text area; "fixed" (the default) is the column widths' own sum. */
export function tableStyle(table: Table): Record<string, string> {
  const style: Record<string, string> = { tableLayout: table.widthMode === "auto" ? "auto" : "fixed" };
  if (table.widthMode === "percent") {
    style.width = `${table.widthPercent ?? 100}%`;
  } else if (table.widthMode !== "auto") {
    const totalCm = table.columnWidthsCm.reduce((sum, cm) => sum + cm, 0);
    style.width = `${totalCm}cm`;
  }
  if (table.alignment === "center") {
    style.marginLeft = "auto";
    style.marginRight = "auto";
  } else if (table.alignment === "right") {
    style.marginLeft = "auto";
    style.marginRight = "0";
  }
  return style;
}

/** One `<col>`'s width (AC-23): a fixed-width table (the default) uses the
 * column's own absolute width; a percent/auto-width table instead uses its
 * *share* of the total as a relative percentage, so the columns scale down
 * together with the table's own width/auto sizing instead of forcing it
 * wide enough to fit their absolute centimetre widths regardless. */
export function tableColumnWidth(table: Table, index: number): string {
  const widthCm = table.columnWidthsCm[index] ?? 0;
  if (table.widthMode !== "percent" && table.widthMode !== "auto") {
    return `${widthCm}cm`;
  }
  const totalCm = table.columnWidthsCm.reduce((sum, cm) => sum + cm, 0);
  return totalCm > 0 ? `${(widthCm / totalCm) * 100}%` : "auto";
}

/** An image's size and wrap-mode rendering (SPEC-024 AC-13, AC-14):
 * "inline" is a plain block, no different from a paragraph's own box;
 * "square" floats left with text wrapping around it; "behind"/"infront"
 * float out of flow entirely (positioned at their block's own slot via the
 * zero-size anchor DocImageBlock.vue renders around them), stacked behind
 * or in front of the surrounding text by z-index. */
export function imageDisplayStyle(image: Image): Record<string, string> {
  const style: Record<string, string> = {
    width: `${image.widthCm}cm`,
    height: `${image.heightCm}cm`,
  };
  if (image.wrap === "square") {
    style.float = "left";
    style.marginRight = "0.35cm";
  } else if (image.wrap === "behind" || image.wrap === "infront") {
    style.position = "absolute";
    style.left = "0";
    style.top = "0";
    style.zIndex = image.wrap === "behind" ? "0" : "5";
  }
  return style;
}

/** A shape's size and position rendering (SPEC-051): unlike Image, a Shape
 * carries a real `offsetXCm`/`offsetYCm`, so every non-`inline` wrap mode
 * positions absolutely at that offset from its anchor (see the anchor
 * span's own `position: relative`, DocShapeBlock.vue) — `"square"`/
 * `"topAndBottom"` do not reflow surrounding text around the box (the
 * "honest, predictable" simplification SPEC-024's own "square" float
 * already makes for a left-edge case; a shape's offset can be anywhere on
 * the page, where a CSS float cannot follow it). `"behind"` paints under
 * the text layer, the other three float above it. `"inline"` ignores the
 * offset entirely, per "Default position — when offsets are absent". */
export function shapeDisplayStyle(shape: Shape): Record<string, string> {
  const heightCm = Math.max(shape.heightCm, 0.05);
  const style: Record<string, string> = {
    width: `${shape.widthCm}cm`,
    height: `${heightCm}cm`,
  };
  if (shape.wrap !== "inline") {
    style.position = "absolute";
    style.left = `${shape.offsetXCm ?? 0}cm`;
    style.top = `${shape.offsetYCm ?? 0}cm`;
    // Base tier matches Image's own behind/in-front split; the document-wide
    // zOrder (reorderShapeZ assigns a dense 0..N-1 across every drawing,
    // shapes and images alike) breaks ties *within* a tier so bring-to-
    // front/forward/send-backward/back (AC-29..33) actually change the
    // painted stacking order, not just the model. The two tiers are kept far
    // enough apart that no realistic document's zOrder range can cross
    // them — "behind" must always paint under the text regardless of how
    // high its zOrder climbs relative to an "in front" shape's own.
    const order = shape.zOrder ?? 0;
    style.zIndex = String(shape.wrap === "behind" ? order - 100000 : 100000 + order);
  }
  return style;
}

/** A chart's size and position rendering (SPEC-052) — identical logic to
 * shapeDisplayStyle since Chart extends the same DrawingPlacement fields;
 * kept as its own function (not a shared helper Image/Shape/Chart all call)
 * per this module's existing per-block-kind convention. */
export function chartDisplayStyle(chart: Chart): Record<string, string> {
  const heightCm = Math.max(chart.heightCm, 0.05);
  const style: Record<string, string> = {
    width: `${chart.widthCm}cm`,
    height: `${heightCm}cm`,
  };
  if (chart.wrap !== "inline") {
    style.position = "absolute";
    style.left = `${chart.offsetXCm ?? 0}cm`;
    style.top = `${chart.offsetYCm ?? 0}cm`;
    const order = chart.zOrder ?? 0;
    style.zIndex = String(chart.wrap === "behind" ? order - 100000 : 100000 + order);
  }
  return style;
}
