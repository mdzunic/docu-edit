// Maps the document model's formatting onto inline CSS for the canvas.
// Kept framework-free (no Vue imports) so it's testable in isolation, and
// split out of DocumentView.vue because that file's <script setup> is
// linted by ESLint's plain espree parser, which rejects TypeScript type
// annotations (see AppShell.vue's comment on the same constraint).
import type { BorderSide, Image, ListDefinition, Paragraph, Run, RunFormatting, Table, TableCell, TableRow } from "./documentModel";
import { listIndentCm } from "./lists";
import { definedTextIndentCm } from "./listDefinitions";
// The one highlight palette (OOXML name -> hex), shared with both file
// formats' readers/writers (SPEC-013) — the model stores the OOXML colour
// name as-is.
import { HIGHLIGHT_HEX } from "./odf/properties";

/** The HTML tag a paragraph should render as: h1-h6 for a resolved heading
 * level (clamped to h6), otherwise a plain paragraph. */
export function headingTag(level: number | undefined): string {
  if (!level) {
    return "p";
  }
  return `h${Math.min(Math.max(level, 1), 6)}`;
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
    style.fontFamily = formatting.fontFamily;
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

export function paragraphStyle(
  paragraph: Paragraph,
  listDefinitions: readonly ListDefinition[] = [],
): Record<string, string> {
  const formatting = paragraph.formatting;
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
