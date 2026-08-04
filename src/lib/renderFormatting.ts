// Maps the document model's formatting onto inline CSS for the canvas.
// Kept framework-free (no Vue imports) so it's testable in isolation, and
// split out of DocumentView.vue because that file's <script setup> is
// linted by ESLint's plain espree parser, which rejects TypeScript type
// annotations (see AppShell.vue's comment on the same constraint).
import type { Paragraph, Run, RunFormatting } from "./documentModel";
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

export function runStyle(formatting: RunFormatting): Record<string, string> {
  const style: Record<string, string> = {};
  if (formatting.bold) {
    style.fontWeight = "bold";
  }
  if (formatting.italic) {
    style.fontStyle = "italic";
  }
  const decorations: string[] = [];
  if (formatting.underline) {
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

export function paragraphStyle(paragraph: Paragraph): Record<string, string> {
  const formatting = paragraph.formatting;
  const style: Record<string, string> = { whiteSpace: "pre-wrap" };
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
