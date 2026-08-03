// Maps the document model's formatting onto inline CSS for the canvas.
// Kept framework-free (no Vue imports) so it's testable in isolation, and
// split out of DocumentView.vue because that file's <script setup> is
// linted by ESLint's plain espree parser, which rejects TypeScript type
// annotations (see AppShell.vue's comment on the same constraint).
import type { Paragraph, Run, RunFormatting } from "./documentModel";

// The 15 standard Word highlight colours (`w:highlight`'s closed value set)
// mapped to CSS — the document model stores the OOXML colour name as-is.
const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: "#ffff00",
  green: "#00ff00",
  cyan: "#00ffff",
  magenta: "#ff00ff",
  blue: "#0000ff",
  red: "#ff0000",
  darkBlue: "#00008b",
  darkCyan: "#008b8b",
  darkGreen: "#006400",
  darkMagenta: "#8b008b",
  darkRed: "#8b0000",
  darkYellow: "#808000",
  darkGray: "#a9a9a9",
  lightGray: "#d3d3d3",
  black: "#000000",
};

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
  if (formatting.highlight) {
    style.backgroundColor = HIGHLIGHT_COLORS[formatting.highlight] ?? formatting.highlight;
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
  if (formatting.indentLeftCm) {
    style.marginLeft = `${formatting.indentLeftCm}cm`;
  }
  if (formatting.indentRightCm) {
    style.marginRight = `${formatting.indentRightCm}cm`;
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
