// The product-neutral border vocabulary (SPEC-074 AC-1..3). `BorderSide` and
// `BorderLineStyle` started life inside core/tables.ts for SPEC-023's cell
// borders; SPEC-074 gives paragraphs and pages the same edges, so the types
// move here and core/tables.ts re-exports them — every shipped import keeps
// working (AC-2). Nothing here knows about a document model, a table or a
// paragraph: a border is a style, a width and a colour, whatever it is drawn
// around.

/** A border line's rendering: the standard presets plus width and colour.
 * `"none"` paints nothing, distinct from an absent `BorderSide` (unset —
 * falls back to whatever the renderer/writer defaults to).
 *
 * SPEC-023 shipped the first five; SPEC-074 adds `thick`, `dashDot` and
 * `dashDotDot` (AC-3) — enough for what documents actually use, without
 * reproducing OOXML's ~180-value `ST_Border` enumeration. */
export type BorderLineStyle =
  | "none"
  | "single"
  | "double"
  | "dashed"
  | "dotted"
  | "thick"
  | "dashDot"
  | "dashDotDot";

/** Every line style, in the order a picker offers them. */
export const BORDER_LINE_STYLES: readonly BorderLineStyle[] = [
  "none",
  "single",
  "thick",
  "double",
  "dashed",
  "dotted",
  "dashDot",
  "dashDotDot",
];

/** One edge's border (SPEC-023 AC-10..12, SPEC-074): style, width in points
 * and a CSS-ready colour ("#rrggbb", or "auto" for the source's automatic
 * colour). */
export interface BorderSide {
  style: BorderLineStyle;
  widthPt: number;
  color: string;
}

/** What `"auto"` paints as — both formats' automatic border colour is black
 * in every consumer that matters, and the renderer needs a real CSS value. */
export const AUTO_BORDER_COLOR = "#000000";

/** A border side's CSS colour: `"auto"` resolves to black, anything else is
 * already CSS-ready. */
export function borderColorCss(color: string): string {
  return color === "auto" ? AUTO_BORDER_COLOR : color;
}

/** The CSS `border-style` keyword for a model line style. CSS has no
 * dash-dot patterns, so both collapse to `dashed` — the same nearest-
 * equivalent rule the `.odt` writer applies (SPEC-074 AC-32); `thick` is a
 * solid line whose weight comes from its width. */
export function borderStyleCss(style: BorderLineStyle): string {
  switch (style) {
    case "none":
      return "none";
    case "double":
      return "double";
    case "dotted":
      return "dotted";
    case "dashed":
    case "dashDot":
    case "dashDotDot":
      return "dashed";
    default:
      return "solid";
  }
}

/** A border side as a CSS shorthand ("1pt solid #000000"), or `"none"` for
 * an unset or explicitly-none edge. Shared by every renderer that paints a
 * border into a style object. */
export function borderSideCss(side: BorderSide | undefined): string {
  if (!side || side.style === "none" || side.widthPt <= 0) {
    return "none";
  }
  // A double line needs at least 3px of box to show its gap; browsers
  // collapse a hairline double into a single rule otherwise.
  const widthPt = side.style === "double" ? Math.max(side.widthPt, 2.25) : side.widthPt;
  return `${roundWidthPt(widthPt)}pt ${borderStyleCss(side.style)} ${borderColorCss(side.color)}`;
}

/** Border widths are held as points with sub-point precision from both
 * formats' own units (eighths of a point in OOXML); rounding to three
 * decimals keeps CSS and written XML stable across a round trip without
 * accumulating float noise. */
export function roundWidthPt(widthPt: number): number {
  return Math.round(widthPt * 1000) / 1000;
}

/** Whether a side actually paints: set, not `"none"`, and wider than zero. */
export function borderSideVisible(side: BorderSide | undefined): boolean {
  return side !== undefined && side.style !== "none" && side.widthPt > 0;
}

/** Structural equality for one edge — the comparison the paragraph merge
 * rule (SPEC-074 AC-12) and the direct-formatting diff (SPEC-019) both
 * need. Two absent sides are equal; an absent side never equals a present
 * one, even a `"none"` one (an explicit off still overrides a style). */
export function borderSidesEqual(a: BorderSide | undefined, b: BorderSide | undefined): boolean {
  if (a === undefined || b === undefined) {
    return a === b;
  }
  return a.style === b.style && roundWidthPt(a.widthPt) === roundWidthPt(b.widthPt) && a.color === b.color;
}

/** Typed pass-through for a border line style coming from a `.vue`
 * template's plain string (espree bars the annotation in `.vue` scripts —
 * the same constraint core/tables.ts's asBorderLineStyle works around; that
 * name is kept there for SPEC-023's importers). */
export function asBorderLineStyleValue(value: string): BorderLineStyle {
  return value as BorderLineStyle;
}
