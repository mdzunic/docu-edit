// The frame → odp XML serialiser (SPEC-081): what a rebuilt `.odp` `draw:page`
// contains. The page keeps its original `draw:page` attributes and its
// preserved `presentation:notes`/`anim:*`; its frames become `draw:frame`/
// shape elements with GENERATED automatic styles collected by a StyleRegistry
// and spliced into content.xml's `office:automatic-styles` (existing styles are
// never rewritten — the reserved `de-` prefix cannot collide with a source's
// names, and `seedFrom` bumps the counters past any that somehow do). Foreign
// and group frames re-emit verbatim.
//
// ODF's own reader derives placeholder bullets from the frame's presentation
// class (its outline style on the master page), so a body placeholder's
// plain level-0 disc paragraphs still render without explicit list markup.
// Every other list (SPEC-097) is real markup: a run of list paragraphs wraps
// in `<text:list text:style-name="de-L{n}">`, one `text:list-item` each,
// nested by level, over a minted `text:list-style`; a body placeholder's
// NON-list paragraphs sit in a `text:list-header` (explicitly unnumbered).
// Paragraph spacing/indents join the paragraph style, letter spacing the text
// style, the text-body properties the frame's graphic style, and links wrap
// runs in `text:a` / whole frames in `draw:a`. The three golden-pending ODF
// details (odp/textDepthGolden.ts) are never written.
//
// Framework-free; imports only presentation siblings and src/lib/core.
import { GEOMETRY_TO_DRAW_TYPE } from "../../core/drawingMl";
import { adjustmentToGuideValue, resolvedShapeAdjustments } from "../../core/shapeCatalogue";
import { computeOwnerGrid, type BorderSide } from "../../core/tables";
import { escapeXmlAttribute, escapeXmlText } from "../../core/xml";
import { BULLET_CHARS, type ListMarkerStyle } from "../../core/lists";
import { sanitizeHref } from "../../core/urls";
import {
  DEFAULT_TABLE_CELL_PADDING_CM,
  type ChartFrame,
  type FrameStroke,
  type ImageFrame,
  type MediaFrame,
  type PlaceholderRole,
  type ShapeFrame,
  type Slide,
  type SlideFill,
  type SlideParagraph,
  type SlideRun,
  type SlideTableCell,
  type TableFrame,
  type TextBodyProps,
  type TextFrame,
} from "../model";
import { effectiveHangingCm, effectiveIndentCm, isListParagraph } from "../slideLists";
import { isLinkableKind } from "../slideLinks";
import { paragraphMarkerStyle } from "../text";
import { odpTransitionProps } from "../transitions";
import {
  DEFAULT_THEME,
  isThemeColor,
  resolveColor,
  resolveFont,
  type DocumentTheme,
  type ThemeColorRef,
} from "../../core/theme";
import { themeGraphicParent, themeTextParent } from "./theme";
import { writeOdpAnim } from "./anim";
import { odpDisplayProps, type OdpPageFieldPlan } from "./fields";

const ROLE_TO_CLASS: Record<PlaceholderRole, string> = {
  title: "title",
  subtitle: "subtitle",
  body: "body",
  other: "outline",
};

/** A length rounded to 3 decimals (`r3`), printed without trailing zeros. */
function r3(value: number): number {
  return Math.round(value * 1000) / 1000;
}

/** Points → centimetres (pt ÷ 72 × 2.54), r3 — the odp wire's length unit. */
function ptToCm(pt: number): number {
  return r3((pt / 72) * 2.54);
}

/** A preserved SPEC-097 property slot (`odpParaProps`/`odpBodyProps`) split
 * into its attribute string and its verbatim child elements: the attribute
 * part never contains "<" (values are escaped), so the first "<" starts the
 * children. */
function splitPreservedProps(slot: string | undefined): { attrs: string[]; children: string } {
  if (!slot) {
    return { attrs: [], children: "" };
  }
  const at = slot.indexOf("<");
  const attrPart = at >= 0 ? slot.slice(0, at) : slot;
  const children = at >= 0 ? slot.slice(at) : "";
  return { attrs: attrPart.match(/[\w.-]+:[\w.-]+="[^"]*"/g) ?? [], children };
}

/** Appends preserved attributes after the modelled ones, skipping any whose
 * name a modelled property already emitted (modelled wins — never one
 * attribute twice). */
function appendPreserved(props: string[], preserved: string[]): void {
  const names = new Set(props.flatMap((p) => Array.from(p.matchAll(/([\w.-]+:[\w.-]+)="/g), (m) => m[1])));
  for (const attr of preserved) {
    const name = attr.slice(0, attr.indexOf("="));
    if (!names.has(name)) {
      names.add(name);
      props.push(attr);
    }
  }
}

/** A properties element: self-closing, or wrapping its child elements. */
function propsElement(tag: string, props: string[], children: string): string {
  const attrs = props.length > 0 ? ` ${props.join(" ")}` : "";
  return children ? `<${tag}${attrs}>${children}</${tag}>` : `<${tag}${attrs}/>`;
}

function alignValue(align: NonNullable<SlideParagraph["align"]>): string {
  if (align === "center") return "center";
  if (align === "right") return "end";
  if (align === "justify") return "justify";
  return "start";
}

/** Collects the automatic styles a page rebuild generates, de-duplicated by
 * their property signature and named in the reserved `de-` namespace. */
/** Whether a ref carries no transforms — only these have a faithful
 * `de-theme-*` parent in odp (SPEC-094: a transformed ref writes the literal
 * alone and loses its reference on re-open — the stated loss). */
function isPlainRef(ref: ThemeColorRef): boolean {
  return ref.lumMod === undefined && ref.lumOff === undefined && ref.shade === undefined && ref.tint === undefined;
}

export class StyleRegistry {
  /** The deck theme refs resolve against; Office when the deck has none. */
  private readonly theme: DocumentTheme;

  constructor(theme: DocumentTheme = DEFAULT_THEME) {
    this.theme = theme;
  }

  private readonly paraStyles: string[] = [];
  private readonly textStyles: string[] = [];
  private readonly graphicStyles: string[] = [];
  private readonly drawingPageStyles: string[] = [];
  private readonly tableStyles: string[] = [];
  private readonly paraCache = new Map<string, string>();
  private readonly textCache = new Map<string, string>();
  private readonly graphicCache = new Map<string, string>();
  private readonly drawingPageCache = new Map<string, string>();
  private readonly tableCache = new Map<string, string>();
  private readonly listStyles: string[] = [];
  private readonly listCache = new Map<string, string>();
  private pCount = 0;
  private lCount = 0;
  private tCount = 0;
  private gCount = 0;
  private dpCount = 0;
  private colCount = 0;
  private rowCount = 0;
  private ceCount = 0;

  /** Advances every counter past any `de-`-prefixed names already present in
   * `source`, so a (pathological) source using the reserved prefix cannot
   * collide with a generated name. */
  seedFrom(source: string): void {
    const maxOf = (letter: string): number => {
      let max = 0;
      for (const m of source.matchAll(new RegExp(`de-${letter}(\\d+)`, "g"))) {
        max = Math.max(max, Number(m[1]));
      }
      return max;
    };
    this.pCount = maxOf("P");
    this.tCount = maxOf("T");
    this.gCount = maxOf("G");
    this.dpCount = maxOf("DP");
    this.lCount = maxOf("L");
  }

  /** A paragraph's automatic style (SPEC-081, widened by SPEC-097): the
   * alignment; the text edge `fo:margin-left` (slideLists' effective indent —
   * for every list item, any level > 0 and any explicit indent) and the
   * hanging `fo:text-indent` (negative; always for a list item); paragraph
   * spacing in cm; the line height (`fo:line-height` as a percentage for a
   * multiple or a length for "exact", `style:line-height-at-least` for the
   * at-least default); then the preserved `odpParaProps`. */
  paragraphStyle(para: SlideParagraph): string | undefined {
    const props: string[] = [];
    if (para.align) props.push(`fo:text-align="${alignValue(para.align)}"`);
    const list = isListParagraph(para);
    if (list || (para.level ?? 0) > 0 || para.indentCm !== undefined) {
      props.push(`fo:margin-left="${r3(effectiveIndentCm(para))}cm"`);
    }
    const hanging = effectiveHangingCm(para);
    if (list || hanging !== 0) props.push(`fo:text-indent="${r3(-hanging)}cm"`);
    if (para.spaceBeforePt) props.push(`fo:margin-top="${ptToCm(para.spaceBeforePt)}cm"`);
    if (para.spaceAfterPt) props.push(`fo:margin-bottom="${ptToCm(para.spaceAfterPt)}cm"`);
    if (para.lineSpacingPt !== undefined) {
      props.push(
        para.lineSpacingRule === "exact"
          ? `fo:line-height="${r3(para.lineSpacingPt)}pt"`
          : `style:line-height-at-least="${r3(para.lineSpacingPt)}pt"`,
      );
    } else if (para.lineSpacing !== undefined && para.lineSpacing !== 1) {
      props.push(`fo:line-height="${r3(para.lineSpacing * 100)}%"`);
    }
    const preserved = splitPreservedProps(para.odpParaProps);
    appendPreserved(props, preserved.attrs);
    if (props.length === 0 && !preserved.children) return undefined;
    const propsXml = propsElement("style:paragraph-properties", props, preserved.children);
    const cached = this.paraCache.get(propsXml);
    if (cached) return cached;
    const name = `de-P${++this.pCount}`;
    this.paraCache.set(propsXml, name);
    this.paraStyles.push(`<style:style style:name="${name}" style:family="paragraph">${propsXml}</style:style>`);
    return name;
  }

  /** Mints (or reuses) a `de-L{n}` `text:list-style` (SPEC-097) from its
   * level styles (1-based level → `text:list-level-style-*` markup),
   * de-duplicated by that signature. */
  listStyle(levels: Map<number, string>): string {
    const key = [...levels.entries()]
      .toSorted((a, b) => a[0] - b[0])
      .map(([, xml]) => xml)
      .join("");
    const cached = this.listCache.get(key);
    if (cached) return cached;
    const name = `de-L${++this.lCount}`;
    this.listCache.set(key, name);
    this.listStyles.push(`<text:list-style style:name="${name}">${key}</text:list-style>`);
    return name;
  }

  textStyle(run: SlideRun): string | undefined {
    const props: string[] = [];
    if (run.bold) props.push(`fo:font-weight="bold"`);
    if (run.italic) props.push(`fo:font-style="italic"`);
    if (run.underline) {
      props.push(`style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"`);
    }
    if (run.fontFamily) {
      // A theme font ref resolves to its family name — the stated odp loss
      // (SPEC-094): ODF has no font placeholder, so the ref does not survive.
      const f = escapeXmlAttribute(resolveFont(run.fontFamily, this.theme));
      props.push(`style:font-name="${f}" fo:font-family="${f}"`);
    }
    if (run.sizePt !== undefined) props.push(`fo:font-size="${run.sizePt}pt"`);
    if (run.letterSpacingPt) props.push(`fo:letter-spacing="${ptToCm(run.letterSpacingPt)}cm"`);
    // A colour ref writes its resolved literal (other readers see the right
    // colour); an UNTRANSFORMED ref also parents the automatic style on the
    // `de-theme-text-{slot}` scheme style, which the reader maps back to a ref.
    let parent = "";
    if (run.color) {
      props.push(`fo:color="${escapeXmlAttribute(resolveColor(run.color, this.theme))}"`);
      if (isThemeColor(run.color) && isPlainRef(run.color)) {
        parent = ` style:parent-style-name="${themeTextParent(run.color.theme)}"`;
      }
    }
    if (props.length === 0) return undefined;
    const key = `${parent}::${props.join(" ")}`;
    const cached = this.textCache.get(key);
    if (cached) return cached;
    const name = `de-T${++this.tCount}`;
    this.textCache.set(key, name);
    this.textStyles.push(
      `<style:style style:name="${name}" style:family="text"${parent}><style:text-properties ${props.join(" ")}/></style:style>`,
    );
    return name;
  }

  /** A frame's graphic style: fill, stroke and (SPEC-097) the text body —
   * `fo:padding-*` for stored insets, `draw:textarea-vertical-align` for a
   * stored anchor, `draw:auto-grow-height="true"` for "resizeShape",
   * `fo:wrap-option="no-wrap"`, a `style:columns` child, then the preserved
   * `odpBodyProps`. Autofit "shrink" writes nothing: its ODF attribute is
   * golden-pending (textDepthGolden.ts). */
  graphicStyle(style: { fill?: SlideFill; stroke?: FrameStroke; body?: TextBodyProps; bodyProps?: string }): string | undefined {
    const props: string[] = [];
    let parent = "";
    if (style.fill?.type === "solid") {
      const color = style.fill.color;
      props.push(`draw:fill="solid" draw:fill-color="${escapeXmlAttribute(resolveColor(color, this.theme))}"`);
      if (isThemeColor(color) && isPlainRef(color)) {
        parent = ` style:parent-style-name="${themeGraphicParent(color.theme)}"`;
      }
    } else if (style.fill?.type === "none") props.push(`draw:fill="none"`);
    if (style.stroke) {
      // A stroke ref has no faithful parent (the scheme styles carry fill), so
      // it writes its resolved literal — part of the stated odp loss.
      props.push(
        `draw:stroke="solid" svg:stroke-color="${escapeXmlAttribute(resolveColor(style.stroke.color, this.theme))}" svg:stroke-width="${(style.stroke.widthPt / 72 * 2.54).toFixed(3)}cm"`,
      );
    }
    const body = style.body;
    if (body?.insetsCm) {
      const { left, top, right, bottom } = body.insetsCm;
      props.push(`fo:padding-left="${r3(left)}cm" fo:padding-top="${r3(top)}cm" fo:padding-right="${r3(right)}cm" fo:padding-bottom="${r3(bottom)}cm"`);
    }
    if (body?.anchor) props.push(`draw:textarea-vertical-align="${body.anchor}"`);
    if (body?.autofit === "resizeShape") props.push(`draw:auto-grow-height="true"`);
    if (body?.wrap === false) props.push(`fo:wrap-option="no-wrap"`);
    appendPreserved(props, splitPreservedProps(style.bodyProps).attrs);
    const columns = body?.columns
      ? `<style:columns fo:column-count="${body.columns.count}" fo:column-gap="${r3(body.columns.gapCm)}cm"/>`
      : "";
    if (props.length === 0 && !columns) return undefined;
    const propsXml = propsElement("style:graphic-properties", props, columns);
    const key = `${parent}::${propsXml}`;
    const cached = this.graphicCache.get(key);
    if (cached) return cached;
    const name = `de-G${++this.gCount}`;
    this.graphicCache.set(key, name);
    this.graphicStyles.push(`<style:style style:name="${name}" style:family="graphic"${parent}>${propsXml}</style:style>`);
    return name;
  }

  /** Mints (or reuses) a `de-DP{n}` drawing-page style carrying a transition's
   * smil/presentation attributes plus the copied fill/background of the page's
   * original drawing-page style, so setting a transition never disturbs the
   * background (SPEC-085). `baseProps` is the filtered original property string
   * ("" for a page that had none). Returns the minted style name. */
  drawingPageStyle(smilProps: string, baseProps: string): string {
    const props = [baseProps, smilProps].filter((p) => p.length > 0).join(" ");
    const cached = this.drawingPageCache.get(props);
    if (cached) {
      return cached;
    }
    const name = `de-DP${++this.dpCount}`;
    this.drawingPageCache.set(props, name);
    this.drawingPageStyles.push(
      `<style:style style:name="${name}" style:family="drawing-page"><style:drawing-page-properties ${props}/></style:style>`,
    );
    return name;
  }

  /** A `table-column` automatic style carrying a column width (SPEC-091). */
  columnStyle(widthCm: number): string {
    return this.tableFamilyStyle("table-column", `<style:table-column-properties style:column-width="${widthCm}cm"/>`, () => `de-C${++this.colCount}`);
  }

  /** A `table-row` automatic style carrying a (minimum) row height. */
  rowStyle(heightCm: number): string {
    return this.tableFamilyStyle(
      "table-row",
      `<style:table-row-properties style:row-height="${heightCm}cm" style:min-row-height="${heightCm}cm"/>`,
      () => `de-R${++this.rowCount}`,
    );
  }

  /** A `table-cell` automatic style carrying the cell's fill / padding / valign
   * / borders, or undefined when the cell has no modelled non-default look. */
  tableCellStyle(cell: SlideTableCell): string | undefined {
    const props: string[] = [];
    // A cell fill ref writes its resolved literal only — table-cell styles
    // cannot parent the graphic-family scheme styles (the stated odp loss).
    if (cell.fill?.type === "solid") props.push(`fo:background-color="${escapeXmlAttribute(resolveColor(cell.fill.color, this.theme))}"`);
    else if (cell.fill?.type === "none") props.push(`fo:background-color="transparent"`);
    if (cell.paddingCm !== undefined && cell.paddingCm !== DEFAULT_TABLE_CELL_PADDING_CM) {
      props.push(`fo:padding="${cell.paddingCm}cm"`);
    }
    if (cell.verticalAlign === "center") props.push(`style:vertical-align="middle"`);
    else if (cell.verticalAlign === "bottom") props.push(`style:vertical-align="bottom"`);
    if (cell.borders) {
      for (const [edge, side] of [
        ["top", cell.borders.top],
        ["bottom", cell.borders.bottom],
        ["left", cell.borders.left],
        ["right", cell.borders.right],
      ] as const) {
        if (side) props.push(`fo:border-${edge}="${foBorderValue(side)}"`);
      }
    }
    if (props.length === 0) return undefined;
    return this.tableFamilyStyle("table-cell", `<style:table-cell-properties ${props.join(" ")}/>`, () => `de-CE${++this.ceCount}`);
  }

  private tableFamilyStyle(family: string, propsXml: string, mint: () => string): string {
    const key = `${family}::${propsXml}`;
    const cached = this.tableCache.get(key);
    if (cached) return cached;
    const name = mint();
    this.tableCache.set(key, name);
    this.tableStyles.push(`<style:style style:name="${name}" style:family="${family}">${propsXml}</style:style>`);
    return name;
  }

  hasStyles(): boolean {
    return (
      this.paraStyles.length +
        this.textStyles.length +
        this.graphicStyles.length +
        this.drawingPageStyles.length +
        this.tableStyles.length +
        this.listStyles.length >
      0
    );
  }

  /** The generated `<style:style>` (and SPEC-097 `<text:list-style>`)
   * elements, to splice before `</office:automatic-styles>`. */
  xml(): string {
    return [
      ...this.paraStyles,
      ...this.textStyles,
      ...this.graphicStyles,
      ...this.drawingPageStyles,
      ...this.tableStyles,
      ...this.listStyles,
    ].join("");
  }
}

/** A `BorderSide` → an `fo:border-*` value ("0.5pt solid #000000"). */
function foBorderValue(side: BorderSide): string {
  if (side.style === "none") {
    return "none";
  }
  const line = side.style === "single" ? "solid" : side.style;
  return `${side.widthPt}pt ${line} ${side.color}`;
}

const DRAW_NS_ATTR = ` draw:name`;

/** The value of `draw:name` (or another attribute) in a page's attribute
 * string. */
export function drawAttr(attrs: string, attr: string): string | undefined {
  const match = new RegExp(`${attr}="([^"]*)"`).exec(attrs);
  return match ? match[1] : undefined;
}

/** Serialises a rebuilt (dirty or new) page. `master` supplies the
 * `draw:master-page-name` for a from-scratch page; a page opened from a file
 * keeps its original `draw:page` attributes (from `preserved["@page"]`) with a
 * uniqued `draw:name`. */
export function serializePage(
  slide: Slide,
  registry: StyleRegistry,
  usedNames: Set<string>,
  master: string,
  styleSource = "",
  layoutName?: string,
  fieldPlan?: OdpPageFieldPlan,
): string {
  let attrs = slide.preserved?.["@page"];
  if (attrs) {
    const current = drawAttr(attrs, "draw:name") ?? "page";
    const unique = uniqueName(current, usedNames);
    attrs = attrs.replace(/draw:name="[^"]*"/, `draw:name="${escapeXmlAttribute(unique)}"`);
    if (!/draw:name=/.test(attrs)) {
      attrs = `${DRAW_NS_ATTR}="${escapeXmlAttribute(unique)}"${attrs}`;
    }
  } else {
    const name = uniqueName("page", usedNames);
    attrs = ` draw:name="${escapeXmlAttribute(name)}" draw:master-page-name="${escapeXmlAttribute(master)}"`;
  }
  // A layout change rewrites the page's `presentation:presentation-page-layout-name`
  // (adding it when absent), leaving `draw:master-page-name` untouched — layout
  // and master are separate ODF axes (SPEC-089 AC-15).
  if (layoutName !== undefined) {
    attrs = setPresentationPageLayoutName(attrs, layoutName);
  }
  // SPEC-093: a header/footer change stamps this page's declaration references
  // (`use-*-name`, stripped when the field is off) and folds its display flags
  // into the drawing-page style below. Field placeholders are never emitted as
  // frames (AC-80).
  if (fieldPlan) {
    attrs = setPresentationName(attrs, "use-footer-name", fieldPlan.footerName);
    attrs = setPresentationName(attrs, "use-date-time-name", fieldPlan.dateName);
  }
  // A modelled transition and/or a field display change mints a fresh
  // drawing-page style (copying the original style's fill/background, minus any
  // replaced transition or stale display flags) and repoints `draw:style-name`
  // at it (SPEC-085 AC-11 / SPEC-093 AC-79).
  const smilProps = slide.transition ? odpTransitionProps(slide.transition) : "";
  const displayProps = fieldPlan ? odpDisplayProps(fieldPlan) : "";
  if (smilProps || displayProps) {
    const originalStyle = drawAttr(attrs, "draw:style-name");
    const baseProps = originalStyle ? drawingPageBaseProps(styleSource, originalStyle) : "";
    const props = [smilProps, displayProps].filter((p) => p.length > 0).join(" ");
    const styleName = registry.drawingPageStyle(props, baseProps);
    attrs = setDrawStyleName(attrs, styleName);
  }
  // When the page carries modelled animations, its animatable frames get a
  // deterministic wire id (`de-f{n}` in z-order among text/shape/image frames)
  // on both `xml:id` and `draw:id` so the `anim:*` tree's `smil:targetElement`
  // can reference them (SPEC-086). Untouched/animation-free pages emit no ids,
  // so their bytes stay verbatim.
  const wireIdFor = new Map<string, string>();
  if (slide.animations !== undefined) {
    let n = 1;
    for (const frame of slide.frames) {
      const animatable =
        frame.kind === "text" ||
        frame.kind === "shape" ||
        frame.kind === "image" ||
        frame.kind === "table" ||
        (frame.kind === "chart" && !frame.xml);
      if (animatable) {
        wireIdFor.set(frame.id, `de-f${n++}`);
      }
    }
  }
  const frames = slide.frames.map((frame) => frameXml(frame, registry, wireIdFor.get(frame.id))).join("");
  const notes = notesXml(slide.notes);
  // Modelled animations win over any preserved timing (never both).
  const anim = slide.animations !== undefined
    ? writeOdpAnim(slide.animations, (id) => wireIdFor.get(id))
    : (slide.preserved?.["anim"] ?? "");
  return `<draw:page${attrs}>${frames}${notes}${anim}</draw:page>`;
}

/** Sets/replaces `draw:style-name` in a page's attribute string. */
function setDrawStyleName(attrs: string, styleName: string): string {
  if (/draw:style-name="[^"]*"/.test(attrs)) {
    return attrs.replace(/draw:style-name="[^"]*"/, `draw:style-name="${styleName}"`);
  }
  return `${attrs} draw:style-name="${styleName}"`;
}

/** Sets/replaces a `presentation:{localName}` attribute in a page's attribute
 * string, or strips it when `value` is undefined (SPEC-093 `use-*-name`). */
function setPresentationName(attrs: string, localName: string, value: string | undefined): string {
  const stripped = attrs.replace(new RegExp(`\\s*presentation:${localName}="[^"]*"`), "");
  return value === undefined ? stripped : `${stripped} presentation:${localName}="${escapeXmlAttribute(value)}"`;
}

/** Sets/replaces `presentation:presentation-page-layout-name` in a page's
 * attribute string (SPEC-089 AC-15). Appended when absent. */
function setPresentationPageLayoutName(attrs: string, layoutName: string): string {
  const escaped = escapeXmlAttribute(layoutName);
  if (/presentation:presentation-page-layout-name="[^"]*"/.test(attrs)) {
    return attrs.replace(
      /presentation:presentation-page-layout-name="[^"]*"/,
      `presentation:presentation-page-layout-name="${escaped}"`,
    );
  }
  return `${attrs} presentation:presentation-page-layout-name="${escaped}"`;
}

/** The fill/background properties of the drawing-page style named `styleName`
 * in `source`, with any transition (smil/presentation-transition/duration)
 * attributes stripped — what the minted style copies so a transition edit
 * leaves the background untouched. "" when the style is not found. */
function drawingPageBaseProps(source: string, styleName: string): string {
  const escaped = styleName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const styleRe = new RegExp(
    `<style:style\\b[^>]*style:name="${escaped}"[^>]*>([\\s\\S]*?)</style:style>`,
  );
  const styleMatch = styleRe.exec(source);
  const inner = styleMatch
    ? styleMatch[1]
    : new RegExp(`<style:style\\b[^>]*style:name="${escaped}"[^>]*/>`).test(source)
      ? ""
      : "";
  const propsMatch = /<style:drawing-page-properties\b([^>]*?)\/?>/.exec(inner);
  if (!propsMatch) {
    return "";
  }
  const kept: string[] = [];
  for (const attr of propsMatch[1].matchAll(/([\w-]+:[\w-]+)="([^"]*)"/g)) {
    const name = attr[1];
    if (
      name.startsWith("smil:") ||
      name === "presentation:transition-type" ||
      name === "presentation:transition-speed" ||
      name === "presentation:duration" ||
      name === "presentation:display-footer" ||
      name === "presentation:display-page-number" ||
      name === "presentation:display-date-time" ||
      name === "presentation:display-header"
    ) {
      continue; // drop the transition / stale display flags being replaced
    }
    kept.push(`${name}="${attr[2]}"`);
  }
  return kept.join(" ");
}

/** A page's `presentation:notes` regenerated from the model's speaker notes
 * (SPEC-083, AC-16): a fixed-geometry `draw:frame presentation:class="notes"`
 * whose text-box carries one `text:p` per entry (formatting flat). Omitted
 * entirely when the slide has no notes (AC-25). */
function notesXml(notes: string[] | undefined): string {
  if (!notes || notes.length === 0) {
    return "";
  }
  const paragraphs = notes
    .map((line) => (line.length > 0 ? `<text:p>${escapeXmlText(line)}</text:p>` : `<text:p/>`))
    .join("");
  return (
    `<presentation:notes>` +
    `<draw:frame presentation:class="notes" svg:x="1.5cm" svg:y="13cm" svg:width="16cm" svg:height="9cm">` +
    `<draw:text-box>${paragraphs}</draw:text-box></draw:frame>` +
    `</presentation:notes>`
  );
}

function uniqueName(base: string, used: Set<string>): string {
  let candidate = base;
  let n = 2;
  while (used.has(candidate)) {
    candidate = `${base}_${n++}`;
  }
  used.add(candidate);
  return candidate;
}

/** The `xml:id`/`draw:id` attribute string for a frame with a wire id, "" when
 * it has none (SPEC-086 — group/foreign frames re-emit verbatim and never
 * receive one, and animations never target them). */
function idAttrs(wireId: string | undefined): string {
  return wireId ? ` xml:id="${wireId}" draw:id="${wireId}"` : "";
}

/** A linkable frame's element wrapped in `<draw:a>` when it carries an
 * external frame link (SPEC-097 — isLinkableKind: text/image/shape; a
 * `linkSlide` writes nothing, the slide-jump form being golden-pending). */
function withFrameLink(frame: Slide["frames"][number], xml: string): string {
  const href = isLinkableKind(frame.kind) ? writableHref(frame.linkHref) : null;
  if (!href) {
    return xml;
  }
  const title = frame.linkTooltip ? ` office:title="${escapeXmlAttribute(frame.linkTooltip)}"` : "";
  return `<draw:a xlink:href="${escapeXmlAttribute(href)}" xlink:type="simple"${title}>${xml}</draw:a>`;
}

function frameXml(frame: Slide["frames"][number], registry: StyleRegistry, wireId?: string): string {
  switch (frame.kind) {
    case "text":
      return withFrameLink(frame, textFrameXml(frame, registry, wireId));
    case "shape":
      return withFrameLink(frame, shapeFrameXml(frame, registry, wireId));
    case "image":
      return withFrameLink(frame, imageFrameXml(frame, registry, wireId));
    case "table":
      return tableFrameXml(frame, registry, wireId);
    case "chart":
      // Untouched chart re-emits verbatim (byte-identical, AC-21); a touched
      // one emits the canonical slide-flavoured draw:frame + draw:object.
      return frame.xml ?? serializeOdpChartFrame(frame, chartObjectHref(frame.sourcePart), wireId);
    case "group":
      // Verbatim (still carrying its source xml) re-emits unchanged; an
      // authored group rebuilds as draw:g, whose children keep their OWN
      // svg:x/svg:y — a draw:g has no box of its own (SPEC-096).
      return frame.xml ?? `<draw:g>${frame.children.map((child) => frameXml(child, registry)).join("")}</draw:g>`;
    case "media":
      // Untouched clip re-emits verbatim (byte-identical, AC-44); a touched or
      // inserted one rebuilds the canonical draw:frame + draw:plugin (SPEC-095).
      return frame.sourceXml ?? mediaFrameXml(frame, registry, wireId);
    case "foreign":
      return frame.xml; // verbatim (AC-39)
    default:
      return "";
  }
}

/** Serialises a `GroupFrame` (with its generated automatic styles) — the test
 * seam the golden spec pins (SPEC-096, both directions). */
export function serializeOdpGroupFrame(frame: Slide["frames"][number]): { frame: string; autoStyles: string } {
  const registry = new StyleRegistry();
  const xml = frameXml(frame, registry);
  return { frame: xml, autoStyles: registry.xml() };
}

/** The `./Object N` href a chart frame's `draw:object` points at, derived from
 * its `sourcePart` ("Object 1/content.xml" → "./Object 1"); "" for a chart
 * created here (the writer mints the directory and rewrites this). */
function chartObjectHref(sourcePart?: string): string {
  if (!sourcePart) {
    return "";
  }
  return `./${sourcePart.replace(/\/content\.xml$/, "")}`;
}

/** The canonical slide-flavoured `draw:frame` + `draw:object` for a touched
 * chart (SPEC-092 AC-21): `presentation:class="chart"`, geometry via
 * `svg:x/y/width/height` (not `text:anchor-type` — that is the text product's
 * container), holding the `draw:object` reference alone. */
export function serializeOdpChartFrame(frame: ChartFrame, objectHref: string, wireId?: string): string {
  return (
    `<draw:frame presentation:class="chart"${idAttrs(wireId)}${geomAttrs(frame)}>` +
    `<draw:object xlink:href="${escapeXmlAttribute(objectHref)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
    `</draw:frame>`
  );
}

// --- table frame (SPEC-091) -------------------------------------------------

/** How many leading rows form the table's header band: the contiguous leading
 * `isHeader` rows, plus row 0 when the frame carries the `firstRow` look. */
function headerRowCount(frame: TableFrame): number {
  let n = 0;
  while (n < frame.rows.length && (frame.rows[n].isHeader === true || (frame.firstRow === true && n === 0))) {
    n++;
  }
  return n;
}

function tableCellXml(cell: SlideTableCell, registry: StyleRegistry): string {
  if (cell.covered) {
    return `<table:covered-table-cell/>`;
  }
  const styleName = cell.odpCellStyleName ?? registry.tableCellStyle(cell);
  const styleAttr = styleName ? ` table:style-name="${escapeXmlAttribute(styleName)}"` : "";
  const colSpan = cell.colSpan && cell.colSpan > 1 ? ` table:number-columns-spanned="${cell.colSpan}"` : "";
  const rowSpan = cell.rowSpan && cell.rowSpan > 1 ? ` table:number-rows-spanned="${cell.rowSpan}"` : "";
  const body = paragraphsXml(cell.paragraphs, registry, false);
  return `<table:table-cell${styleAttr}${colSpan}${rowSpan}>${body}</table:table-cell>`;
}

function tableFrameXml(frame: TableFrame, registry: StyleRegistry, wireId?: string): string {
  // Normalise covered slots against the owner grid so a stale `covered` flag
  // never emits a spurious covered cell.
  const grid = computeOwnerGrid(frame);
  const columns = frame.columnWidthsCm
    .map((w, c) => {
      const styleName = frame.odpColumnStyleNames?.[c] ?? registry.columnStyle(w);
      return `<table:table-column table:style-name="${escapeXmlAttribute(styleName)}"/>`;
    })
    .join("");
  const rowXml = frame.rows.map((row, r) => {
    const cells = row.cells
      .map((cell, c) => {
        const owner = grid[r][c];
        const covered = owner.originRow !== r || owner.originCol !== c;
        return tableCellXml(covered ? { paragraphs: [], covered: true } : cell, registry);
      })
      .join("");
    const styleName = row.odpRowStyleName ?? (row.heightCm !== undefined ? registry.rowStyle(row.heightCm) : undefined);
    const styleAttr = styleName ? ` table:style-name="${escapeXmlAttribute(styleName)}"` : "";
    return `<table:table-row${styleAttr}>${cells}</table:table-row>`;
  });
  const headerCount = headerRowCount(frame);
  let body: string;
  if (headerCount > 0) {
    body = `<table:table-header-rows>${rowXml.slice(0, headerCount).join("")}</table:table-header-rows>${rowXml.slice(headerCount).join("")}`;
  } else {
    body = rowXml.join("");
  }
  const tableStyleAttr = frame.odpTableStyleName ? ` table:style-name="${escapeXmlAttribute(frame.odpTableStyleName)}"` : "";
  const table = `<table:table${tableStyleAttr}>${columns}${body}</table:table>`;
  return `<draw:frame${idAttrs(wireId)}${geomAttrs(frame)}>${table}</draw:frame>`;
}

/** Serialises a `TableFrame` to its `draw:frame` markup plus the generated
 * automatic styles it references — the test seam the golden spec pins
 * (SPEC-091 AC-16, both directions). */
export function serializeOdpTableFrame(frame: TableFrame): { frame: string; autoStyles: string } {
  const registry = new StyleRegistry();
  const xml = tableFrameXml(frame, registry, undefined);
  return { frame: xml, autoStyles: registry.xml() };
}

function geomAttrs(frame: Slide["frames"][number]): string {
  const geom = ` svg:x="${frame.xCm}cm" svg:y="${frame.yCm}cm" svg:width="${frame.widthCm}cm" svg:height="${frame.heightCm}cm"`;
  // ODF draw:transform rotate() is counter-clockwise radians about the centre.
  const transform = frame.rotationDeg
    ? ` draw:transform="rotate(${((-frame.rotationDeg * Math.PI) / 180).toFixed(6)})"`
    : "";
  return `${geom}${transform}`;
}

function textFrameXml(frame: TextFrame, registry: StyleRegistry, wireId?: string): string {
  const cls = frame.role ? ROLE_TO_CLASS[frame.role] : undefined;
  const classAttr = cls ? ` presentation:class="${cls}"` : "";
  const styleName = registry.graphicStyle({ fill: frame.fill, stroke: frame.stroke, body: frame.body, bodyProps: frame.odpBodyProps });
  const styleAttr = styleName ? ` draw:style-name="${styleName}"` : "";
  const body = paragraphsXml(frame.paragraphs, registry, cls !== undefined && IMPLICIT_BULLET_CLASSES.has(cls));
  return `<draw:frame${classAttr}${idAttrs(wireId)}${styleAttr}${geomAttrs(frame)}><draw:text-box>${body}</draw:text-box></draw:frame>`;
}

function imageFrameXml(frame: ImageFrame, registry: StyleRegistry, wireId?: string): string {
  const styleName = registry.graphicStyle({ stroke: frame.stroke });
  const styleAttr = styleName ? ` draw:style-name="${styleName}"` : "";
  return (
    `<draw:frame${idAttrs(wireId)}${styleAttr}${geomAttrs(frame)}>` +
    `<draw:image xlink:href="${escapeXmlAttribute(frame.mediaId)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
    `</draw:frame>`
  );
}

/** The `draw:param` playback-flag rows for a media frame (SPEC-095), in
 * AutoPlay/Loop/Mute/Volume order; "" when it carries no non-default flag. */
function mediaParamsXml(frame: MediaFrame): string {
  const rows: string[] = [];
  if (frame.autoStart) {
    rows.push(`<draw:param draw:name="AutoPlay" draw:value="true"/>`);
  }
  if (frame.loop) {
    rows.push(`<draw:param draw:name="Loop" draw:value="true"/>`);
  }
  if (frame.muted) {
    rows.push(`<draw:param draw:name="Mute" draw:value="true"/>`);
  }
  if (frame.volume !== undefined) {
    rows.push(`<draw:param draw:name="Volume" draw:value="${frame.volume}"/>`);
  }
  return rows.join("");
}

/** A rebuilt media `draw:frame` (SPEC-095): a `draw:plugin` with the clip href,
 * `draw:mime-type`, the playback-flag `draw:param` rows, and a sibling
 * `draw:image` poster; `presentation:play-full-screen` on the frame. An
 * untouched clip re-emits its verbatim `sourceXml` instead (see `frameXml`). */
function mediaFrameXml(frame: MediaFrame, registry: StyleRegistry, wireId?: string): string {
  const styleName = registry.graphicStyle({ stroke: frame.stroke });
  const styleAttr = styleName ? ` draw:style-name="${styleName}"` : "";
  const fullScreen = frame.fullScreen ? ` presentation:play-full-screen="true"` : "";
  const href = frame.mediaId ?? frame.linkTarget ?? "";
  const mime = frame.contentType ? ` draw:mime-type="${escapeXmlAttribute(frame.contentType)}"` : "";
  const poster = frame.posterMediaId
    ? `<draw:image xlink:href="${escapeXmlAttribute(frame.posterMediaId)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`
    : "";
  return (
    `<draw:frame${idAttrs(wireId)}${styleAttr}${geomAttrs(frame)}${fullScreen}>` +
    `<draw:plugin xlink:href="${escapeXmlAttribute(href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"${mime}>` +
    `${mediaParamsXml(frame)}</draw:plugin>${poster}` +
    `</draw:frame>`
  );
}

function paragraphXml(para: SlideParagraph, registry: StyleRegistry): string {
  const styleName = registry.paragraphStyle(para);
  const styleAttr = styleName ? ` text:style-name="${styleName}"` : "";
  return `<text:p${styleAttr}>${runsXml(para.runs, registry)}</text:p>`;
}

/** The external href a link writes, or null: sanitised, and never an
 * in-deck `#…` target — the slide-jump link form is golden-pending
 * (textDepthGolden.ts), so `linkSlide` writes no link at all. */
function writableHref(href: string | undefined): string | null {
  if (!href || href.trim().startsWith("#")) {
    return null;
  }
  return sanitizeHref(href);
}

/** A paragraph's runs, consecutive runs sharing one external link (href +
 * tooltip) wrapped together in a `text:a` (SPEC-097). */
function runsXml(runs: SlideRun[], registry: StyleRegistry): string {
  let out = "";
  let i = 0;
  while (i < runs.length) {
    const href = writableHref(runs[i].linkHref);
    if (!href) {
      out += runXml(runs[i], registry);
      i += 1;
      continue;
    }
    const tooltip = runs[i].linkTooltip;
    let inner = "";
    while (i < runs.length && writableHref(runs[i].linkHref) === href && runs[i].linkTooltip === tooltip) {
      inner += runXml(runs[i], registry);
      i += 1;
    }
    const title = tooltip ? ` office:title="${escapeXmlAttribute(tooltip)}"` : "";
    out += `<text:a xlink:href="${escapeXmlAttribute(href)}" xlink:type="simple"${title}>${inner}</text:a>`;
  }
  return out;
}

function runXml(run: SlideRun, registry: StyleRegistry): string {
  const text = escapeXmlText(run.text);
  const styleName = registry.textStyle(run);
  return styleName ? `<text:span text:style-name="${styleName}">${text}</text:span>` : text;
}

/** The presentation classes whose placeholders ODF readers bullet from the
 * master's outline style (the reader's body role): their plain level-0 disc
 * paragraphs need no list markup, and their non-list paragraphs need an
 * explicit `text:list-header`. */
const IMPLICIT_BULLET_CLASSES = new Set(["outline", "body", "text", "notes"]);

/** A paragraph a body placeholder bullets by itself: a level-0 disc with no
 * custom glyph, picture or other style. */
function isPlaceholderBullet(para: SlideParagraph): boolean {
  return (
    para.bullet === "disc" &&
    (para.level ?? 0) === 0 &&
    (para.bulletStyle === undefined || para.bulletStyle === "disc") &&
    para.bulletChar === undefined &&
    para.bulletImageId === undefined
  );
}

/** A text body's paragraphs (SPEC-097): plain paragraphs stay bare `text:p`
 * (in a body placeholder, inside one `text:list-header` per run so they read
 * back unbulleted); each maximal run of list paragraphs becomes real list
 * markup — except a body placeholder's run of plain level-0 discs, which the
 * placeholder bullets by itself (SPEC-081's bytes). */
function paragraphsXml(paragraphs: SlideParagraph[], registry: StyleRegistry, implicitBullets: boolean): string {
  let out = "";
  let i = 0;
  while (i < paragraphs.length) {
    const list = isListParagraph(paragraphs[i]);
    let j = i;
    while (j < paragraphs.length && isListParagraph(paragraphs[j]) === list) {
      j += 1;
    }
    const run = paragraphs.slice(i, j);
    if (!list) {
      const body = run.map((para) => paragraphXml(para, registry)).join("");
      out += implicitBullets ? `<text:list><text:list-header>${body}</text:list-header></text:list>` : body;
    } else if (implicitBullets && run.every(isPlaceholderBullet)) {
      out += run.map((para) => paragraphXml(para, registry)).join("");
    } else {
      out += listRunXml(run, registry);
    }
    i = j;
  }
  return out;
}

const NUM_FORMATS: Partial<Record<ListMarkerStyle, string>> = {
  decimal: "1",
  lowerAlpha: "a",
  upperAlpha: "A",
  lowerRoman: "i",
  upperRoman: "I",
};

/** A list level's `style:list-level-properties`: label-alignment mode with
 * the text edge (`fo:margin-left`) and the hanging marker box (negative
 * `fo:text-indent`), the form Impress writes. */
function listLevelPropsXml(indentCm: number, hangingCm: number): string {
  return (
    `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">` +
    `<style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="${r3(-hangingCm)}cm" fo:margin-left="${r3(indentCm)}cm"/>` +
    `</style:list-level-properties>`
  );
}

/** What a list level's marker is — the part two paragraphs at one level of
 * one `text:list` must share (a picture, a glyph, or a number format). */
function markerKey(para: SlideParagraph): string {
  if (para.bulletImageId) return `image:${para.bulletImageId}`;
  const style = paragraphMarkerStyle(para) ?? "disc";
  if (para.bullet === "number") return `number:${NUM_FORMATS[style] ?? "1"}`;
  return `bullet:${para.bulletChar ?? BULLET_CHARS[style] ?? BULLET_CHARS.disc}`;
}

/** One `text:list-level-style-*` (1-based `level`) for the first paragraph
 * at that level: `-image` (xlink:href into Pictures/), `-number`
 * (`style:num-format`, `style:num-suffix="."`, `text:start-value` when ≠ 1 —
 * the settled level-style attribute) or `-bullet` (`text:bullet-char`). */
function levelStyleXml(level: number, para: SlideParagraph, start: number | undefined): string {
  const props = listLevelPropsXml(effectiveIndentCm(para), effectiveHangingCm(para));
  if (para.bulletImageId) {
    return (
      `<text:list-level-style-image text:level="${level}" xlink:href="${escapeXmlAttribute(para.bulletImageId)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad">` +
      `${props}</text:list-level-style-image>`
    );
  }
  const style = paragraphMarkerStyle(para) ?? "disc";
  if (para.bullet === "number") {
    const startAttr = start !== undefined && start !== 1 ? ` text:start-value="${start}"` : "";
    return (
      `<text:list-level-style-number text:level="${level}" style:num-format="${NUM_FORMATS[style] ?? "1"}" style:num-suffix="."${startAttr}>` +
      `${props}</text:list-level-style-number>`
    );
  }
  const char = para.bulletChar ?? BULLET_CHARS[style] ?? BULLET_CHARS.disc;
  return `<text:list-level-style-bullet text:level="${level}" text:bullet-char="${escapeXmlAttribute(char)}">${props}</text:list-level-style-bullet>`;
}

interface ListPlan {
  paragraphs: SlideParagraph[];
  /** 0-based level → its marker key, start value and level-style markup. */
  levels: Map<number, { marker: string; start: number | undefined; xml: string }>;
}

/** Splits a run of list paragraphs into `text:list`s, one list style each. A
 * level's marker and start value come from its first paragraph; a later
 * paragraph whose marker differs — or which opens a fresh nested list there
 * (a new numbering run) with a different start — begins a new list. A
 * start-at on an item that only continues a run is the golden-pending
 * restart (textDepthGolden.ts) and is not written (the level's value reads
 * back for it). */
function planLists(run: SlideParagraph[]): ListPlan[] {
  const plans: ListPlan[] = [];
  let current: ListPlan = { paragraphs: [], levels: new Map() };
  let depth = -1;
  for (const para of run) {
    const level = Math.min(Math.max(para.level ?? 0, 0), 8);
    const marker = markerKey(para);
    const numberStart = para.bullet === "number" && !para.bulletImageId ? (para.startAt ?? 1) : undefined;
    const entry = current.levels.get(level);
    const opensRun = level > depth;
    if (entry && (entry.marker !== marker || (opensRun && entry.start !== numberStart))) {
      plans.push(current);
      current = { paragraphs: [], levels: new Map() };
      depth = -1;
    }
    if (!current.levels.has(level)) {
      current.levels.set(level, { marker, start: numberStart, xml: levelStyleXml(level + 1, para, numberStart) });
    }
    current.paragraphs.push(para);
    depth = level;
  }
  plans.push(current);
  return plans;
}

/** A run of list paragraphs as `text:list` markup (SPEC-097): one
 * `text:list-item` per paragraph, a deeper level nested as a `text:list`
 * inside the open item (an empty item bridges a skipped level). */
function listRunXml(run: SlideParagraph[], registry: StyleRegistry): string {
  return planLists(run)
    .map((plan) => {
      const levels = new Map([...plan.levels].map(([level, entry]) => [level + 1, entry.xml] as const));
      const styleName = registry.listStyle(levels);
      let out = `<text:list text:style-name="${styleName}">`;
      let depth = 0;
      const itemOpen: boolean[] = [false];
      for (const para of plan.paragraphs) {
        const level = Math.min(Math.max(para.level ?? 0, 0), 8);
        while (depth > level) {
          if (itemOpen[depth]) out += `</text:list-item>`;
          out += `</text:list>`;
          itemOpen.pop();
          depth -= 1;
        }
        while (depth < level) {
          if (!itemOpen[depth]) {
            out += `<text:list-item>`;
            itemOpen[depth] = true;
          }
          out += `<text:list>`;
          depth += 1;
          itemOpen[depth] = false;
        }
        if (itemOpen[depth]) out += `</text:list-item>`;
        out += `<text:list-item>${paragraphXml(para, registry)}`;
        itemOpen[depth] = true;
      }
      while (depth >= 0) {
        if (itemOpen[depth]) out += `</text:list-item>`;
        out += `</text:list>`;
        depth -= 1;
      }
      return out;
    })
    .join("");
}

function shapeFrameXml(frame: ShapeFrame, registry: StyleRegistry, wireId?: string): string {
  const styleName = registry.graphicStyle({ fill: frame.fill, stroke: frame.stroke, body: frame.body, bodyProps: frame.odpBodyProps });
  const styleAttr = styleName ? ` draw:style-name="${styleName}"` : "";
  const ids = idAttrs(wireId);
  const geom = geomAttrs(frame);
  const text = frame.paragraphs && frame.paragraphs.length > 0
    ? `<draw:text-box>${paragraphsXml(frame.paragraphs, registry, false)}</draw:text-box>`
    : "";
  if (frame.geometry === "rect") {
    return `<draw:rect${ids}${styleAttr}${geom}>${text}</draw:rect>`;
  }
  if (frame.geometry === "ellipse") {
    return `<draw:ellipse${ids}${styleAttr}${geom}>${text}</draw:ellipse>`;
  }
  const drawType = GEOMETRY_TO_DRAW_TYPE[frame.geometry] ?? "non-primitive";
  return (
    `<draw:custom-shape${ids}${styleAttr}${geom}>${text}` +
    `<draw:enhanced-geometry draw:type="${escapeXmlAttribute(drawType)}"${modifiersAttr(frame)}/></draw:custom-shape>`
  );
}

/** A shape's stored adjustments as `draw:modifiers` (SPEC-096): the same
 * 100000ths scale `adjustmentToGuideValue` uses for OOXML guides —
 * deliberately OUR OWN convention, promising a faithful round trip through
 * this reader/writer pair rather than byte-parity with Impress. "" when the
 * shape stores none. */
function modifiersAttr(frame: ShapeFrame): string {
  if (!frame.adjustments || frame.adjustments.length === 0) {
    return "";
  }
  const resolved = resolvedShapeAdjustments(frame.geometry, frame.adjustments);
  if (resolved.length === 0) {
    return "";
  }
  return ` draw:modifiers="${resolved.map((v) => adjustmentToGuideValue(v)).join(" ")}"`;
}

/** Splices generated automatic styles into a content.xml string: before the
 * last `</office:automatic-styles>`, or (if the file lacks the element) a fresh
 * `<office:automatic-styles>…</office:automatic-styles>` before `<office:body`. */
export function spliceAutomaticStyles(content: string, stylesXml: string): string {
  if (stylesXml.length === 0) return content;
  const close = "</office:automatic-styles>";
  const at = content.lastIndexOf(close);
  if (at >= 0) {
    return content.slice(0, at) + stylesXml + content.slice(at);
  }
  return content.replace(/<office:body/, `<office:automatic-styles>${stylesXml}</office:automatic-styles><office:body`);
}
