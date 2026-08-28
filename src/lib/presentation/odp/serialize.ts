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
// class (its outline style on the master page), so disc bullets on placeholders
// render without explicit list markup; explicit-list/numbered authoring rides a
// future lists pass (SPEC-081 out of scope).
//
// Framework-free; imports only presentation siblings and src/lib/core.
import { GEOMETRY_TO_DRAW_TYPE } from "../../core/drawingMl";
import { escapeXmlAttribute, escapeXmlText } from "../../core/xml";
import type {
  FrameStroke,
  ImageFrame,
  PlaceholderRole,
  ShapeFrame,
  Slide,
  SlideFill,
  SlideParagraph,
  SlideRun,
  TextFrame,
} from "../model";
import { odpTransitionProps } from "../transitions";
import { writeOdpAnim } from "./anim";

const ROLE_TO_CLASS: Record<PlaceholderRole, string> = {
  title: "title",
  subtitle: "subtitle",
  body: "body",
  other: "outline",
};

const LEVEL_INDENT_CM = 1.27;

function alignValue(align: NonNullable<SlideParagraph["align"]>): string {
  if (align === "center") return "center";
  if (align === "right") return "end";
  if (align === "justify") return "justify";
  return "start";
}

/** Collects the automatic styles a page rebuild generates, de-duplicated by
 * their property signature and named in the reserved `de-` namespace. */
export class StyleRegistry {
  private readonly paraStyles: string[] = [];
  private readonly textStyles: string[] = [];
  private readonly graphicStyles: string[] = [];
  private readonly drawingPageStyles: string[] = [];
  private readonly paraCache = new Map<string, string>();
  private readonly textCache = new Map<string, string>();
  private readonly graphicCache = new Map<string, string>();
  private readonly drawingPageCache = new Map<string, string>();
  private pCount = 0;
  private tCount = 0;
  private gCount = 0;
  private dpCount = 0;

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
  }

  paragraphStyle(para: SlideParagraph): string | undefined {
    const props: string[] = [];
    if (para.align) props.push(`fo:text-align="${alignValue(para.align)}"`);
    if (para.level) props.push(`fo:margin-left="${(para.level * LEVEL_INDENT_CM).toFixed(3)}cm"`);
    if (props.length === 0) return undefined;
    const key = props.join(" ");
    const cached = this.paraCache.get(key);
    if (cached) return cached;
    const name = `de-P${++this.pCount}`;
    this.paraCache.set(key, name);
    this.paraStyles.push(
      `<style:style style:name="${name}" style:family="paragraph"><style:paragraph-properties ${key}/></style:style>`,
    );
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
      const f = escapeXmlAttribute(run.fontFamily);
      props.push(`style:font-name="${f}" fo:font-family="${f}"`);
    }
    if (run.sizePt !== undefined) props.push(`fo:font-size="${run.sizePt}pt"`);
    if (run.color) props.push(`fo:color="${escapeXmlAttribute(run.color)}"`);
    if (props.length === 0) return undefined;
    const key = props.join(" ");
    const cached = this.textCache.get(key);
    if (cached) return cached;
    const name = `de-T${++this.tCount}`;
    this.textCache.set(key, name);
    this.textStyles.push(
      `<style:style style:name="${name}" style:family="text"><style:text-properties ${key}/></style:style>`,
    );
    return name;
  }

  graphicStyle(style: { fill?: SlideFill; stroke?: FrameStroke }): string | undefined {
    const props: string[] = [];
    if (style.fill?.type === "solid") props.push(`draw:fill="solid" draw:fill-color="${escapeXmlAttribute(style.fill.color)}"`);
    else if (style.fill?.type === "none") props.push(`draw:fill="none"`);
    if (style.stroke) {
      props.push(
        `draw:stroke="solid" svg:stroke-color="${escapeXmlAttribute(style.stroke.color)}" svg:stroke-width="${(style.stroke.widthPt / 72 * 2.54).toFixed(3)}cm"`,
      );
    }
    if (props.length === 0) return undefined;
    const key = props.join(" ");
    const cached = this.graphicCache.get(key);
    if (cached) return cached;
    const name = `de-G${++this.gCount}`;
    this.graphicCache.set(key, name);
    this.graphicStyles.push(
      `<style:style style:name="${name}" style:family="graphic"><style:graphic-properties ${key}/></style:style>`,
    );
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

  hasStyles(): boolean {
    return (
      this.paraStyles.length + this.textStyles.length + this.graphicStyles.length + this.drawingPageStyles.length > 0
    );
  }

  /** The generated `<style:style>` elements, to splice before
   * `</office:automatic-styles>`. */
  xml(): string {
    return [...this.paraStyles, ...this.textStyles, ...this.graphicStyles, ...this.drawingPageStyles].join("");
  }
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
  // A modelled transition mints a fresh drawing-page style (copying the
  // original style's fill/background) and repoints `draw:style-name` at it, so
  // the replaced original's smil attributes never re-emit (SPEC-085 AC-11).
  if (slide.transition) {
    const originalStyle = drawAttr(attrs, "draw:style-name");
    const baseProps = originalStyle ? drawingPageBaseProps(styleSource, originalStyle) : "";
    const styleName = registry.drawingPageStyle(odpTransitionProps(slide.transition), baseProps);
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
      if (frame.kind === "text" || frame.kind === "shape" || frame.kind === "image") {
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
      name === "presentation:duration"
    ) {
      continue; // drop the transition being replaced
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

function frameXml(frame: Slide["frames"][number], registry: StyleRegistry, wireId?: string): string {
  switch (frame.kind) {
    case "text":
      return textFrameXml(frame, registry, wireId);
    case "shape":
      return shapeFrameXml(frame, registry, wireId);
    case "image":
      return imageFrameXml(frame, registry, wireId);
    case "group":
    case "foreign":
      return frame.xml; // verbatim (AC-39)
    default:
      return "";
  }
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
  const styleName = registry.graphicStyle({ fill: frame.fill, stroke: frame.stroke });
  const styleAttr = styleName ? ` draw:style-name="${styleName}"` : "";
  const body = frame.paragraphs.map((para) => paragraphXml(para, registry)).join("");
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

function paragraphXml(para: SlideParagraph, registry: StyleRegistry): string {
  const styleName = registry.paragraphStyle(para);
  const styleAttr = styleName ? ` text:style-name="${styleName}"` : "";
  const inner = para.runs.map((run) => runXml(run, registry)).join("");
  return `<text:p${styleAttr}>${inner}</text:p>`;
}

function runXml(run: SlideRun, registry: StyleRegistry): string {
  const text = escapeXmlText(run.text);
  const styleName = registry.textStyle(run);
  return styleName ? `<text:span text:style-name="${styleName}">${text}</text:span>` : text;
}

function shapeFrameXml(frame: ShapeFrame, registry: StyleRegistry, wireId?: string): string {
  const styleName = registry.graphicStyle({ fill: frame.fill, stroke: frame.stroke });
  const styleAttr = styleName ? ` draw:style-name="${styleName}"` : "";
  const ids = idAttrs(wireId);
  const geom = geomAttrs(frame);
  const text = frame.paragraphs && frame.paragraphs.length > 0
    ? `<draw:text-box>${frame.paragraphs.map((p) => paragraphXml(p, registry)).join("")}</draw:text-box>`
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
    `<draw:enhanced-geometry draw:type="${escapeXmlAttribute(drawType)}"/></draw:custom-shape>`
  );
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
