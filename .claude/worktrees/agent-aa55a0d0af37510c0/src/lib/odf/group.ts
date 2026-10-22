// Reads LibreOffice's group shape — a `draw:g` holding several drawings —
// which SPEC-051 never modelled on either side: the model's own "group" is a
// flat `groupOrdinal` shared by sibling drawings, not a container. So a
// `draw:g` is FLATTENED here, exactly the way docx/group.ts flattens a
// `wpg:wgp`: every member becomes an ordinary Shape block in document order,
// tagged with one fresh ordinal, which gives it the model's existing group
// move/resize/rotate/ungroup behaviour with no new block type.
//
// ODF makes the coordinates easier than OOXML does. A `draw:g` has no child
// coordinate space of its own — no `chOff`/`chExt` twin — so a member's
// `svg:x`/`svg:y` are already in the group's own frame and need no mapping,
// only the group's optional `draw:transform` translation added on top.
//
// Same three consequences as the .docx side, and for the same reasons (see
// docx/group.ts's header): nesting collapses into ONE ordinal, a same-format
// save writes the members back as sibling drawings rather than a `draw:g`,
// and a member no reader models aborts the whole group so the caller can
// preserve it verbatim (SPEC-008) rather than dropping content.
import type { Image, ImageWrapMode, Shape } from "../documentModel";
import type { MediaPart } from "../core/media";
import { DRAW_NAMESPACE, SVG_NAMESPACE, TEXT_NAMESPACE } from "../core/odfXml";
import { odfLengthToCm } from "./units";
import { parseDrawFrame } from "./drawing";
import { parseShapeElement, type OdfShapeTextContext } from "./shape";

// The same editor-only extension namespace shape.ts/connector.ts use to
// round-trip `groupOrdinal`.
const DE_NAMESPACE = "urn:docu-edit:shape";

/** One flattened `draw:g`: its members as ordinary drawing blocks in document
 * order, plus the media parts its member pictures brought with them. The
 * caller stamps every block with the shared `groupOrdinal`. */
export interface ParsedOdfGroup {
  blocks: (Shape | Image)[];
  media: MediaPart[];
}

/** What a group member needs beyond its own markup — the same documents and
 * paragraph parser `parseShapeElement` takes. */
export interface OdfGroupContext {
  contentDoc: Document;
  stylesDoc: Document | undefined;
  text: OdfShapeTextContext;
  /** The package parts and wrap-style table a member `draw:frame` picture
   * resolves against — the same pair parseDrawFrame takes. */
  parts: Map<string, Uint8Array>;
  graphicWrapStyles: Map<string, ImageWrapMode>;
}

/** The group's own placement, shared by every member: a member inside a
 * `draw:g` carries no anchor of its own. */
interface GroupPlacement {
  wrap: Shape["wrap"];
  /** The group's `draw:transform` translation, in cm — zero when it has none. */
  translateXCm: number;
  translateYCm: number;
}

function drawAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(DRAW_NAMESPACE, name);
}

function svgAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(SVG_NAMESPACE, name);
}

/** A `draw:transform` that is a single `translate(x y)` — the only transform
 * form a flat model can absorb, since it folds into the members' offsets.
 * Returns null for any other transform (a rotate, a skew, a matrix, a chain),
 * which aborts the group so it stays preserved rather than being drawn at
 * coordinates the transform would have moved.
 *
 * Absent `draw:transform` is not a failure — it is the common case, and
 * yields a zero translation. */
function parseTranslate(transform: string | null): { xCm: number; yCm: number } | null {
  if (transform === null || transform.trim().length === 0) {
    return { xCm: 0, yCm: 0 };
  }
  const match = /^\s*translate\s*\(\s*([^\s,()]+)\s*[, ]\s*([^\s,()]+)\s*\)\s*$/.exec(transform);
  if (!match) {
    return null;
  }
  const xCm = odfLengthToCm(match[1]);
  const yCm = odfLengthToCm(match[2]);
  if (xCm === undefined || yCm === undefined) {
    return null;
  }
  return { xCm, yCm };
}

/** A member's own position inside the group, in cm. A `draw:line` carries
 * `svg:x1`/`svg:y1` instead of `svg:x`/`svg:y`, matching how
 * `parseShapeElement` derives a line's own box origin; a `draw:frame`
 * picture uses the ordinary pair. */
function memberOffsetCm(el: Element): { xCm: number; yCm: number } | null {
  const isLine = el.localName === "line";
  const xCm = odfLengthToCm(svgAttr(el, isLine ? "x1" : "x"));
  const yCm = odfLengthToCm(svgAttr(el, isLine ? "y1" : "y"));
  if (xCm === undefined || yCm === undefined) {
    return null;
  }
  return { xCm, yCm };
}

/** Flattens one `draw:g` (or a nested one, same content model) into member
 * blocks, appending to `out`. Returns false the moment a child is markup no
 * reader models, which aborts the whole group. */
function collectMembers(
  group: Element,
  placement: GroupPlacement,
  ctx: OdfGroupContext,
  out: ParsedOdfGroup,
): boolean {
  for (const child of Array.from(group.children)) {
    if (child.namespaceURI !== DRAW_NAMESPACE) {
      // `svg:title`/`svg:desc` describe the group itself and carry no
      // geometry; anything else is markup with no member reader.
      if (child.namespaceURI === SVG_NAMESPACE && (child.localName === "desc" || child.localName === "title")) {
        continue;
      }
      return false;
    }

    if (child.localName === "g") {
      // A nested group: its own translation composes with this one's, and its
      // members join the SAME flat ordinal — the model cannot express a group
      // inside a group, and one group that moves as a whole is closer to the
      // source than dropping the inner grouping.
      const nested = parseTranslate(drawAttr(child, "transform"));
      if (!nested) {
        return false;
      }
      const nestedPlacement: GroupPlacement = {
        wrap: placement.wrap,
        translateXCm: placement.translateXCm + nested.xCm,
        translateYCm: placement.translateYCm + nested.yCm,
      };
      if (!collectMembers(child, nestedPlacement, ctx, out)) {
        return false;
      }
      continue;
    }

    const offset = memberOffsetCm(child);
    if (!offset) {
      return false;
    }

    // A `draw:frame` picture member. This took its whole group down with it
    // until Image gained real placement — an image had no position anywhere
    // in the model, so a flattened one would have painted at the group's
    // origin. A frame holding a chart or a formula object rather than an
    // image yields null from parseDrawFrame and still aborts the group.
    if (child.localName === "frame") {
      const parsedFrame = parseDrawFrame(child, ctx.parts, ctx.graphicWrapStyles);
      if (!parsedFrame) {
        return false;
      }
      placeMemberBlock(parsedFrame.image, offset, placement);
      out.blocks.push(parsedFrame.image);
      out.media.push(parsedFrame.media);
      continue;
    }

    if (child.localName !== "custom-shape" && child.localName !== "line") {
      // A `draw:connector`, a nested object frame: no group-member reader, so
      // the whole group stays preserved.
      return false;
    }

    const parsed = parseShapeElement(child, ctx.contentDoc, ctx.stylesDoc, ctx.text);
    if (!parsed) {
      return false;
    }
    placeMemberBlock(parsed.shape, offset, placement);
    out.blocks.push(parsed.shape);
  }
  return true;
}

/** Applies the placement every flattened member gets. A member's own graphic
 * style has no anchor, so parseShapeElement/parseDrawFrame read it as inline
 * and skip its offsets; the group owns the placement, so it is written here
 * from the member's own coordinates plus the group's translation.
 * `groupOrdinal` is stamped by the caller, which owns the document-wide
 * counter. */
function placeMemberBlock(
  block: Shape | Image,
  offset: { xCm: number; yCm: number },
  placement: GroupPlacement,
): void {
  block.wrap = placement.wrap;
  block.offsetXCm = placement.translateXCm + offset.xCm;
  block.offsetYCm = placement.translateYCm + offset.yCm;
  block.relativeToH = "column";
  block.relativeToV = "paragraph";
}

/** Parses a `draw:g` into its flattened members, or null when it holds markup
 * no member reader models — the caller then preserves it verbatim.
 *
 * A group anchored `as-char` becomes `"square"`-wrapped members: the model's
 * grouping is floating-only (shapes.ts's groupShapes converts an inline
 * member the same way), and members that stayed inline would each land in the
 * text flow instead of one positioned group. */
export function parseDrawGroup(el: Element, ctx: OdfGroupContext): ParsedOdfGroup | null {
  if (el.namespaceURI !== DRAW_NAMESPACE || el.localName !== "g") {
    return null;
  }
  const translate = parseTranslate(drawAttr(el, "transform"));
  if (!translate) {
    return null;
  }
  const anchorType = el.getAttributeNS(TEXT_NAMESPACE, "anchor-type");
  const placement: GroupPlacement = {
    wrap: anchorType === null || anchorType === "as-char" ? "square" : "infront",
    translateXCm: translate.xCm,
    translateYCm: translate.yCm,
  };
  const parsed: ParsedOdfGroup = { blocks: [], media: [] };
  if (!collectMembers(el, placement, ctx, parsed)) {
    return null;
  }
  // An empty group renders nothing and groups nothing; leaving it to the
  // opaque path keeps the source markup rather than erasing it.
  if (parsed.blocks.length === 0) {
    return null;
  }
  return parsed;
}

/** One more than the largest `de:group` ordinal already in `content.xml` —
 * the seed for ordinals minted for imported `draw:g` groups, so a group
 * LibreOffice added to a document this editor had saved cannot land on an
 * ordinal that document already uses. The .odt twin of docx/reader.ts's
 * firstFreeGroupOrdinal. */
export function firstFreeOdfGroupOrdinal(root: Element): number {
  let max = 0;
  for (const localName of ["custom-shape", "line", "frame", "connector"]) {
    const nodes = root.getElementsByTagNameNS(DRAW_NAMESPACE, localName);
    for (let i = 0; i < nodes.length; i++) {
      const value = Number(nodes[i].getAttributeNS(DE_NAMESPACE, "group"));
      if (Number.isFinite(value)) {
        max = Math.max(max, value);
      }
    }
  }
  return max + 1;
}
