// Reads Word's real group shape — a `wpg:wgp` inside a `w:drawing`'s
// `a:graphicData` — which SPEC-051 never modelled: its own "group" is a flat
// `groupOrdinal` tag shared by sibling drawings (see shape.ts's DE_NAMESPACE
// note), not a container. So a `wpg:wgp` is FLATTENED here: every member
// becomes an ordinary Shape/Image block in document order, positioned in page space
// and tagged with one fresh `groupOrdinal`, which gives it the model's
// existing group move/resize/rotate/ungroup behaviour with no new block type.
//
// Three consequences worth knowing:
//
//   * `wps:wsp` (shapes, text boxes) and `pic:pic` (pictures) members flatten.
//     A connector, a chart frame or a nested canvas keeps the whole group
//     preserved — see the fall-through at the end of collectMembers.
//
//   * Nesting is lost. A `wpg:grpSp` inside the group flattens into the SAME
//     ordinal rather than a nested one — the flat model cannot express a group
//     inside a group, and one flat group that moves as a whole is closer to
//     the source than dropping the inner grouping's members.
//   * A same-format save writes the members back as N sibling `w:drawing`s
//     carrying `de:group`, not as a `wpg:wgp`. Positions, sizes and the
//     editor's own grouping survive; Word sees N ungrouped drawings. Writing a
//     real `wpg:wgp` back would need a group writer for every group the editor
//     can make, not just imported ones — deliberately out of scope here.
//
// All-or-nothing by design: if any member is markup no reader models, the
// whole group returns null and the caller preserves the original run verbatim
// (SPEC-008). Parsing the members we understand and dropping the rest would
// discard document content, which the preservation rule forbids.
import type { Image, Shape } from "../documentModel";
import type { MediaPart } from "../core/media";
import { firstChildElement, type Relationship } from "../core/xml";
import { emuToCm } from "../core/units";
import {
  mapGroupChild,
  parseGroupChildMapping,
  parseXfrm,
  type GroupChildMapping,
} from "../core/drawingMl";
import { blipMediaResolver, imageFromPic } from "./drawing";
import { shapeFromWsp, type ShapeTextContext } from "./shape";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WPS_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const WPG_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

/** One flattened `wpg:wgp`: its members as ordinary drawing blocks in
 * document order, plus the media parts its member pictures brought with them.
 * The caller stamps every block with the shared `groupOrdinal`. */
export interface ParsedGroupDrawing {
  blocks: (Shape | Image)[];
  media: MediaPart[];
}

/** What a group member needs beyond its own markup: the paragraph parser a
 * text box's contents are read through, and the package lookups a member
 * picture's blip resolves against. */
export interface GroupDrawingContext extends ShapeTextContext {
  mainRels: Map<string, Relationship>;
  mainPartPath: string;
  parts: Map<string, Uint8Array>;
}

/** The group's own placement, read from its `wp:anchor`/`wp:inline` and shared
 * by every member: a member has no container of its own. */
interface GroupPlacement {
  wrap: Shape["wrap"];
  offsetXCm: number;
  offsetYCm: number;
  zOrder?: number;
}

/** A member box straight off its `a:xfrm`, in its parent group's child
 * coordinate space. */
interface ChildXfrm {
  xEmu: number;
  yEmu: number;
  cxEmu: number;
  cyEmu: number;
}

/** Maps child coordinates through unchanged — the mapping a group with no
 * usable `a:xfrm`/`a:chExt` gets, which is what an unscaled group means. */
const IDENTITY_MAPPING: GroupChildMapping = {
  xEmu: 0,
  yEmu: 0,
  cxEmu: 1,
  cyEmu: 1,
  chXEmu: 0,
  chYEmu: 0,
  chCxEmu: 1,
  chCyEmu: 1,
};

function childXfrmOf(spPr: Element | undefined): ChildXfrm | null {
  const xfrm = spPr ? parseXfrm(firstChildElement(spPr, A_NAMESPACE, "xfrm")) : undefined;
  if (
    !xfrm ||
    xfrm.xEmu === undefined ||
    xfrm.yEmu === undefined ||
    xfrm.cxEmu === undefined ||
    xfrm.cyEmu === undefined
  ) {
    return null;
  }
  return { xEmu: xfrm.xEmu, yEmu: xfrm.yEmu, cxEmu: xfrm.cxEmu, cyEmu: xfrm.cyEmu };
}

/** The outermost group's mapping, rebased so the group's own top-left is the
 * origin: mapping a member through it then yields the member's offset *within*
 * the group box, which is exactly what a page-space offset is built from. */
function rootedMapping(grpSpPr: Element | undefined): GroupChildMapping {
  const mapping = grpSpPr ? parseGroupChildMapping(grpSpPr) : undefined;
  return mapping ? { ...mapping, xEmu: 0, yEmu: 0 } : IDENTITY_MAPPING;
}

/** Composes a nested group's mapping with its parent's into one mapping from
 * the nested group's child coordinates straight to the outer group's box —
 * the flat equivalent of the pptx reader's recursive remap. Both the nested
 * group's origin and its scale are themselves expressed in the outer group's
 * child space, so both are carried through the outer scale. */
function composeMapping(outer: GroupChildMapping, nested: GroupChildMapping): GroupChildMapping {
  const sx = outer.cxEmu / outer.chCxEmu;
  const sy = outer.cyEmu / outer.chCyEmu;
  return {
    xEmu: outer.xEmu + (nested.xEmu - outer.chXEmu) * sx,
    yEmu: outer.yEmu + (nested.yEmu - outer.chYEmu) * sy,
    cxEmu: nested.cxEmu * sx,
    cyEmu: nested.cyEmu * sy,
    chXEmu: nested.chXEmu,
    chYEmu: nested.chYEmu,
    chCxEmu: nested.chCxEmu,
    chCyEmu: nested.chCyEmu,
  };
}

interface MemberBox {
  offsetXCm: number;
  offsetYCm: number;
  widthCm: number;
  heightCm: number;
}

/** Places one member in page space: through the group mapping to get its
 * offset within the group box, then out from the group's own page offset. */
function placeMember(child: ChildXfrm, mapping: GroupChildMapping, group: GroupPlacement): MemberBox {
  const mapped = mapGroupChild(mapping, child.xEmu, child.yEmu, child.cxEmu, child.cyEmu);
  return {
    offsetXCm: group.offsetXCm + emuToCm(mapped.xEmu),
    offsetYCm: group.offsetYCm + emuToCm(mapped.yEmu),
    widthCm: emuToCm(mapped.cxEmu),
    heightCm: emuToCm(mapped.cyEmu),
  };
}

/** Applies the placement every flattened member gets: the group's wrap and
 * z-order, and its own mapped box. The `groupOrdinal` that keeps the members
 * moving together is stamped by the caller, which owns the document-wide
 * counter — parsing must stay free of side effects so the reader can call it
 * twice (once to detect a group run, once to build its blocks) without
 * burning an ordinal on the first pass. */
function placeBlock(block: Shape | Image, box: MemberBox, group: GroupPlacement): void {
  block.widthCm = box.widthCm;
  block.heightCm = box.heightCm;
  block.wrap = group.wrap;
  block.offsetXCm = box.offsetXCm;
  block.offsetYCm = box.offsetYCm;
  block.relativeToH = "column";
  block.relativeToV = "paragraph";
  if (group.zOrder !== undefined) {
    // Every member takes the group's own z-order; among themselves they stack
    // by document order, which is the order they are emitted in.
    block.zOrder = group.zOrder;
  }
}

/** Flattens one `wpg:wgp` (or a nested `wpg:grpSp`, same content model) into
 * member blocks, appending to `out`. Returns false the moment a child is
 * markup no reader models, which aborts the whole group. */
function collectMembers(
  wgp: Element,
  mapping: GroupChildMapping,
  group: GroupPlacement,
  ctx: GroupDrawingContext,
  out: ParsedGroupDrawing,
): boolean {
  for (const child of Array.from(wgp.children)) {
    if (child.namespaceURI === WPG_NAMESPACE) {
      if (child.localName === "cNvGrpSpPr" || child.localName === "grpSpPr") {
        continue; // The group's own properties, already read by the caller.
      }
      if (child.localName !== "grpSp") {
        return false;
      }
      const nestedGrpSpPr = firstChildElement(child, WPG_NAMESPACE, "grpSpPr");
      const nested = nestedGrpSpPr ? parseGroupChildMapping(nestedGrpSpPr) : undefined;
      if (!nested) {
        // No usable child mapping means its members cannot be placed; keeping
        // the whole group verbatim beats guessing where they go.
        return false;
      }
      if (!collectMembers(child, composeMapping(mapping, nested), group, ctx, out)) {
        return false;
      }
      continue;
    }

    if (child.namespaceURI === WPS_NAMESPACE && child.localName === "wsp") {
      const childXfrm = childXfrmOf(firstChildElement(child, WPS_NAMESPACE, "spPr"));
      if (!childXfrm) {
        return false;
      }
      const box = placeMember(childXfrm, mapping, group);
      const shape = shapeFromWsp(child, { widthCm: box.widthCm, heightCm: box.heightCm, wrap: group.wrap }, ctx);
      if (!shape) {
        return false;
      }
      // A member's alt text is on its own `wps:cNvPr` — a top-level shape's
      // lives on the container's `wp:docPr`, which a member does not have.
      const altText = firstChildElement(child, WPS_NAMESPACE, "cNvPr")?.getAttribute("descr") || undefined;
      if (altText) {
        shape.altText = altText;
      }
      placeBlock(shape, box, group);
      out.blocks.push(shape);
      continue;
    }

    // A `pic:pic` member — the picture in Word's most common group. This was
    // the one member kind that took its whole group down with it, because an
    // Image had no position anywhere in the model; it flattens now that
    // Image carries the same placement Shape/Chart/Connector do.
    if (child.namespaceURI === PIC_NAMESPACE && child.localName === "pic") {
      const childXfrm = childXfrmOf(firstChildElement(child, PIC_NAMESPACE, "spPr"));
      if (!childXfrm) {
        return false;
      }
      const box = placeMember(childXfrm, mapping, group);
      const parsed = imageFromPic(
        child,
        { widthCm: box.widthCm, heightCm: box.heightCm, wrap: group.wrap },
        { mainRels: ctx.mainRels, mainPartPath: ctx.mainPartPath, parts: ctx.parts },
      );
      if (!parsed) {
        return false;
      }
      // A member picture's alt text is on its own `pic:cNvPr`, the twin of a
      // shape member's `wps:cNvPr`.
      const nvPicPr = firstChildElement(child, PIC_NAMESPACE, "nvPicPr");
      const cNvPr = nvPicPr ? firstChildElement(nvPicPr, PIC_NAMESPACE, "cNvPr") : undefined;
      const altText = cNvPr?.getAttribute("descr") || undefined;
      if (altText) {
        parsed.image.altText = altText;
      }
      placeBlock(parsed.image, box, group);
      out.blocks.push(parsed.image);
      out.media.push(parsed.media);
      continue;
    }

    // A `wps:cxnSp` connector, a `wpg:graphicFrame` chart/table, an
    // `mc:AlternateContent` nested inside the group — none has a group-member
    // reader yet, so the whole group stays preserved verbatim.
    return false;
  }
  return true;
}

/** The group's own wrap/position/z-order, read from its container the way
 * shape.ts reads a top-level shape's. Null for a wrap mode the model does not
 * carry, which keeps the group preserved verbatim.
 *
 * A `wp:inline` group becomes `"square"`-wrapped members at the origin: the
 * model's grouping is floating-only (shapes.ts's groupShapes converts an
 * inline member the same way), and members that kept `"inline"` would each
 * become a block in the text flow instead of one positioned group. */
function groupPlacement(container: Element, isAnchor: boolean): GroupPlacement | null {
  if (!isAnchor) {
    return { wrap: "square", offsetXCm: 0, offsetYCm: 0 };
  }
  let wrap: Shape["wrap"];
  if (firstChildElement(container, WP_NAMESPACE, "wrapSquare")) {
    wrap = "square";
  } else if (firstChildElement(container, WP_NAMESPACE, "wrapTopAndBottom")) {
    wrap = "topAndBottom";
  } else if (firstChildElement(container, WP_NAMESPACE, "wrapNone")) {
    wrap = container.getAttribute("behindDoc") === "1" ? "behind" : "infront";
  } else {
    return null;
  }
  const placement: GroupPlacement = {
    wrap,
    offsetXCm: posOffsetCm(container, "positionH") ?? 0,
    offsetYCm: posOffsetCm(container, "positionV") ?? 0,
  };
  const relativeHeight = container.getAttribute("relativeHeight");
  if (relativeHeight) {
    placement.zOrder = Number(relativeHeight) - 251658240;
  }
  return placement;
}

function posOffsetCm(container: Element, positionName: "positionH" | "positionV"): number | undefined {
  const position = firstChildElement(container, WP_NAMESPACE, positionName);
  const posOffset = position ? firstChildElement(position, WP_NAMESPACE, "posOffset") : undefined;
  if (!posOffset?.textContent) {
    return undefined;
  }
  const emu = Number(posOffset.textContent);
  return Number.isFinite(emu) ? emuToCm(emu) : undefined;
}

/** Parses a `w:drawing` holding a `wpg:wgp` into its flattened members, or
 * null when the drawing is not a group, or is one holding markup no member
 * reader models — the caller then preserves the whole run verbatim. */
export function parseGroupDrawing(drawingEl: Element, ctx: GroupDrawingContext): ParsedGroupDrawing | null {
  const containers = Array.from(drawingEl.children).filter((el) => el.namespaceURI === WP_NAMESPACE);
  if (containers.length !== 1) {
    return null;
  }
  const container = containers[0];
  const isAnchor = container.localName === "anchor";
  if (!isAnchor && container.localName !== "inline") {
    return null;
  }

  const graphic = firstChildElement(container, A_NAMESPACE, "graphic");
  const graphicData = graphic ? firstChildElement(graphic, A_NAMESPACE, "graphicData") : undefined;
  const wgp = graphicData ? firstChildElement(graphicData, WPG_NAMESPACE, "wgp") : undefined;
  if (!wgp) {
    return null;
  }

  const placement = groupPlacement(container, isAnchor);
  if (!placement) {
    return null;
  }

  const parsed: ParsedGroupDrawing = { blocks: [], media: [] };
  // SPEC-073: member picture fills resolve through the same media walk as
  // member pictures, deduplicated into this group's own media list.
  const memberCtx: GroupDrawingContext = ctx.resolveBlip
    ? ctx
    : { ...ctx, resolveBlip: blipMediaResolver(ctx, parsed.media) };
  const mapping = rootedMapping(firstChildElement(wgp, WPG_NAMESPACE, "grpSpPr"));
  if (!collectMembers(wgp, mapping, placement, memberCtx, parsed)) {
    return null;
  }
  // An empty group carries nothing to render and nothing to group; leaving it
  // to the opaque path keeps the source markup rather than erasing it.
  if (parsed.blocks.length === 0) {
    return null;
  }
  return parsed;
}
