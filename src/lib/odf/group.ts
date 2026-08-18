// Writes a `DrawingGroup` container (SPEC-057) as a `draw:g` holding its
// members as child elements. `draw:g` defines no geometry of its own — its
// children are positioned in the anchor's coordinate space directly — so each
// member is emitted at `container offset + member offset`. Rotation is baked
// first (ODF `draw:g` has no rotation attribute; LibreOffice rewrites the
// children, and so do we — see drawingGroups.ts's bakeGroupRotationDeep). A
// sealed container re-emits its preserved `draw:g` subtree verbatim instead.
// Image and chart members are not yet serialised as group children.
import { isDrawingGroup, isShape, type DrawingGroup, type DrawingMember, type Shape } from "../documentModel";
import { bakeGroupRotationDeep } from "../drawingGroups";
import { escapeXmlAttribute } from "../core/xml";
import { shapeElementXml, type OdfShapeTextWriteContext } from "./shape";

const DRAW_G_XMLNS =
  ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';

/** Registers/looks up the styles group members need (member shapes carry
 * their own graphic style). Injected so this module stays independent of the
 * writer's AutomaticStyles concrete type. */
export interface OdfGroupWriteContext extends OdfShapeTextWriteContext {
  shapeStyleNameFor: (shape: Shape) => string;
  graphicStyleNameFor: (wrap: DrawingGroup["wrap"]) => string | undefined;
}

/** A shape member shifted to its absolute position in the anchor frame. */
function shiftedShape(shape: Shape, baseXCm: number, baseYCm: number): Shape {
  return {
    ...shape,
    wrap: "square",
    offsetXCm: baseXCm + (shape.offsetXCm ?? 0),
    offsetYCm: baseYCm + (shape.offsetYCm ?? 0),
  };
}

/** The child elements of a container (already rotation-baked), positioned
 * absolutely at `baseX/baseY + member offset`. */
function membersXml(members: DrawingMember[], baseXCm: number, baseYCm: number, ctx: OdfGroupWriteContext): string {
  return members
    .map((member) => {
      if (isShape(member)) {
        const shifted = shiftedShape(member, baseXCm, baseYCm);
        return shapeElementXml(shifted, ctx.shapeStyleNameFor(shifted), ctx);
      }
      if (isDrawingGroup(member)) {
        return groupElementXml(member, baseXCm + (member.offsetXCm ?? 0), baseYCm + (member.offsetYCm ?? 0), ctx);
      }
      // Image and chart members are not yet serialised as group children.
      return "";
    })
    .join("");
}

/** A `draw:g` element (nested or top-level), with its members positioned at
 * `baseX/baseY`. Assumes rotation is already baked. */
function groupElementXml(group: DrawingGroup, baseXCm: number, baseYCm: number, ctx: OdfGroupWriteContext): string {
  const titleXml = group.altText ? `<svg:title>${escapeXmlAttribute(group.altText)}</svg:title><svg:desc>${escapeXmlAttribute(group.altText)}</svg:desc>` : "";
  const zAttr = group.zOrder !== undefined ? ` draw:z-index="${group.zOrder}"` : "";
  const styleName = ctx.graphicStyleNameFor(group.wrap);
  const styleAttr = styleName !== undefined ? ` draw:style-name="${escapeXmlAttribute(styleName)}"` : "";
  const anchorType = group.wrap === "inline" ? "as-char" : "paragraph";
  return (
    `<draw:g${DRAW_G_XMLNS} text:anchor-type="${anchorType}"${styleAttr}${zAttr}>` +
    titleXml +
    membersXml(group.members, baseXCm, baseYCm, ctx) +
    "</draw:g>"
  );
}

/** Serialises a `DrawingGroup` into a `draw:g` element (wrapped by the caller
 * in `text:p`), members at `container offset + member offset` (SPEC-057
 * AC-63). Rotation is baked first (AC-64) on a serialization-time copy — the
 * open model is never mutated. A sealed container re-emits its preserved
 * subtree verbatim. */
export function groupDrawXml(group: DrawingGroup, ctx: OdfGroupWriteContext): string {
  if (group.preservedSubtreeXml !== undefined) {
    return group.preservedSubtreeXml;
  }
  const baked = bakeGroupRotationDeep(group);
  return groupElementXml(baked, baked.offsetXCm ?? 0, baked.offsetYCm ?? 0, ctx);
}
