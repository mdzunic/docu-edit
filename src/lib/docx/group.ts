// Writes a `DrawingGroup` container (SPEC-057) as a `wpg:wgp` inside a
// `wp:anchor`/`wp:inline`, the same anchor a single shape uses (drawing.ts /
// shape.ts). An unsealed container writes an **identity child space**
// (`chOff` 0,0, `chExt = ext`), members as children in member order at their
// offsets, `@rot` when rotated. A sealed container re-emits its preserved
// `wpg:wgp` subtree verbatim instead — position/anchor patches are applied by
// the caller's same-format preservation path. Charts as members are not yet
// serialised here (they need per-member chart-part planning); such a member is
// skipped, which the group writer logs is a known gap.
import {
  isChart,
  isDrawingGroup,
  isImage,
  isShape,
  type DrawingGroup,
  type DrawingMember,
} from "../documentModel";
import { cmToEmu } from "../core/units";
import { escapeXmlAttribute } from "../core/xml";
import { shapeGroupChildXml, type ShapeTextWriteContext } from "./shape";
import { imageGroupChildXml } from "./drawing";
import { chartGroupChildXml } from "./chart";
import { EMPTY_CHART_PLAN, type ChartWritePlan } from "./chartWritePlan";
import type { MediaWritePlan } from "./media";

const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPG_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

function degToOoxmlRot(deg: number): number {
  return Math.round(deg * 60000);
}

/** A monotonic document-unique id allocator — every `cNvPr` in a group tree
 * gets a fresh id at write time (SPEC-057: ids are never modelled). */
export interface DocPrIdAllocator {
  next(): number;
}

/** Where in the document a container sits, so a Chart member's chart-part
 * plan entry can be looked up by (section, block, member path). */
interface GroupLoc {
  chartPlan: ChartWritePlan;
  section: number;
  block: number;
  path: number[];
}

/** The members of a container as `wpg:wgp` child elements, in member order. */
function membersXml(members: DrawingMember[], ctx: ShapeTextWriteContext, mediaPlan: MediaWritePlan, ids: DocPrIdAllocator, loc: GroupLoc): string {
  return members
    .map((member, index) => {
      const offX = cmToEmu(member.offsetXCm ?? 0);
      const offY = cmToEmu(member.offsetYCm ?? 0);
      if (isShape(member)) {
        return shapeGroupChildXml(member, ids.next(), offX, offY, ctx);
      }
      if (isImage(member)) {
        const relId = mediaPlan.relIdFor(member.mediaId);
        return relId ? imageGroupChildXml(member, relId, ids.next(), offX, offY) : "";
      }
      if (isChart(member)) {
        // A chart member (SPEC-057 AC-55/AC-59): the write plan allocated a
        // chart part + relationship for it, keyed by its member path.
        const entry = loc.chartPlan.planForMember(loc.section, loc.block, [...loc.path, index]);
        return entry ? chartGroupChildXml(member, entry.relId, ids.next(), offX, offY) : "";
      }
      if (isDrawingGroup(member)) {
        return nestedGroupXml(member, ctx, mediaPlan, ids, { ...loc, path: [...loc.path, index] });
      }
      return "";
    })
    .join("");
}

/** The `a:xfrm` for a group/nested-group with an identity child space. */
function groupXfrmXml(group: DrawingGroup, offXEmu: number, offYEmu: number): string {
  const cx = cmToEmu(group.widthCm);
  const cy = cmToEmu(group.heightCm);
  const rotAttr = group.rotationDeg ? ` rot="${degToOoxmlRot(group.rotationDeg)}"` : "";
  return (
    `<a:xfrm${rotAttr}>` +
    `<a:off x="${offXEmu}" y="${offYEmu}"/><a:ext cx="${cx}" cy="${cy}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/>` +
    "</a:xfrm>"
  );
}

/** A nested container as a `wpg:grpSp` child. */
function nestedGroupXml(group: DrawingGroup, ctx: ShapeTextWriteContext, mediaPlan: MediaWritePlan, ids: DocPrIdAllocator, loc: GroupLoc): string {
  const id = ids.next();
  const descrAttr = group.altText ? ` descr="${escapeXmlAttribute(group.altText)}"` : "";
  const offX = cmToEmu(group.offsetXCm ?? 0);
  const offY = cmToEmu(group.offsetYCm ?? 0);
  return (
    `<wpg:grpSp xmlns:wpg="${WPG_NAMESPACE}" xmlns:a="${A_NAMESPACE}">` +
    `<wpg:cNvPr id="${id}" name="Group ${id}"${descrAttr}/><wpg:cNvGrpSpPr/>` +
    `<wpg:grpSpPr>${groupXfrmXml(group, offX, offY)}</wpg:grpSpPr>` +
    membersXml(group.members, ctx, mediaPlan, ids, loc) +
    "</wpg:grpSp>"
  );
}

/** The `wp:wrap*` element for a wrap mode plus `behindDoc`, mirroring the
 * shape writer. */
function anchorWrap(wrap: DrawingGroup["wrap"]): { behindDoc: string; wrapXml: string } {
  if (wrap === "square") {
    return { behindDoc: "0", wrapXml: '<wp:wrapSquare wrapText="bothSides"/>' };
  }
  if (wrap === "topAndBottom") {
    return { behindDoc: "0", wrapXml: "<wp:wrapTopAndBottom/>" };
  }
  if (wrap === "behind") {
    return { behindDoc: "1", wrapXml: "<wp:wrapNone/>" };
  }
  return { behindDoc: "0", wrapXml: "<wp:wrapNone/>" };
}

/** Serialises a `DrawingGroup` into a `w:drawing` element (wrapped by the
 * caller in `<w:r>` in `<w:p>`), holding one `wpg:wgp` with the members as
 * children (SPEC-057 AC-59). A sealed container re-emits its preserved subtree
 * inside the anchor instead of a fresh identity-space group. */
export function groupDrawingXml(group: DrawingGroup, ctx: ShapeTextWriteContext, mediaPlan: MediaWritePlan, ids: DocPrIdAllocator, chartPlan: ChartWritePlan = EMPTY_CHART_PLAN, section = 0, block = 0): string {
  const cx = cmToEmu(group.widthCm);
  const cy = cmToEmu(group.heightCm);
  const id = ids.next();
  const name = `Group ${id}`;
  const descrAttr = group.altText ? ` descr="${escapeXmlAttribute(group.altText)}"` : "";
  const docPr = `<wp:docPr id="${id}" name="${escapeXmlAttribute(name)}"${descrAttr}/>`;

  let graphicDataInner: string;
  if (group.preservedSubtreeXml !== undefined) {
    // Sealed: re-emit the source subtree verbatim (same-format save).
    graphicDataInner = group.preservedSubtreeXml;
  } else {
    graphicDataInner =
      `<wpg:wgp xmlns:wpg="${WPG_NAMESPACE}">` +
      "<wpg:cNvGrpSpPr/>" +
      `<wpg:grpSpPr>${groupXfrmXml(group, 0, 0)}</wpg:grpSpPr>` +
      membersXml(group.members, ctx, mediaPlan, ids, { chartPlan, section, block, path: [] }) +
      "</wpg:wgp>";
  }
  const graphic =
    `<a:graphic xmlns:a="${A_NAMESPACE}">` +
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    graphicDataInner +
    "</a:graphicData>" +
    "</a:graphic>";

  if (group.wrap === "inline") {
    return (
      `<w:drawing xmlns:wp="${WP_NAMESPACE}">` +
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      docPr +
      "<wp:cNvGraphicFramePr/>" +
      graphic +
      "</wp:inline>" +
      "</w:drawing>"
    );
  }

  const { behindDoc, wrapXml } = anchorWrap(group.wrap);
  const offsetXEmu = cmToEmu(group.offsetXCm ?? 0);
  const offsetYEmu = cmToEmu(group.offsetYCm ?? 0);
  const relativeHeight = 251658240 + (group.zOrder ?? 0);
  return (
    `<w:drawing xmlns:wp="${WP_NAMESPACE}">` +
    `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"` +
    ` relativeHeight="${relativeHeight}" behindDoc="${behindDoc}" locked="0"` +
    ' layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${offsetXEmu}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${offsetYEmu}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    wrapXml +
    docPr +
    "<wp:cNvGraphicFramePr/>" +
    graphic +
    "</wp:anchor>" +
    "</w:drawing>"
  );
}
