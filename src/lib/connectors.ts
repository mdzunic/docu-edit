// SPEC-070: connector structure commands and attachment book-keeping — the
// document-model half of the feature (the pure routing lives in
// core/connectors.ts). A Connector is a block-level DrawingPlacement addressed
// like a Shape (shapes.ts); this module owns inserting one, editing its kind /
// outline / arrowheads / wrap / alt-text / adjustment, attaching and detaching
// its endpoints, and — the memory that makes the feature — re-indexing every
// attachment when blocks move and freezing an endpoint at its last routed
// position when its target is deleted.
import {
  cloneContainers,
  isChart,
  isConnector,
  isImage,
  isShape,
  splitParagraphAt,
  type Block,
  type Chart,
  type Connector,
  type ConnectorAttachment,
  type ConnectorKind,
  type DocumentModel,
  type Image,
  type ImageWrapMode,
  type Paragraph,
  type Section,
  type Shape,
} from "./documentModel";
import type { ArrowheadKind, ShapeOutline } from "./core/shapes";
import {
  clampSite,
  clipCentreSite,
  connectionSitePoint,
  routeConnector,
  type ConnectorBox,
  type ConnectorEndpoint,
  type RoutedPath,
} from "./core/connectors";
import { rotatePoint, scalePolygon, shapePolygon, type Point } from "./core/shapeGeometry";
import {
  anchoredParagraph,
  ordinalForParagraph,
  releaseAnchorOrdinal,
  resolveAnchorBlockIndex,
} from "./drawingAnchors";
import { blockLength, clampCaret, type CaretPosition } from "./textEditing";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";
import { DEFAULT_DRAWING_POSITION } from "./shapes";

/** Where a connector command applies — a `{section, block}` pair, mirroring
 * ShapeLocation (the two block kinds keep independent location types). */
export interface ConnectorLocation {
  section: number;
  block: number;
}

/** Typed `null` seed for a connector-location ref in `.vue` scripts. */
export function noConnectorLocation(): ConnectorLocation | null {
  return null;
}

/** The default outline a freshly drawn connector carries — a thin dark line,
 * matching the shape default's stroke. */
export const DEFAULT_CONNECTOR_OUTLINE: ShapeOutline = { color: "#2e5395", widthPt: 1, dash: "solid" };

function emptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

/** The connector a location addresses, or null. */
export function connectorAt(model: DocumentModel, location: ConnectorLocation): Connector | null {
  const target = model.sections[location.section]?.blocks[location.block];
  return isConnector(target) ? target : null;
}

// --- endpoint resolution (model-level, pure) --------------------------------

/** A drawing block whose box a connector can attach to. */
type AttachableBlock = Shape | Image | Chart;

function isAttachable(block: Block | undefined): block is AttachableBlock {
  return isShape(block) || isImage(block) || isChart(block);
}

/** A drawing's placement box in section-relative cm (anchor offsets share an
 * origin for a fixture / same-anchor arrangement — the view uses live DOM
 * measurement for the cross-anchor case). */
function drawingBox(block: AttachableBlock): ConnectorBox {
  return {
    x: block.offsetXCm ?? 0,
    y: block.offsetYCm ?? 0,
    w: block.widthCm,
    h: block.heightCm,
    rotationDeg: block.rotationDeg,
  };
}

/** A shape's outline polygon in absolute section-relative cm, honouring size
 * and rotation — used to clip a centre-site (8) attachment to the edge. */
function shapeOutlineAbsolute(block: AttachableBlock, box: ConnectorBox): Point[] | null {
  if (!isShape(block)) {
    return null;
  }
  const unit = shapePolygon(block.geometry);
  if (!unit) {
    return null;
  }
  const scaled = scalePolygon(unit, box.w, box.h);
  return scaled.map((p) => {
    const r = rotatePoint(p, box.w, box.h, box.rotationDeg ?? 0);
    return { x: box.x + r.x, y: box.y + r.y };
  });
}

/** The resolved section-space anchor point (and outward normal, when
 * attached) of one connector endpoint, given the point the other endpoint
 * sits at (needed to clip a centre attachment). */
function resolveEndpoint(
  section: Section,
  connector: Connector,
  which: "start" | "end",
  otherPoint: Point,
): ConnectorEndpoint {
  const attachment = which === "start" ? connector.startAttachment : connector.endAttachment;
  const origin = { x: connector.offsetXCm ?? 0, y: connector.offsetYCm ?? 0 };
  const free = which === "start" ? connector.start : connector.end;
  if (!attachment) {
    return { point: { x: origin.x + free.xCm, y: origin.y + free.yCm } };
  }
  const target = section.blocks[attachment.target.block];
  if (!isAttachable(target)) {
    // Stale/detached target: fall back to the last routed (free) position.
    return { point: { x: origin.x + free.xCm, y: origin.y + free.yCm } };
  }
  const box = drawingBox(target);
  const site = clampSite(attachment.site);
  if (site === 8) {
    const outline = shapeOutlineAbsolute(target, box);
    const centre = connectionSitePoint(box, 8);
    if (outline) {
      return clipCentreSite(centre, otherPoint, outline);
    }
    return { point: centre };
  }
  const point = connectionSitePoint(box, site);
  // The outward normal is the difference between the site and the centre,
  // recovered from the box so rotation is honoured.
  const centre = connectionSitePoint(box, 8);
  const nx = point.x - centre.x;
  const ny = point.y - centre.y;
  const len = Math.hypot(nx, ny) || 1;
  return { point, normal: { x: nx / len, y: ny / len } };
}

/** Both resolved endpoints of a connector in section-relative cm — the pure,
 * model-level counterpart to the view's DOM measurement. Two passes so a
 * centre attachment can clip toward the far endpoint. */
export function resolveConnectorEndpoints(
  section: Section,
  connector: Connector,
): { start: ConnectorEndpoint; end: ConnectorEndpoint } {
  const origin = { x: connector.offsetXCm ?? 0, y: connector.offsetYCm ?? 0 };
  const rawStart = { x: origin.x + connector.start.xCm, y: origin.y + connector.start.yCm };
  const rawEnd = { x: origin.x + connector.end.xCm, y: origin.y + connector.end.yCm };
  const start = resolveEndpoint(section, connector, "start", rawEnd);
  const end = resolveEndpoint(section, connector, "end", start.point);
  // Re-resolve start now that end is known (matters only for centre clipping).
  const start2 = resolveEndpoint(section, connector, "start", end.point);
  return { start: start2, end };
}

/** The routed path of a connector in section-relative cm — endpoints resolved
 * from the model, then routed by kind. */
export function routeConnectorInSection(section: Section, connector: Connector): RoutedPath {
  const { start, end } = resolveConnectorEndpoints(section, connector);
  return routeConnector(connector.kind, start, end, connector.adjustments?.[0] ?? 0.5);
}

// --- attachment re-indexing and freezing ------------------------------------

/** How block indices moved after an edit: old section-block index → new index,
 * or a value < 0 for a block that was removed. Any index not present is
 * treated as unchanged. */
export type BlockIndexMapping = Map<number, number>;

/** Freezes one endpoint at its current routed position (writing it into
 * `start`/`end`) and drops its attachment — so a connector whose target is
 * gone stays where it last was. */
function freezeEndpoint(section: Section, connector: Connector, which: "start" | "end"): void {
  const { start, end } = resolveConnectorEndpoints(section, connector);
  const point = which === "start" ? start.point : end.point;
  const origin = { x: connector.offsetXCm ?? 0, y: connector.offsetYCm ?? 0 };
  const local = { xCm: point.x - origin.x, yCm: point.y - origin.y };
  if (which === "start") {
    connector.start = local;
    delete connector.startAttachment;
  } else {
    connector.end = local;
    delete connector.endAttachment;
  }
}

/** Rewrites every connector attachment in `section` after blocks move, and
 * detaches (freezing at its last position) any endpoint whose target is gone
 * ("Identity without ids"). Pure: returns a new Section, leaves the input
 * untouched. Called by any command that inserts, deletes or reorders blocks. */
export function reindexAttachments(section: Section, mapping: BlockIndexMapping): Section {
  const blocks = section.blocks.map((block) => {
    if (!isConnector(block)) {
      return block;
    }
    return { ...block } as Connector;
  });
  const next: Section = { ...section, blocks };
  for (const block of next.blocks) {
    if (!isConnector(block)) {
      continue;
    }
    for (const which of ["start", "end"] as const) {
      const attachment = which === "start" ? block.startAttachment : block.endAttachment;
      if (!attachment) {
        continue;
      }
      const mapped = mapping.get(attachment.target.block);
      if (mapped === undefined) {
        continue; // unchanged
      }
      if (mapped < 0) {
        freezeEndpoint(next, block, which);
      } else {
        const moved: ConnectorAttachment = { ...attachment, target: { ...attachment.target, block: mapped } };
        if (which === "start") {
          block.startAttachment = moved;
        } else {
          block.endAttachment = moved;
        }
      }
    }
  }
  return next;
}

/** The mapping for inserting `delta` net blocks at `atBlock`: every original
 * block at or after `atBlock` shifts up by `delta`; nothing is removed. (An
 * insert that splits/replaces a non-drawing block at `atBlock` never touches a
 * connector target, which is always a drawing.) */
export function insertMapping(atBlock: number, delta: number, blockCount: number): BlockIndexMapping {
  const mapping: BlockIndexMapping = new Map();
  if (delta === 0) {
    return mapping;
  }
  for (let i = atBlock; i < blockCount; i++) {
    mapping.set(i, i + delta);
  }
  return mapping;
}

/** The mapping for deleting `count` blocks at `atBlock`: those blocks are
 * removed (-1, so any connector attached to one freezes and detaches); every
 * block after them shifts down by `count`. */
export function deleteMapping(atBlock: number, count: number, blockCount: number): BlockIndexMapping {
  const mapping: BlockIndexMapping = new Map();
  for (let i = atBlock; i < blockCount; i++) {
    mapping.set(i, i < atBlock + count ? -1 : i - count);
  }
  return mapping;
}

/** Re-indexes a section in place after an INSERT of `delta` net blocks at
 * `atBlock` (the pre-splice `atBlock`). */
export function reindexSectionAfterInsert(
  model: DocumentModel,
  section: number,
  atBlock: number,
  delta: number,
): void {
  const target = model.sections[section];
  if (!target || delta === 0) {
    return;
  }
  const preCount = target.blocks.length - delta;
  model.sections[section] = reindexAttachments(target, insertMapping(atBlock, delta, preCount));
}

/** Re-indexes a section in place after a DELETE of `count` blocks at
 * `atBlock` — survivors shift, connectors attached to a removed block freeze. */
export function reindexSectionAfterDelete(
  model: DocumentModel,
  section: number,
  atBlock: number,
  count: number,
): void {
  const target = model.sections[section];
  if (!target) {
    return;
  }
  const preCount = target.blocks.length + count;
  model.sections[section] = reindexAttachments(target, deleteMapping(atBlock, count, preCount));
}

/** Detaches and freezes every connector endpoint in `section` attached to
 * `targetBlock` — the AC-11 behaviour when a shape is deleted, expressed so a
 * caller that removes a block can preserve its connectors. */
export function freezeConnectorsAttachedTo(section: Section, targetBlock: number): Section {
  const mapping: BlockIndexMapping = new Map([[targetBlock, -1]]);
  return reindexAttachments(section, mapping);
}

// --- commands ---------------------------------------------------------------

/** What insertConnector needs beyond where to place it. */
export interface NewConnector {
  kind: ConnectorKind;
  widthCm: number;
  heightCm: number;
  offsetXCm?: number;
  offsetYCm?: number;
  start: { xCm: number; yCm: number };
  end: { xCm: number; yCm: number };
  startAttachment?: ConnectorAttachment;
  endAttachment?: ConnectorAttachment;
  outline?: ShapeOutline;
  startArrow?: ArrowheadKind;
  endArrow?: ArrowheadKind;
}

/** Inserts a connector at the caret, splitting the caret paragraph around it
 * exactly like insertShape. A connector always floats (`wrap: "square"`). */
export function insertConnector(
  model: DocumentModel,
  selection: ModelSelection,
  spec: NewConnector,
): CommandResult {
  const next = cloneContainers(model);
  const at = clampCaret(next, selection.focus);
  if (at.cell) {
    return { model: next, selection: caretSelection(at) };
  }
  const blocks = next.sections[at.section].blocks;
  const target = blocks[at.block];
  const connector: Connector = {
    type: "connector",
    kind: spec.kind,
    widthCm: spec.widthCm,
    heightCm: spec.heightCm,
    wrap: "square",
    relativeToH: DEFAULT_DRAWING_POSITION.relativeToH,
    relativeToV: DEFAULT_DRAWING_POSITION.relativeToV,
    offsetXCm: spec.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm,
    offsetYCm: spec.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm,
    start: spec.start,
    end: spec.end,
    outline: spec.outline ?? { ...DEFAULT_CONNECTOR_OUTLINE },
    endArrow: spec.endArrow ?? "triangle",
  };
  if (spec.startArrow !== undefined) {
    connector.startArrow = spec.startArrow;
  }
  if (spec.startAttachment) {
    connector.startAttachment = spec.startAttachment;
  }
  if (spec.endAttachment) {
    connector.endAttachment = spec.endAttachment;
  }

  const replacement: Block[] = [];
  let blockOffset: number;
  if (target.type !== "paragraph") {
    if (at.offset === 0) {
      replacement.push(connector, target);
      blockOffset = 0;
    } else {
      replacement.push(target, connector, emptyParagraph());
      blockOffset = 1;
    }
  } else {
    const ordinal = ordinalForParagraph(next, target);
    connector.anchorOrdinal = ordinal;
    if (at.offset === 0) {
      replacement.push(connector, anchoredParagraph(target, ordinal));
      blockOffset = 0;
    } else if (at.offset >= blockLength(target)) {
      replacement.push(anchoredParagraph(target, ordinal), connector, emptyParagraph());
      blockOffset = 1;
    } else {
      const [before, after] = splitParagraphAt(target, at.offset);
      replacement.push(anchoredParagraph(before, ordinal), connector, after);
      blockOffset = 1;
    }
  }
  blocks.splice(at.block, 1, ...replacement);
  reindexSectionAfterInsert(next, at.section, at.block, replacement.length - 1);
  const connectorIndex = at.block + blockOffset;
  return {
    model: next,
    selection: caretSelection({ section: at.section, block: connectorIndex + 1, offset: 0 }),
  };
}

function editConnector(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
  mutate: (connector: Connector) => void,
): CommandResult {
  const next = cloneContainers(model);
  const connector = connectorAt(next, location);
  if (!connector) {
    return { model: next, selection };
  }
  const updated: Connector = { ...connector };
  mutate(updated);
  next.sections[location.section].blocks.splice(location.block, 1, updated);
  return { model: next, selection };
}

export function setConnectorKind(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
  kind: ConnectorKind,
): CommandResult {
  return editConnector(model, selection, location, (c) => {
    c.kind = kind;
  });
}

export function setConnectorOutline(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
  outline: ShapeOutline,
): CommandResult {
  return editConnector(model, selection, location, (c) => {
    c.outline = outline;
  });
}

export function setConnectorArrowheads(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
  ends: { startArrow?: ArrowheadKind; endArrow?: ArrowheadKind },
): CommandResult {
  return editConnector(model, selection, location, (c) => {
    if (ends.startArrow !== undefined) {
      c.startArrow = ends.startArrow;
    }
    if (ends.endArrow !== undefined) {
      c.endArrow = ends.endArrow;
    }
  });
}

export function setConnectorWrap(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
  wrap: ImageWrapMode,
): CommandResult {
  return editConnector(model, selection, location, (c) => {
    c.wrap = wrap;
  });
}

export function setConnectorAltText(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
  altText: string,
): CommandResult {
  return editConnector(model, selection, location, (c) => {
    if (altText) {
      c.altText = altText;
    } else {
      delete c.altText;
    }
  });
}

/** Sets the elbow/curved midpoint adjustment fraction (AC-16). */
export function setConnectorAdjustment(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
  fraction: number,
): CommandResult {
  return editConnector(model, selection, location, (c) => {
    c.adjustments = [Math.min(1, Math.max(0, fraction))];
  });
}

/** Re-attaches, detaches, or moves one endpoint of a connector (AC-15). When
 * `attachment` is provided the endpoint binds to that site; when `point`
 * (section-relative cm) is provided the endpoint becomes free at that point;
 * passing neither is a no-op. The connector's placement box is grown to
 * contain both endpoints so its own bounds stay sensible. */
export function setConnectorEndpoint(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
  which: "start" | "end",
  change: { attachment?: ConnectorAttachment | null; point?: { xCm: number; yCm: number } },
): CommandResult {
  return editConnector(model, selection, location, (c) => {
    if (change.attachment === null) {
      // Detach, freezing at the current routed position.
      const section = model.sections[location.section];
      freezeEndpoint(section, c, which);
      return;
    }
    if (change.attachment) {
      if (which === "start") {
        c.startAttachment = change.attachment;
      } else {
        c.endAttachment = change.attachment;
      }
    }
    if (change.point) {
      const origin = { x: c.offsetXCm ?? 0, y: c.offsetYCm ?? 0 };
      const local = { xCm: change.point.xCm - origin.x, yCm: change.point.yCm - origin.y };
      if (which === "start") {
        c.start = local;
        if (!change.attachment) {
          delete c.startAttachment;
        }
      } else {
        c.end = local;
        if (!change.attachment) {
          delete c.endAttachment;
        }
      }
    }
  });
}

/** Deletes a connector block (AC-15 handles endpoints; a connector itself is
 * deleted like any drawing). */
export function deleteConnector(
  model: DocumentModel,
  selection: ModelSelection,
  location: ConnectorLocation,
): CommandResult {
  const next = cloneContainers(model);
  const connector = connectorAt(next, location);
  if (!connector) {
    return { model: next, selection };
  }
  const blocks = next.sections[location.section].blocks;
  blocks.splice(location.block, 1);
  releaseAnchorOrdinal(next, connector.anchorOrdinal);
  reindexSectionAfterDelete(next, location.section, location.block, 1);
  const afterBlocks = next.sections[location.section].blocks;
  if (afterBlocks.length === 0) {
    afterBlocks.push(emptyParagraph());
  }
  const landing = Math.min(location.block, afterBlocks.length - 1);
  return {
    model: next,
    selection: caretSelection({ section: location.section, block: landing, offset: 0 }),
  };
}

/** Re-export so a caller resolving an anchor block index for a connector uses
 * the same helper shapes.ts does. */
export { resolveAnchorBlockIndex };
