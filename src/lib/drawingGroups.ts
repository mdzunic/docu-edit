// Native drawing-group container operations (SPEC-057), framework-free like
// every lib module. Replaces SPEC-051's flat `groupOrdinal` grouping (see
// shapes.ts) with a `DrawingGroup` block that holds its members by
// containment. Everything here is pure geometry/structure over the model —
// the gesture layer (src/views/document) supplies laid-out page rectangles so
// this module stays layout-free.
import {
  cloneContainers,
  isChart,
  isDrawingGroup,
  isImage,
  isShape,
  type Block,
  type DocumentModel,
  type DrawingGroup,
  type DrawingMember,
  type ImageWrapMode,
  type Shape,
} from "./documentModel";
import { MIN_SHAPE_SIZE_CM, normalizeRotationDeg, NUDGE_FINE_PT_CM, NUDGE_PT_CM, type NudgeDirection } from "./core/shapes";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";
import { DEFAULT_DRAWING_POSITION } from "./shapes";

/** A page-space rectangle in cm, in whatever frame the caller measured (the
 * anchor frame for grouping). Position is the top-left corner. */
export interface RectCm {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

/** Addresses a member inside a container by index path from the container
 * block (SPEC-057 "Addressing a member"): each `path` element indexes
 * `members`. An empty path addresses the container itself. */
export interface DrawingLocation {
  section: number;
  block: number;
  /** Index path into nested `members`; empty/absent = the container block. */
  path?: number[];
}

/** Typed `null` seed for a drawing-location ref in `.vue` scripts (espree
 * bars the annotation — see shapes.ts's noShapeLocation). */
export function noDrawingLocation(): DrawingLocation | null {
  return null;
}

/** A member's own placement fields carry offsets, size and rotation; these
 * are the only fields group geometry reads. `wrap`, `relativeToH/V` and
 * `zOrder` are meaningless inside a container. */
function memberOffsetX(member: DrawingMember): number {
  return member.offsetXCm ?? 0;
}
function memberOffsetY(member: DrawingMember): number {
  return member.offsetYCm ?? 0;
}

/** The union of the members' **unrotated** boxes (`offset + size`) — member
 * rotation is ignored for bounds, matching how both formats' group extents
 * treat rotated children (SPEC-057 "Normalized form"). Members are never
 * empty, so this is always defined. */
export function membersUnionBox(members: DrawingMember[]): RectCm {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const member of members) {
    const x = memberOffsetX(member);
    const y = memberOffsetY(member);
    left = Math.min(left, x);
    top = Math.min(top, y);
    right = Math.max(right, x + member.widthCm);
    bottom = Math.max(bottom, y + member.heightCm);
  }
  return { xCm: left, yCm: top, widthCm: right - left, heightCm: bottom - top };
}

/** Renormalizes a container so `min(offsetX) === 0`, `min(offsetY) === 0` and
 * `width/height` equal the members' union — folding the union origin into the
 * container's **own** offsets so nothing moves on screen (SPEC-057
 * "Normalized form"). A rotated container keeps its stored box: the box is the
 * rotation frame, and re-deriving it mid-rotation would move the pivot — so
 * this is a no-op when `rotationDeg` is non-zero. Returns a new container;
 * does not mutate the input. */
export function normalizeGroup(group: DrawingGroup): DrawingGroup {
  const next: DrawingGroup = { ...group, members: group.members.map((m) => ({ ...m })) };
  if (normalizeRotationDeg(next.rotationDeg ?? 0) !== 0) {
    return next;
  }
  const union = membersUnionBox(next.members);
  next.members = next.members.map((m) => ({
    ...m,
    offsetXCm: memberOffsetX(m) - union.xCm,
    offsetYCm: memberOffsetY(m) - union.yCm,
  }));
  next.offsetXCm = (next.offsetXCm ?? DEFAULT_DRAWING_POSITION.offsetXCm) + union.xCm;
  next.offsetYCm = (next.offsetYCm ?? DEFAULT_DRAWING_POSITION.offsetYCm) + union.yCm;
  next.widthCm = union.widthCm;
  next.heightCm = union.heightCm;
  return next;
}

/** Rotates a point clockwise by `deg` about the origin (matching the
 * `rotationDeg` sense the painter uses). */
function rotateAboutOrigin(dx: number, dy: number, deg: number): { dx: number; dy: number } {
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  // Clockwise in screen coordinates (y down): [cos, -sin; sin, cos].
  return { dx: dx * cos - dy * sin, dy: dx * sin + dy * cos };
}

/** `bakeGroupRotation` (SPEC-057 "Rotation and the bake transform") — for a
 * container with rotation θ and box centre C (in its own frame): each
 * member's centre moves to `C + R(θ)·(centre − C)`, each member's
 * `rotationDeg` becomes `(rotationDeg + θ) mod 360`, the container's rotation
 * clears, and the container renormalizes. Nested member containers are baked
 * by rotating their centres and adding θ to their own `rotationDeg` — **not**
 * by recursing into their members. Used by the `.odt` writer (on a
 * serialization-time copy — saving never mutates the open model), by Ungroup,
 * and by the cross-format equality mask. A no-op when rotation is 0. */
export function bakeGroupRotation(group: DrawingGroup): DrawingGroup {
  const theta = normalizeRotationDeg(group.rotationDeg ?? 0);
  if (theta === 0) {
    const next: DrawingGroup = { ...group, members: group.members.map((m) => ({ ...m })) };
    delete next.rotationDeg;
    return next;
  }
  // Box centre in the container's own frame (offsets are member-relative).
  const cx = group.widthCm / 2;
  const cy = group.heightCm / 2;
  const members = group.members.map((member) => {
    const memberCx = memberOffsetX(member) + member.widthCm / 2;
    const memberCy = memberOffsetY(member) + member.heightCm / 2;
    const rotated = rotateAboutOrigin(memberCx - cx, memberCy - cy, theta);
    const newCx = cx + rotated.dx;
    const newCy = cy + rotated.dy;
    const spun: DrawingMember = {
      ...member,
      offsetXCm: newCx - member.widthCm / 2,
      offsetYCm: newCy - member.heightCm / 2,
    };
    const memberRot = normalizeRotationDeg((member.rotationDeg ?? 0) + theta);
    if (memberRot === 0) {
      delete spun.rotationDeg;
    } else {
      spun.rotationDeg = memberRot;
    }
    return spun;
  });
  const baked: DrawingGroup = { ...group, members };
  delete baked.rotationDeg;
  return normalizeGroup(baked);
}

/** Recursively bakes every container in a subtree, outermost first, until no
 * container carries rotation — the `.odt` writer's serialization transform
 * (SPEC-057). Operates on a deep copy; never mutates the input. */
export function bakeGroupRotationDeep(group: DrawingGroup): DrawingGroup {
  const baked = bakeGroupRotation(group);
  const members = baked.members.map((m) => (isDrawingGroup(m) ? bakeGroupRotationDeep(m) : { ...m }));
  return { ...baked, members };
}

/** Scales every member (recursively into nested containers) by the per-axis
 * factors — the interior-scaling half of a container resize (SPEC-057
 * "Interior scaling"). Member `rotationDeg` is unchanged; scaling applies to
 * a member's unrotated box (aspect changes, no shear). */
export function scaleMembers(members: DrawingMember[], sx: number, sy: number): DrawingMember[] {
  return members.map((member) => {
    const scaled: DrawingMember = {
      ...member,
      offsetXCm: memberOffsetX(member) * sx,
      offsetYCm: memberOffsetY(member) * sy,
      widthCm: member.widthCm * sx,
      heightCm: member.heightCm * sy,
    };
    if (isDrawingGroup(scaled)) {
      scaled.members = scaleMembers(scaled.members, sx, sy);
    }
    return scaled;
  });
}

/** Strips a drawing of the anchor-level fields that are meaningless inside a
 * container (SPEC-057): `relativeToH/V` and `zOrder` are removed, `wrap`
 * becomes the fixed placeholder `"square"`, and the offset is set to the
 * member's position within the container box. Everything else — size,
 * rotation, fill, text, media, nested members — is preserved. */
function asMember(block: DrawingMember, offsetXCm: number, offsetYCm: number): DrawingMember {
  const member: DrawingMember = { ...block, wrap: "square", offsetXCm, offsetYCm };
  delete member.relativeToH;
  delete member.relativeToV;
  delete member.zOrder;
  delete (member as { groupOrdinal?: number }).groupOrdinal;
  return member;
}

function zOrderOf(block: { zOrder?: number }): number {
  return block.zOrder ?? 0;
}

/** Builds a container from selected drawings and their laid-out page
 * rectangles (SPEC-057 "Creating a group"). `entries` are in document order;
 * the container is returned along with the document-order index of the first
 * member (its block position). Member order in the container is paint order —
 * ascending `zOrder`, ties by document order. The container box is the union
 * of the rectangles; member offsets are their rectangle's position within it;
 * the container anchors where the first member did, offsets landing its box on
 * the union origin. */
export function buildContainer(entries: { block: DrawingMember; rect: RectCm }[]): DrawingGroup {
  // Paint order: ascending zOrder, ties by current (document) order.
  const ordered = entries
    .map((entry, index) => ({ ...entry, index }))
    .sort((a, b) => {
      const za = zOrderOf(a.block);
      const zb = zOrderOf(b.block);
      if (za !== zb) {
        return za - zb;
      }
      return a.index - b.index;
    });
  const union = {
    xCm: Math.min(...ordered.map((e) => e.rect.xCm)),
    yCm: Math.min(...ordered.map((e) => e.rect.yCm)),
  };
  const right = Math.max(...ordered.map((e) => e.rect.xCm + e.rect.widthCm));
  const bottom = Math.max(...ordered.map((e) => e.rect.yCm + e.rect.heightCm));
  const first = entries[0].block;
  const firstWrap: ImageWrapMode = first.wrap === "inline" ? "square" : first.wrap;
  const maxZ = Math.max(...entries.map((e) => zOrderOf(e.block)));
  const members = ordered.map((e) => asMember(e.block, e.rect.xCm - union.xCm, e.rect.yCm - union.yCm));
  const container: DrawingGroup = {
    type: "drawingGroup",
    members,
    widthCm: right - union.xCm,
    heightCm: bottom - union.yCm,
    wrap: firstWrap,
    relativeToH: first.relativeToH ?? DEFAULT_DRAWING_POSITION.relativeToH,
    relativeToV: first.relativeToV ?? DEFAULT_DRAWING_POSITION.relativeToV,
    offsetXCm: union.xCm,
    offsetYCm: union.yCm,
  };
  if (maxZ !== 0) {
    container.zOrder = maxZ;
  }
  return container;
}

/** A member's block-list location plus its laid-out page rectangle. The
 * gesture layer supplies the rectangles (the same geometry the painter used),
 * keeping this module layout-free. */
export interface GroupEntry {
  section: number;
  block: number;
  rect: RectCm;
}

/** Groups two or more selected drawings into one container at the first
 * member's block position (SPEC-057 AC-3, AC-4). Members must all sit in the
 * same section; the caller (the Group control) enforces that. Leaves the new
 * container selected. */
export function groupDrawings(model: DocumentModel, selection: ModelSelection, entries: GroupEntry[]): CommandResult {
  const next = cloneContainers(model);
  if (entries.length < 2) {
    return { model: next, selection };
  }
  const section = entries[0].section;
  const sorted = [...entries].sort((a, b) => a.block - b.block);
  const blocks = next.sections[section].blocks;
  const built = sorted.map((entry) => {
    const block = blocks[entry.block];
    return { block: block as DrawingMember, rect: entry.rect };
  });
  const container = buildContainer(built);
  const firstBlock = sorted[0].block;
  // Remove the member blocks (highest index first so earlier indices hold).
  const indices = sorted.map((e) => e.block).sort((a, b) => b - a);
  for (const index of indices) {
    blocks.splice(index, 1);
  }
  blocks.splice(firstBlock, 0, container);
  return {
    model: next,
    selection: caretSelection({ section, block: firstBlock, offset: 0 }),
  };
}

/** Restores a member as a standalone drawing when a container dissolves — the
 * inverse of `asMember` (SPEC-057 "Ungrouping"): each member inherits the
 * container's `wrap`, `relativeToH/V` and `zOrder`, anchors to the
 * container's paragraph, and gets offsets `container offset + member
 * offset`. */
function emergeMember(member: DrawingMember, container: DrawingGroup): DrawingMember {
  const emerged: DrawingMember = {
    ...member,
    wrap: container.wrap,
    offsetXCm: (container.offsetXCm ?? 0) + memberOffsetX(member),
    offsetYCm: (container.offsetYCm ?? 0) + memberOffsetY(member),
  };
  if (container.relativeToH !== undefined) {
    emerged.relativeToH = container.relativeToH;
  }
  if (container.relativeToV !== undefined) {
    emerged.relativeToV = container.relativeToV;
  }
  if (container.zOrder !== undefined) {
    emerged.zOrder = container.zOrder;
  }
  return emerged;
}

/** Dissolves a container one level (SPEC-057 "Ungrouping"): splices its
 * members into the block list at the container's position, in member order,
 * each re-anchored to the container's paragraph with unchanged on-screen
 * geometry. A rotated container bakes its rotation first; an `inline`
 * container converts to `square` (members re-entering the flow cannot hold
 * offsets). Leaves the emerged members multi-selected, so Ctrl/Cmd+G inverts
 * it immediately. A nested member container comes out as a container — Ungroup
 * peels one level. */
export function ungroupDrawing(model: DocumentModel, selection: ModelSelection, section: number, block: number): CommandResult & { emerged?: { section: number; block: number }[] } {
  const next = cloneContainers(model);
  const target = next.sections[section]?.blocks[block];
  if (!isDrawingGroup(target)) {
    return { model: next, selection };
  }
  let container = target;
  if (normalizeRotationDeg(container.rotationDeg ?? 0) !== 0) {
    container = bakeGroupRotation(container);
  }
  if (container.wrap === "inline") {
    container = { ...container, wrap: "square" };
  }
  const emergedBlocks = container.members.map((m) => emergeMember(m, container));
  next.sections[section].blocks.splice(block, 1, ...emergedBlocks);
  const emerged = emergedBlocks.map((_, i) => ({ section, block: block + i }));
  return {
    model: next,
    selection: caretSelection({ section, block, offset: 0 }),
    emerged,
  };
}

// --- interior member edits --------------------------------------------------

/** Reads the member a path addresses inside a container, or null. */
export function memberAtPath(container: DrawingGroup, path: number[]): DrawingMember | null {
  let current: DrawingMember = container;
  for (const index of path) {
    if (!isDrawingGroup(current)) {
      return null;
    }
    const child: DrawingMember | undefined = current.members[index];
    if (!child) {
      return null;
    }
    current = child;
  }
  return current;
}

/** Returns a copy of `container` with the member at `path` replaced by
 * `replacement` (or removed when `replacement` is null), renormalizing every
 * container on the path from the edited member up so nothing jumps on screen
 * (SPEC-057 "Normalized form"). A container reduced to one member is signalled
 * by returning that lone member for the caller to dissolve. */
function updateMemberAtPath(container: DrawingGroup, path: number[], replacement: DrawingMember | null): DrawingMember {
  if (path.length === 0) {
    return replacement ?? container;
  }
  const [head, ...rest] = path;
  const members = [...container.members];
  if (rest.length === 0) {
    if (replacement === null) {
      members.splice(head, 1);
    } else {
      members.splice(head, 1, replacement);
    }
  } else {
    const child = members[head];
    if (!isDrawingGroup(child)) {
      return container;
    }
    const updatedChild = updateMemberAtPath(child, rest, replacement);
    // A nested container reduced to a lone member dissolves into it.
    members.splice(head, 1, updatedChild);
  }
  let next: DrawingGroup = { ...container, members };
  if (next.members.length === 1) {
    // Reduced to one member: dissolve, folding the container's own offset in.
    return emergeMember(next.members[0], next);
  }
  next = normalizeGroup(next);
  return next;
}

/** Replaces the container block at `section`/`block` with the result of an
 * interior edit, dissolving it if the edit reduced it to one member. */
function commitContainer(next: DocumentModel, section: number, block: number, updated: DrawingMember): void {
  next.sections[section].blocks.splice(block, 1, updated as Block);
}

/** Moves a selected member inside its container by a page-space delta, then
 * renormalizes (SPEC-057 AC-26). */
export function moveMember(model: DocumentModel, selection: ModelSelection, location: DrawingLocation, dxCm: number, dyCm: number): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return { model: next, selection };
  }
  const member = memberAtPath(container, path);
  if (!member) {
    return { model: next, selection };
  }
  const moved: DrawingMember = { ...member, offsetXCm: memberOffsetX(member) + dxCm, offsetYCm: memberOffsetY(member) + dyCm };
  commitContainer(next, location.section, location.block, updateMemberAtPath(container, path, moved));
  return { model: next, selection };
}

/** Arrow-key nudge / fine nudge of a selected member (SPEC-057 AC-27). */
export function nudgeMember(model: DocumentModel, selection: ModelSelection, location: DrawingLocation, direction: NudgeDirection, fine: boolean): CommandResult {
  const step = fine ? NUDGE_FINE_PT_CM : NUDGE_PT_CM;
  const dx = direction === "left" ? -step : direction === "right" ? step : 0;
  const dy = direction === "up" ? -step : direction === "down" ? step : 0;
  return moveMember(model, selection, location, dx, dy);
}

/** Resizes a selected member inside its container (SPEC-057 AC-28), floored at
 * `MIN_SHAPE_SIZE_CM` on both axes, then renormalizes. */
export function resizeMember(model: DocumentModel, selection: ModelSelection, location: DrawingLocation, widthCm: number, heightCm: number): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return { model: next, selection };
  }
  const member = memberAtPath(container, path);
  if (!member) {
    return { model: next, selection };
  }
  const resized: DrawingMember = { ...member, widthCm: Math.max(MIN_SHAPE_SIZE_CM, widthCm), heightCm: Math.max(MIN_SHAPE_SIZE_CM, heightCm) };
  commitContainer(next, location.section, location.block, updateMemberAtPath(container, path, resized));
  return { model: next, selection };
}

/** Rotates a selected member about its own centre (SPEC-057 AC-29), then
 * renormalizes. */
export function rotateMember(model: DocumentModel, selection: ModelSelection, location: DrawingLocation, rotationDeg: number): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return { model: next, selection };
  }
  const member = memberAtPath(container, path);
  if (!member) {
    return { model: next, selection };
  }
  const rotated: DrawingMember = { ...member };
  const normalized = normalizeRotationDeg(rotationDeg);
  if (normalized === 0) {
    delete rotated.rotationDeg;
  } else {
    rotated.rotationDeg = normalized;
  }
  commitContainer(next, location.section, location.block, updateMemberAtPath(container, path, rotated));
  return { model: next, selection };
}

/** Sets a selected member's alt text (SPEC-057) — its shape toolbar's alt-text
 * field edits the member in place. */
export function setMemberAltText(model: DocumentModel, selection: ModelSelection, location: DrawingLocation, altText: string): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return { model: next, selection };
  }
  const member = memberAtPath(container, path);
  if (!member) {
    return { model: next, selection };
  }
  const updated: DrawingMember = { ...member };
  if (altText) {
    updated.altText = altText;
  } else {
    delete updated.altText;
  }
  commitContainer(next, location.section, location.block, updateMemberAtPath(container, path, updated));
  return { model: next, selection };
}

// --- member text-frame editing seam (SPEC-057 AC-25) ------------------------
// A grouped text box's text is editable in place through the exact SPEC-051
// shape-text lifecycle: these mirror commands.shapes.ts's withShapeModel /
// commitShapeModel, but address the shape MEMBER a DrawingLocation path points
// at rather than a top-level block. Entering edit mode populates an empty
// frame first (ensureMemberTextFrame), just as handleShapeEditText does for a
// bare shape — so every character/paragraph feature body text accepts (lists,
// tables, fields, IME) works unchanged against a member's text.

/** The shape member a location addresses, or null when the path is empty or
 * points at a non-shape (an image/chart/nested container holds no text). */
export function memberShapeAt(model: DocumentModel, location: DrawingLocation): Shape | null {
  const container = model.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return null;
  }
  const member = memberAtPath(container, path);
  return member && isShape(member) ? member : null;
}

/** Addresses a shape member's text frame as a throwaway single-section
 * document (SPEC-057 AC-25) — withShapeModel's seam, reused for a member. An
 * empty (or absent) frame starts from one empty paragraph. */
export function withMemberTextModel(model: DocumentModel, location: DrawingLocation): DocumentModel {
  const shape = memberShapeAt(model, location);
  const blocks: Block[] =
    shape && shape.text && shape.text.blocks.length > 0
      ? shape.text.blocks
      : [{ type: "paragraph", runs: [] }];
  return { ...model, sections: [{ blocks }] };
}

/** Replaces the shape member at `path` inside `container` with the same shape
 * carrying `blocks` as its text (vAlign untouched), renormalizing exactly as
 * every other interior edit does. A text edit never changes member count or
 * geometry, so the renormalized container is geometrically identical. */
function withMemberText(container: DrawingGroup, path: number[], blocks: Block[]): DrawingMember {
  const shape = memberAtPath(container, path);
  if (!shape || !isShape(shape)) {
    return container;
  }
  const updated: Shape = { ...shape, text: { blocks, vAlign: shape.text?.vAlign ?? "top" } };
  return updateMemberAtPath(container, path, updated);
}

/** Folds a `withMemberTextModel` command result back into the member's text
 * frame (SPEC-057 AC-25) — commitShapeModel's sibling for a member: the
 * synthetic section's blocks become the member shape's `text.blocks`, the real
 * sections come back from `before`, and every other model-level change the
 * inner command made (e.g. a new list definition) is kept from `result`. */
export function commitMemberTextModel(before: DocumentModel, location: DrawingLocation, result: DocumentModel): DocumentModel {
  const path = location.path ?? [];
  const blocks = result.sections[0]?.blocks ?? [];
  const sections = before.sections.map((section, sectionIndex) => {
    if (sectionIndex !== location.section) {
      return section;
    }
    return {
      ...section,
      blocks: section.blocks.map((block, blockIndex) =>
        blockIndex === location.block && isDrawingGroup(block) ? (withMemberText(block, path, blocks) as Block) : block,
      ),
    };
  });
  return { ...result, sections };
}

/** Populates an empty shape member's text frame (SPEC-057 AC-25) — the
 * member equivalent of ensureShapeTextFrame, so double-clicking a grouped
 * shape that holds no text yet gives it one to edit. A no-op for a member
 * that is not a text-less shape. */
export function ensureMemberTextFrame(model: DocumentModel, selection: ModelSelection, location: DrawingLocation): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return { model: next, selection };
  }
  const shape = memberAtPath(container, path);
  if (!shape || !isShape(shape) || shape.text) {
    return { model: next, selection };
  }
  commitContainer(next, location.section, location.block, withMemberText(container, path, [{ type: "paragraph", runs: [] }]));
  return { model: next, selection };
}

/** Deletes a selected member from its container (SPEC-057 AC-31). A container
 * reduced to one member dissolves into it (AC-32). */
export function deleteMember(model: DocumentModel, selection: ModelSelection, location: DrawingLocation): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return { model: next, selection };
  }
  commitContainer(next, location.section, location.block, updateMemberAtPath(container, path, null));
  return {
    model: next,
    selection: caretSelection({ section: location.section, block: location.block, offset: 0 }),
  };
}

/** The 0.5cm duplicate offset shared with shapes.ts's duplicateShape. */
export const DUPLICATE_OFFSET_CM = 0.5;

/** Duplicates a selected member as a sibling inserted immediately after its
 * source in `members`, offset by 0.5cm (SPEC-057), then renormalizes. */
export function duplicateMember(model: DocumentModel, selection: ModelSelection, location: DrawingLocation): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return { model: next, selection };
  }
  const member = memberAtPath(container, path);
  if (!member) {
    return { model: next, selection };
  }
  const copy: DrawingMember = deepCloneMember({
    ...member,
    offsetXCm: memberOffsetX(member) + DUPLICATE_OFFSET_CM,
    offsetYCm: memberOffsetY(member) + DUPLICATE_OFFSET_CM,
  });
  const parentPath = path.slice(0, -1);
  const memberIndex = path[path.length - 1];
  const parent = memberAtPath(container, parentPath);
  if (!parent || !isDrawingGroup(parent)) {
    return { model: next, selection };
  }
  const newMembers = [...parent.members];
  newMembers.splice(memberIndex + 1, 0, copy);
  const updatedParent: DrawingGroup = normalizeGroup({ ...parent, members: newMembers });
  const updated = parentPath.length === 0 ? updatedParent : replaceAtPath(container, parentPath, updatedParent);
  commitContainer(next, location.section, location.block, updated);
  return {
    model: next,
    selection: caretSelection({ section: location.section, block: location.block, offset: 0 }),
  };
}

/** The member path a freshly duplicated member lands at — a sibling
 * immediately after its source (SPEC-057). The UI reads this to keep the
 * duplicate selected. */
export function duplicatedMemberPath(location: DrawingLocation): number[] {
  const path = location.path ?? [];
  const parentPath = path.slice(0, -1);
  const memberIndex = path[path.length - 1] ?? 0;
  return [...parentPath, memberIndex + 1];
}

/** The four reorder verbs acting within `members` (SPEC-057 AC-33..AC-36):
 * inside a container, paint order is member order, so these mutate array
 * order only. */
export type MemberReorder = "front" | "forward" | "backward" | "back";

export function reorderMember(model: DocumentModel, selection: ModelSelection, location: DrawingLocation, command: MemberReorder): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[location.section]?.blocks[location.block];
  const path = location.path ?? [];
  if (!isDrawingGroup(container) || path.length === 0) {
    return { model: next, selection };
  }
  const parentPath = path.slice(0, -1);
  const memberIndex = path[path.length - 1];
  const parent = memberAtPath(container, parentPath);
  if (!parent || !isDrawingGroup(parent)) {
    return { model: next, selection };
  }
  const members = [...parent.members];
  let targetIndex = memberIndex;
  if (command === "front") {
    targetIndex = members.length - 1;
  } else if (command === "back") {
    targetIndex = 0;
  } else if (command === "forward") {
    targetIndex = Math.min(members.length - 1, memberIndex + 1);
  } else {
    targetIndex = Math.max(0, memberIndex - 1);
  }
  const [moved] = members.splice(memberIndex, 1);
  members.splice(targetIndex, 0, moved);
  const updatedParent: DrawingGroup = { ...parent, members };
  const updated = parentPath.length === 0 ? updatedParent : replaceAtPath(container, parentPath, updatedParent);
  commitContainer(next, location.section, location.block, updated);
  return { model: next, selection };
}

/** Replaces the container at `path` (a container itself) with `replacement`
 * without the dissolve/normalize logic of `updateMemberAtPath` — used when the
 * caller already produced the fully-formed replacement subtree. */
function replaceAtPath(container: DrawingGroup, path: number[], replacement: DrawingGroup): DrawingGroup {
  if (path.length === 0) {
    return replacement;
  }
  const [head, ...rest] = path;
  const members = [...container.members];
  const child = members[head];
  if (!isDrawingGroup(child)) {
    return container;
  }
  members.splice(head, 1, rest.length === 0 ? replacement : replaceAtPath(child, rest, replacement));
  return { ...container, members };
}

/** Deep-clones a member subtree so a duplicate shares no nested arrays with
 * its source. */
function deepCloneMember(member: DrawingMember): DrawingMember {
  if (isDrawingGroup(member)) {
    return { ...member, members: member.members.map(deepCloneMember) };
  }
  return JSON.parse(JSON.stringify(member)) as DrawingMember;
}

// --- container resize / rotate ----------------------------------------------

/** Resizes a selected container's own box, scaling every member with it
 * (SPEC-057 "Interior scaling"). Both container axes are floored at
 * `MIN_SHAPE_SIZE_CM`; members need no floor of their own (the factors are
 * never zero). */
export function resizeGroupContainer(model: DocumentModel, selection: ModelSelection, section: number, block: number, widthCm: number, heightCm: number): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[section]?.blocks[block];
  if (!isDrawingGroup(container)) {
    return { model: next, selection };
  }
  const newWidth = Math.max(MIN_SHAPE_SIZE_CM, widthCm);
  const newHeight = Math.max(MIN_SHAPE_SIZE_CM, heightCm);
  const sx = newWidth / container.widthCm;
  const sy = newHeight / container.heightCm;
  const updated: DrawingGroup = { ...container, widthCm: newWidth, heightCm: newHeight, members: scaleMembers(container.members, sx, sy) };
  commitContainer(next, section, block, updated);
  return { model: next, selection };
}

/** Rotates a selected container by setting its own `rotationDeg` — it does not
 * touch members (SPEC-057 "container rotation"). */
export function rotateGroupContainer(model: DocumentModel, selection: ModelSelection, section: number, block: number, rotationDeg: number): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[section]?.blocks[block];
  if (!isDrawingGroup(container)) {
    return { model: next, selection };
  }
  const normalized = normalizeRotationDeg(rotationDeg);
  const updated: DrawingGroup = { ...container };
  if (normalized === 0) {
    delete updated.rotationDeg;
  } else {
    updated.rotationDeg = normalized;
  }
  commitContainer(next, section, block, updated);
  return { model: next, selection };
}

// --- container-level edits (the container as one drawing) -------------------

/** Moves a selected container by a page-space delta — the container is one
 * drawing, so this is a plain offset shift (SPEC-057 AC-10). */
export function moveGroupContainer(model: DocumentModel, selection: ModelSelection, section: number, block: number, dxCm: number, dyCm: number): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[section]?.blocks[block];
  if (!isDrawingGroup(container)) {
    return { model: next, selection };
  }
  const updated: DrawingGroup = {
    ...container,
    offsetXCm: (container.offsetXCm ?? 0) + dxCm,
    offsetYCm: (container.offsetYCm ?? 0) + dyCm,
  };
  commitContainer(next, section, block, updated);
  return { model: next, selection };
}

/** Arrow-key nudge / fine nudge of a selected container (SPEC-057 AC-11). */
export function nudgeGroupContainer(model: DocumentModel, selection: ModelSelection, section: number, block: number, direction: NudgeDirection, fine: boolean): CommandResult {
  const step = fine ? NUDGE_FINE_PT_CM : NUDGE_PT_CM;
  const dx = direction === "left" ? -step : direction === "right" ? step : 0;
  const dy = direction === "up" ? -step : direction === "down" ? step : 0;
  return moveGroupContainer(model, selection, section, block, dx, dy);
}

/** Deletes a selected container from the block list (SPEC-057 AC-12). */
export function deleteGroupContainer(model: DocumentModel, selection: ModelSelection, section: number, block: number): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[section]?.blocks[block];
  if (!isDrawingGroup(container)) {
    return { model: next, selection };
  }
  next.sections[section].blocks.splice(block, 1);
  return { model: next, selection: caretSelection({ section, block: Math.max(0, block - 1), offset: 0 }) };
}

/** Duplicates a selected container as a sibling right after it, offset by
 * 0.5cm (SPEC-057 AC-13). The duplicate lands at block+1 and is left
 * selected by the caller. */
export function duplicateGroupContainer(model: DocumentModel, selection: ModelSelection, section: number, block: number): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[section]?.blocks[block];
  if (!isDrawingGroup(container)) {
    return { model: next, selection };
  }
  const copy = deepCloneMember({
    ...container,
    offsetXCm: (container.offsetXCm ?? 0) + DUPLICATE_OFFSET_CM,
    offsetYCm: (container.offsetYCm ?? 0) + DUPLICATE_OFFSET_CM,
  }) as DrawingGroup;
  next.sections[section].blocks.splice(block + 1, 0, copy as Block);
  return { model: next, selection: caretSelection({ section, block: block + 1, offset: 0 }) };
}

/** Sets a selected container's alt text (SPEC-057 AC-49). A sealed container's
 * alt text is read-only; the caller enforces that. */
export function setGroupAltText(model: DocumentModel, selection: ModelSelection, section: number, block: number, altText: string): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[section]?.blocks[block];
  if (!isDrawingGroup(container)) {
    return { model: next, selection };
  }
  const updated: DrawingGroup = { ...container };
  if (altText) {
    updated.altText = altText;
  } else {
    delete updated.altText;
  }
  commitContainer(next, section, block, updated);
  return { model: next, selection };
}

/** Sets a selected container's wrap mode (SPEC-057) — the container is one
 * drawing, so this is the same wrap the surrounding text flows around. */
export function setGroupWrap(model: DocumentModel, selection: ModelSelection, section: number, block: number, wrap: ImageWrapMode): CommandResult {
  const next = cloneContainers(model);
  const container = next.sections[section]?.blocks[block];
  if (!isDrawingGroup(container)) {
    return { model: next, selection };
  }
  commitContainer(next, section, block, { ...container, wrap });
  return { model: next, selection };
}

// --- legacy migration -------------------------------------------------------

/** Collapses drawings sharing a legacy `groupOrdinal` into native containers
 * (SPEC-057 "Legacy migration"), per section, after a section's blocks are
 * assembled. Drawings sharing an ordinal collapse into one container at the
 * first member's block index; the box is the union of the members' raw
 * offsets treated as same-frame values; member offsets re-express against it.
 * An ordinal with a single member yields no container. Mutates and returns the
 * block list. */
export function migrateLegacyGroups(blocks: Block[]): Block[] {
  const ordinals = new Map<number, number[]>();
  blocks.forEach((block, index) => {
    const ordinal = (block as { groupOrdinal?: number }).groupOrdinal;
    if (ordinal !== undefined && (isShape(block) || isImage(block) || isChart(block) || isDrawingGroup(block))) {
      const list = ordinals.get(ordinal) ?? [];
      list.push(index);
      ordinals.set(ordinal, list);
    }
  });
  // Build containers for ordinals with two or more members; drop the ordinal
  // tag from single-member ordinals. Work highest-first so splices don't shift
  // pending indices.
  const groups = [...ordinals.entries()].filter(([, idxs]) => idxs.length >= 2).sort((a, b) => b[1][0] - a[1][0]);
  // First, strip the tag from lone-ordinal members.
  for (const [ordinal, idxs] of ordinals) {
    if (idxs.length < 2) {
      const only = blocks[idxs[0]] as { groupOrdinal?: number };
      delete only.groupOrdinal;
      void ordinal;
    }
  }
  for (const [, idxs] of groups) {
    const memberBlocks = idxs.map((i) => blocks[i] as DrawingMember);
    const entries = memberBlocks.map((block) => ({
      block,
      rect: {
        xCm: block.offsetXCm ?? 0,
        yCm: block.offsetYCm ?? 0,
        widthCm: block.widthCm,
        heightCm: block.heightCm,
      },
    }));
    const container = buildContainer(entries);
    const firstIndex = idxs[0];
    // Remove members (highest index first), then insert the container where the
    // first member was.
    const sortedDesc = [...idxs].sort((a, b) => b - a);
    for (const i of sortedDesc) {
      blocks.splice(i, 1);
    }
    blocks.splice(firstIndex, 0, container);
  }
  return blocks;
}
