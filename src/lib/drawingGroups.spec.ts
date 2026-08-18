// SPEC-057: native drawing-group container operations. Pure model/geometry
// tests — no component mounting, mirroring shapes.spec.ts's conventions.
import { describe, it, expect } from "vitest";
import {
  bakeGroupRotation,
  bakeGroupRotationDeep,
  buildContainer,
  deleteMember,
  duplicateMember,
  groupDrawings,
  membersUnionBox,
  memberAtPath,
  migrateLegacyGroups,
  moveMember,
  normalizeGroup,
  reorderMember,
  resizeGroupContainer,
  resizeMember,
  rotateGroupContainer,
  scaleMembers,
  ungroupDrawing,
  type GroupEntry,
} from "./drawingGroups";
import { caretSelection, type ModelSelection } from "./commands";
import {
  createDocument,
  isDrawingGroup,
  type Block,
  type Chart,
  type DocumentModel,
  type DrawingGroup,
  type DrawingMember,
  type Image,
  type Paragraph,
  type Shape,
} from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function doc(...blocks: Block[]): DocumentModel {
  const model = createDocument("group-fixture");
  model.sections = [{ blocks }];
  return model;
}

function shape(overrides: Partial<Shape> = {}): Shape {
  return { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square", ...overrides };
}

function image(overrides: Partial<Image> = {}): Image {
  return {
    type: "image",
    mediaId: "m1",
    naturalWidthPx: 100,
    naturalHeightPx: 100,
    widthCm: 4,
    heightCm: 3,
    wrap: "square",
    ...overrides,
  };
}

function chart(overrides: Partial<Chart> = {}): Chart {
  return {
    type: "chart",
    chartType: "columnClustered",
    categories: ["a"],
    series: [{ name: "s", values: [1], color: "#123456" }],
    options: { legend: "none", dataLabels: false, gridlines: true },
    widthCm: 5,
    heightCm: 4,
    wrap: "square",
    ...overrides,
  };
}

const anySelection: ModelSelection = caretSelection({ section: 0, block: 0, offset: 0 });

/** A container of two shapes at explicit member offsets, already normalized. */
function pairContainer(overrides: Partial<DrawingGroup> = {}): DrawingGroup {
  return {
    type: "drawingGroup",
    widthCm: 8,
    heightCm: 5,
    wrap: "square",
    offsetXCm: 2,
    offsetYCm: 2,
    members: [
      shape({ widthCm: 3, heightCm: 2, offsetXCm: 0, offsetYCm: 0 }),
      shape({ widthCm: 3, heightCm: 2, offsetXCm: 5, offsetYCm: 3 }),
    ],
    ...overrides,
  };
}

describe("membersUnionBox (SPEC-057 Normalized form)", () => {
  it("spans the members' unrotated boxes", () => {
    const box = membersUnionBox([
      shape({ offsetXCm: 0, offsetYCm: 0, widthCm: 3, heightCm: 2 }),
      shape({ offsetXCm: 5, offsetYCm: 3, widthCm: 3, heightCm: 2 }),
    ]);
    expect(box).toEqual({ xCm: 0, yCm: 0, widthCm: 8, heightCm: 5 });
  });

  it("ignores member rotation for bounds", () => {
    const box = membersUnionBox([shape({ offsetXCm: 1, offsetYCm: 1, widthCm: 2, heightCm: 2, rotationDeg: 45 })]);
    expect(box).toEqual({ xCm: 1, yCm: 1, widthCm: 2, heightCm: 2 });
  });
});

describe("normalizeGroup (SPEC-057 Normalized form)", () => {
  it("folds a non-zero member-offset origin into the container's own offsets", () => {
    const group = normalizeGroup({
      ...pairContainer(),
      members: [
        shape({ offsetXCm: 2, offsetYCm: 1, widthCm: 3, heightCm: 2 }),
        shape({ offsetXCm: 7, offsetYCm: 4, widthCm: 3, heightCm: 2 }),
      ],
    });
    expect(Math.min(...group.members.map((m) => m.offsetXCm ?? -1))).toBe(0);
    expect(Math.min(...group.members.map((m) => m.offsetYCm ?? -1))).toBe(0);
    // Origin (2,1) folded into the container offset (was 2,2).
    expect(group.offsetXCm).toBe(4);
    expect(group.offsetYCm).toBe(3);
    expect(group.widthCm).toBe(8);
    expect(group.heightCm).toBe(5);
  });

  it("does not move members on screen (absolute position preserved)", () => {
    const before = { ...pairContainer(), members: [shape({ offsetXCm: 3, offsetYCm: 2 }), shape({ offsetXCm: 6, offsetYCm: 5 })] };
    const beforeAbs = before.members.map((m) => [(before.offsetXCm ?? 0) + (m.offsetXCm ?? 0), (before.offsetYCm ?? 0) + (m.offsetYCm ?? 0)]);
    const after = normalizeGroup(before);
    const afterAbs = after.members.map((m) => [(after.offsetXCm ?? 0) + (m.offsetXCm ?? 0), (after.offsetYCm ?? 0) + (m.offsetYCm ?? 0)]);
    expect(afterAbs).toEqual(beforeAbs);
  });

  it("is a no-op for a rotated container (keeps its stored box)", () => {
    const rotated = { ...pairContainer({ rotationDeg: 30 }), members: [shape({ offsetXCm: 3, offsetYCm: 2 }), shape({ offsetXCm: 6, offsetYCm: 5 })] };
    const after = normalizeGroup(rotated);
    expect(after.offsetXCm).toBe(rotated.offsetXCm);
    expect(after.widthCm).toBe(rotated.widthCm);
    expect(after.members[0].offsetXCm).toBe(3);
  });
});

describe("scaleMembers (SPEC-057 Interior scaling)", () => {
  it("scales every member's offset and size per axis", () => {
    const scaled = scaleMembers([shape({ offsetXCm: 2, offsetYCm: 1, widthCm: 4, heightCm: 2 })], 2, 3);
    expect(scaled[0]).toMatchObject({ offsetXCm: 4, offsetYCm: 3, widthCm: 8, heightCm: 6 });
  });

  it("leaves member rotation unchanged", () => {
    const scaled = scaleMembers([shape({ rotationDeg: 45, widthCm: 2, heightCm: 2 })], 2, 2);
    expect(scaled[0].rotationDeg).toBe(45);
  });

  it("recurses into nested containers", () => {
    const nested: DrawingGroup = {
      type: "drawingGroup",
      widthCm: 4,
      heightCm: 4,
      wrap: "square",
      offsetXCm: 0,
      offsetYCm: 0,
      members: [shape({ offsetXCm: 1, offsetYCm: 1, widthCm: 2, heightCm: 2 })],
    };
    const scaled = scaleMembers([nested], 2, 2)[0];
    expect(isDrawingGroup(scaled)).toBe(true);
    if (isDrawingGroup(scaled)) {
      expect(scaled.widthCm).toBe(8);
      expect(scaled.members[0]).toMatchObject({ offsetXCm: 2, offsetYCm: 2, widthCm: 4, heightCm: 4 });
    }
  });
});

describe("bakeGroupRotation (SPEC-057 Rotation and the bake transform)", () => {
  it("clears container rotation and preserves centre for a 180° turn", () => {
    // Two members symmetric about the box centre; a 180° bake swaps them and
    // renormalizes back to the same box.
    const group: DrawingGroup = {
      type: "drawingGroup",
      widthCm: 6,
      heightCm: 2,
      wrap: "square",
      offsetXCm: 0,
      offsetYCm: 0,
      rotationDeg: 180,
      members: [
        shape({ offsetXCm: 0, offsetYCm: 0, widthCm: 2, heightCm: 2 }),
        shape({ offsetXCm: 4, offsetYCm: 0, widthCm: 2, heightCm: 2 }),
      ],
    };
    const baked = bakeGroupRotation(group);
    expect(baked.rotationDeg).toBeUndefined();
    expect(baked.widthCm).toBeCloseTo(6);
    expect(baked.heightCm).toBeCloseTo(2);
    // Each member gains 180° and swaps places.
    expect(baked.members[0].rotationDeg).toBe(180);
    expect(baked.members[0].offsetXCm).toBeCloseTo(4);
    expect(baked.members[1].offsetXCm).toBeCloseTo(0);
  });

  it("is a no-op when rotation is zero (but clears the field)", () => {
    const baked = bakeGroupRotation(pairContainer());
    expect(baked.rotationDeg).toBeUndefined();
    expect(baked.members[0].offsetXCm).toBe(0);
  });

  it("deep-bakes nested containers until none carries rotation", () => {
    const nested: DrawingGroup = {
      type: "drawingGroup",
      widthCm: 2,
      heightCm: 2,
      wrap: "square",
      offsetXCm: 0,
      offsetYCm: 0,
      rotationDeg: 90,
      members: [shape({ offsetXCm: 0, offsetYCm: 0, widthCm: 2, heightCm: 2 })],
    };
    const outer: DrawingGroup = {
      type: "drawingGroup",
      widthCm: 4,
      heightCm: 4,
      wrap: "square",
      offsetXCm: 0,
      offsetYCm: 0,
      rotationDeg: 90,
      members: [nested, shape({ offsetXCm: 2, offsetYCm: 2, widthCm: 2, heightCm: 2 })],
    };
    const baked = bakeGroupRotationDeep(outer);
    const anyRotation = (g: DrawingGroup): boolean =>
      (g.rotationDeg ?? 0) !== 0 || g.members.some((m) => isDrawingGroup(m) && anyRotation(m));
    expect(anyRotation(baked)).toBe(false);
  });
});

describe("buildContainer (SPEC-057 Creating a group)", () => {
  it("puts the box on the union of the member rectangles and offsets members within it", () => {
    const entries = [
      { block: shape({ widthCm: 3, heightCm: 2 }), rect: { xCm: 2, yCm: 2, widthCm: 3, heightCm: 2 } },
      { block: shape({ widthCm: 3, heightCm: 2 }), rect: { xCm: 7, yCm: 5, widthCm: 3, heightCm: 2 } },
    ];
    const container = buildContainer(entries);
    expect(container.offsetXCm).toBe(2);
    expect(container.offsetYCm).toBe(2);
    expect(container.widthCm).toBe(8);
    expect(container.heightCm).toBe(5);
    expect(container.members[0]).toMatchObject({ offsetXCm: 0, offsetYCm: 0 });
    expect(container.members[1]).toMatchObject({ offsetXCm: 5, offsetYCm: 3 });
  });

  it("clears member anchor fields and sets the placeholder wrap", () => {
    const entries = [
      { block: shape({ relativeToH: "page", relativeToV: "page", zOrder: 4 }), rect: { xCm: 0, yCm: 0, widthCm: 3, heightCm: 2 } },
      { block: image(), rect: { xCm: 4, yCm: 0, widthCm: 4, heightCm: 3 } },
    ];
    const container = buildContainer(entries);
    for (const member of container.members) {
      expect(member.wrap).toBe("square");
      expect(member.relativeToH).toBeUndefined();
      expect(member.relativeToV).toBeUndefined();
      expect(member.zOrder).toBeUndefined();
    }
  });

  it("orders members by paint order (ascending zOrder)", () => {
    const entries = [
      { block: shape({ zOrder: 5, widthCm: 2, heightCm: 2 }), rect: { xCm: 0, yCm: 0, widthCm: 2, heightCm: 2 } },
      { block: image({ zOrder: 1, widthCm: 2, heightCm: 2 }), rect: { xCm: 3, yCm: 0, widthCm: 2, heightCm: 2 } },
    ];
    const container = buildContainer(entries);
    expect(container.members[0].type).toBe("image");
    expect(container.members[1].type).toBe("shape");
    // Container z-order is the highest member value.
    expect(container.zOrder).toBe(5);
  });
});

describe("groupDrawings / ungroupDrawing (SPEC-057 AC-3..AC-8)", () => {
  it("replaces the selected drawings with one container at the first block position", () => {
    const model = doc(paragraph("a"), shape({ widthCm: 3, heightCm: 2 }), paragraph("b"), image({ widthCm: 4, heightCm: 3 }));
    const entries: GroupEntry[] = [
      { section: 0, block: 1, rect: { xCm: 1, yCm: 1, widthCm: 3, heightCm: 2 } },
      { section: 0, block: 3, rect: { xCm: 6, yCm: 4, widthCm: 4, heightCm: 3 } },
    ];
    const result = groupDrawings(model, anySelection, entries);
    const blocks = result.model.sections[0].blocks;
    // The container lands at index 1 (first member's position); the image is
    // removed from index 3.
    expect(blocks[1].type).toBe("drawingGroup");
    const container = blocks[1] as DrawingGroup;
    expect(container.members).toHaveLength(2);
    // No member moved: absolute positions equal the input rectangles.
    const abs = container.members.map((m) => [(container.offsetXCm ?? 0) + (m.offsetXCm ?? 0), (container.offsetYCm ?? 0) + (m.offsetYCm ?? 0)]);
    expect(abs).toContainEqual([1, 1]);
    expect(abs).toContainEqual([6, 4]);
  });

  it("ungroups one level, re-emerging members at unchanged geometry", () => {
    const model = doc(paragraph("a"), pairContainer());
    const result = ungroupDrawing(model, anySelection, 0, 1);
    const blocks = result.model.sections[0].blocks;
    // Two members spliced in at the container's position.
    expect(blocks).toHaveLength(3);
    expect(blocks[1].type).toBe("shape");
    expect(blocks[2].type).toBe("shape");
    // Member 1's absolute position was (2,2); member 2's was (7,5).
    expect(blocks[1]).toMatchObject({ offsetXCm: 2, offsetYCm: 2 });
    expect(blocks[2]).toMatchObject({ offsetXCm: 7, offsetYCm: 5 });
  });

  it("group then ungroup preserves each member's absolute geometry", () => {
    const model = doc(shape({ widthCm: 3, heightCm: 2 }), image({ widthCm: 4, heightCm: 3 }));
    const entries: GroupEntry[] = [
      { section: 0, block: 0, rect: { xCm: 1, yCm: 1, widthCm: 3, heightCm: 2 } },
      { section: 0, block: 1, rect: { xCm: 6, yCm: 4, widthCm: 4, heightCm: 3 } },
    ];
    const grouped = groupDrawings(model, anySelection, entries);
    const ungrouped = ungroupDrawing(grouped.model, anySelection, 0, 0);
    const blocks = ungrouped.model.sections[0].blocks;
    const positions = blocks.map((b) => [(b as DrawingMember).offsetXCm, (b as DrawingMember).offsetYCm]);
    expect(positions).toContainEqual([1, 1]);
    expect(positions).toContainEqual([6, 4]);
  });

  it("nests a selected container as one member; ungroup peels one level", () => {
    const model = doc(pairContainer(), shape({ widthCm: 2, heightCm: 2 }));
    const entries: GroupEntry[] = [
      { section: 0, block: 0, rect: { xCm: 2, yCm: 2, widthCm: 8, heightCm: 5 } },
      { section: 0, block: 1, rect: { xCm: 12, yCm: 2, widthCm: 2, heightCm: 2 } },
    ];
    const grouped = groupDrawings(model, anySelection, entries);
    const outer = grouped.model.sections[0].blocks[0] as DrawingGroup;
    // A member is still the inner container.
    expect(outer.members.some((m) => isDrawingGroup(m))).toBe(true);
    // Ungroup peels one level: the inner container re-emerges as a container.
    const ungrouped = ungroupDrawing(grouped.model, anySelection, 0, 0);
    expect(ungrouped.model.sections[0].blocks.some((b) => isDrawingGroup(b))).toBe(true);
  });
});

describe("interior member edits (SPEC-057 AC-26..AC-36)", () => {
  it("moves a member and renormalizes with no visual jump", () => {
    const model = doc(pairContainer());
    // Move the first member up-left by (1,1): its absolute pos was (2,2) -> (1,1).
    const result = moveMember(model, anySelection, { section: 0, block: 0, path: [0] }, -1, -1);
    const container = result.model.sections[0].blocks[0] as DrawingGroup;
    const firstAbs = [(container.offsetXCm ?? 0) + (container.members[0].offsetXCm ?? 0), (container.offsetYCm ?? 0) + (container.members[0].offsetYCm ?? 0)];
    expect(firstAbs).toEqual([1, 1]);
    // Still normalized.
    expect(Math.min(...container.members.map((m) => m.offsetXCm ?? 1))).toBe(0);
  });

  it("resizes a member floored at the minimum", () => {
    const model = doc(pairContainer());
    const result = resizeMember(model, anySelection, { section: 0, block: 0, path: [0] }, 0.01, 0.01);
    const container = result.model.sections[0].blocks[0] as DrawingGroup;
    const target = container.members.find((m) => m.widthCm < 1);
    expect(target?.widthCm).toBe(0.2);
  });

  it("deletes a member; a container reduced to one member dissolves into it (AC-32)", () => {
    const model = doc(pairContainer());
    const result = deleteMember(model, anySelection, { section: 0, block: 0, path: [1] });
    // Container had 2 members; deleting one leaves a lone member -> dissolved.
    expect(result.model.sections[0].blocks[0].type).toBe("shape");
  });

  it("duplicates a member as a sibling after its source", () => {
    const model = doc(pairContainer());
    const result = duplicateMember(model, anySelection, { section: 0, block: 0, path: [0] });
    const container = result.model.sections[0].blocks[0] as DrawingGroup;
    expect(container.members).toHaveLength(3);
  });

  it("reorders a member within its siblings (send to back)", () => {
    const model = doc(pairContainer());
    const result = reorderMember(model, anySelection, { section: 0, block: 0, path: [1] }, "back");
    const container = result.model.sections[0].blocks[0] as DrawingGroup;
    // The second member (offsetXCm 5) moved to index 0; offsets are unchanged
    // — inside a container paint order is member order.
    expect(container.members[0].offsetXCm).toBe(5);
  });
});

describe("container resize/rotate (SPEC-057 AC-19/AC-20)", () => {
  it("resizing scales every member per axis", () => {
    const model = doc(pairContainer());
    const result = resizeGroupContainer(model, anySelection, 0, 0, 16, 10);
    const container = result.model.sections[0].blocks[0] as DrawingGroup;
    expect(container.widthCm).toBe(16);
    expect(container.heightCm).toBe(10);
    // Second member was at offset (5,3), size (3,2); scaled x2 -> (10,6),(6,4).
    expect(container.members[1]).toMatchObject({ offsetXCm: 10, offsetYCm: 6, widthCm: 6, heightCm: 4 });
  });

  it("rotating sets the container's own rotationDeg without touching members", () => {
    const model = doc(pairContainer());
    const result = rotateGroupContainer(model, anySelection, 0, 0, 45);
    const container = result.model.sections[0].blocks[0] as DrawingGroup;
    expect(container.rotationDeg).toBe(45);
    expect(container.members[0].rotationDeg).toBeUndefined();
  });
});

describe("memberAtPath", () => {
  it("resolves a nested member by index path", () => {
    const inner = pairContainer();
    const outer: DrawingGroup = {
      type: "drawingGroup",
      widthCm: 10,
      heightCm: 6,
      wrap: "square",
      offsetXCm: 0,
      offsetYCm: 0,
      members: [inner, shape({ widthCm: 2, heightCm: 2 })],
    };
    expect(memberAtPath(outer, [0, 1])?.type).toBe("shape");
    expect(memberAtPath(outer, [1])?.type).toBe("shape");
    expect(memberAtPath(outer, [5])).toBeNull();
  });
});

describe("migrateLegacyGroups (SPEC-057 Legacy migration)", () => {
  it("collapses drawings sharing an ordinal into one container at the first block index", () => {
    const blocks: Block[] = [
      paragraph("a"),
      shape({ widthCm: 3, heightCm: 2, offsetXCm: 1, offsetYCm: 1, groupOrdinal: 7 }),
      paragraph("b"),
      shape({ widthCm: 3, heightCm: 2, offsetXCm: 6, offsetYCm: 4, groupOrdinal: 7 }),
    ];
    migrateLegacyGroups(blocks);
    const container = blocks.find((b) => isDrawingGroup(b)) as DrawingGroup;
    expect(container).toBeTruthy();
    expect(container.members).toHaveLength(2);
    // No stray groupOrdinal survives on members.
    for (const member of container.members) {
      expect((member as { groupOrdinal?: number }).groupOrdinal).toBeUndefined();
    }
  });

  it("leaves a lone-ordinal drawing ungrouped and strips its tag", () => {
    const blocks: Block[] = [shape({ groupOrdinal: 3 }), paragraph("x")];
    migrateLegacyGroups(blocks);
    expect(blocks[0].type).toBe("shape");
    expect((blocks[0] as { groupOrdinal?: number }).groupOrdinal).toBeUndefined();
  });

  it("migrates a chart member's ordinal too", () => {
    const blocks: Block[] = [
      shape({ groupOrdinal: 2, offsetXCm: 0, offsetYCm: 0 }),
      chart({ groupOrdinal: 2, offsetXCm: 5, offsetYCm: 0 }),
    ];
    migrateLegacyGroups(blocks);
    const container = blocks.find((b) => isDrawingGroup(b)) as DrawingGroup;
    expect(container.members.map((m) => m.type).sort()).toEqual(["chart", "shape"]);
  });
});
