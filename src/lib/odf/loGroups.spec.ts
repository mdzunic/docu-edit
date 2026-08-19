// SPEC-057 — lo-groups.odt is the hand-built native `draw:g` fixture
// (scripts/build-spec057-fixtures.mjs). Opening it exercises the group reader
// against genuine native markup: a plain pair, a nested group, a mixed
// shape/image/text-box group, a group whose rotation is baked into member
// positions, and a preserved-member (sealed) group. AC-60..62, AC-67.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseOdt } from "./reader";
import { isDrawingGroup, isImage, isSealedGroup, isShape, type DrawingGroup } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

async function openGroups(): Promise<DrawingGroup[]> {
  const parsed = await parseOdt(new Uint8Array(readFileSync(join(fixturesDir, "lo-groups.odt"))));
  return parsed.sections.flatMap((s) => s.blocks).filter(isDrawingGroup);
}

describe("SPEC-057 — opening lo-groups.odt", () => {
  it("yields five native containers in document order (AC-60)", async () => {
    const groups = await openGroups();
    expect(groups).toHaveLength(5);
    expect(groups.every((g) => g.members.length >= 2)).toBe(true);
  });

  it("group 0 is the plain pair (AC-60)", async () => {
    const [plain] = await openGroups();
    expect(plain.altText).toBe("Plain pair");
    expect(plain.members.every(isShape)).toBe(true);
    expect((plain.members[0] as { offsetXCm?: number }).offsetXCm).toBeCloseTo(0, 4);
    expect((plain.members[1] as { offsetXCm?: number }).offsetXCm).toBeCloseTo(3, 4);
  });

  it("group 1 nests a container as one member (AC-61)", async () => {
    const nested = (await openGroups())[1];
    const inner = nested.members.find(isDrawingGroup) as DrawingGroup;
    expect(inner).toBeDefined();
    expect(inner.members).toHaveLength(2);
  });

  it("group 2 mixes a shape, a draw:frame image and a text box (AC-62)", async () => {
    const mixed = (await openGroups())[2];
    expect(mixed.members.some(isImage)).toBe(true);
    expect(mixed.members.some((m) => isShape(m) && m.text !== undefined)).toBe(true);
  });

  it("group 4 holds a preserved member and so is sealed (AC-67)", async () => {
    const preserved = (await openGroups())[4];
    expect(isSealedGroup(preserved)).toBe(true);
  });
});
