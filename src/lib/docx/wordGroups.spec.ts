// SPEC-057 — word-groups.docx is the hand-built native `wpg:wgp` fixture
// (scripts/build-spec057-fixtures.mjs). Opening it exercises the group reader
// against genuine native markup: a plain pair (with a non-identity child
// coordinate space), a nested group, a mixed shape/image/text-box group, a
// rotated container, and a preserved-member (sealed) group. AC-53..58, AC-67.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx } from "./reader";
import { isDrawingGroup, isImage, isSealedGroup, isShape, type DrawingGroup } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

async function openGroups(): Promise<DrawingGroup[]> {
  const parsed = await parseDocx(new Uint8Array(readFileSync(join(fixturesDir, "word-groups.docx"))));
  return parsed.sections.flatMap((s) => s.blocks).filter(isDrawingGroup);
}

describe("SPEC-057 — opening word-groups.docx", () => {
  it("yields five native containers in document order (AC-53)", async () => {
    const groups = await openGroups();
    expect(groups).toHaveLength(5);
    expect(groups.every((g) => g.members.length >= 2)).toBe(true);
  });

  it("group 0 is the plain pair, read through a non-identity child space (AC-58)", async () => {
    const [plain] = await openGroups();
    expect(plain.altText).toBe("Plain pair");
    expect(plain.members).toHaveLength(2);
    expect(plain.members.every(isShape)).toBe(true);
    // child space is 2× the box, so a 4-unit child ext renders as a 2cm member.
    expect((plain.members[0] as { widthCm: number }).widthCm).toBeCloseTo(2, 4);
    expect((plain.members[1] as { offsetXCm?: number }).offsetXCm).toBeCloseTo(3, 4);
  });

  it("group 1 nests a container as one member (AC-56)", async () => {
    const nested = (await openGroups())[1];
    expect(nested.members).toHaveLength(2);
    const inner = nested.members.find(isDrawingGroup) as DrawingGroup;
    expect(inner).toBeDefined();
    expect(inner.members).toHaveLength(2);
  });

  it("group 2 mixes a shape, an image and a text box (AC-54)", async () => {
    const mixed = (await openGroups())[2];
    expect(mixed.members.some(isImage)).toBe(true);
    expect(mixed.members.some((m) => isShape(m) && m.text !== undefined)).toBe(true);
    const image = mixed.members.find(isImage)!;
    expect(image.mediaId).toContain("image1.png");
  });

  it("group 3 carries container rotation (AC-57)", async () => {
    const rotated = (await openGroups())[3];
    expect(rotated.rotationDeg).toBe(30);
  });

  it("group 4 holds a preserved member and so is sealed (AC-67)", async () => {
    const preserved = (await openGroups())[4];
    expect(isSealedGroup(preserved)).toBe(true);
    expect(preserved.members.some((m) => isShape(m) && m.unsupportedGeometryXml !== undefined)).toBe(true);
  });
});
