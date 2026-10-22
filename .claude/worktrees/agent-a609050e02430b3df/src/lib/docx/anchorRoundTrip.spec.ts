// SPEC-078: `.docx` round trips of the explicit anchor pairing — a floating
// shape/chart with a resolvable `anchorOrdinal` writes its w:drawing run
// INSIDE its anchor paragraph's w:p (containment, the native encoding), the
// reader mints the pairing back from that containment, and ordinal-less
// models keep the legacy own-paragraph form byte-shape-identically.
// Mirrors chartRoundTrip.spec.ts's writeDocx/parseDocx conventions.
import { describe, it, expect } from "vitest";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";
import { unzip } from "../core/zip";
import { equalStructural } from "../equality";
import { resolveAnchorBlockIndex } from "../drawingAnchors";
import { createDocument, isShape, type Block, type DocumentModel, type Paragraph, type Shape } from "../documentModel";

function paragraph(text: string, anchorOrdinal?: number): Paragraph {
  const block: Paragraph = { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
  if (anchorOrdinal !== undefined) {
    block.anchorOrdinal = anchorOrdinal;
  }
  return block;
}

/** Cycle-stable floating shape: zOrder 0 because the writer's
 * relativeHeight always reads back as an explicit zOrder, and
 * relativeToH/relativeToV because the reader sets them on every non-inline
 * shape. */
function floatingShape(overrides: Partial<Shape> = {}): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 3,
    heightCm: 2,
    wrap: "square",
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: 1,
    offsetYCm: 1.5,
    zOrder: 0,
    ...overrides,
  };
}

function model(...blocks: Block[]): DocumentModel {
  const base = createDocument("anchor-cycle-doc");
  base.sections = [{ blocks }];
  return base;
}

async function documentXmlOf(bytes: Uint8Array): Promise<string> {
  const parts = await unzip(bytes);
  const documentPart = parts.get("word/document.xml");
  if (!documentPart) {
    throw new Error("no word/document.xml in the package");
  }
  return new TextDecoder().decode(documentPart);
}

/** The `<w:p>…</w:p>` chunks of the body, for asserting which paragraph a
 * drawing landed inside. */
function paragraphChunks(xml: string): string[] {
  return xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [];
}

describe("writeDocx — containment (SPEC-078)", () => {
  it("writes two shapes anchored to one paragraph inside that paragraph's own w:p", async () => {
    const source = model(
      paragraph("Anchor", 1),
      floatingShape({ anchorOrdinal: 1 }),
      floatingShape({ anchorOrdinal: 1, geometry: "ellipse", offsetYCm: 4 }),
      paragraph("After"),
    );
    const { bytes } = await writeDocx(source);
    const xml = await documentXmlOf(bytes);
    const chunks = paragraphChunks(xml);
    const anchorChunk = chunks.find((chunk) => chunk.includes("Anchor"));
    expect(anchorChunk).toBeDefined();
    expect(anchorChunk!.match(/<wp:anchor /g)).toHaveLength(2);
    // The runs come after the paragraph's own text (both drawings follow
    // their anchor in the block list), and no drawing-only paragraph exists.
    expect(anchorChunk!.indexOf("Anchor")).toBeLessThan(anchorChunk!.indexOf("<wp:anchor "));
    expect(chunks.some((chunk) => chunk.startsWith("<w:p><w:r><w:drawing"))).toBe(false);
  });

  it("emits a drawing that PRECEDES its anchor before the paragraph's runs", async () => {
    const source = model(floatingShape({ anchorOrdinal: 1 }), paragraph("Anchor", 1));
    const { bytes } = await writeDocx(source);
    const anchorChunk = paragraphChunks(await documentXmlOf(bytes)).find((chunk) => chunk.includes("Anchor"));
    expect(anchorChunk).toBeDefined();
    expect(anchorChunk!.indexOf("<wp:anchor ")).toBeLessThan(anchorChunk!.indexOf("Anchor"));
  });

  it("keeps the legacy own-paragraph form for an ordinal-less shape, and for a dangling ordinal", async () => {
    for (const legacyShape of [floatingShape(), floatingShape({ anchorOrdinal: 99 })]) {
      const { bytes } = await writeDocx(model(paragraph("Anchor"), legacyShape));
      const chunks = paragraphChunks(await documentXmlOf(bytes));
      const anchorChunk = chunks.find((chunk) => chunk.includes("Anchor"));
      expect(anchorChunk).toBeDefined();
      expect(anchorChunk!.includes("<wp:anchor ")).toBe(false);
      expect(chunks.some((chunk) => chunk.startsWith("<w:p><w:r><w:drawing"))).toBe(true);
    }
  });
});

describe("parseDocx — pairing minted from containment (SPEC-078)", () => {
  it("round-trips two shapes on one paragraph with pairing, offsets and resolution intact", async () => {
    const source = model(
      paragraph("Anchor", 1),
      floatingShape({ anchorOrdinal: 1 }),
      floatingShape({ anchorOrdinal: 1, geometry: "ellipse", offsetYCm: 4 }),
      paragraph("After"),
    );
    const { bytes } = await writeDocx(source);
    const reopened = await parseDocx(bytes);
    const blocks = reopened.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "shape", "shape", "paragraph"]);
    const anchor = blocks[0] as Paragraph;
    expect(anchor.anchorOrdinal).toBe(1);
    expect((blocks[1] as Shape).anchorOrdinal).toBe(1);
    expect((blocks[2] as Shape).anchorOrdinal).toBe(1);
    expect((blocks[1] as Shape).offsetYCm).toBeCloseTo(1.5, 5);
    expect((blocks[2] as Shape).offsetYCm).toBeCloseTo(4, 5);
    expect((blocks[3] as Paragraph).anchorOrdinal).toBeUndefined();
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(0);
    expect(resolveAnchorBlockIndex(blocks, 2)).toBe(0);
  });

  it("preserves the block order of a drawing that precedes its anchor", async () => {
    const source = model(floatingShape({ anchorOrdinal: 1 }), paragraph("Anchor", 1));
    const { bytes } = await writeDocx(source);
    const blocks = (await parseDocx(bytes)).sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["shape", "paragraph"]);
    expect((blocks[0] as Shape).anchorOrdinal).toBe(1);
    expect((blocks[1] as Paragraph).anchorOrdinal).toBe(1);
  });

  it("reads a legacy own-paragraph shape unpaired, resolving by adjacency as before", async () => {
    const { bytes } = await writeDocx(model(paragraph("Anchor"), floatingShape()));
    const blocks = (await parseDocx(bytes)).sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "shape"]);
    expect((blocks[0] as Paragraph).anchorOrdinal).toBeUndefined();
    expect((blocks[1] as Shape).anchorOrdinal).toBeUndefined();
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(0);
  });

  it("absorbs into a LIST paragraph too — no ODT-style carve-out in .docx", async () => {
    const listAnchor = paragraph("Item", 1);
    listAnchor.list = { type: "bullet", level: 1 };
    const source = model(listAnchor, floatingShape({ anchorOrdinal: 1 }));
    const { bytes } = await writeDocx(source);
    const anchorChunk = paragraphChunks(await documentXmlOf(bytes)).find((chunk) => chunk.includes("Item"));
    expect(anchorChunk).toBeDefined();
    expect(anchorChunk!.includes("<wp:anchor ")).toBe(true);
    const blocks = (await parseDocx(bytes)).sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "shape"]);
    expect((blocks[0] as Paragraph).anchorOrdinal).toBe(1);
    expect((blocks[0] as Paragraph).list?.type).toBe("bullet");
    expect((blocks[1] as Shape).anchorOrdinal).toBe(1);
  });

  it("stays structurally stable across a second save-and-reopen", async () => {
    const source = model(
      paragraph("Anchor", 1),
      floatingShape({ anchorOrdinal: 1 }),
      floatingShape({ anchorOrdinal: 1, offsetYCm: 4 }),
      paragraph("After"),
    );
    const first = await parseDocx((await writeDocx(source)).bytes);
    const firstModel: DocumentModel = { ...source, sections: first.sections, defaultStyles: first.defaultStyles };
    const second = await parseDocx((await writeDocx(firstModel)).bytes);
    expect(equalStructural(first.sections, second.sections)).toBe(true);
    expect(first.sections[0].blocks.filter((block) => isShape(block))).toHaveLength(2);
  });
});
