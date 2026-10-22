// SPEC-078: `.odt` round trips of the explicit anchor pairing — the exact
// twin of docx/anchorRoundTrip.spec.ts: a floating shape with a resolvable
// `anchorOrdinal` writes its draw element INSIDE its anchor paragraph's
// text:p, the reader mints the pairing back from that containment,
// ordinal-less models keep the legacy own-paragraph form, and the ODT-only
// carve-out (a LIST-paragraph anchor) stays own-paragraph with the pairing
// re-derived by adjacency on open.
import { describe, it, expect } from "vitest";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
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

/** Cycle-stable floating shape for ODT: no zOrder (ODT does not carry one)
 * and relativeToH/relativeToV because the reader sets them on every
 * non-inline shape. */
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
    ...overrides,
  };
}

function model(...blocks: Block[]): DocumentModel {
  const base = createDocument("anchor-odt-cycle-doc");
  base.sections = [{ blocks }];
  return base;
}

async function contentXmlOf(bytes: Uint8Array): Promise<string> {
  const parts = await unzip(bytes);
  const contentPart = parts.get("content.xml");
  if (!contentPart) {
    throw new Error("no content.xml in the package");
  }
  return new TextDecoder().decode(contentPart);
}

/** The `<text:p …>…</text:p>` chunks of the body (self-closing paragraphs
 * excluded — a paragraph holding a drawing is never empty), for asserting
 * which paragraph a drawing landed inside. */
function paragraphChunks(xml: string): string[] {
  return xml.match(/<text:p[^/>]*>[\s\S]*?<\/text:p>/g) ?? [];
}

describe("writeOdt — containment (SPEC-078)", () => {
  it("writes two shapes anchored to one paragraph inside that paragraph's own text:p", async () => {
    const source = model(
      paragraph("Anchor", 1),
      floatingShape({ anchorOrdinal: 1 }),
      floatingShape({ anchorOrdinal: 1, geometry: "ellipse", offsetYCm: 4 }),
      paragraph("After"),
    );
    const { bytes } = await writeOdt(source);
    const chunks = paragraphChunks(await contentXmlOf(bytes));
    const anchorChunk = chunks.find((chunk) => chunk.includes("Anchor"));
    expect(anchorChunk).toBeDefined();
    expect(anchorChunk!.match(/<draw:custom-shape/g)).toHaveLength(2);
    expect(anchorChunk!.indexOf("Anchor")).toBeLessThan(anchorChunk!.indexOf("<draw:custom-shape"));
    // No drawing-only carrier paragraph remains.
    const drawingOnly = chunks.filter((chunk) => !chunk.includes("Anchor") && chunk.includes("<draw:custom-shape"));
    expect(drawingOnly).toHaveLength(0);
  });

  it("emits a drawing that PRECEDES its anchor before the paragraph's text", async () => {
    const source = model(floatingShape({ anchorOrdinal: 1 }), paragraph("Anchor", 1));
    const { bytes } = await writeOdt(source);
    const anchorChunk = paragraphChunks(await contentXmlOf(bytes)).find((chunk) => chunk.includes("Anchor"));
    expect(anchorChunk).toBeDefined();
    expect(anchorChunk!.indexOf("<draw:custom-shape")).toBeLessThan(anchorChunk!.indexOf("Anchor"));
  });

  it("keeps the legacy own-paragraph form for an ordinal-less shape", async () => {
    const { bytes } = await writeOdt(model(paragraph("Anchor"), floatingShape()));
    const chunks = paragraphChunks(await contentXmlOf(bytes));
    const anchorChunk = chunks.find((chunk) => chunk.includes("Anchor"));
    expect(anchorChunk).toBeDefined();
    expect(anchorChunk!.includes("<draw:custom-shape")).toBe(false);
    expect(chunks.some((chunk) => !chunk.includes("Anchor") && chunk.includes("<draw:custom-shape"))).toBe(true);
  });

  it("keeps the own-paragraph form for a LIST-paragraph anchor — the ODT carve-out", async () => {
    const listAnchor = paragraph("Item", 1);
    listAnchor.list = { type: "bullet", level: 1 };
    const source = model(listAnchor, floatingShape({ anchorOrdinal: 1 }));
    const { bytes } = await writeOdt(source);
    const xml = await contentXmlOf(bytes);
    const chunks = paragraphChunks(xml);
    const itemChunk = chunks.find((chunk) => chunk.includes("Item"));
    expect(itemChunk).toBeDefined();
    expect(itemChunk!.includes("<draw:custom-shape")).toBe(false);
    expect(chunks.some((chunk) => !chunk.includes("Item") && chunk.includes("<draw:custom-shape"))).toBe(true);
    // Reopened: the pairing is gone (never serialized outside containment)
    // but adjacency still resolves the shape to the list paragraph.
    const blocks = (await parseOdt(bytes)).sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "shape"]);
    expect((blocks[1] as Shape).anchorOrdinal).toBeUndefined();
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(0);
  });
});

describe("parseOdt — pairing minted from containment (SPEC-078)", () => {
  it("round-trips two shapes on one paragraph with pairing, offsets and resolution intact", async () => {
    const source = model(
      paragraph("Anchor", 1),
      floatingShape({ anchorOrdinal: 1 }),
      floatingShape({ anchorOrdinal: 1, geometry: "ellipse", offsetYCm: 4 }),
      paragraph("After"),
    );
    const { bytes } = await writeOdt(source);
    const blocks = (await parseOdt(bytes)).sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "shape", "shape", "paragraph"]);
    expect((blocks[0] as Paragraph).anchorOrdinal).toBe(1);
    expect((blocks[1] as Shape).anchorOrdinal).toBe(1);
    expect((blocks[2] as Shape).anchorOrdinal).toBe(1);
    expect((blocks[1] as Shape).offsetYCm).toBeCloseTo(1.5, 3);
    expect((blocks[2] as Shape).offsetYCm).toBeCloseTo(4, 3);
    expect(resolveAnchorBlockIndex(blocks, 1)).toBe(0);
    expect(resolveAnchorBlockIndex(blocks, 2)).toBe(0);
  });

  it("preserves the block order of a drawing that precedes its anchor", async () => {
    const source = model(floatingShape({ anchorOrdinal: 1 }), paragraph("Anchor", 1));
    const { bytes } = await writeOdt(source);
    const blocks = (await parseOdt(bytes)).sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["shape", "paragraph"]);
    expect((blocks[0] as Shape).anchorOrdinal).toBe(1);
    expect((blocks[1] as Paragraph).anchorOrdinal).toBe(1);
  });

  it("stays structurally stable across a second save-and-reopen", async () => {
    const source = model(
      paragraph("Anchor", 1),
      floatingShape({ anchorOrdinal: 1 }),
      floatingShape({ anchorOrdinal: 1, offsetYCm: 4 }),
      paragraph("After"),
    );
    const first = await parseOdt((await writeOdt(source)).bytes);
    const firstModel: DocumentModel = { ...source, sections: first.sections, defaultStyles: first.defaultStyles };
    const second = await parseOdt((await writeOdt(firstModel)).bytes);
    expect(equalStructural(first.sections, second.sections)).toBe(true);
    expect(first.sections[0].blocks.filter((block) => isShape(block))).toHaveLength(2);
  });
});
