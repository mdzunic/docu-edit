// SPEC-070: connector structure commands + attachment book-keeping.
import { describe, expect, it } from "vitest";
import {
  connectorAt,
  deleteConnector,
  insertConnector,
  reindexAttachments,
  resolveConnectorEndpoints,
  routeConnectorInSection,
  setConnectorAdjustment,
  setConnectorArrowheads,
  setConnectorEndpoint,
  setConnectorKind,
  deleteMapping,
  insertMapping,
  type NewConnector,
} from "./connectors";
import { deleteShape, insertShape } from "./shapes";
import { collectDropRecords } from "./opaque";
import { caretSelection, type ModelSelection } from "./commands";
import {
  createDocument,
  isConnector,
  type Connector,
  type DocumentModel,
  type Paragraph,
  type Section,
  type Shape,
} from "./documentModel";

function paragraph(text = ""): Paragraph {
  return { type: "paragraph", runs: text ? [{ text, formatting: {} }] : [] };
}

function shapeAtBox(x: number, y: number, w = 4, h = 2): Shape {
  return { type: "shape", geometry: "rect", widthCm: w, heightCm: h, wrap: "square", offsetXCm: x, offsetYCm: y };
}

function doc(...blocks: DocumentModel["sections"][0]["blocks"]): DocumentModel {
  const model = createDocument("connector-fixture");
  model.sections = [{ blocks }];
  return model;
}

function sel(block: number): ModelSelection {
  return caretSelection({ section: 0, block, offset: 0 });
}

/** Two shapes at blocks 0 and 1, a connector at block 2 attached start->shape0
 * right-mid (site 1), end->shape1 left-mid (site 3). */
function attachedFixture(): { model: DocumentModel; section: Section } {
  const a = shapeAtBox(0, 0); // right-mid at (4,1)
  const b = shapeAtBox(10, 0); // left-mid at (10,1)
  const connector: Connector = {
    type: "connector",
    kind: "straight",
    widthCm: 6,
    heightCm: 2,
    wrap: "square",
    offsetXCm: 4,
    offsetYCm: 0,
    start: { xCm: 0, yCm: 1 },
    end: { xCm: 6, yCm: 1 },
    startAttachment: { target: { block: 0 }, site: 1 },
    endAttachment: { target: { block: 1 }, site: 3 },
    endArrow: "triangle",
  };
  const model = doc(a, b, connector);
  return { model, section: model.sections[0] };
}

describe("resolveConnectorEndpoints", () => {
  it("resolves attached endpoints to their site points", () => {
    const { section } = attachedFixture();
    const connector = section.blocks[2] as Connector;
    const { start, end } = resolveConnectorEndpoints(section, connector);
    expect(start.point).toEqual({ x: 4, y: 1 }); // shape0 right-mid
    expect(end.point).toEqual({ x: 10, y: 1 }); // shape1 left-mid
  });

  it("re-routes when an attached shape moves (geometry is derived)", () => {
    const { section } = attachedFixture();
    const connector = section.blocks[2] as Connector;
    (section.blocks[1] as Shape).offsetXCm = 20; // move shape1 right
    const { end } = resolveConnectorEndpoints(section, connector);
    expect(end.point).toEqual({ x: 20, y: 1 });
  });

  it("falls back to the free position for a detached end", () => {
    const { section } = attachedFixture();
    const connector = section.blocks[2] as Connector;
    delete connector.endAttachment;
    const { end } = resolveConnectorEndpoints(section, connector);
    expect(end.point).toEqual({ x: 10, y: 1 }); // origin(4,0)+end(6,1)
  });

  it("clips a centre attachment to the shape outline", () => {
    const a = shapeAtBox(0, 0, 4, 2); // centre (2,1)
    const connector: Connector = {
      type: "connector",
      kind: "straight",
      widthCm: 6,
      heightCm: 2,
      wrap: "square",
      offsetXCm: 0,
      offsetYCm: 0,
      start: { xCm: 2, yCm: 1 },
      end: { xCm: 10, yCm: 1 },
      startAttachment: { target: { block: 0 }, site: 8 },
    };
    const model = doc(a, connector);
    const { start } = resolveConnectorEndpoints(model.sections[0], connector);
    // Clipped to the right edge of the 4x2 rect at (4,1), not the centre.
    expect(start.point.x).toBeCloseTo(4, 6);
    expect(start.point.y).toBeCloseTo(1, 6);
  });
});

describe("routeConnectorInSection", () => {
  it("routes a straight connector between its resolved endpoints", () => {
    const { section } = attachedFixture();
    const path = routeConnectorInSection(section, section.blocks[2] as Connector);
    expect(path.points).toEqual([
      { x: 4, y: 1 },
      { x: 10, y: 1 },
    ]);
  });
});

describe("reindexAttachments", () => {
  it("shifts surviving targets and detaches a removed one", () => {
    const { section } = attachedFixture();
    // Simulate deleting block 0: block 0 removed, 1->0, 2->1.
    const mapping = new Map([
      [0, -1],
      [1, 0],
      [2, 1],
    ]);
    const next = reindexAttachments(section, mapping);
    const connector = next.blocks[2] as Connector;
    expect(connector.startAttachment).toBeUndefined(); // detached (target gone)
    expect(connector.endAttachment?.target.block).toBe(0); // shifted 1->0
    // Frozen at the last routed position (shape0 right-mid was (4,1)).
    const origin = { x: connector.offsetXCm ?? 0, y: connector.offsetYCm ?? 0 };
    expect(origin.x + connector.start.xCm).toBeCloseTo(4, 6);
    expect(origin.y + connector.start.yCm).toBeCloseTo(1, 6);
  });

  it("leaves the input section untouched", () => {
    const { section } = attachedFixture();
    const before = JSON.parse(JSON.stringify(section));
    reindexAttachments(section, new Map([[0, -1]]));
    expect(section).toEqual(before);
  });
});

describe("insert/delete mappings", () => {
  it("insertMapping shifts blocks at or after atBlock", () => {
    const m = insertMapping(1, 2, 4);
    expect(m.get(0)).toBeUndefined();
    expect(m.get(1)).toBe(3);
    expect(m.get(3)).toBe(5);
  });

  it("deleteMapping removes the deleted block and shifts the rest", () => {
    const m = deleteMapping(1, 1, 4);
    expect(m.get(0)).toBeUndefined();
    expect(m.get(1)).toBe(-1);
    expect(m.get(2)).toBe(1);
    expect(m.get(3)).toBe(2);
  });
});

describe("insertConnector / deleteConnector", () => {
  it("inserts a floating connector and can delete it", () => {
    const model = doc(paragraph("hello"));
    const spec: NewConnector = {
      kind: "elbow",
      widthCm: 5,
      heightCm: 3,
      offsetXCm: 1,
      offsetYCm: 1,
      start: { xCm: 0, yCm: 0 },
      end: { xCm: 5, yCm: 3 },
    };
    const inserted = insertConnector(model, sel(0), spec);
    const connectors = inserted.model.sections[0].blocks.filter(isConnector);
    expect(connectors).toHaveLength(1);
    expect(connectors[0].kind).toBe("elbow");
    expect(connectors[0].endArrow).toBe("triangle"); // default arrow at the end

    const location = { section: 0, block: inserted.model.sections[0].blocks.findIndex(isConnector) };
    const deleted = deleteConnector(inserted.model, inserted.selection, location);
    expect(deleted.model.sections[0].blocks.filter(isConnector)).toHaveLength(0);
  });
});

describe("deleting a shape freezes its connectors (AC-11)", () => {
  it("keeps the connector, detaching the affected endpoint at its last position", () => {
    const { model } = attachedFixture();
    // Delete shape at block 0.
    const result = deleteShape(model, sel(0), { section: 0, block: 0 });
    const connectors = result.model.sections[0].blocks.filter(isConnector);
    expect(connectors).toHaveLength(1); // connector survives
    const connector = connectors[0];
    expect(connector.startAttachment).toBeUndefined(); // start detached
    expect(connector.endAttachment).toBeDefined(); // end still attached
    // The end target shifted from block 1 to block 0.
    expect(connector.endAttachment?.target.block).toBe(0);
    // The detached start freezes at shape0's LAST routed position (right-mid
    // (4,1)) — not at shape1's site, which shifted into the deleted slot. A
    // freeze against the post-splice blocks would land it at (14,1).
    const origin = { x: connector.offsetXCm ?? 0, y: connector.offsetYCm ?? 0 };
    expect(origin.x + connector.start.xCm).toBeCloseTo(4, 6);
    expect(origin.y + connector.start.yCm).toBeCloseTo(1, 6);
  });
});

describe("inserting a shape shifts attachment indices", () => {
  it("keeps a connector pointing at the same shape after an insert before it", () => {
    const { model } = attachedFixture();
    // Insert a new shape at the very start (block 0), pushing everything down.
    const result = insertShape(model, sel(0), { geometry: "rect", widthCm: 2, heightCm: 2, offsetXCm: 0, offsetYCm: 5 });
    const connector = result.model.sections[0].blocks.find(isConnector)!;
    // start was block0, end was block1; both shift up by the inserted count.
    const shapeIndices = result.model.sections[0].blocks
      .map((b, i) => (b.type === "shape" ? i : -1))
      .filter((i) => i >= 0);
    // The two original shapes are still targeted (their new indices).
    expect(connector.startAttachment).toBeDefined();
    expect(connector.endAttachment).toBeDefined();
    expect(shapeIndices).toContain(connector.startAttachment!.target.block);
    expect(shapeIndices).toContain(connector.endAttachment!.target.block);
  });
});

describe("cross-format unexpressible attachment (AC-32)", () => {
  it("reports a shapeGeometry DropRecord for a connector attached to a non-shape on .odt save", () => {
    const image = { type: "image" as const, mediaId: "m1", naturalWidthPx: 10, naturalHeightPx: 10, widthCm: 3, heightCm: 3, wrap: "square" as const };
    const connector: Connector = {
      type: "connector",
      kind: "straight",
      widthCm: 5,
      heightCm: 2,
      wrap: "square",
      start: { xCm: 0, yCm: 0 },
      end: { xCm: 5, yCm: 2 },
      startAttachment: { target: { block: 0 }, site: 1 },
    };
    const model = doc(image, connector);
    model.format = "docx";
    const dropped = collectDropRecords(model, "odt");
    expect(dropped.some((d) => d.kind === "shapeGeometry")).toBe(true);
  });

  it("does not report a drop for a connector attached to a shape (both formats express it)", () => {
    const { model } = attachedFixture();
    model.format = "docx";
    const dropped = collectDropRecords(model, "odt");
    expect(dropped.some((d) => d.kind === "shapeGeometry")).toBe(false);
  });
});

describe("endpoint + formatting commands", () => {
  it("detaches an endpoint, freezing at the routed position", () => {
    const { model } = attachedFixture();
    const result = setConnectorEndpoint(model, sel(2), { section: 0, block: 2 }, "end", { attachment: null });
    const connector = connectorAt(result.model, { section: 0, block: 2 })!;
    expect(connector.endAttachment).toBeUndefined();
    const origin = { x: connector.offsetXCm ?? 0, y: connector.offsetYCm ?? 0 };
    expect(origin.x + connector.end.xCm).toBeCloseTo(10, 6);
  });

  it("re-attaches an endpoint to a different site", () => {
    const { model } = attachedFixture();
    const result = setConnectorEndpoint(model, sel(2), { section: 0, block: 2 }, "end", {
      attachment: { target: { block: 1 }, site: 0 },
    });
    const connector = connectorAt(result.model, { section: 0, block: 2 })!;
    expect(connector.endAttachment?.site).toBe(0);
  });

  it("sets kind, arrowheads and adjustment", () => {
    const { model } = attachedFixture();
    let r = setConnectorKind(model, sel(2), { section: 0, block: 2 }, "curved");
    r = setConnectorArrowheads(r.model, r.selection, { section: 0, block: 2 }, { startArrow: "open", endArrow: "none" });
    r = setConnectorAdjustment(r.model, r.selection, { section: 0, block: 2 }, 0.25);
    const connector = connectorAt(r.model, { section: 0, block: 2 })!;
    expect(connector.kind).toBe("curved");
    expect(connector.startArrow).toBe("open");
    expect(connector.endArrow).toBe("none");
    expect(connector.adjustments).toEqual([0.25]);
  });
});
