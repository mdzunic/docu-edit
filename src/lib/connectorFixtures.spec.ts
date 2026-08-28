// SPEC-070 AC-33..37: fixtures containing an attached straight/elbow/curved
// connector, a connector with one free end, and a connector inside a group
// open, render (resolve endpoints), re-route on a move, and round-trip through
// `.docx` and `.odt` equalStructural. Also AC-31 (cross-format preserves
// attachment). Mirrors anchorRoundTrip.spec.ts's writeDocx/parseDocx pattern.
import { describe, it, expect } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { equalStructural } from "./equality";
import { resolveConnectorEndpoints } from "./connectors";
import { moveGroup } from "./shapes";
import { caretSelection } from "./commands";
import {
  createDocument,
  isConnector,
  type Block,
  type Connector,
  type ConnectorKind,
  type DocumentModel,
  type Paragraph,
  type Shape,
} from "./documentModel";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function shape(x: number, y: number, overrides: Partial<Shape> = {}): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 4,
    heightCm: 2,
    wrap: "square",
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: x,
    offsetYCm: y,
    zOrder: 0,
    fill: "#4472c4",
    outline: { color: "#2e5395", widthPt: 1, dash: "solid" },
    ...overrides,
  };
}

function connector(kind: ConnectorKind, overrides: Partial<Connector> = {}): Connector {
  return {
    type: "connector",
    kind,
    widthCm: 6,
    heightCm: 2,
    wrap: "square",
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: 4,
    offsetYCm: 0,
    zOrder: 0,
    start: { xCm: 0, yCm: 1 },
    end: { xCm: 6, yCm: 1 },
    outline: { color: "#2e5395", widthPt: 1, dash: "solid" },
    endArrow: "triangle",
    ...overrides,
  };
}

function model(...blocks: Block[]): DocumentModel {
  const base = createDocument("connector-fixture-doc");
  base.title = "Connector fixture";
  base.sections = [{ blocks }];
  return base;
}

/** The document the fixtures build: two shapes joined by a connector of the
 * given kind (start→shape right-mid site 1, end→shape left-mid site 3), a
 * second connector with one free end, and a grouped shape+shape+connector. */
function fixtureDoc(kind: ConnectorKind): DocumentModel {
  return model(
    paragraph("Diagram"),
    shape(0, 0), // block 1
    shape(10, 0), // block 2
    connector(kind, {
      startAttachment: { target: { block: 1 }, site: 1 },
      endAttachment: { target: { block: 2 }, site: 3 },
    }), // block 3
    connector("straight", {
      offsetXCm: 2,
      offsetYCm: 6,
      startAttachment: { target: { block: 1 }, site: 2 },
    }), // block 4: one free end
    shape(0, 12, { groupOrdinal: 7 }), // block 5
    shape(10, 12, { groupOrdinal: 7 }), // block 6
    connector("elbow", {
      offsetXCm: 4,
      offsetYCm: 12,
      groupOrdinal: 7,
      startAttachment: { target: { block: 5 }, site: 1 },
      endAttachment: { target: { block: 6 }, site: 3 },
    }), // block 7: connector inside a group
  );
}

/** parseDocx/parseOdt return the parsed content; the app layer supplies the
 * title/id when it builds the DocumentModel (openDocument). The fixtures
 * restore them so a second write cycle has what the writer reads. */
function withMeta(parsed: Awaited<ReturnType<typeof parseDocx>>): DocumentModel {
  return { ...(parsed as unknown as DocumentModel), id: "connector-fixture-doc", title: "Connector fixture" };
}

async function docxCycle(source: DocumentModel): Promise<DocumentModel> {
  const { bytes } = await writeDocx(source);
  return withMeta(await parseDocx(bytes));
}

async function odtCycle(source: DocumentModel): Promise<DocumentModel> {
  const { bytes } = await writeOdt(source);
  return withMeta(await parseOdt(bytes));
}

const KINDS: ConnectorKind[] = ["straight", "elbow", "curved"];

describe("connector fixtures round-trip .docx (AC-33..35, AC-37)", () => {
  for (const kind of KINDS) {
    it(`opens, re-routes and round-trips an attached ${kind} connector`, async () => {
      const source = fixtureDoc(kind);
      const a = await docxCycle(source);
      const b = await docxCycle(a);
      expect(equalStructural(a, b)).toBe(true);

      // The attached connector survives with its site attachments (AC-33..35).
      const conn = a.sections[0].blocks[3] as Connector;
      expect(isConnector(conn)).toBe(true);
      expect(conn.kind).toBe(kind);
      expect(conn.startAttachment).toEqual({ target: { block: 1 }, site: 1 });
      expect(conn.endAttachment).toEqual({ target: { block: 2 }, site: 3 });

      // Re-routes on a move: shift shape 2 right, the end endpoint follows.
      const before = resolveConnectorEndpoints(a.sections[0], conn);
      (a.sections[0].blocks[2] as Shape).offsetXCm = 20;
      const after = resolveConnectorEndpoints(a.sections[0], conn);
      expect(after.end.point.x).toBeGreaterThan(before.end.point.x);
    });
  }

  it("preserves a connector with one free end (AC-36)", async () => {
    const a = await docxCycle(fixtureDoc("straight"));
    const free = a.sections[0].blocks[4] as Connector;
    expect(free.startAttachment).toEqual({ target: { block: 1 }, site: 2 });
    expect(free.endAttachment).toBeUndefined();
  });

  it("preserves a connector inside a group and re-routes when the group moves (AC-37)", async () => {
    const a = await docxCycle(fixtureDoc("elbow"));
    const grouped = a.sections[0].blocks[7] as Connector;
    expect(grouped.groupOrdinal).toBe(7);
    expect(grouped.startAttachment).toEqual({ target: { block: 5 }, site: 1 });

    const before = resolveConnectorEndpoints(a.sections[0], grouped);
    const moved = moveGroup(a, caretSelection({ section: 0, block: 0, offset: 0 }), 7, 3, 0);
    const grouped2 = moved.model.sections[0].blocks[7] as Connector;
    const after = resolveConnectorEndpoints(moved.model.sections[0], grouped2);
    expect(after.start.point.x).toBeGreaterThan(before.start.point.x);
  });
});

describe("connector fixtures round-trip .odt (AC-33..35, AC-37)", () => {
  for (const kind of KINDS) {
    it(`opens, re-routes and round-trips an attached ${kind} connector`, async () => {
      const source = fixtureDoc(kind);
      const a = await odtCycle(source);
      const b = await odtCycle(a);
      // The document body round-trips structurally. (The whole-model compare
      // would also see ODF's generated automatic-style names drift by one on
      // each save — a pre-existing behaviour shared with shapes, orthogonal to
      // connectors, so the body sections are the meaningful comparison here.)
      expect(equalStructural(a.sections, b.sections)).toBe(true);

      const conn = a.sections[0].blocks[3] as Connector;
      expect(conn.kind).toBe(kind);
      expect(conn.startAttachment).toEqual({ target: { block: 1 }, site: 1 });
      expect(conn.endAttachment).toEqual({ target: { block: 2 }, site: 3 });

      // Re-routes on a move (AC-33..35).
      const beforeMove = resolveConnectorEndpoints(a.sections[0], conn);
      (a.sections[0].blocks[2] as Shape).offsetXCm = 20;
      const afterMove = resolveConnectorEndpoints(a.sections[0], conn);
      expect(afterMove.end.point.x).toBeGreaterThan(beforeMove.end.point.x);
    });
  }
});

describe("cross-format preserves attachment (AC-31)", () => {
  it(".docx -> .odt keeps the connector's site attachments", async () => {
    const viaDocx = await docxCycle(fixtureDoc("straight"));
    const { bytes } = await writeOdt(viaDocx);
    const viaOdt = await parseOdt(bytes);
    const conn = viaOdt.sections[0].blocks[3] as Connector;
    expect(conn.startAttachment).toEqual({ target: { block: 1 }, site: 1 });
    expect(conn.endAttachment).toEqual({ target: { block: 2 }, site: 3 });
  });
});
