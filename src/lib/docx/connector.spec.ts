// SPEC-070: the DrawingML connector (wps:cxnSp) reader/writer and the
// arrow-preset untangle. Mirrors shape.spec.ts's fixture conventions.
import { describe, it, expect } from "vitest";
import {
  connectorDrawingXml,
  parseConnectorDrawing,
  resolveDocxConnectorAttachments,
  stampConnectorTargetId,
  type ConnectorTargetIds,
} from "./connector";
import { parseShapeDrawing, shapeDrawingXml, type ShapeTextContext, type ShapeTextWriteContext } from "./shape";
import { parseXml } from "../core/xml";
import { equalStructural } from "../equality";
import { type Connector, type Paragraph, type Section, type Shape } from "../documentModel";

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function drawingElement(xml: string): Element {
  const wrapped = `<root xmlns:w="${WORD_NS}">${xml}</root>`;
  return parseXml(wrapped).documentElement.children[0];
}

const readCtx: ShapeTextContext = {
  parseParagraph: (): Paragraph => ({ type: "paragraph", runs: [] }),
};
const writeCtx: ShapeTextWriteContext = {
  paragraphXml: (): string => "<w:p/>",
};

function baseConnector(overrides: Partial<Connector> = {}): Connector {
  return {
    type: "connector",
    kind: "straight",
    widthCm: 6,
    heightCm: 2,
    wrap: "square",
    offsetXCm: 1,
    offsetYCm: 2,
    start: { xCm: 0, yCm: 0 },
    end: { xCm: 6, yCm: 2 },
    outline: { color: "#2e5395", widthPt: 1, dash: "solid" },
    endArrow: "triangle",
    ...overrides,
  };
}

function roundTrip(connector: Connector, ids: ConnectorTargetIds = {}): Connector | null {
  const xml = connectorDrawingXml(connector, 9, ids);
  return parseConnectorDrawing(drawingElement(xml))?.connector ?? null;
}

describe("connector docx round-trip", () => {
  it("round-trips a free straight connector idempotently (AC-23, AC-27)", () => {
    // model -> write -> read (A) then A -> write -> read (B): a stable
    // representation round-trips equalStructural (the reader normalises
    // zOrder/relativeTo, so compare read-backs, as the fixture tests do).
    const a = roundTrip(baseConnector());
    expect(a).not.toBeNull();
    const b = roundTrip(a!);
    expect(equalStructural(a, b)).toBe(true);
    expect(a!.start).toEqual({ xCm: 0, yCm: 0 });
    expect(a!.end).toEqual({ xCm: 6, yCm: 2 });
  });

  it("writes the correct preset per kind (AC-23)", () => {
    expect(connectorDrawingXml(baseConnector({ kind: "straight" }), 1, {})).toContain('prst="straightConnector1"');
    expect(connectorDrawingXml(baseConnector({ kind: "elbow" }), 1, {})).toContain('prst="bentConnector3"');
    expect(connectorDrawingXml(baseConnector({ kind: "curved" }), 1, {})).toContain('prst="curvedConnector3"');
  });

  it("round-trips elbow with an adjustment", () => {
    const connector = baseConnector({ kind: "elbow", adjustments: [0.25] });
    const back = roundTrip(connector)!;
    expect(back.kind).toBe("elbow");
    expect(back.adjustments).toEqual([0.25]);
  });

  it("round-trips arrowheads at both ends (AC-18)", () => {
    const connector = baseConnector({ startArrow: "open", endArrow: "none" });
    const back = roundTrip(connector)!;
    expect(back.startArrow).toBe("open");
    expect(back.endArrow).toBe("none");
  });

  it.each([
    ["dash", "dash"],
    ["dot", "sysDot"],
    ["dashDot", "dashDot"],
    ["longDash", "lgDash"],
    ["longDashDot", "lgDashDot"],
  ] as const)("round-trips the %s outline dash through a:prstDash val=%s (SPEC-073 AC-2/AC-11/AC-32)", (model, ooxml) => {
    const connector = baseConnector({ outline: { color: "#2e5395", widthPt: 1, dash: model } });
    const xml = connectorDrawingXml(connector, 1, {});
    expect(xml).toContain(`<a:prstDash val="${ooxml}"/>`);
    expect(roundTrip(connector)!.outline?.dash).toBe(model);
  });

  it("a solid dash writes no a:prstDash", () => {
    expect(connectorDrawingXml(baseConnector(), 1, {})).not.toContain("prstDash");
  });

  it("a preserved glow re-emits BEFORE the modelled outerShdw (CT_EffectList order, AC-36)", () => {
    const shadow = { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 };
    const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const connector = baseConnector({
      shadow,
      unsupportedStyleXml: `<a:glow xmlns:a="${A_NS}" rad="63500"><a:srgbClr val="00FF00"/></a:glow>`,
    });
    const xml = connectorDrawingXml(connector, 1, {});
    expect(xml.match(/<a:effectLst>/g)).toHaveLength(1);
    expect(xml.indexOf("<a:glow")).toBeLessThan(xml.indexOf("<a:outerShdw"));
  });

  it("an edited modelled outline wins over a stale preserved a:ln (AC-36 one-source)", () => {
    const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const connector = baseConnector({
      outline: { color: "#ff0000", widthPt: 2, dash: "longDash" },
      unsupportedStyleXml: `<a:ln xmlns:a="${A_NS}" w="25400"><a:gradFill><a:gsLst/></a:gradFill></a:ln>`,
    });
    const xml = connectorDrawingXml(connector, 1, {});
    expect(xml).not.toContain("gradFill");
    expect(xml).toContain('<a:srgbClr val="FF0000"/>');
    expect(xml).toContain('<a:prstDash val="lgDash"/>');
    expect(xml.match(/<a:ln[ >]/g)).toHaveLength(1);
  });

  it("writes a:stCxn/a:endCxn @id/@idx when attached (AC-24)", () => {
    const connector = baseConnector({
      startAttachment: { target: { block: 0 }, site: 1 },
      endAttachment: { target: { block: 1 }, site: 3 },
    });
    const xml = connectorDrawingXml(connector, 9, { start: 4, end: 5 });
    expect(xml).toContain('<a:stCxn id="4" idx="1"/>');
    expect(xml).toContain('<a:endCxn id="5" idx="3"/>');
  });
});

describe("attachment resolution by docPr id (AC-28)", () => {
  it("resolves stCxn/endCxn ids to the target block index", () => {
    // A section: shape (docPr 4) at block 0, shape (docPr 5) at block 1,
    // connector at block 2 attached start->4 end->5.
    const shapeA: Shape = { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square" };
    const shapeB: Shape = { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square" };
    const connector = baseConnector({
      startAttachment: { target: { block: 0 }, site: 1 },
      endAttachment: { target: { block: 1 }, site: 3 },
    });
    const connectorXml = connectorDrawingXml(connector, 9, { start: 4, end: 5 });
    const parsed = parseConnectorDrawing(drawingElement(connectorXml))!.connector;

    const section: Section = { blocks: [shapeA, shapeB, parsed] };
    // Stamp the docPr ids the two shapes were "read" under.
    stampConnectorTargetId(drawingElement(shapeDrawingXml(shapeA, 4, writeCtx)), shapeA);
    stampConnectorTargetId(drawingElement(shapeDrawingXml(shapeB, 5, writeCtx)), shapeB);

    resolveDocxConnectorAttachments({ sections: [section] });
    expect(parsed.startAttachment).toEqual({ target: { block: 0 }, site: 1 });
    expect(parsed.endAttachment).toEqual({ target: { block: 1 }, site: 3 });
  });
});

describe("arrow preset untangle (AC-29, AC-30)", () => {
  it("writes arrow as prst=line + a:tailEnd, not a connector preset", () => {
    const arrow: Shape = {
      type: "shape",
      geometry: "arrow",
      widthCm: 3,
      heightCm: 0,
      wrap: "inline",
      outline: { color: "#000000", widthPt: 1, dash: "solid" },
    };
    const xml = shapeDrawingXml(arrow, 1, writeCtx);
    expect(xml).toContain('prst="line"');
    expect(xml).not.toContain("straightConnector1");
    expect(xml).toContain('<a:tailEnd type="triangle"/>');
    // And it reads back as an arrow.
    const back = parseShapeDrawing(drawingElement(xml), readCtx)!.shape;
    expect(back.geometry).toBe("arrow");
  });

  it("reads a legacy straightConnector1 wsp back as an arrow (AC-30)", () => {
    const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    const legacy =
      `<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:inline><wp:extent cx="1080000" cy="0"/><wp:docPr id="1" name="Shape 1"/>` +
      `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
      `<wps:wsp xmlns:wps="${WPS_NS}"><wps:spPr>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="1080000" cy="0"/></a:xfrm>` +
      `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>` +
      `</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    const back = parseShapeDrawing(drawingElement(legacy), readCtx)!.shape;
    expect(back.geometry).toBe("arrow");
  });
});
