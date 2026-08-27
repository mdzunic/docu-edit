// SPEC-070: the ODF draw:connector reader/writer.
import { describe, it, expect } from "vitest";
import {
  connectorElementXml,
  parseConnectorElement,
  resolveOdfConnectorAttachments,
  stampOdfConnectorTargetId,
  odfDrawId,
} from "./connector";
import { shapeElementXml } from "./shape";
import { parseXml } from "../core/xml";
import { equalStructural } from "../equality";
import { type Connector, type Section, type Shape } from "../documentModel";

const DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

function el(xml: string): Element {
  const wrapped = `<root xmlns:draw="${DRAW_NS}" xmlns:svg="${SVG_NS}" xmlns:text="${TEXT_NS}" xmlns:style="${STYLE_NS}">${xml}</root>`;
  return parseXml(wrapped).documentElement.children[0];
}

function contentDoc(styles: string): Document {
  return parseXml(`<root xmlns:draw="${DRAW_NS}" xmlns:svg="${SVG_NS}" xmlns:style="${STYLE_NS}">${styles}</root>`);
}

function baseConnector(overrides: Partial<Connector> = {}): Connector {
  return {
    type: "connector",
    kind: "straight",
    widthCm: 6,
    heightCm: 2,
    wrap: "square",
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: 1,
    offsetYCm: 2,
    start: { xCm: 0, yCm: 0 },
    end: { xCm: 6, yCm: 2 },
    outline: { color: "#2e5395", widthPt: 1, dash: "solid" },
    endArrow: "triangle",
    ...overrides,
  };
}

function roundTrip(connector: Connector): Connector {
  const style = `<style:style style:name="cn1" style:family="graphic"><style:graphic-properties draw:stroke="solid" svg:stroke-color="#2e5395" svg:stroke-width="0.035cm"/></style:style>`;
  const xml = connectorElementXml(connector, "cn1", {});
  const parsed = parseConnectorElement(el(xml), contentDoc(style), undefined);
  return parsed!.connector;
}

describe("odf connector round-trip", () => {
  it("always writes svg:x1/y1/x2/y2 (AC-27)", () => {
    const xml = connectorElementXml(baseConnector(), "cn1", {});
    expect(xml).toMatch(/svg:x1=/);
    expect(xml).toMatch(/svg:y1=/);
    expect(xml).toMatch(/svg:x2=/);
    expect(xml).toMatch(/svg:y2=/);
  });

  it("writes draw:type per kind (AC-25)", () => {
    expect(connectorElementXml(baseConnector({ kind: "straight" }), "c", {})).toContain('draw:type="line"');
    expect(connectorElementXml(baseConnector({ kind: "elbow" }), "c", {})).toContain('draw:type="standard"');
    expect(connectorElementXml(baseConnector({ kind: "curved" }), "c", {})).toContain('draw:type="curve"');
  });

  it("round-trips a free straight connector idempotently", () => {
    const a = roundTrip(baseConnector());
    const b = roundTrip(a);
    expect(equalStructural(a, b)).toBe(true);
    expect(a.kind).toBe("straight");
    expect(a.endArrow).toBe("triangle");
  });

  it("round-trips elbow kind, adjustment and arrowheads", () => {
    const back = roundTrip(baseConnector({ kind: "elbow", adjustments: [0.25], startArrow: "open", endArrow: "none" }));
    expect(back.kind).toBe("elbow");
    expect(back.adjustments).toEqual([0.25]);
    expect(back.startArrow).toBe("open");
    expect(back.endArrow).toBe("none");
  });

  it("writes draw:start-shape/glue-point when attached (AC-26)", () => {
    const connector = baseConnector({
      startAttachment: { target: { block: 0 }, site: 1 },
      endAttachment: { target: { block: 1 }, site: 3 },
    });
    const xml = connectorElementXml(connector, "cn1", { start: "s0", end: "s1" });
    expect(xml).toContain('draw:start-shape="s0"');
    expect(xml).toContain('draw:start-glue-point="1"');
    expect(xml).toContain('draw:end-shape="s1"');
    expect(xml).toContain('draw:end-glue-point="3"');
  });

  it("resolves start/end shape ids to block indices (AC-28)", () => {
    const shapeA: Shape = { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square" };
    const shapeB: Shape = { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square" };
    const connector = baseConnector({
      startAttachment: { target: { block: 0 }, site: 1 },
      endAttachment: { target: { block: 1 }, site: 3 },
    });
    const idA = odfDrawId(0, 0);
    const idB = odfDrawId(0, 1);
    const connectorXml = connectorElementXml(connector, "cn1", { start: idA, end: idB });
    const parsed = parseConnectorElement(el(connectorXml), contentDoc(""), undefined)!.connector;

    const section: Section = { blocks: [shapeA, shapeB, parsed] };
    stampOdfConnectorTargetId(el(shapeElementXml(shapeA, "s", { paragraphXml: () => "" }, idA)), shapeA);
    stampOdfConnectorTargetId(el(shapeElementXml(shapeB, "s", { paragraphXml: () => "" }, idB)), shapeB);

    resolveOdfConnectorAttachments({ sections: [section] });
    expect(parsed.startAttachment).toEqual({ target: { block: 0 }, site: 1 });
    expect(parsed.endAttachment).toEqual({ target: { block: 1 }, site: 3 });
  });
});
