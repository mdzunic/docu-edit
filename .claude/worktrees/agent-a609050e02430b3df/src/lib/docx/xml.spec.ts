import { describe, it, expect } from "vitest";
import {
  parseXml,
  XmlParseError,
  parseRelationships,
  resolvePartPath,
  relsPathFor,
} from "../core/xml";

describe("parseXml", () => {
  it("parses well-formed XML", () => {
    const doc = parseXml("<root><child/></root>");
    expect(doc.documentElement.tagName).toBe("root");
  });

  it("throws XmlParseError for malformed XML", () => {
    expect(() => parseXml("<root><unclosed></root>")).toThrow(XmlParseError);
  });
});

describe("parseRelationships", () => {
  it("reads Id/Type/Target from a .rels part", () => {
    const xml = `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>`;
    const rels = parseRelationships(xml);
    expect(rels.get("rId1")).toEqual({
      id: "rId1",
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
      target: "word/document.xml",
    });
    expect(rels.get("rId2")?.target).toBe("styles.xml");
  });
});

describe("resolvePartPath", () => {
  it("resolves a relative target against its base part's directory", () => {
    expect(resolvePartPath("word/document.xml", "styles.xml")).toBe("word/styles.xml");
  });

  it("resolves a package-root-relative target", () => {
    expect(resolvePartPath("word/document.xml", "/word/styles.xml")).toBe("word/styles.xml");
  });

  it("resolves .. segments", () => {
    expect(resolvePartPath("word/document.xml", "../docProps/core.xml")).toBe("docProps/core.xml");
  });
});

describe("relsPathFor", () => {
  it("builds the rels path for a nested part", () => {
    expect(relsPathFor("word/document.xml")).toBe("word/_rels/document.xml.rels");
  });

  it("builds the rels path for a root part", () => {
    expect(relsPathFor("document.xml")).toBe("_rels/document.xml.rels");
  });
});
