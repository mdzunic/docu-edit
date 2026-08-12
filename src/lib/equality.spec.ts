import { describe, it, expect } from "vitest";
import {
  canonicalizeXml,
  equalBytes,
  equalSemantic,
  equalStructural,
  firstStructuralDifference,
  resolve,
} from "./equality";
import { createDocument } from "./documentModel";

describe("equalBytes", () => {
  it("is true for identical buffers and false otherwise", () => {
    expect(equalBytes(new Uint8Array([1, 2, 3]), new Uint8Array([1, 2, 3]))).toBe(true);
    expect(equalBytes(new Uint8Array([1, 2, 3]), new Uint8Array([1, 2, 4]))).toBe(false);
    expect(equalBytes(new Uint8Array([1, 2, 3]), new Uint8Array([1, 2]))).toBe(false);
    expect(equalBytes(new Uint8Array(0), new Uint8Array(0))).toBe(true);
  });
});

describe("canonicalizeXml", () => {
  it("ignores attribute order", () => {
    expect(canonicalizeXml('<a x="1" y="2"/>')).toBe(canonicalizeXml('<a y="2" x="1"/>'));
  });

  it("ignores namespace prefixes: the same expanded names are equal", () => {
    const withW = '<w:tbl xmlns:w="http://example.com/ns"><w:tr/></w:tbl>';
    const withOther = '<z:tbl xmlns:z="http://example.com/ns"><z:tr/></z:tbl>';
    expect(canonicalizeXml(withW)).toBe(canonicalizeXml(withOther));
  });

  it("ignores where namespace declarations sit", () => {
    const onRoot = '<a xmlns:n="urn:x"><n:b n:v="1"/></a>';
    const onChild = '<a><n:b xmlns:n="urn:x" n:v="1"/></a>';
    expect(canonicalizeXml(onRoot)).toBe(canonicalizeXml(onChild));
  });

  it("ignores comments but keeps text verbatim", () => {
    expect(canonicalizeXml("<a><!-- note -->text</a>")).toBe(canonicalizeXml("<a>text</a>"));
    expect(canonicalizeXml("<a> text </a>")).not.toBe(canonicalizeXml("<a>text</a>"));
  });

  it("distinguishes real content differences", () => {
    expect(canonicalizeXml('<a x="1"/>')).not.toBe(canonicalizeXml('<a x="2"/>'));
    expect(canonicalizeXml("<a><b/></a>")).not.toBe(canonicalizeXml("<a><c/></a>"));
  });
});

describe("equalStructural / firstStructuralDifference", () => {
  it("is reflexive and ignores ids", () => {
    const doc = createDocument("id-a");
    const other = { ...structuredClone(doc), id: "id-b" };
    expect(equalStructural(doc, other)).toBe(true);
    expect(equalStructural(other, doc)).toBe(true);
  });

  it("reports the path of the first difference", () => {
    const a = createDocument("id");
    const b = structuredClone(a);
    b.sections[0].blocks.push({ type: "paragraph", runs: [{ text: "x", formatting: {} }] });
    expect(firstStructuralDifference(a, b)).toBe("sections[0].blocks.length");

    const c = structuredClone(a);
    c.title = "Other";
    expect(firstStructuralDifference(a, c)).toBe("title");
  });

  it("treats empty objects, empty arrays and absent fields alike", () => {
    const a = { runs: [{ text: "x", formatting: {} }], opaqueInlines: [] };
    const b = { runs: [{ text: "x" }] };
    expect(equalStructural(a, b)).toBe(true);
  });

  it("compares fields named xml canonically, not byte-for-byte", () => {
    const a = { xml: '<w:tbl xmlns:w="urn:x" a="1" b="2"/>' };
    const b = { xml: '<q:tbl xmlns:q="urn:x" b="2" a="1"/>' };
    expect(equalStructural(a, b)).toBe(true);
    const c = { xml: '<w:tbl xmlns:w="urn:x" a="9"/>' };
    expect(firstStructuralDifference(a, c)).toBe("xml");
  });

  it("reports scalar differences at their full path", () => {
    const a = { sections: [{ blocks: [{ type: "paragraph", runs: [{ text: "one" }] }] }] };
    const b = { sections: [{ blocks: [{ type: "paragraph", runs: [{ text: "two" }] }] }] };
    expect(firstStructuralDifference(a, b)).toBe("sections[0].blocks[0].runs[0].text");
  });
});

describe("resolve / equalSemantic (specs/002 'Equality', relation 3)", () => {
  it("strips the style reference but changes no formatting on an already-fully-baked model", () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      {
        type: "paragraph",
        styleId: "Heading1",
        headingLevel: 1,
        formatting: { spaceBeforePt: 20, spaceAfterPt: 6 },
        runs: [{ text: "Title", formatting: { bold: true, fontSizePt: 20 } }],
      },
    ];
    const originalRuns = (doc.sections[0].blocks[0] as { runs: unknown }).runs;
    const resolved = resolve(doc);
    const paragraph = resolved.sections[0].blocks[0] as { styleId?: string; formatting?: unknown; runs: unknown };
    expect(paragraph.styleId).toBeUndefined();
    expect(paragraph.formatting).toEqual({ spaceBeforePt: 20, spaceAfterPt: 6 });
    expect(paragraph.runs).toEqual(originalRuns);
    expect(equalSemantic(doc, doc)).toBe(true); // reflexive
  });

  it("makes a directly-bolded run and one bolded only by its style compare equal", () => {
    const styled = createDocument("id-a");
    styled.sections[0].blocks = [
      { type: "paragraph", styleId: "Heading1", headingLevel: 1, runs: [{ text: "Title", formatting: {} }] },
    ];
    const direct = createDocument("id-b");
    direct.sections[0].blocks = [
      {
        type: "paragraph",
        headingLevel: 1,
        formatting: { spaceBeforePt: 20, spaceAfterPt: 6 },
        runs: [{ text: "Title", formatting: { bold: true, fontSizePt: 20 } }],
      },
    ];
    // Not structurally equal (styleId differs, direct formatting is bare on
    // one side) — this is exactly the gap equalSemantic exists to close.
    expect(equalStructural(styled, direct)).toBe(false);
    expect(equalSemantic(styled, direct)).toBe(true);
  });

  it("still finds a genuine content difference after resolving", () => {
    const a = createDocument("id-a");
    a.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "one", formatting: {} }] }];
    const b = createDocument("id-b");
    b.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "two", formatting: {} }] }];
    expect(equalSemantic(a, b)).toBe(false);
  });
});
