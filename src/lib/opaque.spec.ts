import { describe, it, expect } from "vitest";
import {
  base64ToBytes,
  bytesToBase64,
  collectDropRecords,
  DROP_KIND_LABELS,
  noConversionReport,
} from "./opaque";
import { createDocument } from "./documentModel";

describe("base64 codecs", () => {
  it("round-trips arbitrary binary bytes", () => {
    const bytes = new Uint8Array(512).map((_, i) => (i * 37 + 11) % 256);
    expect(base64ToBytes(bytesToBase64(bytes))).toEqual(bytes);
  });

  it("round-trips empty input", () => {
    expect(base64ToBytes(bytesToBase64(new Uint8Array(0)))).toEqual(new Uint8Array(0));
  });
});

describe("collectDropRecords", () => {
  it("is empty for a blank document and for a same-format save", () => {
    const blank = createDocument("id-1");
    expect(collectDropRecords(blank, "docx")).toEqual([]);
    expect(collectDropRecords(blank, "odt")).toEqual([]);

    const docx = { ...createDocument("id-2"), format: "docx" as const };
    docx.opaque = [{ name: "customXml/item1.xml", data: bytesToBase64(new Uint8Array([1])) }];
    expect(collectDropRecords(docx, "docx")).toEqual([]);
  });

  it("classifies parts and nodes with human-readable locations", () => {
    const model = { ...createDocument("id-3"), format: "docx" as const };
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "one", formatting: {} }] },
      {
        type: "paragraph",
        runs: [{ text: "two", formatting: {} }],
        opaqueInlines: [{ xml: "<w:sdt xmlns:w='urn:w'><w:sdtContent/></w:sdt>", runIndex: 1 }],
      },
    ];
    model.sections[0].opaqueBlocks = [{ xml: "<w:tbl xmlns:w='urn:w'/>", blockIndex: 2 }];
    model.opaque = [
      { name: "customXml/item1.xml", data: "" },
      { name: "word/charts/chart1.xml", data: "" },
      { name: "word/embeddings/oleObject1.bin", data: "" },
      { name: "word/vbaProject.bin", data: "" },
      { name: "word/styles.xml", data: "" },
      { name: "word/_rels/document.xml.rels", data: "" },
    ];

    const dropped = collectDropRecords(model, "odt");
    expect(dropped.map((d) => [d.kind, d.location])).toEqual([
      ["contentControl", "paragraph 2"],
      ["opaqueBlock", "end of document"],
      ["customXml", 'package part "customXml/item1.xml"'],
      ["chart", 'package part "word/charts/chart1.xml"'],
      ["embeddedObject", 'package part "word/embeddings/oleObject1.bin"'],
      ["macro", 'package part "word/vbaProject.bin"'],
    ]);
    for (const record of dropped) {
      expect(DROP_KIND_LABELS[record.kind]).toBeTruthy();
      expect(record.reason).toContain("OpenDocument Text (.odt)");
    }
  });

  it("reports a tracked table/row/cell property change as dropped on a cross-format save to .odt (SPEC-039 AC-19)", () => {
    const model = { ...createDocument("id-4"), format: "docx" as const };
    const change = { changeId: "c1", author: "Ada", date: "2026-01-01T00:00:00.000Z", previous: {} };
    model.sections[0].blocks = [
      {
        type: "table",
        columnWidthsCm: [5, 5],
        changes: { propertyChange: change },
        rows: [
          {
            changes: { propertyChange: change },
            cells: [
              { blocks: [{ type: "paragraph", runs: [] }], changes: { propertyChange: change } },
              { blocks: [{ type: "paragraph", runs: [] }] },
            ],
          },
        ],
      },
    ];

    const dropped = collectDropRecords(model, "odt");
    expect(dropped.map((d) => [d.kind, d.location])).toEqual([
      ["trackedPropertyChange", "table at paragraph 1"],
      ["trackedPropertyChange", "table at paragraph 1, row 1"],
      ["trackedPropertyChange", "table at paragraph 1, row 1, column 1"],
    ]);
    expect(dropped[0].reason).toContain("OpenDocument Text");

    // Not reported for a .docx target (w:tblPrChange/w:trPrChange/
    // w:tcPrChange represent it — AC-15) or a same-format .docx save.
    expect(collectDropRecords(model, "docx")).toEqual([]);
  });
});

describe("noConversionReport", () => {
  it("is the dialog's hidden state", () => {
    expect(noConversionReport()).toBeNull();
  });
});
