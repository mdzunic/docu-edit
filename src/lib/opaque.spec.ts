import { describe, it, expect } from "vitest";
import {
  base64ToBytes,
  bytesToBase64,
  collectDropRecords,
  DROP_KIND_LABELS,
  noConversionReport,
  withoutDropped,
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

  it("reports an unsupported shape geometry as dropped on a cross-format save (SPEC-051 AC-61)", () => {
    const model = { ...createDocument("id-shape-1"), format: "docx" as const };
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "before", formatting: {} }] },
      {
        type: "shape",
        geometry: "rect",
        widthCm: 3,
        heightCm: 2,
        wrap: "square",
        unsupportedGeometryXml: '<a:prstGeom prst="cube"><a:avLst/></a:prstGeom>',
      },
    ];

    const dropped = collectDropRecords(model, "odt");
    expect(dropped.map((d) => [d.kind, d.location])).toEqual([["shapeGeometry", "paragraph 2"]]);
    expect(dropped[0].reason).toContain("OpenDocument Text (.odt)");
    expect(DROP_KIND_LABELS.shapeGeometry).toBe("Shape geometry");
  });

  it("does not report a shape's unsupported geometry on a same-format save", () => {
    const model = { ...createDocument("id-shape-2"), format: "docx" as const };
    model.sections[0].blocks = [
      {
        type: "shape",
        geometry: "rect",
        widthCm: 3,
        heightCm: 2,
        wrap: "square",
        unsupportedGeometryXml: '<a:prstGeom prst="cube"><a:avLst/></a:prstGeom>',
      },
    ];
    expect(collectDropRecords(model, "docx")).toEqual([]);
  });

  it("does not report a shape with a recognised (catalogue) geometry as dropped, even cross-format", () => {
    const model = { ...createDocument("id-shape-3"), format: "docx" as const };
    model.sections[0].blocks = [
      { type: "shape", geometry: "ellipse", widthCm: 3, heightCm: 2, wrap: "square" },
    ];
    expect(collectDropRecords(model, "odt")).toEqual([]);
  });

  it("reports a catalogue geometry with no clean target-format equivalent as dropped (SPEC-069 AC-9)", () => {
    const model = { ...createDocument("id-shape-4"), format: "docx" as const };
    // borderCallout1 is marked approximateIn: ["odt"] — a docx->odt save
    // writes the closest ODF preset AND reports the drop; a same-format save
    // and the reverse direction report nothing.
    model.sections[0].blocks = [
      { type: "shape", geometry: "borderCallout1", widthCm: 3, heightCm: 2, wrap: "square" },
    ];
    expect(collectDropRecords(model, "odt").map((d) => d.kind)).toEqual(["shapeGeometry"]);
    expect(collectDropRecords(model, "docx")).toEqual([]);
  });
});

describe("noConversionReport", () => {
  it("is the dialog's hidden state", () => {
    expect(noConversionReport()).toBeNull();
  });
});

describe("SPEC-050 — collectDropRecords for a .md target", () => {
  it("is empty for a blank document with no unsupported content", () => {
    const blank = createDocument("id-5");
    expect(collectDropRecords(blank, "md")).toEqual([]);
  });

  it("reports paragraph formatting, run formatting and bookmarks even for a document that is already .md-sourced", () => {
    // Markdown is not full-fidelity for content this editor can create —
    // unlike a .docx<->.odt save, "already this format" is not a same-
    // format short-circuit (see collectDropRecords's own doc comment).
    const model = { ...createDocument("id-6"), format: "md" as const };
    model.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [{ text: "centered", formatting: { color: "#ff0000" } }],
        formatting: { alignment: "center" },
        bookmarks: [{ name: "anchor1", offset: 0 }],
      },
    ];
    const dropped = collectDropRecords(model, "md");
    expect(dropped.map((d) => d.kind).sort()).toEqual(["bookmark", "paragraphFormatting", "runFormatting"]);
    for (const record of dropped) {
      expect(DROP_KIND_LABELS[record.kind]).toBeTruthy();
      expect(record.reason).toContain("Markdown (.md)");
    }
  });

  it("does not flag a run's baked style/underlay formatting as unsupported (SPEC-019's Quote style)", () => {
    const model = { ...createDocument("id-7"), format: "docx" as const };
    model.sections[0].blocks = [
      {
        type: "paragraph",
        styleId: "Quote",
        runs: [{ text: "quoted", formatting: { italic: true, color: "#666666", fontFamily: "Calibri" } }],
        formatting: { indentStartCm: 1.27, spaceBeforePt: 10, spaceAfterPt: 10 },
      },
    ];
    expect(collectDropRecords(model, "md")).toEqual([]);
  });

  it("reports a header, non-default page margins and a footnote reference", () => {
    const model = { ...createDocument("id-8"), format: "docx" as const };
    model.sections[0].margins = { topCm: 5, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 };
    model.sections[0].headers = { default: { blocks: [{ type: "paragraph", runs: [] }] } };
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "1", formatting: {}, note: { noteId: "n1", kind: "footnote" } }] },
    ];
    const dropped = collectDropRecords(model, "md");
    expect(dropped.map((d) => d.kind).sort()).toEqual(["headerFooter", "note", "pageSetup"]);
  });
});

describe("SPEC-050 — withoutDropped", () => {
  it("returns the source unchanged when nothing was dropped", () => {
    const model = createDocument("id-9");
    expect(withoutDropped(model, [])).toBe(model);
  });

  it("clears only what the reported kinds cover, leaving the rest of the model alone", () => {
    const model = { ...createDocument("id-10"), format: "md" as const };
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "kept", formatting: { bold: true } }] },
      {
        type: "paragraph",
        runs: [{ text: "centered", formatting: {} }],
        formatting: { alignment: "center" },
      },
    ];
    const dropped = collectDropRecords(model, "md");
    expect(dropped.map((d) => d.kind)).toEqual(["paragraphFormatting"]);
    const masked = withoutDropped(model, dropped);
    expect(masked.sections[0].blocks[0]).toEqual(model.sections[0].blocks[0]);
    expect((masked.sections[0].blocks[1] as { formatting?: unknown }).formatting).toBeUndefined();
  });
});

// --- SPEC-073: shapeStyle drops (AC-34) --------------------------------------

describe("shapeStyle DropRecords (SPEC-073 AC-34)", () => {
  function shapeDoc(overrides: Record<string, unknown>) {
    const model = { ...createDocument("id-style-drop"), format: "docx" as const };
    model.sections[0].blocks = [
      { type: "paragraph", runs: [] },
      { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square", ...overrides } as never,
    ];
    return model;
  }

  it("has a label for the new kind", () => {
    expect(DROP_KIND_LABELS.shapeStyle).toBe("Shape styling");
  });

  it("preserved unmodelled styling drops on any cross-format save, naming it", () => {
    const model = shapeDoc({ unsupportedStyleXml: '<a:scene3d xmlns:a="x"/>' });
    const dropped = collectDropRecords(model, "odt");
    expect(dropped.map((d) => [d.kind, d.location])).toEqual([["shapeStyle", "paragraph 2"]]);
    expect(dropped[0].reason).toContain("unmodelled styling");
  });

  it("preserved unmodelled styling reports nothing on a same-format save (AC-35)", () => {
    const model = shapeDoc({ unsupportedStyleXml: '<a:scene3d xmlns:a="x"/>' });
    expect(collectDropRecords(model, "docx")).toEqual([]);
  });

  it("a compound outline drops to .odt, which cannot express it", () => {
    const model = shapeDoc({ outline: { color: "#000000", widthPt: 1, dash: "solid", compound: "thickThin" } });
    const dropped = collectDropRecords(model, "odt");
    expect(dropped.map((d) => d.kind)).toEqual(["shapeStyle"]);
    expect(dropped[0].reason).toContain("compound");
  });

  it("a gradient's intermediate stops and per-stop transparency each drop to .odt", () => {
    const model = shapeDoc({
      fill: {
        kind: "gradient",
        gradient: {
          type: "linear",
          angleDeg: 90,
          stops: [
            { position: 0, color: "#000000" },
            { position: 0.5, color: "#888888", alpha: 0.5 },
            { position: 1, color: "#ffffff" },
          ],
        },
      },
    });
    const reasons = collectDropRecords(model, "odt").map((d) => d.reason);
    expect(reasons.some((r) => r.includes("intermediate stops"))).toBe(true);
    expect(reasons.some((r) => r.includes("per-stop transparency"))).toBe(true);
  });

  it("a pattern preset outside the closed hatch table drops to .odt; a mapped one does not", () => {
    const unmapped = shapeDoc({ fill: { kind: "pattern", preset: "pct50", foreground: "#000000", background: "#ffffff" } });
    expect(collectDropRecords(unmapped, "odt").map((d) => d.kind)).toEqual(["shapeStyle"]);
    const mapped = shapeDoc({ fill: { kind: "pattern", preset: "ltUpDiag", foreground: "#000000", background: "#ffffff" } });
    expect(collectDropRecords(mapped, "odt")).toEqual([]);
  });

  it("a two-stop opaque gradient, alpha, shadow and arrowheads all convert cleanly (no drops)", () => {
    const model = shapeDoc({
      fill: { kind: "solid", color: "#4472c4", alpha: 0.5 },
      outline: {
        color: "#2e5395",
        widthPt: 1,
        dash: "dashDot",
        alpha: 0.5,
        join: "round",
        cap: "round",
        startArrow: { kind: "open", size: "large" },
      },
      shadow: { color: "#000000", alpha: 0.4, blurPt: 4, distancePt: 3, angleDeg: 90 },
    });
    expect(collectDropRecords(model, "odt")).toEqual([]);
  });

  it("a connector with preserved styling reports a shapeStyle drop cross-format", () => {
    const model = { ...createDocument("id-style-drop-c"), format: "docx" as const };
    model.sections[0].blocks = [
      {
        type: "connector",
        kind: "straight",
        start: { xCm: 0, yCm: 0 },
        end: { xCm: 3, yCm: 0 },
        widthCm: 3,
        heightCm: 0,
        wrap: "infront",
        unsupportedStyleXml: '<a:scene3d xmlns:a="x"/>',
      } as never,
    ];
    expect(collectDropRecords(model, "odt").map((d) => d.kind)).toEqual(["shapeStyle"]);
    expect(collectDropRecords(model, "docx")).toEqual([]);
  });
});
