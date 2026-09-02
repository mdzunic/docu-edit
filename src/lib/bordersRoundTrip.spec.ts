// SPEC-074 — paragraph and page borders/shading through both file formats
// (AC-28..AC-33). One suite covering both writers, because the interesting
// assertions are symmetric: the same model must survive a .docx and a .odt
// save, and the two formats' losses must be exactly the ones the spec names.
import { describe, expect, it } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx, buildDocumentXml } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { equalStructural, firstStructuralDifference } from "./equality";
import { collectBorderStyleDropRecords, collectPageBackgroundDropRecords } from "./opaque";
import { borderRuns, NO_SHADING_FILL } from "./paragraphBorders";
import { paintedParagraphFormatting, paragraphPaintResolver } from "./styles";
import type { BorderSide } from "./core/borders";
import {
  createDefaultStyles,
  type Block,
  type DocumentModel,
  type PageBorders,
  type Paragraph,
  type ParagraphBorders,
  type ParagraphFormatting,
  type Section,
} from "./documentModel";

const SINGLE: BorderSide = { style: "single", widthPt: 1, color: "#000000" };
const DOUBLE: BorderSide = { style: "double", widthPt: 3, color: "#c00000" };
const DOTTED: BorderSide = { style: "dotted", widthPt: 0.5, color: "#0070c0" };

function para(text: string, formatting?: ParagraphFormatting): Paragraph {
  const paragraph: Paragraph = { type: "paragraph", runs: [{ text, formatting: {} }] };
  if (formatting) {
    paragraph.formatting = formatting;
  }
  return paragraph;
}

function docOf(blocks: Block[], section: Partial<Section> = {}): DocumentModel {
  return {
    id: "borders-doc",
    title: "Borders",
    sections: [{ blocks, ...section }],
    defaultStyles: createDefaultStyles(),
  };
}

async function docxRoundTrip(model: DocumentModel): Promise<DocumentModel> {
  const { bytes } = await writeDocx(model);
  const parsed = await parseDocx(bytes);
  return { ...model, sections: parsed.sections, defaultStyles: parsed.defaultStyles, format: "docx" };
}

async function odtRoundTrip(model: DocumentModel): Promise<DocumentModel> {
  const { bytes } = await writeOdt(model);
  const parsed = await parseOdt(bytes);
  return { ...model, sections: parsed.sections, defaultStyles: parsed.defaultStyles, format: "odt" };
}

function firstParagraph(model: DocumentModel): Paragraph {
  return model.sections[0].blocks[0] as Paragraph;
}

// --- .docx paragraph borders and shading (AC-28) -----------------------------

describe("paragraph borders and shading through .docx (AC-28)", () => {
  const boxed: ParagraphBorders = {
    top: SINGLE,
    bottom: SINGLE,
    left: SINGLE,
    right: SINGLE,
    between: DOTTED,
    paddingPt: { top: 6, bottom: 6, left: 8, right: 8 },
  };

  it("writes w:pBdr with every edge, @w:val, @w:sz in eighths and @w:space in points", async () => {
    const xml = buildDocumentXml(docOf([para("Note", { borders: boxed, shadingFill: "#fff2cc" })]));
    expect(xml).toContain("<w:pBdr>");
    expect(xml).toContain('<w:top w:val="single" w:sz="8" w:space="6" w:color="000000"/>');
    expect(xml).toContain('<w:left w:val="single" w:sz="8" w:space="8" w:color="000000"/>');
    expect(xml).toContain('<w:bottom w:val="single" w:sz="8" w:space="6" w:color="000000"/>');
    expect(xml).toContain('<w:right w:val="single" w:sz="8" w:space="8" w:color="000000"/>');
    expect(xml).toContain('<w:between w:val="dotted" w:sz="4" w:space="0" w:color="0070C0"/>');
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>');
  });

  it("emits pBdr and shd before w:tabs, as CT_PPrBase orders them", () => {
    const xml = buildDocumentXml(
      docOf([
        para("Note", {
          borders: { bottom: SINGLE },
          shadingFill: "#eeeeee",
          tabStops: [{ positionCm: 2, alignment: "left" }],
        }),
      ]),
    );
    const pBdr = xml.indexOf("<w:pBdr>");
    const shd = xml.indexOf("<w:shd ");
    const tabs = xml.indexOf("<w:tabs>");
    expect(pBdr).toBeGreaterThan(-1);
    expect(shd).toBeGreaterThan(pBdr);
    expect(tabs).toBeGreaterThan(shd);
  });

  it("round-trips every edge, its padding and the shading", async () => {
    const model = docOf([para("Note", { borders: boxed, shadingFill: "#fff2cc" })]);
    const back = await docxRoundTrip(model);
    expect(firstParagraph(back).formatting?.borders).toEqual(boxed);
    expect(firstParagraph(back).formatting?.shadingFill).toBe("#fff2cc");
  });

  it("round-trips each of the eight line styles", async () => {
    const styles = ["single", "double", "dashed", "dotted", "thick", "dashDot", "dashDotDot"] as const;
    const blocks = styles.map((style) => para(style, { borders: { bottom: { style, widthPt: 2, color: "#336699" } } }));
    const back = await docxRoundTrip(docOf(blocks));
    styles.forEach((style, index) => {
      const paragraph = back.sections[0].blocks[index] as Paragraph;
      expect(paragraph.formatting?.borders?.bottom).toEqual({ style, widthPt: 2, color: "#336699" });
    });
  });

  it("keeps a bottom-rule heading's single edge and adds none of its own", async () => {
    const back = await docxRoundTrip(docOf([para("Heading", { borders: { bottom: SINGLE } })]));
    const borders = firstParagraph(back).formatting?.borders;
    expect(borders?.bottom).toEqual(SINGLE);
    expect(borders?.top).toBeUndefined();
    expect(borders?.left).toBeUndefined();
    expect(borders?.right).toBeUndefined();
  });

  it("leaves an unbordered paragraph with no pBdr and no shd at all", () => {
    const xml = buildDocumentXml(docOf([para("Plain")]));
    expect(xml).not.toContain("<w:pBdr>");
    expect(xml).not.toContain("<w:shd ");
  });
});

// --- .odt paragraph borders and shading (AC-29) ------------------------------

describe("paragraph borders and shading through .odt (AC-29)", () => {
  it("writes fo:border*, fo:padding*, fo:background-color and style:join-border", async () => {
    const model = docOf([
      para("Note", {
        borders: { top: SINGLE, bottom: SINGLE, left: SINGLE, right: SINGLE, paddingPt: { top: 6 } },
        shadingFill: "#fff2cc",
      }),
    ]);
    const { bytes } = await writeOdt(model);
    const content = await contentXmlOf(bytes);
    expect(content).toContain('fo:border-top="1pt solid #000000"');
    expect(content).toContain('fo:padding-top="6pt"');
    expect(content).toContain('fo:padding-bottom="4pt"');
    expect(content).toContain('fo:background-color="#fff2cc"');
    expect(content).toContain('style:join-border="true"');
  });

  it("round-trips every edge, its padding and the shading", async () => {
    const borders: ParagraphBorders = {
      top: DOUBLE,
      bottom: DOUBLE,
      left: DOTTED,
      right: DOTTED,
      paddingPt: { top: 6, bottom: 6, left: 3, right: 3 },
    };
    const back = await odtRoundTrip(docOf([para("Note", { borders, shadingFill: "#f2f2f2" })]));
    expect(firstParagraph(back).formatting?.borders).toEqual(borders);
    expect(firstParagraph(back).formatting?.shadingFill).toBe("#f2f2f2");
  });

  it("carries the between edge ODF has no attribute for, rather than losing it", async () => {
    const borders: ParagraphBorders = { top: SINGLE, bottom: SINGLE, left: SINGLE, right: SINGLE, between: DOTTED };
    const back = await odtRoundTrip(docOf([para("A", borders && { borders }), para("B", { borders })]));
    expect(firstParagraph(back).formatting?.borders?.between).toEqual(DOTTED);
  });
});

// --- page borders and background (AC-30, AC-31) ------------------------------

const PAGE_BORDERS: PageBorders = {
  top: DOUBLE,
  bottom: DOUBLE,
  left: DOUBLE,
  right: DOUBLE,
  offsetPt: 18,
  measureFrom: "text",
};

describe("page borders and background through .docx (AC-30)", () => {
  it("writes w:pgBorders in the sectPr and w:background in the document", () => {
    const xml = buildDocumentXml(
      docOf([para("Certificate")], { pageBorders: PAGE_BORDERS, pageBackground: "#fdf6e3" }),
    );
    expect(xml).toContain('<w:pgBorders w:offsetFrom="text">');
    expect(xml).toContain('<w:top w:val="double" w:sz="24" w:space="18" w:color="C00000"/>');
    expect(xml).toContain('<w:background w:color="FDF6E3"/>');
    // w:background is w:document's first child, where the schema puts it.
    expect(xml.indexOf("<w:background")).toBeLessThan(xml.indexOf("<w:body>"));
  });

  it("marks which pages take the border", () => {
    const xml = buildDocumentXml(
      docOf([para("x")], { pageBorders: { ...PAGE_BORDERS, appliesTo: "exceptFirstPage" } }),
    );
    expect(xml).toContain('w:display="notFirstPage"');
  });

  it("round-trips the border, its offset, its reference edge and the background", async () => {
    const model = docOf([para("Certificate")], { pageBorders: PAGE_BORDERS, pageBackground: "#fdf6e3" });
    const back = await docxRoundTrip(model);
    expect(back.sections[0].pageBorders).toEqual(PAGE_BORDERS);
    expect(back.sections[0].pageBackground).toBe("#fdf6e3");
  });

  it("round-trips a border on only some edges", async () => {
    const borders: PageBorders = { top: SINGLE, bottom: SINGLE, offsetPt: 24, measureFrom: "page" };
    const back = await docxRoundTrip(docOf([para("x")], { pageBorders: borders }));
    expect(back.sections[0].pageBorders).toEqual(borders);
  });
});

describe("page borders and background through .odt (AC-31)", () => {
  it("writes the border and background onto the page layout", async () => {
    const model = docOf([para("Certificate")], { pageBorders: PAGE_BORDERS, pageBackground: "#fdf6e3" });
    const { bytes } = await writeOdt(model);
    const styles = await stylesXmlOf(bytes);
    expect(styles).toContain('fo:border-top="3pt double #c00000"');
    expect(styles).toContain('fo:padding-top="18pt"');
    expect(styles).toContain('fo:background-color="#fdf6e3"');
  });

  it("round-trips the border, its offset, its reference edge and the background", async () => {
    const model = docOf([para("Certificate")], { pageBorders: PAGE_BORDERS, pageBackground: "#fdf6e3" });
    const back = await odtRoundTrip(model);
    expect(back.sections[0].pageBorders).toEqual(PAGE_BORDERS);
    expect(back.sections[0].pageBackground).toBe("#fdf6e3");
  });
});

// --- nearest-equivalent mapping and its DropRecord (AC-32) -------------------

describe("a border style one format lacks (AC-32)", () => {
  it("writes dashDot as ODF's nearest equivalent, dashed", async () => {
    const model = docOf([para("x", { borders: { bottom: { style: "dashDot", widthPt: 1, color: "#000000" } } })]);
    const content = await contentXmlOf((await writeOdt(model)).bytes);
    expect(content).toContain('fo:border-bottom="1pt dashed #000000"');
  });

  it("reports a paragraphFormatting DropRecord for each paragraph that uses one", async () => {
    const model = docOf([
      para("a", { borders: { bottom: { style: "dashDot", widthPt: 1, color: "#000000" } } }),
      para("b", { borders: { top: { style: "dashDotDot", widthPt: 1, color: "#000000" } } }),
      para("c", { borders: { top: SINGLE } }),
    ]);
    const records = collectBorderStyleDropRecords(model);
    expect(records).toHaveLength(2);
    expect(records[0].kind).toBe("paragraphFormatting");
    expect(records[0].location).toBe("paragraph 1");
    expect(records[1].location).toBe("paragraph 2");
    // ...and the .odt writer surfaces them on a save.
    const { dropped } = await writeOdt(model);
    expect(dropped.filter((d) => d.kind === "paragraphFormatting")).toHaveLength(2);
  });

  it("reports nothing for the styles ODF can express", async () => {
    const model = docOf([para("x", { borders: { top: DOUBLE, bottom: DOTTED, left: SINGLE } })]);
    expect(collectBorderStyleDropRecords(model)).toEqual([]);
    const { dropped } = await writeOdt(model);
    expect(dropped.filter((d) => d.kind === "paragraphFormatting")).toEqual([]);
  });

  it("reports a page background .docx cannot carry when the sections disagree", async () => {
    const model: DocumentModel = {
      id: "d",
      title: "t",
      defaultStyles: createDefaultStyles(),
      sections: [
        { blocks: [para("a")], pageBackground: "#ffeeee" },
        { blocks: [para("b")], pageBackground: "#eeffee" },
      ],
    };
    const records = collectPageBackgroundDropRecords(model);
    expect(records).toHaveLength(1);
    expect(records[0].kind).toBe("pageSetup");
    const { dropped } = await writeDocx(model);
    expect(dropped.filter((d) => d.kind === "pageSetup")).toHaveLength(1);
    // The first section's colour is the one that survives.
    const back = await parseDocx((await writeDocx(model)).bytes);
    expect(back.sections[0].pageBackground).toBe("#ffeeee");
  });

  it("reports nothing when every section agrees", async () => {
    const model: DocumentModel = {
      id: "d",
      title: "t",
      defaultStyles: createDefaultStyles(),
      sections: [
        { blocks: [para("a")], pageBackground: "#ffeeee" },
        { blocks: [para("b")], pageBackground: "#ffeeee" },
      ],
    };
    expect(collectPageBackgroundDropRecords(model)).toEqual([]);
  });
});

// --- the AC-33 documents ------------------------------------------------------

/** The spec's fixture content: a boxed note paragraph, a bottom-rule
 * heading, three consecutive identical bordered paragraphs, a shaded code
 * block and a page border. */
function fixtureModel(): DocumentModel {
  const runBorders: ParagraphBorders = {
    top: SINGLE,
    bottom: SINGLE,
    left: SINGLE,
    right: SINGLE,
    between: DOTTED,
    paddingPt: { top: 4, bottom: 4, left: 6, right: 6 },
  };
  const blocks: Block[] = [
    para("Heading with a rule", { borders: { bottom: DOUBLE, paddingPt: { bottom: 2 } } }),
    para("Note: this paragraph is boxed.", {
      borders: { top: SINGLE, bottom: SINGLE, left: SINGLE, right: SINGLE },
      shadingFill: "#fff2cc",
    }),
    para("Merged one", { borders: runBorders }),
    para("Merged two", { borders: runBorders }),
    para("Merged three", { borders: runBorders }),
    para("const answer = 42;", { shadingFill: "#f5f5f5" }),
  ];
  return docOf(blocks, { pageBorders: PAGE_BORDERS, pageBackground: "#fdf6e3" });
}

describe("the SPEC-074 fixtures open, render, save and reopen equalStructural (AC-33)", () => {
  // The AC is about a *file*: open it, save it, reopen it, and the two
  // models must be structurally equal. Comparing the reopened model against
  // the authored one instead would trip over the cm↔twips rounding every
  // page geometry picks up on its first trip through a package, which is
  // SPEC-027's business and not this spec's.
  it("survives a .docx save and reopen", async () => {
    const opened = await docxRoundTrip(fixtureModel());
    const reopened = await docxRoundTrip(opened);
    expect(firstStructuralDifference(opened, reopened) ?? "equal").toBe("equal");
    expect(equalStructural(opened, reopened)).toBe(true);
    // ...and the borders really are still there, not equal-because-absent.
    const note = opened.sections[0].blocks[1] as Paragraph;
    expect(note.formatting?.borders?.top).toEqual(SINGLE);
    expect(note.formatting?.shadingFill).toBe("#fff2cc");
    expect(opened.sections[0].pageBorders?.top).toEqual(DOUBLE);
    expect(opened.sections[0].pageBackground).toBe("#fdf6e3");
  });

  it("survives a .odt save and reopen", async () => {
    const opened = await odtRoundTrip(fixtureModel());
    const reopened = await odtRoundTrip(opened);
    expect(firstStructuralDifference(opened, reopened) ?? "equal").toBe("equal");
    expect(equalStructural(opened, reopened)).toBe(true);
    const merged = opened.sections[0].blocks[2] as Paragraph;
    expect(merged.formatting?.borders?.between).toEqual(DOTTED);
    expect(opened.sections[0].pageBorders?.left).toEqual(DOUBLE);
    expect(opened.sections[0].pageBackground).toBe("#fdf6e3");
  });

  it("carries the three merged paragraphs across formats as one run", async () => {
    const viaOdt = await odtRoundTrip(await docxRoundTrip(fixtureModel()));
    const runs = borderRuns(viaOdt.sections[0].blocks);
    expect(runs.slice(2, 5).map((r) => r.top)).toEqual([true, false, false]);
    expect(runs.slice(2, 5).map((r) => r.bottom)).toEqual([false, false, true]);
    expect(runs.slice(2, 5).map((r) => r.between)).toEqual([true, true, false]);
  });
});

// --- helpers -----------------------------------------------------------------

async function contentXmlOf(bytes: Uint8Array): Promise<string> {
  return partOf(bytes, "content.xml");
}

async function stylesXmlOf(bytes: Uint8Array): Promise<string> {
  return partOf(bytes, "styles.xml");
}

async function partOf(bytes: Uint8Array, name: string): Promise<string> {
  const { unzip } = await import("./core/zip");
  const parts = await unzip(bytes);
  const part = parts.get(name);
  if (!part) {
    throw new Error(`missing ${name}`);
  }
  return new TextDecoder().decode(part);
}

// SPEC-074 AC-17/AC-18 — a style's own borders and shading. They live on the
// style, not copied onto every paragraph that uses it, so what has to survive
// a save is the *style definition*: `w:pBdr`/`w:shd` inside styles.xml's
// `w:pPr`, and the ODF paragraph properties of the named style.
describe("a paragraph style's borders and shading (AC-17, AC-18)", () => {
  function styledDoc(): DocumentModel {
    const boxed = para("Inherits the style's box");
    boxed.styleId = "Callout";
    const model = docOf([boxed]);
    model.styles = [
      { id: "Normal", name: "Normal" },
      {
        id: "Callout",
        name: "Callout",
        basedOn: "Normal",
        paragraphFormatting: {
          borders: { top: SINGLE, bottom: SINGLE, left: SINGLE, right: SINGLE },
          shadingFill: "#fff2cc",
        },
      },
    ];
    return model;
  }

  function calloutOf(model: DocumentModel) {
    return model.styles?.find((style) => style.id === "Callout")?.paragraphFormatting;
  }

  it("survives a .docx round trip on the style itself", async () => {
    const { bytes } = await writeDocx(styledDoc());
    const parsed = await parseDocx(bytes);
    const callout = calloutOf({ ...styledDoc(), styles: parsed.styles });
    expect(callout?.borders?.top).toEqual(SINGLE);
    expect(callout?.borders?.right).toEqual(SINGLE);
    expect(callout?.shadingFill).toBe("#fff2cc");
  });

  it("survives a .odt round trip on the style itself", async () => {
    const { bytes } = await writeOdt(styledDoc());
    const parsed = await parseOdt(bytes);
    const callout = calloutOf({ ...styledDoc(), styles: parsed.styles });
    expect(callout?.borders?.top).toEqual(SINGLE);
    expect(callout?.shadingFill).toBe("#fff2cc");
  });

  it("paints the run through the resolver, not through the paragraph's own formatting", () => {
    const model = styledDoc();
    // The paragraph itself carries nothing — the box is entirely its style's.
    expect((model.sections[0].blocks[0] as Paragraph).formatting).toBeUndefined();
    expect(borderRuns(model.sections[0].blocks)[0].sides).toBe(false);
    expect(
      borderRuns(model.sections[0].blocks, paragraphPaintResolver(model))[0],
    ).toEqual({ top: true, bottom: true, sides: true, between: false, continuesBelow: false });
  });

  // A paragraph that switches the style's box back off (AC-17) carries a
  // value that is present and paints nothing. Both formats spell exactly
  // that, and it has to survive a save: writing nothing would let the
  // style's box come back on the next open.
  describe("a paragraph that overrides the style's box with none", () => {
    function overriddenDoc(): DocumentModel {
      // Normal carries the box, so both paragraphs inherit it — the shape a
      // user gets from "Modify Normal → Borders and shading" in a blank
      // document. The first switches it back off; the second must not.
      const off = para("No box on this one");
      off.formatting = { borders: {}, shadingFill: NO_SHADING_FILL };
      const model = docOf([off, para("Still boxed by Normal")]);
      model.styles = [
        {
          id: "Normal",
          name: "Normal",
          paragraphFormatting: {
            borders: { top: SINGLE, bottom: SINGLE, left: SINGLE, right: SINGLE },
            shadingFill: "#ffd9d9",
          },
        },
      ];
      return model;
    }

    function paintedOf(model: DocumentModel, block = 0): ParagraphFormatting | undefined {
      return paintedParagraphFormatting(model, model.sections[0].blocks[block] as Paragraph);
    }

    it("writes an all-nil w:pBdr and a fill-less w:shd", async () => {
      const xml = buildDocumentXml(overriddenDoc());
      expect(xml).toContain('<w:pBdr><w:top w:val="nil" w:sz="0" w:space="0" w:color="auto"/>');
      expect(xml).toContain('<w:right w:val="nil" w:sz="0" w:space="0" w:color="auto"/></w:pBdr>');
      expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
    });

    it("writes fo:border=\"none\" and a transparent background", async () => {
      const { bytes } = await writeOdt(overriddenDoc());
      const content = await partOf(bytes, "content.xml");
      expect(content).toContain('fo:border="none"');
      expect(content).toContain('fo:background-color="transparent"');
      // Nothing to space, so no padding is written beside it.
      expect(content).not.toContain('fo:border="none" fo:padding');
    });

    it.each([
      ["docx", docxRoundTrip],
      ["odt", odtRoundTrip],
    ])("survives a .%s round trip, so the style's box stays off", async (_format, roundTrip) => {
      const reopened = await roundTrip(overriddenDoc());
      expect(firstParagraph(reopened).formatting).toEqual({
        borders: {},
        shadingFill: NO_SHADING_FILL,
      });
      expect(paintedOf(reopened)?.borders).toEqual({});
      expect(paintedOf(reopened)?.shadingFill).toBe(NO_SHADING_FILL);
      // ...and the paragraph that did not override still takes Normal's box.
      expect(paintedOf(reopened, 1)?.borders?.top).toEqual(SINGLE);
      expect(paintedOf(reopened, 1)?.shadingFill).toBe("#ffd9d9");
    });
  });

  // The same two spellings appear all over ordinary documents, on paragraphs
  // whose style sets no border or fill at all — Word emits `w:shd
  // w:val="clear" w:fill="auto"` and LibreOffice `fo:border="none"` freely.
  // There they override nothing, and the model keeps no husk for them.
  describe("a redundant override in the file", () => {
    it("reads back as no borders and no shading, not as an empty husk", async () => {
      const plain = docOf([para("Nothing to override")]);
      const { bytes } = await writeDocx(plain);
      const withHusk = (await partOf(bytes, "word/document.xml")).replace(
          "<w:r>",
          '<w:pPr><w:pBdr><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>' +
            '<w:right w:val="nil"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="auto"/>' +
            "</w:pPr><w:r>",
        );
      const parsed = await parseDocx(await rezip(bytes, "word/document.xml", withHusk));
      const paragraph = parsed.sections[0].blocks[0] as Paragraph;
      expect(paragraph.formatting?.borders).toBeUndefined();
      expect(paragraph.formatting?.shadingFill).toBeUndefined();
    });
  });
});

/** Replaces one part of a written package, for the reader tests that need
 * XML no writer of this product emits. */
async function rezip(bytes: Uint8Array, name: string, xml: string): Promise<Uint8Array> {
  const { unzip } = await import("./core/zip");
  const { zip } = await import("./core/zipWriter");
  const parts = await unzip(bytes);
  parts.set(name, new TextEncoder().encode(xml));
  return zip([...parts].map(([entry, data]) => ({ name: entry, data })));
}
