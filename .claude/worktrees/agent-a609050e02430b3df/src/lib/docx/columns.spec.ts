// SPEC-030 — .docx column round-trip: w:cols (w:num/w:space/w:sep) for equal
// columns, explicit w:col children for unequal widths, and a reading of an
// authored w:cols.
import { describe, expect, it } from "vitest";
import { parseDocx } from "./reader";
import { buildDocumentXml, writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { zip } from "../core/zipWriter";
import { WORD_NAMESPACE } from "./xml";
import { createDocument, type DocumentModel, type Section } from "../documentModel";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const encoder = new TextEncoder();

function modelWith(section: Section): DocumentModel {
  const model = createDocument("spec-030");
  model.format = "docx";
  model.sections = [section];
  return model;
}

async function docxFromBody(bodyXml: string): Promise<Uint8Array> {
  const documentXml = `${XML_DECLARATION}<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${bodyXml}</w:body></w:document>`;
  return zip([{ name: "word/document.xml", data: encoder.encode(documentXml) }]);
}

describe("docx columns round-trip (SPEC-030)", () => {
  it("writes equal columns as w:num/w:space with no w:col children (AC-16a)", async () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 3, equalWidth: true, gapCm: 1.25, separator: true },
    });
    const xml = buildDocumentXml(model);
    expect(xml).toContain('<w:cols w:num="3"');
    expect(xml).toContain('w:sep="1"');
    expect(xml).not.toContain("<w:col ");
  });

  it("writes unequal columns with explicit w:col widths (AC-16)", async () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 2, equalWidth: false, gapCm: 1, separator: false, widthsCm: [5, 10] },
    });
    const xml = buildDocumentXml(model);
    expect(xml).toContain('<w:cols w:num="2"');
    expect((xml.match(/<w:col /g) ?? []).length).toBe(2);
  });

  it("round-trips equal columns through a write/read cycle", async () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 2, equalWidth: true, gapCm: 1.25, separator: true },
    });
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const cols = reopened.sections[reopened.sections.length - 1].columns;
    expect(cols?.count).toBe(2);
    expect(cols?.equalWidth).toBe(true);
    expect(cols?.separator).toBe(true);
  });

  it("round-trips unequal columns' proportions", async () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 2, equalWidth: false, gapCm: 1, separator: false, widthsCm: [5, 10] },
    });
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const cols = reopened.sections[reopened.sections.length - 1].columns;
    expect(cols?.count).toBe(2);
    expect(cols?.equalWidth).toBe(false);
    expect(cols?.widthsCm).toHaveLength(2);
    // The second column is about twice the first.
    expect((cols!.widthsCm![1] ?? 0) / (cols!.widthsCm![0] ?? 1)).toBeCloseTo(2, 1);
  });

  it("reads an authored two-column w:cols", async () => {
    const bytes = await docxFromBody(
      "<w:p><w:r><w:t>body</w:t></w:r></w:p>" +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
        '<w:cols w:num="2" w:space="708" w:sep="1"/></w:sectPr>',
    );
    const parsed = await parseDocx(bytes);
    const cols = parsed.sections[0].columns;
    expect(cols?.count).toBe(2);
    expect(cols?.separator).toBe(true);
  });

  it("produces schema-valid output with columns", async () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 2, equalWidth: false, gapCm: 1, separator: true, widthsCm: [6, 9] },
    });
    expect(await validateDocx((await writeDocx(model)).bytes)).toEqual([]);
  });
});
