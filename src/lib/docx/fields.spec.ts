// SPEC-053: fields through the .docx reader and writer — every catalogue type
// as a w:fldSimple, the run-split w:fldChar/w:instrText encoding read into
// `cached`, and an unsupported field type preserved byte-for-byte.
import { describe, expect, it } from "vitest";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";
import { zip } from "../core/zipWriter";
import {
  createDocument,
  createFieldSpecRun,
  type DocumentModel,
  type FieldSpec,
  type Paragraph,
} from "../documentModel";

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const encoder = new TextEncoder();

function bodyDocx(body: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`
  );
}

async function makeDocx(documentXml: string): Promise<Uint8Array> {
  return zip([{ name: "word/document.xml", data: encoder.encode(documentXml) }]);
}

function fieldParagraph(field: FieldSpec): Paragraph {
  return { type: "paragraph", runs: [createFieldSpecRun(field)] };
}

function modelWithFields(fields: FieldSpec[]): DocumentModel {
  const model = createDocument("docx-fields");
  model.sections = [{ blocks: fields.map(fieldParagraph) }];
  model.properties = {
    author: "Ada",
    title: "Spec",
    subject: "Fields",
    keywords: "k",
    customProperties: [{ name: "Client", type: "text", value: "Acme" }],
  };
  return model;
}

const CATALOGUE: FieldSpec[] = [
  { kind: "date", format: "d MMMM yyyy", cached: "4 March 2026" },
  { kind: "time", format: "HH:mm", cached: "14:30" },
  { kind: "author", cached: "Ada" },
  { kind: "docTitle", cached: "Spec" },
  { kind: "docSubject", cached: "Fields" },
  { kind: "docKeywords", cached: "k" },
  { kind: "fileName", cached: "report.docx" },
  { kind: "fileName", path: true, cached: "/d/report.docx" },
  { kind: "wordCount", cached: "12" },
  { kind: "docProperty", name: "Client", cached: "Acme" },
  { kind: "page", format: "decimal" },
  { kind: "numpages", format: "decimal" },
  { kind: "ref", target: "_h0_intro", show: "text", cached: "Introduction" },
  { kind: "ref", target: "_h0_intro", show: "position", cached: "above" },
  { kind: "pageRef", target: "_h0_intro", show: "page", cached: "2" },
  // SPEC-055 adds SEQ to the field catalogue: what SPEC-053 preserved as an
  // unsupported field now round-trips as a caption `seq` field.
  { kind: "seq", label: "Figure", format: "decimal", cached: "7" },
];

describe("docx field round-trip (SPEC-053 AC-47, AC-39, AC-40)", () => {
  it("round-trips one field of each catalogue type through write/read", async () => {
    const model = modelWithFields(CATALOGUE);
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const blocks = reopened.sections[0].blocks as Paragraph[];
    expect(blocks).toHaveLength(CATALOGUE.length);
    CATALOGUE.forEach((expected, index) => {
      expect(blocks[index].runs[0].field).toEqual(expected);
    });
  });

  it("keeps an unsupported field's code and cached result byte-for-byte", async () => {
    const model = modelWithFields([{ kind: "unsupported", code: "GREETINGLINE \\* MERGEFORMAT", cached: "Dear Sir" }]);
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const field = (reopened.sections[0].blocks[0] as Paragraph).runs[0].field;
    expect(field).toEqual({ kind: "unsupported", code: "GREETINGLINE \\* MERGEFORMAT", cached: "Dear Sir" });
  });
});

describe("docx complex-field reading (SPEC-053 AC-47)", () => {
  it("reads a run-split w:fldChar/w:instrText DATE field into a date field with cached result", async () => {
    const body =
      "<w:p>" +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> DATE \\@ "d MMMM yyyy" </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      "<w:r><w:t>4 March 2026</w:t></w:r>" +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      "</w:p>";
    const bytes = await makeDocx(bodyDocx(body));
    const parsed = await parseDocx(bytes);
    const run = (parsed.sections[0].blocks[0] as Paragraph).runs[0];
    expect(run.field).toEqual({ kind: "date", format: "d MMMM yyyy", cached: "4 March 2026" });
    // One caret atom, not the literal result text spread across runs.
    expect((parsed.sections[0].blocks[0] as Paragraph).runs).toHaveLength(1);
  });

  it("reads an unknown complex field into an unsupported field", async () => {
    const body =
      "<w:p>" +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      "<w:r><w:instrText>ADDRESSBLOCK</w:instrText></w:r>" +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      "<w:r><w:t>123 Main St</w:t></w:r>" +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      "</w:p>";
    const parsed = await parseDocx(await makeDocx(bodyDocx(body)));
    const run = (parsed.sections[0].blocks[0] as Paragraph).runs[0];
    expect(run.field?.kind).toBe("unsupported");
    expect(run.field).toMatchObject({ cached: "123 Main St" });
  });

  it("reads a compact w:fldSimple AUTHOR field", async () => {
    const body = '<w:p><w:fldSimple w:instr="AUTHOR"><w:r><w:t>Ada</w:t></w:r></w:fldSimple></w:p>';
    const parsed = await parseDocx(await makeDocx(bodyDocx(body)));
    const run = (parsed.sections[0].blocks[0] as Paragraph).runs[0];
    expect(run.field).toEqual({ kind: "author", cached: "Ada" });
  });
});

describe("docx injection safety (SPEC-053 AC-36, AC-38)", () => {
  it("escapes a field's markup cached value and round-trips it verbatim", async () => {
    const markup = '<w:injected/>& "quoted"';
    const model = modelWithFields([{ kind: "docProperty", name: "Client", cached: markup }]);
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const field = (reopened.sections[0].blocks[0] as Paragraph).runs[0].field;
    expect(field).toEqual({ kind: "docProperty", name: "Client", cached: markup });
  });
});
