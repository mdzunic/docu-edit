// SPEC-053: fields through the .odt reader and writer — the text: elements
// named in the catalogue, number:date-style for date/time, text:user-field-decls
// for custom properties, and an unsupported field type preserved verbatim.
import { describe, expect, it } from "vitest";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
import {
  createDocument,
  createFieldSpecRun,
  type DocumentModel,
  type FieldSpec,
  type Paragraph,
} from "../documentModel";

function fieldParagraph(field: FieldSpec): Paragraph {
  return { type: "paragraph", runs: [createFieldSpecRun(field)] };
}

function modelWithFields(fields: FieldSpec[]): DocumentModel {
  const model = createDocument("odt-fields");
  model.format = "odt";
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
  { kind: "fileName", cached: "report.odt" },
  { kind: "wordCount", cached: "12" },
  { kind: "docProperty", name: "Client", cached: "Acme" },
  { kind: "page", format: "decimal" },
  { kind: "numpages", format: "decimal" },
  { kind: "ref", target: "target1", show: "text", cached: "Introduction" },
  { kind: "ref", target: "target1", show: "position", cached: "above" },
  { kind: "pageRef", target: "target1", show: "page", cached: "2" },
];

describe("odt field round-trip (SPEC-053 AC-48)", () => {
  it("round-trips one field of each catalogue type through write/read", async () => {
    const model = modelWithFields(CATALOGUE);
    const { bytes } = await writeOdt(model);
    const reopened = await parseOdt(bytes);
    const blocks = reopened.sections[0].blocks as Paragraph[];
    CATALOGUE.forEach((expected, index) => {
      expect(blocks[index].runs[0].field).toEqual(expected);
    });
  });

  it("declares a text:user-field-decls entry for each docProperty field", async () => {
    const model = modelWithFields([{ kind: "docProperty", name: "Client", cached: "Acme" }]);
    const { bytes } = await writeOdt(model);
    const { unzip } = await import("../core/zip");
    const entries = await unzip(bytes);
    const content = new TextDecoder().decode(entries.get("content.xml")!);
    expect(content).toContain("<text:user-field-decls>");
    expect(content).toContain('text:name="Client"');
    expect(content).toContain('office:string-value="Acme"');
  });

  it("preserves an unsupported text:conditional-text field verbatim and shows its value", async () => {
    const model = createDocument("odt-unsupported");
    model.format = "odt";
    const code = '<text:conditional-text text:condition="ooow:x==1" text:string-value-if-true="Yes" text:string-value-if-false="No">No</text:conditional-text>';
    model.sections = [{ blocks: [fieldParagraph({ kind: "unsupported", code, cached: "No" })] }];
    const { bytes } = await writeOdt(model);
    const reopened = await parseOdt(bytes);
    const field = (reopened.sections[0].blocks[0] as Paragraph).runs[0].field;
    expect(field?.kind).toBe("unsupported");
    expect(field).toMatchObject({ cached: "No" });
  });
});
