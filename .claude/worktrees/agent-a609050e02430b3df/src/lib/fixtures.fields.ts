// The SPEC-053 field fixture model, shared by the generator that writes the
// committed packages (scripts/build-spec053-fixtures.mjs) and the unit spec
// that guards them (fixtures.fields.spec.ts). It builds bytes and nothing
// else — no filesystem access, no side effects on import — so importing it
// can never rewrite a fixture the e2e suite may be reading.
//
// The fixtures carry one field of each catalogue type, one Date field in a
// header, one locked field, a cross-reference to a heading and one to a
// bookmark, a cross-reference whose target has been deleted, a field with a
// picture string outside the format table, and one unsupported type.
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import {
  createDocument,
  createFieldSpecRun,
  type Block,
  type DocumentModel,
  type FieldSpec,
  type Paragraph,
} from "./documentModel";

/** Pinned so a regenerated fixture's dcterms:modified / dc:date does not
 * become whichever day someone re-ran the generator — see SPEC-008's
 * round-trip determinism rule. */
export const FIXED_MODIFIED = new Date(Date.UTC(2026, 2, 4, 14, 30, 0));

// One field of each catalogue type, cached at the values the fixtures ship.
const CATALOGUE: FieldSpec[] = [
  { kind: "date", format: "d MMMM yyyy", cached: "4 March 2026" },
  { kind: "time", format: "HH:mm", cached: "14:30" },
  { kind: "author", cached: "Ada Lovelace" },
  { kind: "docTitle", cached: "Field Fixtures" },
  { kind: "docSubject", cached: "SPEC-053" },
  { kind: "docKeywords", cached: "fields, docx, odt" },
  { kind: "fileName", cached: "word-fields.docx" },
  { kind: "wordCount", cached: "42" },
  { kind: "docProperty", name: "Client", cached: "Acme Corp" },
  { kind: "page", format: "decimal" },
  { kind: "numpages", format: "decimal" },
  // A cross-reference to a heading (its hidden bookmark) and to a bookmark.
  { kind: "ref", target: "_h0_Introduction", show: "text", cached: "Introduction" },
  { kind: "pageRef", target: "clause_liability", show: "page", cached: "2" },
];

// The extra fixture cases the spec enumerates.
const LOCKED: FieldSpec = { kind: "date", format: "d MMMM yyyy", cached: "1 March 2025", locked: true };
const DELETED_TARGET: FieldSpec = { kind: "ref", target: "gone", show: "text", cached: "Error! Reference source not found." };
const UNKNOWN_PICTURE: FieldSpec = { kind: "date", format: "MMMM/yy", cached: "March/26" };
// GREETINGLINE has no model field of its own — a genuinely unsupported type.
// (SPEC-055 promoted SEQ, this fixture's old unsupported instruction, to a real
// caption `seq` field, so it can no longer stand in for the unsupported case.)
const UNSUPPORTED_DOCX: FieldSpec = { kind: "unsupported", code: "GREETINGLINE \\* MERGEFORMAT", cached: "7" };
const UNSUPPORTED_ODT: FieldSpec = {
  kind: "unsupported",
  code: '<text:conditional-text text:condition="ooow:x==1" text:string-value-if-true="Yes" text:string-value-if-false="No">No</text:conditional-text>',
  cached: "No",
};

/** Every catalogue kind the fixtures are required to contain. */
export const FIXTURE_FIELD_KINDS: FieldSpec["kind"][] = [
  "date",
  "time",
  "author",
  "docTitle",
  "docSubject",
  "docKeywords",
  "fileName",
  "wordCount",
  "docProperty",
  "page",
  "numpages",
  "ref",
  "pageRef",
  "unsupported",
];

function fieldParagraph(field: FieldSpec): Paragraph {
  return { type: "paragraph", runs: [createFieldSpecRun(field)] };
}

/** The shared fixture model, parameterised by the format's unsupported field. */
function fixtureModel(unsupported: FieldSpec, id: string, fileName: string): DocumentModel {
  const model = createDocument(id);
  const headingPara: Paragraph = {
    type: "paragraph",
    headingLevel: 1,
    runs: [{ text: "Introduction", formatting: {} }],
    bookmarks: [{ name: "_h0_Introduction", offset: 0, hidden: true }],
  };
  const bookmarkPara: Paragraph = {
    type: "paragraph",
    runs: [{ text: "Liability is capped.", formatting: {} }],
    bookmarks: [{ name: "clause_liability", offset: 0 }],
  };
  const blocks: Block[] = [
    headingPara,
    bookmarkPara,
    ...CATALOGUE.map(fieldParagraph),
    fieldParagraph(LOCKED),
    fieldParagraph(DELETED_TARGET),
    fieldParagraph(UNKNOWN_PICTURE),
    fieldParagraph(unsupported),
  ];
  model.sections = [
    {
      blocks,
      headers: { default: { blocks: [fieldParagraph({ kind: "date", format: "dd/MM/yyyy", cached: "04/03/2026" })] } },
    },
  ];
  model.properties = {
    author: "Ada Lovelace",
    title: "Field Fixtures",
    subject: "SPEC-053",
    keywords: "fields, docx, odt",
    customProperties: [{ name: "Client", type: "text", value: "Acme Corp" }],
  };
  model.title = fileName;
  return model;
}

/** The .docx fixture's model, as the generator and the spec both build it. */
export function docxFixtureModel(): DocumentModel {
  return fixtureModel(UNSUPPORTED_DOCX, "word-fields", "word-fields.docx");
}

/** The .odt fixture's model. */
export function odtFixtureModel(): DocumentModel {
  const model = fixtureModel(UNSUPPORTED_ODT, "lo-fields", "lo-fields.odt");
  model.format = "odt";
  return model;
}

/** File name → bytes for both committed fixtures, built with the pinned
 * `modified` stamp so two builds of an unchanged model agree. */
export async function buildFieldFixtures(): Promise<{ name: string; bytes: Uint8Array }[]> {
  const docx = await writeDocx(docxFixtureModel(), { modified: FIXED_MODIFIED });
  const odt = await writeOdt(odtFixtureModel(), { modified: FIXED_MODIFIED });
  return [
    { name: "word-fields.docx", bytes: docx.bytes },
    { name: "lo-fields.odt", bytes: odt.bytes },
  ];
}
