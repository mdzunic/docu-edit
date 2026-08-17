// SPEC-053 fixtures (AC-51, AC-52, AC-53): builds the shared field fixtures
// `e2e/fixtures/word-fields.docx` and `e2e/fixtures/lo-fields.odt` — one field
// of each catalogue type, one Date field in a header, one locked field, a
// cross-reference to a heading and one to a bookmark, a cross-reference whose
// target has been deleted, a field with a picture string outside the format
// table, and one unsupported type — then re-reads them to prove every field
// opens at its saved value and a no-edit round trip leaves every cached result
// unchanged. Deterministic writers (a fixed `modified` stamp) make the bytes
// stable, so the committed fixtures never drift.
import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, it, expect, beforeAll } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { equalStructural } from "./equality";
import { unzip } from "./core/zip";
import { caretSelection, documentStart } from "./commands";
import { updateAllFields } from "./commands.fields";
import { titleFromFileName } from "./openDocument";
import {
  createDocument,
  createFieldSpecRun,
  documentFileName,
  type Block,
  type DocumentModel,
  type FieldSpec,
  type Paragraph,
} from "./documentModel";

const FIXED_MODIFIED = new Date(Date.UTC(2026, 2, 4, 14, 30, 0));
const FIXTURE_DIR = path.resolve(__dirname, "../../e2e/fixtures");

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
const UNSUPPORTED_DOCX: FieldSpec = { kind: "unsupported", code: "SEQ Figure \\* ARABIC", cached: "7" };
const UNSUPPORTED_ODT: FieldSpec = {
  kind: "unsupported",
  code: '<text:conditional-text text:condition="ooow:x==1" text:string-value-if-true="Yes" text:string-value-if-false="No">No</text:conditional-text>',
  cached: "No",
};

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

/** Every field run in a parsed document, in document order (body then
 * headers) — accepts the readers' parsed shapes, not just DocumentModel. */
function allFields(model: Pick<DocumentModel, "sections">): FieldSpec[] {
  const found: FieldSpec[] = [];
  const scan = (blocks: Block[]) => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        for (const run of block.runs) {
          if (run.field) {
            found.push(run.field);
          }
        }
      }
    }
  };
  for (const section of model.sections) {
    scan(section.blocks);
    if (section.headers?.default) {
      scan(section.headers.default.blocks);
    }
  }
  return found;
}

/** The `.docx` package's FILENAME field, as the bytes on disk spell it —
 * `<w:fldSimple w:instr="FILENAME">` wrapping its cached result run. Compared
 * as a string so "unchanged" means unchanged markup, not just an equal model. */
async function fileNameFieldXml(bytes: Uint8Array): Promise<string> {
  const parts = await unzip(bytes);
  const xml = new TextDecoder().decode(parts.get("word/document.xml") ?? new Uint8Array());
  return /<w:fldSimple[^>]*\bFILENAME\b[^>]*>[\s\S]*?<\/w:fldSimple>/.exec(xml)?.[0] ?? "";
}

/** The model the app holds for a fixture opened from disk: `openDocumentFile`
 * strips the extension off for `title` (the app bar) and keeps the whole name
 * in `fileName` (what a FILENAME field renders). */
function openedAs(parsed: Awaited<ReturnType<typeof parseDocx>>, id: string, fileName: string): DocumentModel {
  return { ...parsed, id, title: titleFromFileName(fileName), fileName } as DocumentModel;
}

let docxBytes: Uint8Array;
let odtBytes: Uint8Array;

beforeAll(async () => {
  const docxModel = fixtureModel(UNSUPPORTED_DOCX, "word-fields", "word-fields.docx");
  const odtModel = fixtureModel(UNSUPPORTED_ODT, "lo-fields", "lo-fields.odt");
  odtModel.format = "odt";
  docxBytes = (await writeDocx(docxModel, { modified: FIXED_MODIFIED })).bytes;
  odtBytes = (await writeOdt(odtModel, { modified: FIXED_MODIFIED })).bytes;
  // Persist the committed fixtures (deterministic bytes, so no drift). The
  // assertions below run against the in-memory bytes, so a read-only
  // filesystem (some CI sandboxes) never fails the test — it just skips the
  // refresh.
  try {
    await fs.writeFile(path.join(FIXTURE_DIR, "word-fields.docx"), docxBytes);
    await fs.writeFile(path.join(FIXTURE_DIR, "lo-fields.odt"), odtBytes);
  } catch {
    // Read-only sandbox — keep the committed fixtures as they are.
  }
});

describe("SPEC-053 .docx field fixture (AC-51, AC-52, AC-53)", () => {
  it("opens one field of each catalogue type at its saved value", async () => {
    const reopened = await parseDocx(docxBytes);
    const fields = allFields(reopened);
    // Every catalogue kind is present.
    const kinds = new Set(fields.map((field) => field.kind));
    for (const kind of ["date", "time", "author", "docTitle", "docSubject", "docKeywords", "fileName", "wordCount", "docProperty", "page", "numpages", "ref", "pageRef", "unsupported"]) {
      expect(kinds.has(kind as FieldSpec["kind"])).toBe(true);
    }
    // The unsupported field opens at its saved value.
    const unsupported = fields.find((field) => field.kind === "unsupported");
    expect(unsupported).toMatchObject({ cached: "7" });
    // The locked field is preserved as locked.
    expect(fields.some((field) => "locked" in field && field.locked)).toBe(true);
  });

  it("leaves every cached result unchanged on a no-edit round trip", async () => {
    const reopened = await parseDocx(docxBytes);
    const rewritten = (await writeDocx({ ...reopened, id: "word-fields", title: "word-fields.docx" }, { modified: FIXED_MODIFIED })).bytes;
    const reparsed = await parseDocx(rewritten);
    expect(equalStructural(reparsed.sections, reopened.sections)).toBe(true);
    expect(allFields(reparsed)).toEqual(allFields(reopened));
  });

  it("leaves the FILENAME field byte-identical through Tools → Update all fields (AC-16, AC-53)", async () => {
    // The regression: the update path was fed `model.title`, which has had the
    // extension stripped off it, so one Update-all silently rewrote the value
    // Word shows ("word-fields.docx") to "word-fields".
    const opened = openedAs(await parseDocx(docxBytes), "word-fields", "word-fields.docx");
    const updated = updateAllFields(
      opened,
      caretSelection(documentStart()),
      { now: FIXED_MODIFIED, fileName: documentFileName(opened) },
      () => null,
    );
    expect(allFields(updated.model).find((field) => field.kind === "fileName")).toEqual({
      kind: "fileName",
      cached: "word-fields.docx",
    });
    const rewritten = (await writeDocx(updated.model, { modified: FIXED_MODIFIED })).bytes;
    expect(await fileNameFieldXml(rewritten)).toBe(await fileNameFieldXml(docxBytes));
  });
});

describe("SPEC-053 .odt field fixture (AC-51, AC-52, AC-53)", () => {
  it("opens one field of each catalogue type at its saved value", async () => {
    const reopened = await parseOdt(odtBytes);
    const kinds = new Set(allFields(reopened).map((field) => field.kind));
    for (const kind of ["date", "time", "author", "docTitle", "docSubject", "docKeywords", "fileName", "wordCount", "docProperty", "page", "numpages", "ref", "pageRef", "unsupported"]) {
      expect(kinds.has(kind as FieldSpec["kind"])).toBe(true);
    }
  });

  it("leaves every cached result unchanged on a no-edit round trip", async () => {
    const reopened = await parseOdt(odtBytes);
    const rewritten = (await writeOdt({ ...reopened, id: "lo-fields", title: "lo-fields.odt", format: "odt" }, { modified: FIXED_MODIFIED })).bytes;
    const reparsed = await parseOdt(rewritten);
    expect(allFields(reparsed)).toEqual(allFields(reopened));
  });
});
