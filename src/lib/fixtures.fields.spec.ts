// SPEC-053 fixtures (AC-51, AC-52, AC-53): guards the committed field
// fixtures `e2e/fixtures/word-fields.docx` and `e2e/fixtures/lo-fields.odt`
// — one field of each catalogue type, one Date field in a header, one locked
// field, a cross-reference to a heading and one to a bookmark, a
// cross-reference whose target has been deleted, a field with a picture
// string outside the format table, and one unsupported type. Every field
// opens at its saved value and a no-edit round trip leaves every cached
// result unchanged.
//
// This spec is READ-ONLY on e2e/fixtures. It used to regenerate both files in
// a beforeAll, which meant every `npm test` rewrote two committed fixtures
// that e2e/SPEC-053.spec.ts uploads — a unit run overlapping an e2e run in
// the same checkout could have Playwright read a file mid-write. Generation
// now lives in scripts/build-spec053-fixtures.mjs, and the assertions below
// run against the bytes that are actually committed, so they test the
// artifact the e2e suite loads rather than an in-memory rebuild of it.
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
  FIXED_MODIFIED,
  FIXTURE_FIELD_KINDS,
  buildFieldFixtures,
} from "./fixtures.fields";
import { documentFileName, type Block, type DocumentModel, type FieldSpec } from "./documentModel";

const FIXTURE_DIR = path.resolve(__dirname, "../../e2e/fixtures");

const REGENERATE_HINT =
  "The committed fixture no longer matches src/lib/fixtures.fields.ts. " +
  "Regenerate it with `npx vite-node scripts/build-spec053-fixtures.mjs` and commit the result.";

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
  docxBytes = await fs.readFile(path.join(FIXTURE_DIR, "word-fields.docx"));
  odtBytes = await fs.readFile(path.join(FIXTURE_DIR, "lo-fields.odt"));
});

// The drift guard the old in-place rewrite provided implicitly: if someone
// edits the fixture model without re-running the generator, the committed
// packages and the model part company and this fails with the command to fix
// it. Compared structurally rather than byte-for-byte on purpose — the
// writers compress through the platform's CompressionStream, whose exact
// deflate output is not guaranteed to match across Node versions, so byte
// equality would fail for reasons that have nothing to do with drift.
describe("SPEC-053 committed field fixtures match the fixture model", () => {
  it("is regenerated from src/lib/fixtures.fields.ts", async () => {
    const built = await buildFieldFixtures();
    const committed: Record<string, Uint8Array> = {
      "word-fields.docx": docxBytes,
      "lo-fields.odt": odtBytes,
    };
    for (const { name, bytes } of built) {
      const parse = name.endsWith(".odt") ? parseOdt : parseDocx;
      const fresh = await parse(bytes);
      const shipped = await parse(committed[name]);
      expect(equalStructural(shipped.sections, fresh.sections), `${name}: ${REGENERATE_HINT}`).toBe(true);
      expect(allFields(shipped), `${name}: ${REGENERATE_HINT}`).toEqual(allFields(fresh));
    }
  });
});

describe("SPEC-053 .docx field fixture (AC-51, AC-52, AC-53)", () => {
  it("opens one field of each catalogue type at its saved value", async () => {
    const reopened = await parseDocx(docxBytes);
    const fields = allFields(reopened);
    // Every catalogue kind is present.
    const kinds = new Set(fields.map((field) => field.kind));
    for (const kind of FIXTURE_FIELD_KINDS) {
      expect(kinds.has(kind)).toBe(true);
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
    for (const kind of FIXTURE_FIELD_KINDS) {
      expect(kinds.has(kind)).toBe(true);
    }
  });

  it("leaves every cached result unchanged on a no-edit round trip", async () => {
    const reopened = await parseOdt(odtBytes);
    const rewritten = (await writeOdt({ ...reopened, id: "lo-fields", title: "lo-fields.odt", format: "odt" }, { modified: FIXED_MODIFIED })).bytes;
    const reparsed = await parseOdt(rewritten);
    expect(allFields(reparsed)).toEqual(allFields(reopened));
  });
});
