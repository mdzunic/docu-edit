// SPEC-055 fixtures (AC-45..53): guards the committed caption fixtures
// `e2e/fixtures/word-captions.docx` and `e2e/fixtures/lo-captions.odt` — three
// figure captions and two table captions interleaved plus one custom-label
// caption, one caption above its object and the rest below, one caption
// separated from its object by a paragraph, a cross-reference of each format
// and one broken cross-reference, and a generated table of figures.
//
// READ-ONLY on e2e/fixtures like fixtures.fields.spec.ts: generation lives in
// scripts/build-spec055-fixtures.mjs so a `npm test` never rewrites a file the
// e2e suite may be reading; the assertions run against the committed bytes.
import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, it, expect, beforeAll } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { equalStructural } from "./equality";
import { FIXED_MODIFIED, buildCaptionFixtures } from "./fixtures.captions";
import { captionDisplayText, isCaptionParagraph, listCaptions } from "./captions";
import { isTof, type Block, type DocumentModel, type Paragraph, type TableOfFigures } from "./documentModel";

const FIXTURE_DIR = path.resolve(__dirname, "../../e2e/fixtures");

const REGENERATE_HINT =
  "The committed fixture no longer matches src/lib/fixtures.captions.ts. " +
  "Regenerate it with `npx vite-node scripts/build-spec055-fixtures.mjs` and commit the result.";

function refFields(model: Pick<DocumentModel, "sections">) {
  const found = [];
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "paragraph") {
        for (const run of block.runs) {
          if (run.field?.kind === "ref" || run.field?.kind === "pageRef") {
            found.push(run.field);
          }
        }
      }
    }
  }
  return found;
}

function captionParagraphs(model: Pick<DocumentModel, "sections">): Paragraph[] {
  const found: Paragraph[] = [];
  for (const section of model.sections) {
    for (const block of section.blocks as Block[]) {
      if (isCaptionParagraph(block)) {
        found.push(block);
      }
    }
  }
  return found;
}

let docxBytes: Uint8Array;
let odtBytes: Uint8Array;

beforeAll(async () => {
  docxBytes = await fs.readFile(path.join(FIXTURE_DIR, "word-captions.docx"));
  odtBytes = await fs.readFile(path.join(FIXTURE_DIR, "lo-captions.odt"));
});

describe("SPEC-055 committed caption fixtures match the fixture model", () => {
  it("is regenerated from src/lib/fixtures.captions.ts", async () => {
    const built = await buildCaptionFixtures();
    const committed: Record<string, Uint8Array> = {
      "word-captions.docx": docxBytes,
      "lo-captions.odt": odtBytes,
    };
    for (const { name, bytes } of built) {
      const parse = name.endsWith(".odt") ? parseOdt : parseDocx;
      const fresh = await parse(bytes);
      const shipped = await parse(committed[name]);
      expect(equalStructural(shipped.sections, fresh.sections), `${name}: ${REGENERATE_HINT}`).toBe(true);
    }
  });
});

describe("SPEC-055 .docx caption fixture (AC-45, AC-49, AC-50, AC-51)", () => {
  it("opens every caption at its saved number (AC-45)", async () => {
    const reopened = await parseDocx(docxBytes);
    const captions = captionParagraphs(reopened);
    const texts = captions.map(captionDisplayText);
    expect(texts).toContain("Figure 1: Revenue by quarter");
    expect(texts).toContain("Figure 3: Forecast");
    expect(texts).toContain("Table 2: Q2 results");
    expect(texts).toContain("Diagram 1: System architecture");
    // Three figures and two tables, plus the custom label.
    const labels = listCaptions(reopened as DocumentModel).map((c) => c.label);
    expect(labels.filter((l) => l === "Figure")).toHaveLength(3);
    expect(labels.filter((l) => l === "Table")).toHaveLength(2);
    expect(labels).toContain("Diagram");
  });

  it("resolves each cross-reference format and lists the broken one (AC-49, AC-50)", async () => {
    const reopened = await parseDocx(docxBytes);
    const shows = refFields(reopened).map((f) => f.show);
    for (const show of ["labelAndNumber", "full", "numberOnly", "textOnly", "page", "position"]) {
      expect(shows).toContain(show);
    }
    expect(refFields(reopened).some((f) => f.cached === "Error! Reference source not found.")).toBe(true);
  });

  it("regenerates a Word-authored table of figures with three entries (AC-51)", async () => {
    const reopened = await parseDocx(docxBytes);
    const tof = reopened.sections[0].blocks.find(isTof) as TableOfFigures;
    expect(tof.label).toBe("Figure");
    expect(tof.entries).toHaveLength(3);
    expect(tof.entries[0].text).toBe("Figure 1: Revenue by quarter");
  });

  it("leaves captions and the table of figures unchanged on a no-edit round trip (AC-52)", async () => {
    const reopened = await parseDocx(docxBytes);
    const rewritten = (await writeDocx({ ...reopened, id: "word-captions", title: "word-captions.docx" }, { modified: FIXED_MODIFIED })).bytes;
    const reparsed = await parseDocx(rewritten);
    expect(equalStructural(reparsed.sections, reopened.sections)).toBe(true);
  });
});

describe("SPEC-055 .odt caption fixture (AC-47, AC-48, AC-50)", () => {
  it("opens every caption at its saved number (AC-47)", async () => {
    const reopened = await parseOdt(odtBytes);
    const texts = captionParagraphs(reopened).map(captionDisplayText);
    expect(texts).toContain("Figure 2: Year-over-year growth");
    expect(texts).toContain("Table 1: Q1 results");
  });

  it("resolves references and renders the broken one (AC-50)", async () => {
    const reopened = await parseOdt(odtBytes);
    expect(refFields(reopened).some((f) => f.cached === "Error! Reference source not found.")).toBe(true);
  });

  it("regenerates the table of figures (AC-51)", async () => {
    const reopened = await parseOdt(odtBytes);
    const tof = reopened.sections[0].blocks.find(isTof) as TableOfFigures;
    expect(tof.entries).toHaveLength(3);
  });
});
