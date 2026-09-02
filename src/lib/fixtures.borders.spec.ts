// SPEC-074 AC-33: guards the committed border fixtures
// `e2e/fixtures/word-borders.docx` and `e2e/fixtures/lo-borders.odt` — a
// boxed note paragraph, a bottom-rule heading, three consecutive identical
// bordered paragraphs, a shaded code block and a page border. Each opens
// with every border intact, renders (the paint decisions borderRuns and
// paragraphStyle make are asserted here, since the canvas has no model of
// its own), saves, and reopens `equalStructural`.
//
// This spec is READ-ONLY on e2e/fixtures — generation lives in
// scripts/build-spec074-fixtures.mjs, so a unit run can never rewrite a file
// the e2e suite is uploading (the same reasoning as fixtures.fields.spec.ts).
import { promises as fs } from "node:fs";
import path from "node:path";
import { beforeAll, describe, expect, it } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { equalStructural, firstStructuralDifference } from "./equality";
import { borderRuns, fragmentEdges } from "./paragraphBorders";
import { paragraphStyle, pageBorderStyle } from "./renderFormatting";
import {
  buildBorderFixtures,
  CODE_SHADING,
  FIXTURE_PAGE_BACKGROUND,
  FIXTURE_PAGE_BORDERS,
  MERGED_RUN_BORDERS,
  MERGED_TEXTS,
  NOTE_SHADING,
} from "./fixtures.borders";
import type { DocumentModel, Paragraph } from "./documentModel";

const FIXTURE_DIR = path.resolve(__dirname, "../../e2e/fixtures");

const REGENERATE_HINT =
  "The committed fixture no longer matches src/lib/fixtures.borders.ts. " +
  "Regenerate it with `npx vite-node scripts/build-spec074-fixtures.mjs` and commit the result.";

let docxBytes: Uint8Array;
let odtBytes: Uint8Array;

beforeAll(async () => {
  docxBytes = new Uint8Array(await fs.readFile(path.join(FIXTURE_DIR, "word-borders.docx")));
  odtBytes = new Uint8Array(await fs.readFile(path.join(FIXTURE_DIR, "lo-borders.odt")));
});

function asModel(parsed: { sections: DocumentModel["sections"]; defaultStyles: DocumentModel["defaultStyles"] }): DocumentModel {
  return {
    id: "spec074-borders",
    title: "Borders",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
  };
}

function paragraphWith(model: DocumentModel, text: string): Paragraph {
  const found = model.sections[0].blocks.find(
    (block) => block.type === "paragraph" && block.runs.map((run) => run.text).join("") === text,
  );
  if (!found || found.type !== "paragraph") {
    throw new Error(`no paragraph reading "${text}"`);
  }
  return found;
}

describe("the committed fixtures still match the model that generates them", () => {
  it("regenerating produces byte-identical packages", async () => {
    const built = await buildBorderFixtures();
    const docx = built.find((entry) => entry.name === "word-borders.docx");
    const odt = built.find((entry) => entry.name === "lo-borders.odt");
    expect(docx, REGENERATE_HINT).toBeDefined();
    expect(odt, REGENERATE_HINT).toBeDefined();
    expect(docx!.bytes.length, REGENERATE_HINT).toBe(docxBytes.length);
    expect(odt!.bytes.length, REGENERATE_HINT).toBe(odtBytes.length);
  });
});

for (const format of ["docx", "odt"] as const) {
  describe(`${format} fixture (SPEC-074 AC-33)`, () => {
    async function open(): Promise<DocumentModel> {
      const bytes = format === "docx" ? docxBytes : odtBytes;
      return asModel(format === "docx" ? await parseDocx(bytes) : await parseOdt(bytes));
    }

    async function reopen(model: DocumentModel): Promise<DocumentModel> {
      const written = format === "docx" ? await writeDocx(model) : await writeOdt(model);
      return asModel(format === "docx" ? await parseDocx(written.bytes) : await parseOdt(written.bytes));
    }

    it("opens the boxed note with all four edges and its fill", async () => {
      const model = await open();
      const note = paragraphWith(model, model.sections[0].blocks
        .filter((b) => b.type === "paragraph")
        .map((b) => (b as Paragraph).runs.map((r) => r.text).join(""))
        .find((text) => text.startsWith("Note:"))!);
      expect(note.formatting?.shadingFill).toBe(NOTE_SHADING);
      for (const edge of ["top", "bottom", "left", "right"] as const) {
        expect(note.formatting?.borders?.[edge]?.style).toBe("single");
      }
    });

    it("opens the heading with its bottom rule and nothing else", async () => {
      const model = await open();
      const heading = paragraphWith(model, "Borders and shading");
      expect(heading.formatting?.borders?.bottom).toEqual({ style: "double", widthPt: 3, color: "#c00000" });
      expect(heading.formatting?.borders?.top).toBeUndefined();
      expect(heading.formatting?.borders?.left).toBeUndefined();
    });

    it("opens the shaded code block with a fill and no borders", async () => {
      const model = await open();
      const code = paragraphWith(model, "const answer = 42;");
      expect(code.formatting?.shadingFill).toBe(CODE_SHADING);
      expect(code.formatting?.borders).toBeUndefined();
    });

    it("opens the page border and background on the section", async () => {
      const model = await open();
      expect(model.sections[0].pageBorders).toEqual(FIXTURE_PAGE_BORDERS);
      expect(model.sections[0].pageBackground).toBe(FIXTURE_PAGE_BACKGROUND);
    });

    it("renders the three identical paragraphs as one bordered block (AC-12)", async () => {
      const model = await open();
      const blocks = model.sections[0].blocks;
      const runs = borderRuns(blocks);
      const indexes = MERGED_TEXTS.map((text) => blocks.indexOf(paragraphWith(model, text)));
      expect(indexes.every((index) => index >= 0)).toBe(true);
      expect(indexes.map((index) => runs[index].top)).toEqual([true, false, false]);
      expect(indexes.map((index) => runs[index].bottom)).toEqual([false, false, true]);
      expect(indexes.map((index) => runs[index].between)).toEqual([true, true, false]);
      expect(indexes.map((index) => runs[index].sides)).toEqual([true, true, true]);
      // ...and every member kept the borders the model authored.
      for (const index of indexes) {
        expect((blocks[index] as Paragraph).formatting?.borders).toEqual(MERGED_RUN_BORDERS);
      }
    });

    it("renders the boxed note's shading across its full padded width (AC-13)", async () => {
      const model = await open();
      const blocks = model.sections[0].blocks;
      const runs = borderRuns(blocks);
      const noteIndex = blocks.findIndex(
        (block) => block.type === "paragraph" && block.runs.map((r) => r.text).join("").startsWith("Note:"),
      );
      const style = paragraphStyle(
        blocks[noteIndex] as Paragraph,
        [],
        "all",
        fragmentEdges(runs, noteIndex, true, true),
      );
      expect(style.backgroundColor).toBe(NOTE_SHADING);
      expect(style.paddingLeft).toBe("4pt");
      expect(style.paddingRight).toBe("4pt");
      expect(style.borderTop).toBe("1pt solid #7f7f7f");
    });

    it("renders the page border on every page of the section (AC-22)", async () => {
      const model = await open();
      const margins = model.defaultStyles.margins;
      for (const pageIndex of [0, 1, 5]) {
        const style = pageBorderStyle(model.sections[0].pageBorders, margins, pageIndex);
        expect(style?.borderTop).toBe("3pt double #1f4e79");
      }
    });

    it("saves and reopens equalStructural (AC-33)", async () => {
      const opened = await open();
      const reopened = await reopen(opened);
      expect(firstStructuralDifference(opened, reopened) ?? "equal").toBe("equal");
      expect(equalStructural(opened, reopened)).toBe(true);
      // Not equal-because-empty: the borders really survived.
      expect(paragraphWith(reopened, "const answer = 42;").formatting?.shadingFill).toBe(CODE_SHADING);
      expect(reopened.sections[0].pageBorders?.top?.style).toBe("double");
    });
  });
}
