// SPEC-033 AC-17, AC-19, AC-20: the .docx TOC field — a real complex field
// (fldChar begin/separate/end + instrText) with a cached result, so Word
// shows it without an "update fields?" prompt on open — round-tripped
// through this app's own writer/reader, and read back correctly from a
// hand-built fixture matching genuine Word output (including a NESTED
// PAGEREF field per entry's page number, see scripts/build-docx-fixtures.mjs).
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx } from "./reader";
import { writeDocx } from "./writer";
import { validateDocx } from "./validate";
import {
  asToc,
  createDocument,
  isToc,
  type DocumentModel,
  type Paragraph,
  type TableOfContents,
} from "../documentModel";
import {
  defaultTocOptions,
  insertTableOfContents,
  tocHeadingsStale,
  tocPageNumbersStale,
  updateTocEntirely,
  noTocPageNumbers,
} from "../toc";
import { caretSelection } from "../commands";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

function heading(text: string, level: number): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }], headingLevel: level };
}

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function modelWithHeadings(): DocumentModel {
  const model = createDocument("toc-docx-test");
  model.sections = [
    {
      blocks: [
        heading("Chapter One", 1),
        paragraph("Body under chapter one."),
        heading("Section A", 2),
        paragraph("Body under section A."),
      ],
    },
  ];
  return model;
}

describe("docx TOC round trip (SPEC-033 AC-17)", () => {
  it("writes and reads back a TOC's entries, levels and options", async () => {
    const model = modelWithHeadings();
    const inserted = insertTableOfContents(
      model,
      caretSelection({ section: 0, block: 0, offset: 0 }),
      { levels: [1, 2], showPageNumbers: true, hyperlinksOnly: false },
      () => 1,
    ).model;
    inserted.format = "docx";

    const { bytes } = await writeDocx(inserted);
    expect(await validateDocx(bytes)).toEqual([]);

    const parsed = await parseDocx(bytes);
    const toc = parsed.sections[0].blocks.find(isToc) as TableOfContents | undefined;
    expect(toc).toBeDefined();
    expect(toc?.entries.map((e) => e.text)).toEqual(["Chapter One", "Section A"]);
    expect(toc?.entries.map((e) => e.level)).toEqual([1, 2]);
    expect(toc?.entries.every((e) => e.pageNumber === 1)).toBe(true);
    expect(toc?.options.showPageNumbers).toBe(true);
    expect(toc?.options.hyperlinksOnly).toBe(false);
    expect(toc?.options.levels).toEqual([1, 2]);
  });

  it("writes hyperlinks-only entries with no page-number trailer, and reads that back", async () => {
    const model = modelWithHeadings();
    const inserted = insertTableOfContents(
      model,
      caretSelection({ section: 0, block: 0, offset: 0 }),
      { levels: [1], showPageNumbers: false, hyperlinksOnly: true },
      () => 3,
    ).model;
    inserted.format = "docx";

    const { bytes } = await writeDocx(inserted);
    expect(await validateDocx(bytes)).toEqual([]);
    const parsed = await parseDocx(bytes);
    const toc = parsed.sections[0].blocks.find(isToc) as TableOfContents | undefined;
    expect(toc?.entries).toHaveLength(1);
    expect(toc?.entries[0].pageNumber).toBeNull();
    expect(toc?.options.hyperlinksOnly).toBe(true);
  });

  it("writes and reads back the empty-TOC placeholder (AC-22, AC-23)", async () => {
    const model = createDocument("toc-empty-docx");
    model.format = "docx";
    const inserted = insertTableOfContents(
      model,
      caretSelection({ section: 0, block: 0, offset: 0 }),
      defaultTocOptions(),
      noTocPageNumbers(),
    ).model;

    const { bytes } = await writeDocx(inserted);
    expect(await validateDocx(bytes)).toEqual([]);
    const parsed = await parseDocx(bytes);
    const toc = parsed.sections[0].blocks.find(isToc) as TableOfContents | undefined;
    expect(toc?.entries).toEqual([]);
  });
});

describe("docx TOC — opening a Word-authored fixture (SPEC-033 AC-19, AC-20)", () => {
  it("renders the existing cached entries, including a nested PAGEREF page number", async () => {
    const parsed = await parseDocx(loadFixture("word-toc.docx"));
    const toc = parsed.sections[0].blocks.find(isToc) as TableOfContents | undefined;
    expect(toc).toBeDefined();
    expect(toc?.entries).toEqual([
      { text: "Introduction", level: 1, anchor: "_Toc1", pageNumber: 1 },
      { text: "Background", level: 2, anchor: "_Toc2", pageNumber: 1 },
    ]);
    expect(toc?.options.levels).toEqual([1, 2]);
    expect(toc?.options.showPageNumbers).toBe(true);

    // The headings the field's own body content still carries, untouched.
    const headingBlocks = parsed.sections[0].blocks.filter(
      (block): block is Paragraph => block.type === "paragraph" && block.headingLevel !== undefined,
    );
    expect(headingBlocks.map((b) => b.runs.map((r) => r.text).join(""))).toEqual([
      "Introduction",
      "Background",
    ]);
  });

  it("allows regeneration (AC-20): the fixture's TOC is not stale against its own headings, and updates cleanly if a heading changes", async () => {
    const parsed = await parseDocx(loadFixture("word-toc.docx"));
    const model: DocumentModel = {
      id: "toc-fixture-model",
      title: "Fixture",
      sections: parsed.sections,
      defaultStyles: createDocument("seed").defaultStyles,
    };
    const tocIndex = model.sections[0].blocks.findIndex(isToc);
    const toc = asToc(model.sections[0].blocks[tocIndex]);
    expect(tocHeadingsStale(model, toc)).toBe(false);
    expect(tocPageNumbersStale(toc, () => 1)).toBe(false);

    // Rename the "Background" heading — the TOC should now be stale, and
    // "Update entire table" should regenerate cleanly from the fixture's
    // real content.
    const headingIndex = model.sections[0].blocks.findIndex(
      (block): block is Paragraph =>
        block.type === "paragraph" &&
        block.headingLevel === 2 &&
        block.runs.map((r) => r.text).join("") === "Background",
    );
    const headingBlock = model.sections[0].blocks[headingIndex] as Paragraph;
    model.sections[0].blocks[headingIndex] = {
      ...headingBlock,
      runs: [{ text: "Background Renamed", formatting: {} }],
    };
    expect(tocHeadingsStale(model, toc)).toBe(true);

    const updated = updateTocEntirely(
      model,
      caretSelection({ section: 0, block: 0, offset: 0 }),
      0,
      tocIndex,
      () => 1,
    ).model;
    const updatedToc = asToc(updated.sections[0].blocks[tocIndex]);
    expect(updatedToc.entries.map((e) => e.text)).toEqual(["Introduction", "Background Renamed"]);
  });
});
