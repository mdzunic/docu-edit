// SPEC-033 AC-18: the .odt TOC — `text:table-of-content` with an
// `text:index-body` of cached entries — round-tripped through this app's
// own writer/reader.
import { describe, it, expect } from "vitest";
import { parseOdt } from "./reader";
import { writeOdt } from "./writer";
import {
  createDocument,
  isToc,
  type DocumentModel,
  type Paragraph,
  type TableOfContents,
} from "../documentModel";
import { defaultTocOptions, insertTableOfContents, noTocPageNumbers } from "../toc";
import { caretSelection } from "../commands";

function heading(text: string, level: number): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }], headingLevel: level };
}

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function modelWithHeadings(): DocumentModel {
  const model = createDocument("toc-odf-test");
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

describe("odt TOC round trip (SPEC-033 AC-18)", () => {
  it("writes and reads back a TOC's entries, levels and options", async () => {
    const model = modelWithHeadings();
    const inserted = insertTableOfContents(
      model,
      caretSelection({ section: 0, block: 0, offset: 0 }),
      { levels: [1, 2], showPageNumbers: true, hyperlinksOnly: false },
      () => 1,
    ).model;
    inserted.format = "odt";

    const { bytes } = await writeOdt(inserted);
    const parsed = await parseOdt(bytes);
    const toc = parsed.sections[0].blocks.find(isToc) as TableOfContents | undefined;
    expect(toc).toBeDefined();
    expect(toc?.entries.map((e) => e.text)).toEqual(["Chapter One", "Section A"]);
    expect(toc?.entries.map((e) => e.level)).toEqual([1, 2]);
    expect(toc?.entries.every((e) => e.pageNumber === 1)).toBe(true);
    expect(toc?.options.showPageNumbers).toBe(true);
    expect(toc?.options.hyperlinksOnly).toBe(false);
  });

  it("writes hyperlinks-only entries with no page-number trailer, and reads that back", async () => {
    const model = modelWithHeadings();
    const inserted = insertTableOfContents(
      model,
      caretSelection({ section: 0, block: 0, offset: 0 }),
      { levels: [1], showPageNumbers: false, hyperlinksOnly: true },
      () => 3,
    ).model;
    inserted.format = "odt";

    const { bytes } = await writeOdt(inserted);
    const parsed = await parseOdt(bytes);
    const toc = parsed.sections[0].blocks.find(isToc) as TableOfContents | undefined;
    expect(toc?.entries).toHaveLength(1);
    expect(toc?.entries[0].pageNumber).toBeNull();
    expect(toc?.options.hyperlinksOnly).toBe(true);
  });

  it("writes and reads back the empty-TOC placeholder (AC-22, AC-23)", async () => {
    const model = createDocument("toc-empty-odt");
    model.format = "odt";
    const inserted = insertTableOfContents(
      model,
      caretSelection({ section: 0, block: 0, offset: 0 }),
      defaultTocOptions(),
      noTocPageNumbers(),
    ).model;

    const { bytes } = await writeOdt(inserted);
    const parsed = await parseOdt(bytes);
    const toc = parsed.sections[0].blocks.find(isToc) as TableOfContents | undefined;
    expect(toc?.entries).toEqual([]);
  });
});
