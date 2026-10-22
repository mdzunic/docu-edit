import { describe, it, expect } from "vitest";
import { writeMd, mdFileName, MD_MIME_TYPE } from "./writer";
import { createDocument } from "../documentModel";
import { parseMd } from "./reader";

async function textOf(model: Parameters<typeof writeMd>[0]) {
  const { bytes } = await writeMd(model);
  return new TextDecoder().decode(bytes);
}

describe("SPEC-050 — writeMd", () => {
  it("emits an ATX heading for a Heading1 paragraph", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      { type: "paragraph", styleId: "Heading1", headingLevel: 1, runs: [{ text: "Title", formatting: {} }] },
    ];
    expect(await textOf(doc)).toBe("# Title\n");
  });

  it("emits canonical bold/italic/strikethrough/code spellings regardless of source formatting flags", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [
          { text: "bold", formatting: { bold: true } },
          { text: " ", formatting: {} },
          { text: "italic", formatting: { italic: true } },
          { text: " ", formatting: {} },
          { text: "gone", formatting: { strikethrough: true } },
          { text: " ", formatting: {} },
          { text: "code", formatting: {}, styleId: "CodeChar" },
        ],
      },
    ];
    expect(await textOf(doc)).toBe("**bold** _italic_ ~~gone~~ `code`\n");
  });

  it("escapes literal Markdown-significant characters in plain text", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "1 * 2 = [two] and a `backtick`", formatting: {} }] },
    ];
    const text = await textOf(doc);
    const reparsed = await parseMd(new TextEncoder().encode(text));
    expect(reparsed.sections[0].blocks[0]).toMatchObject({
      runs: [{ text: "1 * 2 = [two] and a `backtick`" }],
    });
  });

  it("emits a fenced code block with the style's language", async () => {
    const doc = createDocument("id");
    doc.styles = [{ id: "Code:python", name: "Code", basedOn: "Code", language: "python" }];
    doc.sections[0].blocks = [
      { type: "paragraph", styleId: "Code:python", runs: [{ text: "x = 1", formatting: {} }] },
      { type: "paragraph", styleId: "Code:python", runs: [{ text: "print(x)", formatting: {} }] },
    ];
    expect(await textOf(doc)).toBe("```python\nx = 1\nprint(x)\n```\n");
  });

  it("emits a blockquote with one > per physical line", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      { type: "paragraph", styleId: "Quote", runs: [{ text: "line one\nline two", formatting: {} }] },
    ];
    expect(await textOf(doc)).toBe("> line one\\\n> line two\n");
  });

  it("emits a thematic break for the HorizontalRule style", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [{ type: "paragraph", styleId: "HorizontalRule", runs: [] }];
    expect(await textOf(doc)).toBe("---\n");
  });

  it("emits nested bulleted and numbered lists (blank-line separated — every block joins that way, and a blank line between list items does not break adjacency on re-parse)", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "one", formatting: {} }], list: { type: "bullet", level: 1 } },
      { type: "paragraph", runs: [{ text: "nested", formatting: {} }], list: { type: "bullet", level: 2 } },
    ];
    expect(await textOf(doc)).toBe("- one\n\n  - nested\n");
  });

  it("emits a GFM pipe table with the first row as header, escaping literal pipes", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      {
        type: "table",
        columnWidthsCm: [5, 5],
        rows: [
          {
            isHeader: true,
            cells: [
              { blocks: [{ type: "paragraph", runs: [{ text: "A", formatting: {} }] }] },
              { blocks: [{ type: "paragraph", runs: [{ text: "B", formatting: {} }] }] },
            ],
          },
          {
            cells: [
              { blocks: [{ type: "paragraph", runs: [{ text: "a | b", formatting: {} }] }] },
              { blocks: [{ type: "paragraph", runs: [{ text: "2", formatting: {} }] }] },
            ],
          },
        ],
      },
    ];
    expect(await textOf(doc)).toBe("| A | B |\n| --- | --- |\n| a \\| b | 2 |\n");
  });

  it("emits an inline image as a data URI", async () => {
    const doc = createDocument("id");
    doc.media = [{ mediaId: "m1", contentType: "image/png", data: "QUJD" }];
    doc.sections[0].blocks = [
      {
        type: "image",
        mediaId: "m1",
        naturalWidthPx: 10,
        naturalHeightPx: 10,
        widthCm: 1,
        heightCm: 1,
        wrap: "inline",
        altText: "a dot",
      },
    ];
    expect(await textOf(doc)).toBe("![a dot](data:image/png;base64,QUJD)\n");
  });

  it("emits front matter verbatim, first, unparsed", async () => {
    const doc = createDocument("id");
    doc.frontMatter = "---\ntitle: Hi\n---";
    doc.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "body", formatting: {} }] }];
    expect(await textOf(doc)).toBe("---\ntitle: Hi\n---\n\nbody\n");
  });

  it("re-emits an opaque block verbatim, anchored in place", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "before", formatting: {} }] },
      { type: "paragraph", runs: [{ text: "after", formatting: {} }] },
    ];
    doc.sections[0].opaqueBlocks = [{ xml: "<div>raw</div>", blockIndex: 1 }];
    expect(await textOf(doc)).toBe("before\n\n<div>raw</div>\n\nafter\n");
  });

  it("skips a manual break block (no Markdown representation)", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "before", formatting: {} }] },
      { type: "break", kind: "page" },
      { type: "paragraph", runs: [{ text: "after", formatting: {} }] },
    ];
    expect(await textOf(doc)).toBe("before\n\nafter\n");
  });

  it("emits a table of contents as its rendered text, not as a field", async () => {
    const doc = createDocument("id");
    doc.sections[0].blocks = [
      {
        type: "toc",
        options: { levels: [1], showPageNumbers: false, hyperlinksOnly: false },
        entries: [
          { text: "Intro", level: 1, anchor: "intro", pageNumber: null },
          { text: "Conclusion", level: 1, anchor: "conclusion", pageNumber: null },
        ],
      },
    ];
    expect(await textOf(doc)).toBe("Intro\n\nConclusion\n");
  });
});

describe("SPEC-050 — mdFileName / MD_MIME_TYPE", () => {
  it("appends .md to the title", () => {
    expect(mdFileName("Quarterly Report")).toBe("Quarterly Report.md");
  });

  it("falls back to the default title when blank", () => {
    expect(mdFileName("   ")).toBe("Untitled document.md");
  });

  it("is the registered Markdown content type", () => {
    expect(MD_MIME_TYPE).toBe("text/markdown");
  });
});
