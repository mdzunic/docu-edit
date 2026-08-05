import { describe, it, expect } from "vitest";
import { createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { blockText, type CaretPosition } from "./textEditing";
import { caretSelection, type ModelSelection } from "./commands";
import { equalStructural } from "./equality";
import {
  clipboardFlavours,
  DOC_MODEL_MIME,
  executePaste,
  extractFragment,
  fragmentHtml,
  fragmentText,
  matchDestination,
  parseFragmentPayload,
  pasteBlocks,
  pastePlainText,
  serializeFragment,
} from "./clipboard";

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function doc(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("clipboard-test");
  model.sections = [{ blocks }];
  return model;
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function snapshot(model: DocumentModel): string {
  return JSON.stringify(model);
}

describe("extractFragment", () => {
  it("returns an empty fragment for a collapsed selection", () => {
    const model = doc(paragraph(["hello"]));
    expect(extractFragment(model, caretSelection(caret(0, 2)))).toEqual([]);
  });

  it("splits boundary runs at the selection edges", () => {
    const model = doc(paragraph(["hello ", { bold: true }], ["world"]));
    const fragment = extractFragment(model, range(caret(0, 3), caret(0, 8)));
    expect(fragment).toEqual([
      {
        type: "paragraph",
        runs: [
          { text: "lo ", formatting: { bold: true } },
          { text: "wo", formatting: {} },
        ],
      },
    ]);
  });

  it("captures whole paragraphs with their formatting, style and heading", () => {
    const model = doc(
      paragraph(["Title"]),
      paragraph(["Body", { italic: true }]),
    );
    model.sections[0].blocks[0].headingLevel = 1;
    model.sections[0].blocks[0].styleId = "Heading1";
    model.sections[0].blocks[1].formatting = { alignment: "center" };
    const fragment = extractFragment(model, range(caret(0, 0), caret(1, 4)));
    expect(fragment).toHaveLength(2);
    expect(fragment[0].headingLevel).toBe(1);
    expect(fragment[0].styleId).toBe("Heading1");
    expect(fragment[1].formatting).toEqual({ alignment: "center" });
  });

  it("excludes preservation side-lists and does not alias the model's runs", () => {
    const model = doc(paragraph(["hello"]));
    model.sections[0].blocks[0].opaqueInlines = [{ xml: "<x/>", runIndex: 0 }];
    const fragment = extractFragment(model, range(caret(0, 0), caret(0, 5)));
    expect(fragment[0].opaqueInlines).toBeUndefined();
    expect(fragment[0].runs[0]).not.toBe(model.sections[0].blocks[0].runs[0]);
  });
});

describe("clipboard flavours", () => {
  it("linearises paragraphs to newline-joined plain text (AC-7)", () => {
    const fragment = [paragraph(["one"]), paragraph(["two"])];
    expect(fragmentText(fragment)).toBe("one\ntwo");
  });

  it("generates minimal HTML with headings, toggles and span styles (AC-8)", () => {
    const heading: Paragraph = { ...paragraph(["Title"]), headingLevel: 2 };
    const body = paragraph(
      ["bold", { bold: true }],
      [" red", { color: "#ff0000", fontFamily: "Arial", fontSizePt: 14 }],
    );
    body.formatting = { alignment: "center" };
    const html = fragmentHtml([heading, body]);
    expect(html).toBe(
      '<h2>Title</h2><p style="text-align:center"><b>bold</b>' +
        '<span style="color:#ff0000;font-family:Arial;font-size:14pt"> red</span></p>',
    );
  });

  it("escapes markup-significant characters and renders breaks", () => {
    const html = fragmentHtml([paragraph(['a<b>&"\nc'])]);
    expect(html).toBe("<p>a&lt;b&gt;&amp;\"<br>c</p>");
  });

  it("bundles all three flavours under the internal MIME type (AC-9)", () => {
    const fragment = [paragraph(["hi"])];
    const flavours = clipboardFlavours(fragment);
    expect(flavours.text).toBe("hi");
    expect(flavours.html).toBe("<p>hi</p>");
    const parsed = JSON.parse(flavours.payload);
    expect(parsed.mimeType).toBe(DOC_MODEL_MIME);
    expect(parsed.version).toBe(1);
    expect(parsed.blocks).toEqual(fragment);
  });
});

describe("parseFragmentPayload", () => {
  it("round-trips a serialised fragment", () => {
    const fragment = [
      { ...paragraph(["x", { bold: true, color: "#ff0000" }]), headingLevel: 3 },
    ];
    expect(parseFragmentPayload(serializeFragment(fragment))).toEqual(fragment);
  });

  it("rejects malformed JSON, wrong mime type and wrong version", () => {
    expect(parseFragmentPayload("not json")).toBeNull();
    expect(
      parseFragmentPayload(JSON.stringify({ mimeType: "text/plain", version: 1, blocks: [] })),
    ).toBeNull();
    expect(
      parseFragmentPayload(JSON.stringify({ mimeType: DOC_MODEL_MIME, version: 2, blocks: [] })),
    ).toBeNull();
    expect(
      parseFragmentPayload(JSON.stringify({ mimeType: DOC_MODEL_MIME, version: 1 })),
    ).toBeNull();
  });

  it("rebuilds blocks from whitelisted fields only", () => {
    const raw = JSON.stringify({
      mimeType: DOC_MODEL_MIME,
      version: 1,
      blocks: [
        {
          type: "paragraph",
          runs: [
            {
              text: "hi",
              formatting: {
                bold: true,
                onclick: "alert(1)",
                color: "javascript:alert(1)",
                fontSizePt: "12",
              },
            },
          ],
          formatting: { alignment: "center", position: "absolute" },
          headingLevel: 99,
          evil: { xml: "<script/>" },
        },
      ],
    });
    expect(parseFragmentPayload(raw)).toEqual([
      {
        type: "paragraph",
        runs: [{ text: "hi", formatting: { bold: true } }],
        formatting: { alignment: "center" },
      },
    ]);
  });

  it("rejects payloads whose blocks are not paragraphs", () => {
    const raw = JSON.stringify({
      mimeType: DOC_MODEL_MIME,
      version: 1,
      blocks: [{ type: "table", runs: [] }],
    });
    expect(parseFragmentPayload(raw)).toBeNull();
  });
});

describe("pasteBlocks", () => {
  it("splices a single block's runs into the destination paragraph", () => {
    const model = doc(paragraph(["heldo"]));
    const before = snapshot(model);
    const result = pasteBlocks(model, caretSelection(caret(0, 3)), [
      paragraph(["lo wor", { bold: true }]),
    ]);
    expect(blockText(result.model.sections[0].blocks[0])).toBe("hello wordo");
    expect(result.model.sections[0].blocks[0].runs).toEqual([
      { text: "hel", formatting: {} },
      { text: "lo wor", formatting: { bold: true } },
      { text: "do", formatting: {} },
    ]);
    expect(result.selection.focus).toEqual(caret(0, 9));
    expect(snapshot(model)).toBe(before);
  });

  it("keeps the destination's paragraph formatting for a single-block splice", () => {
    const model = doc(paragraph(["text"]));
    model.sections[0].blocks[0].formatting = { alignment: "right" };
    const pasted: Paragraph = { ...paragraph(["X"]), headingLevel: 1 };
    const result = pasteBlocks(model, caretSelection(caret(0, 2)), [pasted]);
    expect(result.model.sections[0].blocks[0].formatting).toEqual({ alignment: "right" });
    expect(result.model.sections[0].blocks[0].headingLevel).toBeUndefined();
  });

  it("replaces an empty destination paragraph wholesale, keeping the pasted formatting", () => {
    const model = doc(paragraph([""]));
    const pasted: Paragraph = { ...paragraph(["Heading"]), headingLevel: 2 };
    const result = pasteBlocks(model, caretSelection(caret(0, 0)), [pasted]);
    expect(result.model.sections[0].blocks).toHaveLength(1);
    expect(result.model.sections[0].blocks[0].headingLevel).toBe(2);
    expect(result.selection.focus).toEqual(caret(0, 7));
  });

  it("pasting whole paragraphs into an empty paragraph restores them exactly (AC-10..13)", () => {
    const source = doc(
      { ...paragraph(["Title"]), headingLevel: 1, styleId: "Heading1" },
      paragraph(["plain ", {}], ["bold", { bold: true }]),
      { ...paragraph(["centered"]), formatting: { alignment: "center" } } as Paragraph,
    );
    const fragment = extractFragment(source, range(caret(0, 0), caret(2, 8)));
    const payload = parseFragmentPayload(serializeFragment(fragment));
    expect(payload).not.toBeNull();
    const target = doc(paragraph([""]));
    const result = pasteBlocks(target, caretSelection(caret(0, 0)), payload ?? []);
    expect(equalStructural(result.model.sections[0].blocks, source.sections[0].blocks)).toBe(
      true,
    );
    expect(result.selection.focus).toEqual(caret(2, 8));
  });

  it("splits the destination paragraph around a multi-block paste (AC-35)", () => {
    const model = doc(paragraph(["helloworld"]));
    model.sections[0].blocks[0].formatting = { alignment: "justify" };
    const result = pasteBlocks(model, caretSelection(caret(0, 5)), [
      paragraph(["A"]),
      { ...paragraph(["B"]), headingLevel: 2 },
      paragraph(["C"]),
    ]);
    const blocks = result.model.sections[0].blocks;
    expect(blocks.map(blockText)).toEqual(["helloA", "B", "Cworld"]);
    // The split halves keep the destination's paragraph formatting; the
    // middle block lands verbatim.
    expect(blocks[0].formatting).toEqual({ alignment: "justify" });
    expect(blocks[1].headingLevel).toBe(2);
    expect(blocks[2].formatting).toEqual({ alignment: "justify" });
    expect(result.selection.focus).toEqual(caret(2, 1));
  });

  it("pasting at a paragraph start keeps the first block intact", () => {
    const model = doc(paragraph(["world"]));
    const result = pasteBlocks(model, caretSelection(caret(0, 0)), [
      { ...paragraph(["Head"]), headingLevel: 3 },
      paragraph(["tail "]),
    ]);
    const blocks = result.model.sections[0].blocks;
    expect(blocks.map(blockText)).toEqual(["Head", "tail world"]);
    expect(blocks[0].headingLevel).toBe(3);
    expect(blocks[1].headingLevel).toBeUndefined();
  });

  it("pasting at a paragraph end appends the later blocks verbatim", () => {
    const model = doc(paragraph(["hello"]));
    const result = pasteBlocks(model, caretSelection(caret(0, 5)), [
      paragraph([" there"]),
      { ...paragraph(["Next"]), headingLevel: 2 },
    ]);
    const blocks = result.model.sections[0].blocks;
    expect(blocks.map(blockText)).toEqual(["hello there", "Next"]);
    expect(blocks[1].headingLevel).toBe(2);
    expect(result.selection.focus).toEqual(caret(1, 4));
  });

  it("replaces a non-collapsed selection (AC-34)", () => {
    const model = doc(paragraph(["one two three"]));
    const result = pasteBlocks(model, range(caret(0, 4), caret(0, 7)), [
      paragraph(["TWO", { bold: true }]),
    ]);
    expect(blockText(result.model.sections[0].blocks[0])).toBe("one TWO three");
    expect(result.selection.focus).toEqual(caret(0, 7));
  });

  it("repeated pastes of one fragment never share objects", () => {
    const fragment = [paragraph(["x"]), paragraph(["y"])];
    const model = doc(paragraph([""]));
    const first = pasteBlocks(model, caretSelection(caret(0, 0)), fragment);
    const second = pasteBlocks(first.model, first.selection, fragment);
    const blocks = second.model.sections[0].blocks;
    expect(blocks[0]).not.toBe(blocks[1]);
    expect(blocks.map(blockText)).toEqual(["x", "yx", "y"]);
  });
});

describe("pastePlainText", () => {
  it("inserts text that inherits the caret run's formatting (AC-28)", () => {
    const model = doc(paragraph(["bold", { bold: true }]));
    const result = pastePlainText(model, caretSelection(caret(0, 4)), "er");
    expect(result.model.sections[0].blocks[0].runs).toEqual([
      { text: "bolder", formatting: { bold: true } },
    ]);
  });

  it("turns newlines into paragraph breaks that keep the caret paragraph's formatting", () => {
    const model = doc(paragraph(["start"]));
    model.sections[0].blocks[0].formatting = { alignment: "center" };
    const result = pastePlainText(model, caretSelection(caret(0, 5)), "one\r\ntwo\rthree");
    const blocks = result.model.sections[0].blocks;
    expect(blocks.map(blockText)).toEqual(["startone", "two", "three"]);
    expect(blocks[1].formatting).toEqual({ alignment: "center" });
    expect(blocks[2].formatting).toEqual({ alignment: "center" });
    expect(result.selection.focus).toEqual(caret(2, 5));
  });

  it("replaces a non-collapsed selection", () => {
    const model = doc(paragraph(["one two three"]));
    const result = pastePlainText(model, range(caret(0, 4), caret(0, 7)), "TWO");
    expect(blockText(result.model.sections[0].blocks[0])).toBe("one TWO three");
  });

  it("honours the caller's pending format the way typing would", () => {
    const model = doc(paragraph(["x"]));
    const result = pastePlainText(model, caretSelection(caret(0, 1)), "y", { italic: true });
    expect(result.model.sections[0].blocks[0].runs).toEqual([
      { text: "x", formatting: {} },
      { text: "y", formatting: { italic: true } },
    ]);
  });
});

describe("matchDestination", () => {
  it("re-resolves block and run formatting to the destination", () => {
    const fragment = [
      { ...paragraph(["Title", { bold: true, color: "#ff0000" }]), headingLevel: 1 },
      { ...paragraph(["body", { italic: true }]), formatting: { alignment: "right" } } as Paragraph,
    ];
    const destination: Paragraph = {
      ...paragraph(["dest"]),
      formatting: { alignment: "center" },
      styleId: "Quote",
    };
    const matched = matchDestination(fragment, destination, { underline: true });
    expect(matched).toEqual([
      {
        type: "paragraph",
        runs: [{ text: "Title", formatting: { underline: true } }],
        formatting: { alignment: "center" },
        styleId: "Quote",
      },
      {
        type: "paragraph",
        runs: [{ text: "body", formatting: { underline: true } }],
        formatting: { alignment: "center" },
        styleId: "Quote",
      },
    ]);
  });
});

describe("executePaste", () => {
  const fragment = [
    { ...paragraph(["Big", { bold: true }]), headingLevel: 1 },
    paragraph(["small"]),
  ];

  it("mode source pastes the fragment verbatim", () => {
    const model = doc(paragraph([""]));
    const result = executePaste(model, caretSelection(caret(0, 0)), { fragment, text: "" }, "source");
    expect(result.model.sections[0].blocks[0].headingLevel).toBe(1);
    expect(result.model.sections[0].blocks[0].runs[0].formatting).toEqual({ bold: true });
  });

  it("mode match takes the destination's paragraph and typed formatting", () => {
    const model = doc(paragraph(["ctx", { italic: true }]));
    const result = executePaste(
      model,
      caretSelection(caret(0, 3)),
      { fragment, text: "" },
      "match",
    );
    const blocks = result.model.sections[0].blocks;
    expect(blocks[0].headingLevel).toBeUndefined();
    // Every pasted run now carries what a freshly typed character would.
    expect(blocks[0].runs).toEqual([{ text: "ctxBig", formatting: { italic: true } }]);
    expect(blocks[1].runs).toEqual([{ text: "small", formatting: { italic: true } }]);
  });

  it("mode text feeds the fragment's text through the plain-text path", () => {
    const model = doc(paragraph([""]));
    const result = executePaste(model, caretSelection(caret(0, 0)), { fragment, text: "" }, "text");
    const blocks = result.model.sections[0].blocks;
    expect(blocks.map(blockText)).toEqual(["Big", "small"]);
    expect(blocks[0].headingLevel).toBeUndefined();
    expect(blocks[0].runs[0].formatting).toEqual({});
  });

  it("a payload without a fragment always pastes as plain text", () => {
    const model = doc(paragraph([""]));
    const result = executePaste(
      model,
      caretSelection(caret(0, 0)),
      { fragment: null, text: "a\nb" },
      "source",
    );
    expect(result.model.sections[0].blocks.map(blockText)).toEqual(["a", "b"]);
  });

  it("never mutates the input model", () => {
    const model = doc(paragraph(["stable"]));
    const before = snapshot(model);
    executePaste(model, caretSelection(caret(0, 3)), { fragment, text: "t" }, "source");
    executePaste(model, caretSelection(caret(0, 3)), { fragment, text: "t" }, "match");
    executePaste(model, caretSelection(caret(0, 3)), { fragment, text: "t" }, "text");
    expect(snapshot(model)).toBe(before);
  });
});
