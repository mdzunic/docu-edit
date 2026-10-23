// SPEC-106 — slide Markdown: the grammar row by row (separators, titles,
// subtitle vs body, bullets and their nesting, data-URI images, notes, front
// matter), SPEC-050's inline formatting through md/reader.ts's additive
// `parseInlineRuns`, every drop case (reported, never guessed), the writer's
// canonical spelling and its report of what Markdown cannot carry, and the
// `.md` → deck → `.md` round trip under `equalStructural` on the model
// (never `equalBytes` on the text — SPEC-050's rule).
import { describe, expect, it } from "vitest";
import { firstStructuralDifference, equalStructural } from "../equality";
import { parseInlineRuns } from "../md/reader";
import { groupFrames, insertShapeFrame, insertTableFrame } from "../presentation/frames";
import type { ImageFrame, PresentationModel, SlideFrame, TextFrame } from "../presentation/model";
import { titleBodyFrames, titleSlideFrames } from "../presentation/slides";
import type { DeckDropKind } from "./deckDrops";
import { DocumentGenerationError } from "./errors";
import { PresentationBuilder } from "./presentation";
import { parseSlideMarkdown, parseSlideMarkdownWithReport, writeSlideMarkdown } from "./slideMarkdown";

const PNG_4X2 = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAAAAABawyK/AAAAC0lEQVR4nGNggAEAAAoAAX+AdF4AAAAASUVORK5CYII=";
const PNG_URI = `data:image/png;base64,${PNG_4X2}`;

/** The spec's own example deck. */
const EXAMPLE = `---
title: Q3 review
---
# Q3 review
Prepared 30 September

---

## Highlights
- Revenue up 12%
  - Europe up 20%

Notes: Lead with Europe.
`;

function texts(frame: SlideFrame | undefined): string[] {
  return frame && frame.kind === "text" ? frame.paragraphs.map((p) => p.runs.map((run) => run.text).join("")) : [];
}

function byRole(model: PresentationModel, slide: number, role: string): TextFrame | undefined {
  return model.slides[slide].frames.find((frame): frame is TextFrame => frame.kind === "text" && frame.role === role);
}

function geometry(frame: { xCm: number; yCm: number; widthCm: number; heightCm: number }): number[] {
  return [frame.xCm, frame.yCm, frame.widthCm, frame.heightCm];
}

function kinds(dropped: { kind: DeckDropKind }[]): DeckDropKind[] {
  return dropped.map((record) => record.kind);
}

describe("parseInlineRuns — md/reader.ts's additive export (AC-15)", () => {
  it("is SPEC-050's inline grammar over one line", () => {
    const runs = parseInlineRuns("plain **bold** _it_ *also it* ~~gone~~ [link](https://example.com) \\*lit\\*");
    expect(runs.map((run) => [run.text, run.formatting])).toEqual([
      ["plain ", {}],
      ["bold", { bold: true }],
      [" ", {}],
      ["it", { italic: true }],
      [" ", {}],
      ["also it", { italic: true }],
      [" ", {}],
      ["gone", { strikethrough: true }],
      [" ", {}],
      ["link", { linkHref: "https://example.com" }],
      [" *lit*", {}],
    ]);
  });

  it("resolves a code span against the CodeChar style, as parseMd does", () => {
    const [code] = parseInlineRuns("`x`");
    expect(code.text).toBe("x");
    expect(code.styleId).toBe("CodeChar");
    expect(code.formatting.fontFamily).toBeTruthy();
  });
});

describe("the slide grammar (AC-14)", () => {
  it("parses the spec's example: a title slide, then a bullet slide with notes", () => {
    const { deck, dropped } = parseSlideMarkdownWithReport(EXAMPLE);
    expect(dropped).toEqual([]);
    expect(deck.title).toBe("Q3 review");
    expect(deck.slides).toHaveLength(2);
    const [title, subtitle] = deck.slides[0].frames as TextFrame[];
    const [expectedTitle, expectedSubtitle] = titleSlideFrames(33.867, 19.05);
    expect(geometry(title)).toEqual(geometry(expectedTitle));
    expect(geometry(subtitle)).toEqual(geometry(expectedSubtitle));
    expect(texts(title)).toEqual(["Q3 review"]);
    expect(texts(subtitle)).toEqual(["Prepared 30 September"]);
    const [slideTitle, body] = deck.slides[1].frames as TextFrame[];
    const [bodyTitle, bodyRect] = titleBodyFrames(33.867, 19.05);
    expect(geometry(slideTitle)).toEqual(geometry(bodyTitle));
    expect(geometry(body)).toEqual(geometry(bodyRect));
    expect(body.paragraphs).toEqual([
      { runs: [{ text: "Revenue up 12%" }], bullet: "disc" },
      { runs: [{ text: "Europe up 20%" }], bullet: "disc", level: 1 },
    ]);
    expect(deck.slides[1].notes).toEqual(["Lead with Europe."]);
    expect(deck.opaque).toBeUndefined();
  });

  it("splits slides at --- lines between blocks; leading, trailing and front-matter --- mint no slide", () => {
    expect(parseSlideMarkdown("# A\n---\n# B\n---\n# C\n").slides).toHaveLength(3);
    expect(parseSlideMarkdown("---\n\n# A\n").slides).toHaveLength(1);
    expect(parseSlideMarkdown("# A\n\n---\n\n").slides).toHaveLength(1);
    expect(parseSlideMarkdown("---\ntitle: T\n---\n# A").slides).toHaveLength(1);
    expect(parseSlideMarkdown("").slides).toEqual([]);
    // An empty chunk between two separators is a blank slide.
    const blank = parseSlideMarkdown("# A\n\n---\n\n---\n\n# C");
    expect(blank.slides).toHaveLength(3);
    expect(blank.slides[1].frames).toEqual([]);
  });

  it("does not split inside a fenced block", () => {
    const { deck, dropped } = parseSlideMarkdownWithReport("# A\n\n```yaml\n---\nkey: value\n```\n\n- after");
    expect(deck.slides).toHaveLength(1);
    expect(kinds(dropped)).toEqual(["codeFence"]);
    expect(texts(byRole(deck, 0, "body"))).toEqual(["after"]);
  });

  it("# and ## both fill the title; a slide that opens with anything else has no title frame", () => {
    const deck = parseSlideMarkdown("# One\n---\n# Two\n---\nJust text\n\n- and a row");
    expect(texts(byRole(deck, 0, "title"))).toEqual(["One"]);
    expect(texts(byRole(deck, 1, "title"))).toEqual(["Two"]);
    expect(byRole(deck, 2, "title")).toBeUndefined();
    expect(deck.slides[2].frames).toHaveLength(1);
    expect(texts(byRole(deck, 2, "body"))).toEqual(["Just text", "and a row"]);
  });

  it("plain paragraphs before any list are the first slide's subtitle, body paragraphs after it", () => {
    const deck = parseSlideMarkdown("# Deck\nLine one\n\nLine two\n- row\n\nAfter the list\n---\n## Next\nPlain body\n\n- row");
    expect(texts(byRole(deck, 0, "subtitle"))).toEqual(["Line one", "Line two"]);
    expect(texts(byRole(deck, 0, "body"))).toEqual(["row", "After the list"]);
    expect(byRole(deck, 0, "body")!.paragraphs[1].bullet).toBeUndefined();
    expect(byRole(deck, 1, "subtitle")).toBeUndefined();
    expect(texts(byRole(deck, 1, "body"))).toEqual(["Plain body", "row"]);
  });

  it("a first slide that opens with a list after its title is a title + body slide", () => {
    const deck = parseSlideMarkdown("# Agenda\n- one\n- two");
    expect(deck.slides[0].frames.map((frame) => (frame as TextFrame).role)).toEqual(["title", "body"]);
    expect(geometry(deck.slides[0].frames[0])).toEqual(geometry(titleBodyFrames(33.867, 19.05)[0]));
  });

  it("joins a paragraph's soft-wrapped lines and splits it at hard breaks (one slide paragraph per rendered line)", () => {
    const deck = parseSlideMarkdown("# Deck\n---\n## T\nsoft\nwrapped  \nhard\\\nbreaks");
    expect(texts(byRole(deck, 1, "body"))).toEqual(["soft wrapped", "hard", "breaks"]);
  });

  it("-, * and + rows are disc bullets, 1. rows numbered, nested by two-space indent and clamped at level 8", () => {
    const deck = parseSlideMarkdown(
      ["## Rows", "- dash", "* star", "+ plus", "1. one", "2) two", "  - two spaces", "   - three spaces", "    - four", " ".repeat(30) + "- deep"].join("\n"),
    );
    expect(byRole(deck, 0, "body")!.paragraphs.map((p) => [p.runs[0].text, p.bullet, p.level ?? 0])).toEqual([
      ["dash", "disc", 0],
      ["star", "disc", 0],
      ["plus", "disc", 0],
      ["one", "number", 0],
      ["two", "number", 0],
      ["two spaces", "disc", 1],
      ["three spaces", "disc", 1],
      ["four", "disc", 2],
      ["deep", "disc", 8],
    ]);
  });

  it("maps SPEC-050's inline formatting onto slide runs", () => {
    const deck = parseSlideMarkdown("## T\n- **bold** and _italic_ and [a link](https://example.com)");
    expect(byRole(deck, 0, "body")!.paragraphs[0].runs).toEqual([
      { text: "bold", bold: true },
      { text: " and " },
      { text: "italic", italic: true },
      { text: " and " },
      { text: "a link", linkHref: "https://example.com" },
    ]);
  });

  it("drops what a slide run has no field for, counted per slide", () => {
    const { deck, dropped } = parseSlideMarkdownWithReport("## T\n- ~~struck~~ plain ~~again~~\n- `code`");
    expect(texts(byRole(deck, 0, "body"))).toEqual(["struck plain again", "code"]);
    // Runs a dropped mark no longer tells apart are one run.
    expect(byRole(deck, 0, "body")!.paragraphs[0].runs).toEqual([{ text: "struck plain again" }]);
    expect(dropped).toEqual([
      { kind: "runFormatting", location: "slide 1", reason: "strikethrough dropped from 2 runs: a slide run has no strikethrough" },
      { kind: "runFormatting", location: "slide 1", reason: "code-span formatting dropped from 1 run: a slide run has no code-span formatting" },
    ]);
  });

  it("a standalone data-URI image is an image frame fitted into the body area", () => {
    const { deck, dropped } = parseSlideMarkdownWithReport(`# Deck\n---\n## Pipeline\n\n![](${PNG_URI})`);
    expect(dropped).toEqual([]);
    const [title, image] = deck.slides[1].frames;
    expect((title as TextFrame).role).toBe("title");
    expect(image.kind).toBe("image");
    const body = titleBodyFrames(33.867, 19.05)[1];
    expect(image.heightCm).toBe(body.heightCm);
    expect(image.widthCm).toBe(22.86);
    expect((image as ImageFrame).mediaId).toBe("ppt/media/image1.png");
    expect(deck.media).toEqual([{ mediaId: "ppt/media/image1.png", contentType: "image/png", data: PNG_4X2 }]);
    // A percent-encoded (non-base64) data URI decodes to the same bytes.
    const percent = Array.from(atob(PNG_4X2), (char) => `%${char.charCodeAt(0).toString(16).padStart(2, "0")}`).join("");
    const decoded = parseSlideMarkdown(`![](data:image/png,${percent})`);
    expect(decoded.slides[0].frames.map((frame) => frame.kind)).toEqual(["image"]);
    expect(decoded.media![0].data).toBe(PNG_4X2);
  });

  it("never fetches: a remote image is dropped as remoteImage; an unsupported or broken one as media", () => {
    const { deck, dropped } = parseSlideMarkdownWithReport(
      "## T\n\n![chart](https://example.com/chart.png)\n\n![x](data:image/svg+xml;base64,PHN2Zy8+)\n\n![y](data:image/png;base64,AAAA)\n\n![alt text](" + PNG_URI + ")",
    );
    expect(deck.slides[0].frames.filter((frame) => frame.kind === "image")).toHaveLength(1);
    expect(dropped.map((record) => [record.kind, record.location])).toEqual([
      ["remoteImage", "slide 1"],
      ["media", "slide 1"],
      ["media", "slide 1"],
      ["media", "slide 1"],
    ]);
    expect(dropped[0].reason).toContain("nothing is fetched");
    expect(dropped[3].reason).toContain("alt text");
  });

  it("Notes: opens speaker notes to the end of the slide", () => {
    const deck = parseSlideMarkdown(
      "## A\n- row\n\nNotes: First line.\nSecond line, **verbatim**.\n\n- not a bullet\n\n\n---\n## B\nNotes:\n\nAfter a blank.\n---\n## C\nNotes:\n",
    );
    expect(deck.slides[0].notes).toEqual(["First line.", "Second line, **verbatim**.", "", "- not a bullet"]);
    expect(texts(byRole(deck, 0, "body"))).toEqual(["row"]);
    expect(deck.slides[1].notes).toEqual(["", "After a blank."]);
    expect("notes" in deck.slides[2]).toBe(false);
  });

  it("front matter sets title, format and size; options fill in what it does not", () => {
    const deck = parseSlideMarkdown("---\ntitle: \"Q3: review\"\nformat: odp\nsize: 4:3\n---\n# A");
    expect([deck.title, deck.format, deck.slideWidthCm, deck.slideHeightCm]).toEqual(["Q3: review", "odp", 25.4, 19.05]);
    const fallback = parseSlideMarkdown("# A", { title: "From options", format: "odp", size: "4:3" });
    expect([fallback.title, fallback.format, fallback.slideWidthCm]).toEqual(["From options", "odp", 25.4]);
    const mixed = parseSlideMarkdown("---\ntitle: 'It''s mine' # comment\n---\n# A", { title: "Fallback" });
    expect(mixed.title).toBe("It's mine");
    expect(parseSlideMarkdown("# A").format).toBe("pptx");
  });

  it("reports any other front-matter key once, and a value the deck cannot take", () => {
    const { deck, dropped } = parseSlideMarkdownWithReport(
      "---\nauthor: me\ntags:\n  - a\nauthor: again\nformat: docx\nsize: 16:10\ntitle:\n  nested: x\n---\n# A",
    );
    expect(deck.format).toBe("pptx");
    expect(deck.slideWidthCm).toBe(33.867);
    expect(deck.title).toBe("Untitled presentation");
    expect(dropped.map((record) => [record.kind, record.location])).toEqual(Array.from({ length: 5 }, () => ["frontMatterKey", "front matter"]));
    expect(dropped.map((record) => record.reason)).toEqual([
      '"author" is not a slide front-matter key (title, format, size): not kept',
      '"tags" is not a slide front-matter key (title, format, size): not kept',
      'format "docx" is not "pptx" or "odp": not kept',
      'size "16:10" is not "16:9" or "4:3": not kept',
      '"title" takes a single value, not a nested one: not kept',
    ]);
  });

  it("validates its arguments", () => {
    expect(() => parseSlideMarkdown(42 as never)).toThrow(DocumentGenerationError);
    expect(() => parseSlideMarkdown("# A", { theme: "dark" } as never)).toThrow("options.theme: unknown property");
    expect(() => parseSlideMarkdown("# A", { format: "key" } as never)).toThrow(/^options\.format: expected one of/);
  });
});

describe("anything outside the grammar is reported, never guessed (AC-16)", () => {
  it("drops block quotes, tables, fences, raw HTML, reference definitions and ### headings with a record each", () => {
    const { deck, dropped } = parseSlideMarkdownWithReport(
      [
        "## Kept title",
        "> a quote",
        "> continued",
        "",
        "| a | b |",
        "| --- | --- |",
        "| 1 | 2 |",
        "",
        "~~~",
        "code",
        "~~~",
        "",
        "<div>html</div>",
        "",
        "[ref]: https://example.com",
        "",
        "### Too deep",
        "",
        "## A second heading",
        "",
        "***",
        "",
        "- the one kept row",
      ].join("\n"),
    );
    expect(kinds(dropped)).toEqual(["blockQuote", "table", "codeFence", "rawHtml", "rawHtml", "heading", "heading", "rawHtml"]);
    expect(dropped.every((record) => record.location === "slide 1")).toBe(true);
    expect(texts(byRole(deck, 0, "title"))).toEqual(["Kept title"]);
    expect(texts(byRole(deck, 0, "body"))).toEqual(["the one kept row"]);
  });

  it("names each slide by its 1-based number", () => {
    const { dropped } = parseSlideMarkdownWithReport("# A\n---\n## B\n> q\n---\n## C\n### deep");
    expect(dropped.map((record) => `${record.kind}@${record.location}`)).toEqual(["blockQuote@slide 2", "heading@slide 3"]);
  });

  it("parseSlideMarkdown returns the same deck without the report", () => {
    const text = "## T\n> q\n- row";
    expect(equalStructural(parseSlideMarkdown(text), parseSlideMarkdownWithReport(text).deck)).toBe(true);
  });
});

describe("writeSlideMarkdown (AC-17)", () => {
  it("writes the canonical spelling: front matter, ATX titles, - bullets, two-space nesting, images, notes", () => {
    const deck = new PresentationBuilder({ title: "Q3 review" })
      .addTitleSlide("Q3 review", "Prepared 30 September")
      .addBulletSlide("Highlights", ["Revenue up 12%", ["Europe up 20%", { runs: [{ text: "APAC", bold: true }, " ", { text: "flat", italic: true }] }]])
      .setNotes(["Lead with Europe.", "Then APAC."])
      .addBulletSlide("Steps", ["First", "Second"], { style: "number" })
      .addImageSlide("Pipeline", { data: PNG_4X2, contentType: "image/png" })
      .toDeck();
    const { text, dropped } = writeSlideMarkdown(deck);
    expect(dropped).toEqual([]);
    expect(text).toBe(
      [
        "---",
        "title: Q3 review",
        "format: pptx",
        "size: 16:9",
        "---",
        "",
        "# Q3 review",
        "",
        "Prepared 30 September",
        "",
        "---",
        "",
        "## Highlights",
        "",
        "- Revenue up 12%",
        "  - Europe up 20%",
        "  - **APAC** _flat_",
        "",
        "Notes: Lead with Europe.",
        "Then APAC.",
        "",
        "---",
        "",
        "## Steps",
        "",
        "1. First",
        "2. Second",
        "",
        "---",
        "",
        "## Pipeline",
        "",
        `![](${PNG_URI})`,
        "",
      ].join("\n"),
    );
    // Written by the inverse, read back by the grammar: the same deck.
    expect(firstStructuralDifference(parseSlideMarkdown(text), deck)).toBeNull();
  });

  it("escapes text that would otherwise read as syntax", () => {
    const deck = new PresentationBuilder()
      .addBulletSlide("C# and *stars* #", [
        "# not a heading",
        "1. not numbered",
        "Notes: not notes",
        "---",
        "<b>not html</b>",
        "[not](a link)",
        "ends with a backslash \\",
      ])
      .setNotes(["---", "\\leading backslash"])
      .toDeck();
    const { text } = writeSlideMarkdown(deck);
    expect(text).toContain("\n# C# and \\*stars\\* \\#\n");
    expect(text).toContain("Notes: \\---\n\\\\leading backslash");
    expect(firstStructuralDifference(parseSlideMarkdown(text), deck)).toBeNull();
  });

  it("keeps a blank first or last slide a slide", () => {
    const deck = new PresentationBuilder().addBlankSlide().addTitleSlide("Middle").addBlankSlide().addBlankSlide().toDeck();
    const { text } = writeSlideMarkdown(deck);
    const back = parseSlideMarkdown(text);
    expect(back.slides.map((slide) => slide.frames.length)).toEqual([0, 2, 0, 0]);
  });

  it("reports every construct Markdown cannot carry, naming its slide", () => {
    const deck = new PresentationBuilder({ size: { width: { cm: 30 }, height: { cm: 20 } } })
      .addTitleSlide([{ text: "Coloured", color: "#ff0000", fontSizePt: 40 }])
      .addBlankSlide()
      .addTextBox([{ text: "Box", fontFamily: "Georgia", underline: true }], { x: { cm: 1 }, y: { cm: 1 }, width: { cm: 5 }, height: { cm: 2 } })
      .addImage({ data: PNG_4X2, contentType: "image/png", x: { cm: 1 }, y: { cm: 5 }, width: { cm: 4 } })
      .setBackground("#000000")
      .toDeck();
    insertShapeFrame(deck, 1, "rect", { xCm: 10, yCm: 1, widthCm: 2, heightCm: 2 });
    insertShapeFrame(deck, 1, "ellipse", { xCm: 13, yCm: 1, widthCm: 2, heightCm: 2 });
    insertShapeFrame(deck, 1, "rect", { xCm: 16, yCm: 1, widthCm: 2, heightCm: 2 });
    const shapes = deck.slides[1].frames.filter((frame) => frame.kind === "shape");
    groupFrames(deck, 1, [shapes[1].id, shapes[2].id]);
    insertTableFrame(deck, 1, 2, 2);
    deck.slides[1].frames.push(
      { id: "f", kind: "foreign", xml: "<x/>", label: "Object", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 },
      { id: "m", kind: "media", media: "video", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 },
      { id: "c", kind: "chart", chartType: "bar", categories: [], series: [], xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 } as never,
    );
    deck.slides[0].transition = { type: "fade" };
    deck.slides[0].animations = [
      { frameId: deck.slides[0].frames[0].id, effect: "fadeIn", trigger: "click" },
      { frameId: deck.slides[0].frames[0].id, effect: "pulse", trigger: "afterPrevious" },
    ];
    (deck.slides[0].frames[0] as TextFrame).fill = { type: "solid", color: "#eeeeee" };
    (deck.slides[0].frames[0] as TextFrame).stroke = { color: "#000000", widthPt: 1 };

    const { text, dropped } = writeSlideMarkdown(deck);
    const at = (location: string) => dropped.filter((record) => record.location === location).map((record) => record.kind);
    expect(at("deck")).toEqual(["slideGeometry"]);
    expect(at("slide 1").toSorted()).toEqual(
      ["animation", "animation", "fill", "runFormatting", "runFormatting", "stroke", "transition"].toSorted(),
    );
    expect(at("slide 2").toSorted()).toEqual(
      [
        "background",
        "chart",
        "fill",
        "foreignFrame",
        "group",
        "media",
        "runFormatting",
        "runFormatting",
        "shape",
        "slideGeometry",
        "slideGeometry",
        "stroke",
        "table",
      ].toSorted(),
    );
    const runDrops = dropped.filter((record) => record.kind === "runFormatting").map((record) => record.reason);
    expect(runDrops).toEqual([
      "size dropped from 1 run: Markdown has no spelling for it",
      "colour dropped from 1 run: Markdown has no spelling for it",
      "underline dropped from 1 run: Markdown has no spelling for it",
      "font dropped from 1 run: Markdown has no spelling for it",
    ]);
    // What Markdown can carry is still written: the text and the picture.
    expect(text).toContain("# Coloured");
    expect(text).toContain("\nBox\n");
    expect(text).toContain(`![](${PNG_URI})`);
    expect(text).not.toContain("size:");
  });
});

describe(".md → deck → .md round trip (AC-18)", () => {
  const SOURCE = [
    "---",
    "title: Kitchen sink",
    "format: odp",
    "size: 4:3",
    "unknown: dropped",
    "---",
    "",
    "# Kitchen *sink*",
    "Subtitle with __bold__ and a [link](https://example.com/a_(b))",
    "",
    "---",
    "",
    "## Rows",
    "* star",
    "    + deep plus",
    "3. three",
    "4. four",
    "",
    "Closing ~~struck~~ paragraph  ",
    "and a hard break",
    "",
    "> dropped quote",
    "",
    `![a picture](${PNG_URI})`,
    "",
    "Notes: speak",
    "slowly",
    "",
    "---",
    "No title here, C# 1. and \\- escapes",
    "",
    "---",
    "",
    "---",
    "## Last ##",
    "Notes:",
    "",
  ].join("\n");

  it("satisfies equalStructural on the model, and the written text is canonical (a fixed point)", () => {
    const first = parseSlideMarkdownWithReport(SOURCE);
    expect(first.dropped.length).toBeGreaterThan(0);
    const written = writeSlideMarkdown(first.deck);
    expect(written.dropped).toEqual([]);
    const second = parseSlideMarkdownWithReport(written.text);
    expect(second.dropped).toEqual([]);
    expect(firstStructuralDifference(first.deck, second.deck)).toBeNull();
    expect(equalStructural(first.deck, second.deck)).toBe(true);
    expect(writeSlideMarkdown(second.deck).text).toBe(written.text);
    // equalBytes on the text is explicitly NOT the relation: the writer
    // normalises spelling (SPEC-050's rule).
    expect(written.text).not.toBe(SOURCE);
    expect(first.deck.slides).toHaveLength(5);
    expect(first.deck.format).toBe("odp");
    expect(first.deck.slides[3].frames).toEqual([]);
  });
});
