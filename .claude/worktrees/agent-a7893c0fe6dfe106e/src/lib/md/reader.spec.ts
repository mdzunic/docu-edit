import { describe, expect, it, vi } from "vitest";
import { parseMd } from "./reader";
import { isImage, isTable, type Paragraph, type Table } from "../documentModel";

function bytes(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

function paragraphs(blocks: unknown[]): Paragraph[] {
  return blocks as Paragraph[];
}

describe("SPEC-050 — parseMd", () => {
  it("parses an ATX heading", async () => {
    const parsed = await parseMd(bytes("# Hello world\n"));
    const [p] = paragraphs(parsed.sections[0].blocks);
    expect(p.headingLevel).toBe(1);
    expect(p.styleId).toBe("Heading1");
    // Effective (baked) formatting, matching the built-in Heading 1 style
    // (bold/size) plus the document-default font family underlay — every
    // reader in this codebase bakes resolved formatting onto runs alongside
    // the style reference (see reader.ts's bakeRun).
    expect(p.runs).toEqual([
      { text: "Hello world", formatting: { bold: true, fontSizePt: 20, fontFamily: "Calibri" } },
    ]);
  });

  it("parses bold/italic/strikethrough/code inline formatting", async () => {
    const parsed = await parseMd(bytes("**bold** *italic* ~~gone~~ `code`\n"));
    const [p] = paragraphs(parsed.sections[0].blocks);
    expect(p.runs.map((r) => r.text)).toEqual(["bold", " ", "italic", " ", "gone", " ", "code"]);
    expect(p.runs[0].formatting.bold).toBe(true);
    expect(p.runs[2].formatting.italic).toBe(true);
    expect(p.runs[4].formatting.strikethrough).toBe(true);
    expect(p.runs[6].styleId).toBe("CodeChar");
    // Baked formatting, not just the style reference (equality.ts's
    // resolve() and every other reader in this codebase relies on the
    // effective formatting living on the run itself — see reader.ts's
    // bakeRun/codeCharFormatting).
    expect(p.runs[6].formatting.fontFamily).toBe("Courier New");
  });

  it("bakes CodeChar's monospace formatting for an inline code span even when the document has no thematic break or languaged fenced block", async () => {
    // Regression: CodeChar used to only end up in the document's style
    // catalogue as a side effect of a thematic break or a languaged fence
    // seeding it (reader.ts's ensureMdStyles) — a plain paragraph with
    // nothing but an inline code span rendered it in the default
    // proportional font instead of monospace.
    const parsed = await parseMd(bytes("Run `npm test` to test.\n"));
    const [p] = paragraphs(parsed.sections[0].blocks);
    const codeRun = p.runs.find((r) => r.styleId === "CodeChar");
    expect(codeRun?.text).toBe("npm test");
    expect(codeRun?.formatting.fontFamily).toBe("Courier New");
    expect(parsed.styles.some((s) => s.id === "CodeChar")).toBe(true);
  });

  it("parses a hyperlink", async () => {
    const parsed = await parseMd(bytes("see [a link](https://example.com/x)\n"));
    const [p] = paragraphs(parsed.sections[0].blocks);
    const linkRun = p.runs.find((r) => r.formatting.linkHref);
    expect(linkRun?.text).toBe("a link");
    expect(linkRun?.formatting.linkHref).toBe("https://example.com/x");
  });

  // A link target reaches the `:href` a rendered run binds (DocPageBlock.vue),
  // so the Markdown reader launders it exactly as the .docx and .odt readers
  // do — a dangerous scheme leaves the text in place but unlinked.
  it.each([
    ["javascript:alert(1)"],
    ["JaVaScRiPt:alert(1)"],
    ["java\tscript:alert(1)"],
    ["  javascript:alert(1)"],
    ["data:text/html,<script>alert(1)</script>"],
    ["vbscript:msgbox(1)"],
  ])("strips a dangerous link scheme, keeping the text: %s", async (href) => {
    const parsed = await parseMd(bytes(`see [a link](${href})\n`));
    const [p] = paragraphs(parsed.sections[0].blocks);
    expect(p.runs.map((r) => r.text).join("")).toBe("see a link");
    expect(p.runs.every((r) => r.formatting.linkHref === undefined)).toBe(true);
  });

  it("keeps a safe link's other formatting when the scheme is stripped", async () => {
    const parsed = await parseMd(bytes("**[bold link](javascript:alert(1))**\n"));
    const [p] = paragraphs(parsed.sections[0].blocks);
    const run = p.runs.find((r) => r.text === "bold link");
    expect(run?.formatting.bold).toBe(true);
    expect(run?.formatting.linkHref).toBeUndefined();
  });

  it("parses a two-space hard line break as embedded \\n", async () => {
    const parsed = await parseMd(bytes("line one  \nline two\n"));
    const [p] = paragraphs(parsed.sections[0].blocks);
    expect(p.runs.map((r) => r.text).join("")).toBe("line one\nline two");
  });

  it("parses a fenced code block with an info string into Code:<lang> paragraphs", async () => {
    const parsed = await parseMd(bytes("```js\nconst x = 1;\nconsole.log(x);\n```\n"));
    const ps = paragraphs(parsed.sections[0].blocks);
    expect(ps).toHaveLength(2);
    expect(ps[0].styleId).toBe("Code:js");
    expect(ps[0].runs[0].text).toBe("const x = 1;");
    const style = parsed.styles.find((s) => s.id === "Code:js");
    expect(style?.language).toBe("js");
    expect(style?.basedOn).toBe("Code");
    // seeding pulls in the full built-in set too, so Normal/Heading1/etc still resolve
    expect(parsed.styles.some((s) => s.id === "Normal")).toBe(true);
  });

  it("parses a blockquote", async () => {
    const parsed = await parseMd(bytes("> quoted text\n> more\n"));
    const [p] = paragraphs(parsed.sections[0].blocks);
    expect(p.styleId).toBe("Quote");
    expect(p.runs.map((r) => r.text).join("")).toBe("quoted text more");
  });

  it("parses a thematic break", async () => {
    const parsed = await parseMd(bytes("above\n\n---\n\nbelow\n"));
    const ps = paragraphs(parsed.sections[0].blocks);
    expect(ps[1].styleId).toBe("HorizontalRule");
    expect(ps[1].runs).toEqual([]);
  });

  it("parses nested bulleted/numbered lists", async () => {
    const parsed = await parseMd(bytes("- one\n  - nested\n1. first\n2. second\n"));
    const ps = paragraphs(parsed.sections[0].blocks);
    expect(ps[0].list).toEqual({ type: "bullet", level: 1 });
    expect(ps[1].list).toEqual({ type: "bullet", level: 2 });
    expect(ps[2].list).toEqual({ type: "number", level: 1 });
    // Only the first item of a contiguous run at a level is where
    // lib/lists.ts's markersForBlocks ever reads `startAt` (numbering is
    // adjacency-derived, not identity-derived) — a later item's own literal
    // number is not carried, matching what a .docx/.odt round trip of an
    // ad-hoc (unnumbered-definition) list preserves.
    expect(ps[3].list).toEqual({ type: "number", level: 1 });
  });

  it("carries startAt on a numbered list that does not begin at 1", async () => {
    const parsed = await parseMd(bytes("5. fifth\n6. sixth\n"));
    const ps = paragraphs(parsed.sections[0].blocks);
    expect(ps[0].list).toEqual({ type: "number", level: 1, startAt: 5 });
    expect(ps[1].list).toEqual({ type: "number", level: 1 });
  });

  it("parses a GFM pipe table", async () => {
    const parsed = await parseMd(bytes("| A | B |\n| --- | --- |\n| 1 | 2 |\n"));
    const [table] = parsed.sections[0].blocks;
    expect(isTable(table)).toBe(true);
    const t = table as Table;
    expect(t.rows).toHaveLength(2);
    expect(t.rows[0].isHeader).toBe(true);
    expect(t.rows[0].cells.map((c) => c.blocks[0].runs.map((r) => r.text).join(""))).toEqual(["A", "B"]);
    expect(t.rows[1].cells.map((c) => c.blocks[0].runs.map((r) => r.text).join(""))).toEqual(["1", "2"]);
  });

  it("parses a standalone data-URI image", async () => {
    const png1x1 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
    const parsed = await parseMd(bytes(`![alt text](data:image/png;base64,${png1x1})\n`));
    const [block] = parsed.sections[0].blocks;
    expect(isImage(block)).toBe(true);
    expect(parsed.media).toHaveLength(1);
    expect(parsed.media[0].contentType).toBe("image/png");
  });

  // The image path fetches on the reader's behalf off a URL the document
  // supplied, so it takes the same scheme allow-list /open?url= does.
  it.each([
    ["file:///etc/passwd"],
    ["blob:https://example.com/2b3c4d5e-0000-4000-8000-000000000000"],
    ["ftp://example.com/logo.png"],
    ["javascript:alert(1)"],
  ])("never fetches an image src with a non-http(s) scheme: %s", async (src) => {
    const fetchSpy = vi.fn();
    vi.stubGlobal("fetch", fetchSpy);
    try {
      const parsed = await parseMd(bytes(`![alt](${src})\n`), { baseUrl: "https://example.com/doc.md" });
      expect(fetchSpy).not.toHaveBeenCalled();
      expect(parsed.media).toHaveLength(0);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("preserves an unfetchable image reference rather than dropping it (SPEC-008)", async () => {
    vi.stubGlobal("fetch", vi.fn());
    try {
      const parsed = await parseMd(bytes("![alt](file:///etc/passwd)\n"), {
        baseUrl: "https://example.com/doc.md",
      });
      expect(parsed.sections[0].opaqueBlocks).toHaveLength(1);
      expect(parsed.sections[0].opaqueBlocks?.[0].xml).toContain("file:///etc/passwd");
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("still fetches a relative image src resolved against an http(s) base", async () => {
    const png = Uint8Array.from(atob(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
    ), (c) => c.charCodeAt(0));
    const fetchSpy = vi.fn().mockResolvedValue({
      ok: true,
      headers: new Headers({ "content-type": "image/png" }),
      arrayBuffer: () => Promise.resolve(png.buffer),
    });
    vi.stubGlobal("fetch", fetchSpy);
    try {
      const parsed = await parseMd(bytes("![alt](images/photo.png)\n"), {
        baseUrl: "https://example.com/docs/doc.md",
      });
      expect(fetchSpy).toHaveBeenCalledWith("https://example.com/docs/images/photo.png");
      expect(parsed.media).toHaveLength(1);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("preserves front matter verbatim and off the body", async () => {
    const parsed = await parseMd(bytes("---\ntitle: Hi\n---\n\n# Heading\n"));
    expect(parsed.frontMatter).toBe("---\ntitle: Hi\n---");
    const [p] = paragraphs(parsed.sections[0].blocks);
    expect(p.headingLevel).toBe(1);
  });

  it("preserves a raw HTML block as an opaque block", async () => {
    const parsed = await parseMd(bytes("before\n\n<div>\n  <p>raw</p>\n</div>\n\nafter\n"));
    expect(parsed.sections[0].opaqueBlocks).toHaveLength(1);
    expect(parsed.sections[0].opaqueBlocks?.[0].xml).toContain("<div>");
  });
});
