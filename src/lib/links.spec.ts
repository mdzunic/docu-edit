import { describe, it, expect } from "vitest";
import { asParagraph, createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { caretSelection, type ModelSelection } from "./commands";
import { blockText, insertText, deleteRange, splitParagraph } from "./textEditing";
import {
  anchorCaret,
  applyLink,
  autolinkAt,
  bookmarkCaret,
  deleteBookmark,
  detectBareUrl,
  headingAnchor,
  insertBookmark,
  insertLinkText,
  isDangerousUrl,
  isFetchableUrl,
  isLocalMachineUrl,
  linkAtCaret,
  linkTargetLabel,
  linkTargetOf,
  listBookmarks,
  listHeadings,
  removeLink,
  sanitizeHref,
  urlScheme,
  validateBookmarkName,
  withLinkTarget,
} from "./links";

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function doc(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("links-test");
  model.sections = [{ blocks }];
  return model;
}

function selection(anchor: [number, number], focus: [number, number]): ModelSelection {
  return {
    anchor: { section: 0, block: anchor[0], offset: anchor[1] },
    focus: { section: 0, block: focus[0], offset: focus[1] },
  };
}

// --- URL safety (AC-14..18): the security test the spec requires ------------

describe("dangerous URL schemes (security, AC-14..18)", () => {
  it("rejects javascript: URLs", () => {
    expect(isDangerousUrl("javascript:alert(1)")).toBe(true);
    expect(sanitizeHref("javascript:alert(1)")).toBeNull();
  });

  it("rejects data: URLs", () => {
    expect(isDangerousUrl("data:text/html,<script>alert(1)</script>")).toBe(true);
    expect(sanitizeHref("data:text/html,x")).toBeNull();
  });

  it("rejects vbscript: URLs", () => {
    expect(isDangerousUrl("vbscript:msgbox(1)")).toBe(true);
    expect(sanitizeHref("vbscript:msgbox(1)")).toBeNull();
  });

  it("rejects case and whitespace disguises the way browsers parse them", () => {
    expect(isDangerousUrl("JaVaScRiPt:alert(1)")).toBe(true);
    expect(isDangerousUrl("  javascript:alert(1)")).toBe(true);
    expect(isDangerousUrl("java\tscript:alert(1)")).toBe(true);
    expect(isDangerousUrl("java\nscript:alert(1)")).toBe(true);
    expect(isDangerousUrl("\u0000javascript:alert(1)")).toBe(true);
  });

  it("allows http, https, mailto and relative targets", () => {
    expect(isDangerousUrl("https://example.com")).toBe(false);
    expect(isDangerousUrl("http://example.com/a?b=c")).toBe(false);
    expect(isDangerousUrl("mailto:a@example.com")).toBe(false);
    expect(isDangerousUrl("example.com/data:x")).toBe(false);
    expect(sanitizeHref(" https://example.com ")).toBe("https://example.com");
  });

  it("parses the scheme only when one is present", () => {
    expect(urlScheme("https://example.com")).toBe("https");
    expect(urlScheme("no-scheme-here")).toBeNull();
    expect(urlScheme("path/with:colon")).toBeNull();
  });

  it("never autolinks a dangerous URL (stripped on paste, AC-17)", () => {
    expect(detectBareUrl("javascript:alert(1)", 19)).toBeNull();
    expect(detectBareUrl("data:text/html,x", 16)).toBeNull();
    expect(detectBareUrl("vbscript:msgbox(1)", 18)).toBeNull();
    const model = doc(paragraph(["javascript:alert(1) "]));
    expect(autolinkAt(model, { section: 0, block: 0 }, 19)).toBeNull();
  });

  it("strips a dangerous target instead of storing it via withLinkTarget callers", () => {
    // Readers call sanitizeHref before withLinkTarget; a null target leaves
    // the formatting linkless.
    expect(withLinkTarget({ bold: true }, null)).toEqual({ bold: true });
  });
});

// --- autolink detection (AC-5, AC-6) -----------------------------------------

describe("isFetchableUrl (what the app will request on the reader's behalf)", () => {
  it("allows the two web schemes", () => {
    expect(isFetchableUrl("https://example.com/a.docx")).toBe(true);
    expect(isFetchableUrl("http://intranet.example/a.docx")).toBe(true);
  });

  it.each([
    ["data:text/html,<script>alert(1)</script>"],
    ["blob:https://example.com/2b3c4d5e-0000-4000-8000-000000000000"],
    ["javascript:alert(1)"],
    ["vbscript:msgbox(1)"],
    ["file:///etc/passwd"],
    ["ftp://example.com/a.docx"],
  ])("refuses %s", (url) => {
    expect(isFetchableUrl(url)).toBe(false);
  });

  it("refuses anything that is not an absolute URL", () => {
    // Callers resolving a relative reference must check the RESOLVED url —
    // see resolveImageSrc in md/reader.ts.
    expect(isFetchableUrl("images/photo.png")).toBe(false);
    expect(isFetchableUrl("/absolute/path.png")).toBe(false);
    expect(isFetchableUrl("")).toBe(false);
    expect(isFetchableUrl("not a url at all")).toBe(false);
  });

  it("is stricter than sanitizeHref: mailto may be linked but never fetched", () => {
    expect(sanitizeHref("mailto:a@example.com")).toBe("mailto:a@example.com");
    expect(isFetchableUrl("mailto:a@example.com")).toBe(false);
  });

  // SPEC-113's CSP admits plain-http loopback for a local model server, so
  // nothing but this rule keeps an untrusted document (or a link into
  // /open?url=) from reaching the user's own services.
  it.each([
    ["http://localhost:8080/api/whatever"],
    ["http://LOCALHOST:11434/v1/models"],
    ["http://localhost./a.png"],
    ["https://localhost/a.png"],
    ["http://admin.localhost:3001/a.png"],
    ["http://127.0.0.1:8080/a.png"],
    ["http://127.12.0.9/a.png"],
    ["http://2130706433/a.png"],
    ["http://0x7f.1/a.png"],
    ["http://0.0.0.0:8080/a.png"],
    ["http://[::1]:8080/a.png"],
    ["http://[0:0:0:0:0:0:0:1]/a.png"],
    ["http://[::]/a.png"],
    ["http://[::ffff:127.0.0.1]:8080/a.png"],
  ])("refuses this computer: %s", (url) => {
    expect(isLocalMachineUrl(url)).toBe(true);
    expect(isFetchableUrl(url, "https://docu-edit.example")).toBe(false);
  });

  it.each([
    ["http://localhost.example.com/a.png"],
    ["http://notlocalhost/a.png"],
    ["http://128.0.0.1/a.png"],
    ["http://[::2]/a.png"],
    ["https://example.com/127.0.0.1/a.png"],
  ])("does not mistake a look-alike for this computer: %s", (url) => {
    expect(isLocalMachineUrl(url)).toBe(false);
    expect(isFetchableUrl(url, "https://docu-edit.example")).toBe(true);
  });

  it("still fetches from the app's own origin when that is on this computer (a dev server)", () => {
    expect(isFetchableUrl("http://localhost:5173/fixtures/photo.png", "http://localhost:5173")).toBe(true);
    // Only that origin: another port on the same host is another service.
    expect(isFetchableUrl("http://localhost:8080/photo.png", "http://localhost:5173")).toBe(false);
    expect(isFetchableUrl("https://localhost:5173/photo.png", "http://localhost:5173")).toBe(false);
    // The page's own origin is the default.
    expect(isFetchableUrl(`${location.origin}/photo.png`)).toBe(true);
  });
});

describe("detectBareUrl", () => {
  it("matches an http(s) URL ending at the offset", () => {
    const text = "See http://example.com";
    expect(detectBareUrl(text, text.length)).toEqual({
      start: 4,
      end: 22,
      href: "http://example.com",
    });
  });

  it("prefixes https:// onto www. shorthand", () => {
    const text = "www.example.com";
    expect(detectBareUrl(text, text.length)).toEqual({
      start: 0,
      end: 15,
      href: "https://www.example.com",
    });
  });

  it("peels trailing punctuation and leading wrappers", () => {
    const text = "(https://example.com/a).";
    expect(detectBareUrl(text, text.length)).toEqual({
      start: 1,
      end: 22,
      href: "https://example.com/a",
    });
  });

  it("ignores plain words and partial schemes", () => {
    expect(detectBareUrl("hello world", 11)).toBeNull();
    expect(detectBareUrl("http://", 7)).toBeNull();
    expect(detectBareUrl("www.nodots", 10)).toBeNull();
  });
});

// --- link commands -----------------------------------------------------------

describe("applyLink / linkAtCaret / removeLink", () => {
  it("wraps the selected text in a hyperlink (AC-3)", () => {
    const model = doc(paragraph(["Visit example now"]));
    const result = applyLink(model, selection([0, 6], [0, 13]), { href: "https://example.com" });
    const runs = asParagraph(result.model.sections[0].blocks[0]).runs;
    expect(runs).toEqual([
      { text: "Visit ", formatting: {} },
      { text: "example", formatting: { linkHref: "https://example.com" } },
      { text: " now", formatting: {} },
    ]);
  });

  it("reads the link extent around a caret (AC-8)", () => {
    const model = doc(
      paragraph(["Visit "], ["example", { linkHref: "https://example.com" }], [" now"]),
    );
    const extent = linkAtCaret(model, { section: 0, block: 0, offset: 9 });
    expect(extent).not.toBeNull();
    expect(extent!.from).toBe(6);
    expect(extent!.to).toBe(13);
    expect(extent!.text).toBe("example");
    expect(extent!.target).toEqual({ href: "https://example.com" });
  });

  it("returns null outside a link", () => {
    const model = doc(
      paragraph(["Visit "], ["example", { linkHref: "https://example.com" }], [" now"]),
    );
    expect(linkAtCaret(model, { section: 0, block: 0, offset: 3 })).toBeNull();
  });

  it("spans differently formatted runs sharing one target", () => {
    const model = doc(
      paragraph(
        ["one", { linkHref: "https://example.com" }],
        ["two", { linkHref: "https://example.com", bold: true }],
        ["out", { linkHref: "https://other.com" }],
      ),
    );
    const extent = linkAtCaret(model, { section: 0, block: 0, offset: 2 });
    expect(extent!.from).toBe(0);
    expect(extent!.to).toBe(6);
  });

  it("removes the whole link at a collapsed caret, keeping text and formatting (AC-11, AC-33, AC-34)", () => {
    const model = doc(
      paragraph(
        ["Visit "],
        ["bold link", { linkHref: "https://example.com", bold: true, fontSizePt: 14 }],
      ),
    );
    const result = removeLink(model, caretSelection({ section: 0, block: 0, offset: 10 }));
    const runs = asParagraph(result.model.sections[0].blocks[0]).runs;
    expect(runs).toEqual([
      { text: "Visit ", formatting: {} },
      { text: "bold link", formatting: { bold: true, fontSizePt: 14 } },
    ]);
    expect(blockText(result.model.sections[0].blocks[0])).toBe("Visit bold link");
  });

  it("inserts new linked text at a collapsed caret (AC-2)", () => {
    const model = doc(paragraph(["before after"]));
    const result = insertLinkText(
      model,
      caretSelection({ section: 0, block: 0, offset: 7 }),
      "example",
      { href: "https://example.com" },
    );
    expect(blockText(result.model.sections[0].blocks[0])).toBe("before exampleafter");
    const runs = asParagraph(result.model.sections[0].blocks[0]).runs;
    expect(runs[1]).toEqual({ text: "example", formatting: { linkHref: "https://example.com" } });
    expect(result.selection.focus.offset).toBe(14);
  });

  it("replaces a selection with the dialog's edited text", () => {
    const model = doc(paragraph(["Visit here now"]));
    const result = insertLinkText(model, selection([0, 6], [0, 10]), "example.com", {
      href: "https://example.com",
    });
    expect(blockText(result.model.sections[0].blocks[0])).toBe("Visit example.com now");
  });

  it("does not modify the input model (undo contract)", () => {
    const model = doc(paragraph(["Visit example"]));
    const before = JSON.stringify(model);
    applyLink(model, selection([0, 6], [0, 13]), { href: "https://example.com" });
    removeLink(model, caretSelection({ section: 0, block: 0, offset: 8 }));
    expect(JSON.stringify(model)).toBe(before);
  });
});

describe("autolinkAt (AC-5..7)", () => {
  it("links the bare URL and returns the model only — the caller keeps its own caret", () => {
    const model = doc(paragraph(["See http://example.com "]));
    const next = autolinkAt(model, { section: 0, block: 0 }, 22);
    expect(next).not.toBeNull();
    const runs = asParagraph(next!.sections[0].blocks[0]).runs;
    expect(runs).toEqual([
      { text: "See ", formatting: {} },
      { text: "http://example.com", formatting: { linkHref: "http://example.com" } },
      { text: " ", formatting: {} },
    ]);
    // the input model is untouched, so the caller's two-step undo works
    expect(asParagraph(model.sections[0].blocks[0]).runs).toHaveLength(1);
  });

  it("returns null when the word before the offset is not a URL", () => {
    const model = doc(paragraph(["plain words "]));
    expect(autolinkAt(model, { section: 0, block: 0 }, 11)).toBeNull();
  });

  it("never re-links text already inside a link", () => {
    const model = doc(paragraph(["http://example.com", { linkHref: "http://example.com" }]));
    expect(autolinkAt(model, { section: 0, block: 0 }, 18)).toBeNull();
  });
});

describe("typing at a link boundary", () => {
  it("does not extend the link when typing at its end", () => {
    const model = doc(paragraph(["go: "], ["link", { linkHref: "https://example.com" }]));
    insertText(model, { section: 0, block: 0, offset: 8 }, "x");
    const runs = asParagraph(model.sections[0].blocks[0]).runs;
    expect(runs[runs.length - 1]).toEqual({ text: "x", formatting: {} });
  });

  it("keeps the link when typing strictly inside it", () => {
    const model = doc(paragraph(["link", { linkHref: "https://example.com" }]));
    insertText(model, { section: 0, block: 0, offset: 2 }, "x");
    const runs = asParagraph(model.sections[0].blocks[0]).runs;
    expect(runs).toEqual([
      { text: "lixnk", formatting: { linkHref: "https://example.com" } },
    ]);
  });
});

// --- bookmarks (AC-22..28) ---------------------------------------------------

describe("bookmarks", () => {
  it("inserts a bookmark at the caret and lists it in document order (AC-22, AC-23)", () => {
    const model = doc(paragraph(["first"]), paragraph(["second"]));
    const one = insertBookmark(model, caretSelection({ section: 0, block: 1, offset: 3 }), "later");
    const two = insertBookmark(one.model, caretSelection({ section: 0, block: 0, offset: 2 }), "earlier");
    expect(listBookmarks(two.model)).toEqual([
      { name: "earlier", address: { section: 0, block: 0 }, offset: 2 },
      { name: "later", address: { section: 0, block: 1 }, offset: 3 },
    ]);
  });

  it("resolves a bookmark to its caret for jumping (AC-24)", () => {
    const model = doc(paragraph(["hello world"]));
    const result = insertBookmark(model, caretSelection({ section: 0, block: 0, offset: 6 }), "mark");
    expect(bookmarkCaret(result.model, "mark")).toEqual({ section: 0, block: 0, offset: 6 });
    expect(bookmarkCaret(result.model, "missing")).toBeNull();
  });

  it("deletes a bookmark without touching the text (AC-25)", () => {
    const model = doc(paragraph(["hello"]));
    const added = insertBookmark(model, caretSelection({ section: 0, block: 0, offset: 2 }), "mark");
    const removed = deleteBookmark(added.model, caretSelection({ section: 0, block: 0, offset: 0 }), "mark");
    expect(listBookmarks(removed.model)).toEqual([]);
    expect(blockText(removed.model.sections[0].blocks[0])).toBe("hello");
    expect(asParagraph(removed.model.sections[0].blocks[0]).bookmarks).toBeUndefined();
  });

  it("validates names: empty, spaces, uniqueness (AC-26..28)", () => {
    expect(validateBookmarkName("", [])).toContain("empty");
    expect(validateBookmarkName("  ", [])).toContain("empty");
    expect(validateBookmarkName("has space", [])).toContain("spaces");
    expect(validateBookmarkName("dup", ["dup"])).toContain("already exists");
    expect(validateBookmarkName("ok_name-1", ["other"])).toBe("");
  });
});

describe("bookmark anchors track edits", () => {
  it("shift with text inserted before them", () => {
    const model = doc(paragraph(["hello world"]));
    const withMark = insertBookmark(model, caretSelection({ section: 0, block: 0, offset: 6 }), "mark").model;
    insertText(withMark, { section: 0, block: 0, offset: 0 }, "abc");
    expect(listBookmarks(withMark)[0].offset).toBe(9);
  });

  it("stay put when text is typed exactly at the anchor", () => {
    const model = doc(paragraph(["hello"]));
    const withMark = insertBookmark(model, caretSelection({ section: 0, block: 0, offset: 2 }), "mark").model;
    insertText(withMark, { section: 0, block: 0, offset: 2 }, "xx");
    expect(listBookmarks(withMark)[0].offset).toBe(2);
  });

  it("distribute across a paragraph split without duplication", () => {
    const model = doc(paragraph(["one two"]));
    let next = insertBookmark(model, caretSelection({ section: 0, block: 0, offset: 1 }), "a").model;
    next = insertBookmark(next, caretSelection({ section: 0, block: 0, offset: 6 }), "b").model;
    splitParagraph(next, { section: 0, block: 0, offset: 4 });
    expect(listBookmarks(next)).toEqual([
      { name: "a", address: { section: 0, block: 0 }, offset: 1 },
      { name: "b", address: { section: 0, block: 1 }, offset: 2 },
    ]);
  });

  it("collapse into a same-paragraph deletion instead of vanishing", () => {
    const model = doc(paragraph(["abcdef"]));
    const withMark = insertBookmark(model, caretSelection({ section: 0, block: 0, offset: 4 }), "mark").model;
    deleteRange(withMark, { section: 0, block: 0, offset: 1 }, { section: 0, block: 0, offset: 5 });
    expect(listBookmarks(withMark)[0].offset).toBe(1);
  });

  it("survive a cross-paragraph merge", () => {
    const model = doc(paragraph(["first"]), paragraph(["second"]));
    let next = insertBookmark(model, caretSelection({ section: 0, block: 0, offset: 2 }), "a").model;
    next = insertBookmark(next, caretSelection({ section: 0, block: 1, offset: 4 }), "b").model;
    deleteRange(next, { section: 0, block: 0, offset: 4 }, { section: 0, block: 1, offset: 2 });
    expect(listBookmarks(next)).toEqual([
      { name: "a", address: { section: 0, block: 0 }, offset: 2 },
      { name: "b", address: { section: 0, block: 0 }, offset: 6 },
    ]);
  });
});

// --- heading targets (AC-19..21) ---------------------------------------------

describe("heading link targets", () => {
  function headed(text: string, level: number): Paragraph {
    return { ...paragraph([text]), headingLevel: level, styleId: `Heading${level}` };
  }

  it("lists headings with deterministic anchors", () => {
    const model = doc(headed("Introduction", 1), paragraph(["body"]), headed("Details", 2));
    const headings = listHeadings(model);
    expect(headings).toHaveLength(2);
    expect(headings[0]).toMatchObject({
      text: "Introduction",
      level: 1,
      anchor: headingAnchor("Introduction", 0),
      address: { section: 0, block: 0 },
    });
    expect(headings[1].anchor).toBe(headingAnchor("Details", 1));
  });

  it("derives anchors without spaces, stable per text and ordinal", () => {
    expect(headingAnchor("Hello World!", 3)).toBe("_h3_Hello_World");
    expect(headingAnchor("", 0)).toBe("_h0_heading");
    expect(/\s/.test(headingAnchor("a b c", 1))).toBe(false);
  });

  it("resolves anchors: bookmarks first, then headings", () => {
    const model = doc(headed("Intro", 1), paragraph(["text here"]));
    const anchor = listHeadings(model)[0].anchor;
    expect(anchorCaret(model, anchor)).toEqual({ section: 0, block: 0, offset: 0 });
    const withMark = insertBookmark(model, caretSelection({ section: 0, block: 1, offset: 5 }), "mark").model;
    expect(anchorCaret(withMark, "mark")).toEqual({ section: 0, block: 1, offset: 5 });
    expect(anchorCaret(model, "unknown")).toBeNull();
  });

  it("labels targets for the chip: URL, heading text or bookmark name", () => {
    const model = doc(headed("Intro", 1));
    expect(linkTargetLabel(model, { href: "https://example.com" })).toBe("https://example.com");
    expect(linkTargetLabel(model, { anchor: listHeadings(model)[0].anchor })).toBe("Intro");
    expect(linkTargetLabel(model, { anchor: "somemark" })).toBe("somemark");
  });

  it("reads a run's target", () => {
    expect(linkTargetOf({ linkHref: "https://x.com" })).toEqual({ href: "https://x.com" });
    expect(linkTargetOf({ linkAnchor: "mark" })).toEqual({ anchor: "mark" });
    expect(linkTargetOf({ bold: true })).toBeNull();
  });
});
