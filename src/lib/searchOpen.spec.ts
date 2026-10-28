// SPEC-114 — a search result opens AT its match: the route and `n` a result
// links to, counted by each product's own find surface.
import { beforeEach, describe, expect, it, vi } from "vitest";
import { findParams } from "./findLink";

beforeEach(() => {
  vi.resetModules();
  vi.doMock("./idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

const para = (text: string) => ({ runs: [{ text }] });

async function modules() {
  const store = await import("./documentStore");
  const open = await import("./searchOpen");
  const { createDocument } = await import("./documentModel");
  const { documentText } = await import("./documentText");
  const { deckText } = await import("./presentation/deckText");
  return { store, open, createDocument, documentText, deckText };
}

describe("findParams", () => {
  it("reads find and a 1-based n, safely", () => {
    expect(findParams({ find: " rotterdam ", n: "3" })).toEqual({ find: "rotterdam", n: 3 });
    expect(findParams({})).toEqual({ find: "", n: 1 });
    expect(findParams({ find: ["a"], n: "zero" })).toEqual({ find: "", n: 1 });
    expect(findParams({ find: "x", n: "0" })).toEqual({ find: "x", n: 1 });
  });
});

describe("searchResultLocation", () => {
  it("opens a text document's find bar at the nth match the result names", async () => {
    const { store, open, createDocument, documentText } = await modules();
    const doc = createDocument("doc-1");
    doc.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "Rotterdam in the intro.", formatting: {} }] },
      { type: "paragraph", runs: [{ text: "The Rotterdam figures rose.", formatting: {} }] },
    ];
    store.storeDocument(doc);
    const text = documentText(doc).text;
    // The index's first body match is the second paragraph's.
    const at = text.indexOf("Rotterdam figures");
    expect(open.searchResultLocation("doc-1", "rotterdam figures", at)).toEqual({ path: "/doc/doc-1", query: { find: "rotterdam figures" } });
    expect(open.searchResultLocation("doc-1", "rotterdam", at)).toEqual({ path: "/doc/doc-1", query: { find: "rotterdam", n: "2" } });
    expect(open.searchResultLocation("doc-1", "rotterdam", 0)).toEqual({ path: "/doc/doc-1", query: { find: "rotterdam" } });
    // A filters-only query has nothing to find: the plain editor route.
    expect(open.searchResultLocation("doc-1", "", -1)).toBe("/doc/doc-1");
  });

  it("opens a deck's find at the first hit on the result's slide", async () => {
    const { store, open, deckText } = await modules();
    const box = { xCm: 0, yCm: 0, widthCm: 10, heightCm: 2 };
    const deck = {
      kind: "presentation" as const,
      id: "deck-1",
      title: "Deck",
      format: "pptx" as const,
      slideWidthCm: 33.867,
      slideHeightCm: 19.05,
      slides: [
        { id: "s1", frames: [{ kind: "text" as const, id: "t1", ...box, paragraphs: [para("Berlin first")] }] },
        { id: "s2", frames: [{ kind: "text" as const, id: "t2", ...box, paragraphs: [para("Open Berlin office")] }] },
      ],
    };
    store.storePresentation(deck);
    const text = deckText(deck).text;
    expect(open.searchResultLocation("deck-1", "Berlin", text.indexOf("Berlin office"))).toEqual({
      path: "/slides/deck-1",
      query: { find: "Berlin", n: "2" },
    });
  });

  it("hands a PDF its query, routed to the media viewer", async () => {
    const { store, open } = await modules();
    const { assembleMediaModel } = await import("./media/open");
    const { buildScannedPdfBytes } = await import("./media/fixtures/build");
    const media = await assembleMediaModel("scan", "pdf", buildScannedPdfBytes());
    store.storeMedia(media);
    expect(open.searchResultLocation(media.id, "chapter", 10)).toEqual({ path: `/media/${media.id}`, query: { find: "chapter" } });
  });
});
