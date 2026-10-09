// The editor routes' kind guards (SPEC-107 adds /media/:id): an id stored as
// one kind, navigated to another kind's route, is sent to its own route —
// in every direction — through documentKinds.editorPathFor; an unknown id
// stays put. The guards are called directly, so no view is loaded.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { buildPngFixtureBytes } from "./lib/media/fixtures/build";

beforeEach(() => {
  vi.resetModules();
  vi.doMock("./lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

afterEach(() => {
  vi.doUnmock("./lib/idb");
});

async function setup() {
  const store = await import("./lib/documentStore");
  const { routes } = await import("./router");
  const { assembleMediaModel } = await import("./lib/media/open");
  const { createBlankPresentation } = await import("./lib/presentation/slides");
  store.getOrCreateDocument("text-1");
  const deck = createBlankPresentation();
  store.storePresentation(deck);
  const media = await assembleMediaModel("Shot", "png", await buildPngFixtureBytes());
  store.storeMedia(media);
  /** Runs `path`'s beforeEnter guard for `id`. */
  function enter(path: string, id: string): unknown {
    const guard = routes.find((route) => route.path === path)?.beforeEnter as (to: unknown) => unknown;
    return guard({ params: { id, n: "1" } });
  }
  return { enter, deckId: deck.id, mediaId: media.id };
}

describe("editor route kind guards", () => {
  it("send a media id from the text and presentation routes to /media/:id", async () => {
    const { enter, mediaId } = await setup();
    expect(enter("/doc/:id", mediaId)).toBe(`/media/${mediaId}`);
    expect(enter("/slides/:id", mediaId)).toBe(`/media/${mediaId}`);
    expect(enter("/slides/:id/present/:n", mediaId)).toBe(`/media/${mediaId}`);
    expect(enter("/media/:id", mediaId)).toBe(true);
  });

  it("send a text or presentation id away from /media/:id to its own route", async () => {
    const { enter, deckId } = await setup();
    expect(enter("/media/:id", "text-1")).toBe("/doc/text-1");
    expect(enter("/media/:id", deckId)).toBe(`/slides/${deckId}`);
    expect(enter("/doc/:id", deckId)).toBe(`/slides/${deckId}`);
    expect(enter("/slides/:id", "text-1")).toBe("/doc/text-1");
    expect(enter("/doc/:id", "text-1")).toBe(true);
    expect(enter("/slides/:id", deckId)).toBe(true);
  });

  it("let an unknown id through to every route", async () => {
    const { enter } = await setup();
    for (const path of ["/doc/:id", "/slides/:id", "/slides/:id/present/:n", "/media/:id"]) {
      expect(enter(path, "nope")).toBe(true);
    }
  });
});
