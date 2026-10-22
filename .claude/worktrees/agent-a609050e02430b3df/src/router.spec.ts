// The editor routes' kind guards (SPEC-107 adds /media/:id, SPEC-116
// /sheets/:id and the one shared guardKind): an id stored as one kind,
// navigated to another kind's route, is sent to its own route — in every
// direction — through documentKinds.editorPathFor; an unknown id stays put.
// The guards are called directly, so no view is loaded.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { buildPngFixtureBytes } from "./lib/media/fixtures/build";
import { buildMiniBookXlsxBytes } from "./lib/spreadsheet/fixtures/build";
import { editorPathFor, type DocumentKind } from "./lib/documentKinds";

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
  const { readSpreadsheetBytes } = await import("./lib/spreadsheet/open");
  const workbook = await readSpreadsheetBytes(await buildMiniBookXlsxBytes(), "xlsx", "mini-book");
  store.storeSpreadsheet(workbook);
  /** Runs `path`'s beforeEnter guard for `id`. */
  function enter(path: string, id: string): unknown {
    const guard = routes.find((route) => route.path === path)?.beforeEnter as (to: unknown) => unknown;
    return guard({ params: { id, n: "1" } });
  }
  return { enter, deckId: deck.id, mediaId: media.id, workbookId: workbook.id };
}

/** Every editor route and the kind that owns it. */
const ROUTE_KINDS = [
  ["/doc/:id", "text"],
  ["/slides/:id", "presentation"],
  ["/slides/:id/present/:n", "presentation"],
  ["/media/:id", "media"],
  ["/sheets/:id", "spreadsheet"],
] as const;

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
    for (const [path] of ROUTE_KINDS) {
      expect(enter(path, "nope")).toBe(true);
    }
  });

  it("send a workbook id from /doc/:id to /sheets/:id, and a document id the other way", async () => {
    const { enter, workbookId } = await setup();
    expect(enter("/doc/:id", workbookId)).toBe(`/sheets/${workbookId}`);
    expect(enter("/sheets/:id", "text-1")).toBe("/doc/text-1");
  });

  it("guard every kind × every route through one rule (SPEC-116's table)", async () => {
    const { enter, deckId, mediaId, workbookId } = await setup();
    const ids: Record<DocumentKind, string> = {
      text: "text-1",
      presentation: deckId,
      media: mediaId,
      spreadsheet: workbookId,
    };
    for (const [path, owner] of ROUTE_KINDS) {
      for (const [kind, id] of Object.entries(ids) as Array<[DocumentKind, string]>) {
        expect(enter(path, id), `${kind} id at ${path}`).toBe(kind === owner ? true : editorPathFor(kind, id));
      }
    }
  });

  it("route every editor path through the shared guardKind helper", async () => {
    const { routes, guardKind } = await import("./router");
    for (const [path, owner] of ROUTE_KINDS) {
      const guard = routes.find((route) => route.path === path)?.beforeEnter as (to: unknown) => unknown;
      expect(String(guard), path).toBe(String(guardKind(owner)));
    }
  });
});
