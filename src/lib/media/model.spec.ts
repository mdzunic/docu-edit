// MediaModel invariants (SPEC-107 AC-3): the model is plain JSON — no Maps,
// no typed arrays, no class instances — so the store persists it through the
// same IndexedDB machinery as the other two products, and the summaries an
// open computes survive the round trip.
import { describe, expect, it } from "vitest";
import { assembleMediaModel } from "./open";
import { MEDIA_CONTENT_TYPES } from "./model";
import {
  buildPagesPdfBytes,
  buildPngFixtureBytes,
  PNG_FIXTURE_HEIGHT,
  PNG_FIXTURE_WIDTH,
} from "./fixtures/build";

describe("MediaModel", () => {
  it("is plain JSON: JSON.parse(JSON.stringify(model)) deep-equals it (image)", async () => {
    const model = await assembleMediaModel("shot", "png", await buildPngFixtureBytes());
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
  });

  it("is plain JSON for a PDF, summaries included", async () => {
    const model = await assembleMediaModel("pages", "pdf", buildPagesPdfBytes());
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
  });

  it("assembles an image model: payload in media, dimensions summarised", async () => {
    const bytes = await buildPngFixtureBytes();
    const model = await assembleMediaModel("shot", "png", bytes);
    expect(model.kind).toBe("media");
    expect(model.format).toBe("png");
    expect(model.contentType).toBe(MEDIA_CONTENT_TYPES.png);
    expect(model.byteLength).toBe(bytes.length);
    expect(model.media).toHaveLength(1);
    expect(model.media[0].mediaId).toBe(model.sourceMediaId);
    expect(model.image).toEqual({ widthPx: PNG_FIXTURE_WIDTH, heightPx: PNG_FIXTURE_HEIGHT });
    expect(model.pdf).toBeUndefined();
  });

  it("assembles a PDF model: structural summary and /Info properties", async () => {
    const model = await assembleMediaModel("pages", "pdf", buildPagesPdfBytes());
    expect(model.pdf?.pageCount).toBe(3);
    expect(model.pdf?.version).toBe("1.7");
    expect(model.pdf?.reconstructed).toBe(false);
    expect(model.pdf?.pages[1].rotate).toBe(90);
    expect(model.image).toBeUndefined();
    expect(model.properties?.title).toBe("SPEC-107 pages");
    expect(model.properties?.author).toBe("docu-edit factory");
    expect(model.properties?.created).toBe("2026-08-30 12:00");
    expect(model.properties?.extraFields?.[0]).toMatchObject({ name: "Producer", value: "docu-edit" });
  });
});
