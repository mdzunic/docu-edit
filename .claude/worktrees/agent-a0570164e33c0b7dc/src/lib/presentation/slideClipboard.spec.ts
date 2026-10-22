// SPEC-099 — the slide clipboard: both flavours, the size cap, the whitelist
// parser against hostile payloads, and nonce precedence.
import { afterEach, describe, expect, it } from "vitest";
import type { ImageFrame, PresentationModel, TextFrame } from "./model";
import { readPresentationBytes } from "./open";
import { buildPptxFixtureBytes, FIXTURE_SLIDES } from "./fixtures/build";
import {
  buildSlideClipboardPayload,
  clearSlideClipboardStore,
  latestSlideCopy,
  MAX_SLIDE_CLIPBOARD_BYTES,
  parseSlideClipboardPayload,
  resolveSlidePaste,
  SLIDES_MIME,
  slideClipboardFlavours,
  storedSlideCopy,
} from "./slideClipboard";

async function fixture(): Promise<PresentationModel> {
  return readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
}

afterEach(() => clearSlideClipboardStore());

describe("slideClipboardFlavours", () => {
  it("writes the outline as text/plain and the labelled payload, remembering it in-tab", async () => {
    const model = await fixture();
    const flavours = slideClipboardFlavours(model, [1, 0]);
    expect(flavours.text).toBe(
      [FIXTURE_SLIDES[0].title, ...FIXTURE_SLIDES[0].lines].join("\n") +
        "\n\n" +
        [FIXTURE_SLIDES[1].title, "Ship presentations MVP", "Broaden format support", "Hire two engineers", "Open Berlin office"].join("\n"),
    );
    const parsed = JSON.parse(flavours.payload!);
    expect(parsed.mimeType).toBe(SLIDES_MIME);
    expect(parsed.version).toBe(1);
    expect(parsed.nonce).toBe(flavours.nonce);
    expect(parsed.slides).toHaveLength(2);
    expect(storedSlideCopy(flavours.nonce)?.slides).toHaveLength(2);
    expect(latestSlideCopy()?.nonce).toBe(flavours.nonce);
  });

  it("carries only the media the copied slides reference, plus the theme", async () => {
    const model = await fixture();
    expect(buildSlideClipboardPayload(model, [0]).media).toEqual([]);
    const withImage = buildSlideClipboardPayload(model, [3]);
    expect(withImage.media.map((part) => part.mediaId)).toEqual(["ppt/media/image1.png"]);
    expect(withImage.theme).toEqual(model.theme);
  });

  it("omits the private flavour above the cap while the in-tab store still pastes the slides", async () => {
    const model = await fixture();
    model.media = [{ mediaId: "ppt/media/image1.png", contentType: "image/png", data: "A".repeat(MAX_SLIDE_CLIPBOARD_BYTES + 10) }];
    const flavours = slideClipboardFlavours(model, [3]);
    expect(flavours.payload).toBeNull();
    const pasted = resolveSlidePaste({ text: flavours.text });
    expect(pasted?.slides).toHaveLength(1);
    expect(pasted?.media[0].data).toHaveLength(MAX_SLIDE_CLIPBOARD_BYTES + 10);
  });
});

describe("parseSlideClipboardPayload", () => {
  it("round-trips a real copy", async () => {
    const model = await fixture();
    const payload = buildSlideClipboardPayload(model, [1, 3]);
    const { payload: json } = slideClipboardFlavours(model, [1, 3]);
    const parsed = parseSlideClipboardPayload(json!);
    expect(parsed?.slides).toHaveLength(2);
    expect(parsed?.format).toBe("pptx");
    expect((parsed!.slides[0].frames[0] as TextFrame).paragraphs).toEqual((payload.slides[0].frames[0] as TextFrame).paragraphs);
    const image = parsed!.slides[1].frames.find((frame) => frame.kind === "image") as ImageFrame;
    expect(image.mediaId).toBe("ppt/media/image1.png");
    expect(parsed?.media).toHaveLength(1);
  });

  it("rejects anything that is not a v1 SLIDES_MIME payload with slides", () => {
    expect(parseSlideClipboardPayload("not json")).toBeNull();
    expect(parseSlideClipboardPayload(JSON.stringify({ mimeType: "text/plain", version: 1 }))).toBeNull();
    const base = { mimeType: SLIDES_MIME, version: 1, nonce: "n1", format: "pptx", slideWidthCm: 33.867, slideHeightCm: 19.05 };
    expect(parseSlideClipboardPayload(JSON.stringify({ ...base, version: 2, slides: [{ frames: [] }] }))).toBeNull();
    expect(parseSlideClipboardPayload(JSON.stringify({ ...base, slides: [] }))).toBeNull();
    expect(parseSlideClipboardPayload(JSON.stringify({ ...base, format: "docx", slides: [{ frames: [] }] }))).toBeNull();
    expect(parseSlideClipboardPayload(JSON.stringify({ ...base, slideWidthCm: -1, slides: [{ frames: [] }] }))).toBeNull();
  });

  it("rebuilds hostile payloads from the whitelist", () => {
    const hostile = {
      mimeType: SLIDES_MIME,
      version: 1,
      nonce: "n2",
      format: "pptx",
      slideWidthCm: 33.867,
      slideHeightCm: 19.05,
      slides: [
        {
          id: "x",
          injected: "<script>",
          frames: [
            {
              kind: "text",
              id: "f",
              xCm: 1,
              yCm: 1,
              widthCm: 1,
              heightCm: 1,
              onclick: "alert(1)",
              paragraphs: [
                {
                  runs: [
                    { text: "hi", color: "url(javascript:alert(1))", linkHref: "javascript:alert(1)", sizePt: "12" },
                    { text: "ok", color: { theme: "accent2", lumMod: 0.5, evil: 1 }, linkHref: "https://example.com" },
                  ],
                },
              ],
            },
            { kind: "script", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1 },
            { kind: "image", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, mediaId: "../../etc/passwd" },
            { kind: "text", xCm: "0", yCm: 0, widthCm: 1, heightCm: 1, paragraphs: [] },
          ],
        },
      ],
      media: [
        { mediaId: "ppt/media/a.png", contentType: "text/html", data: "PGgxPg==" },
        { mediaId: "ppt/media/b.png", contentType: "image/png", data: "iVBORw0KGgo=" },
      ],
      theme: { name: "T", colors: { dk1: "red" }, fonts: { major: "A", minor: "B" } },
    };
    // A literal `__proto__` key, which only raw JSON can carry.
    const raw = JSON.stringify(hostile).replace(`"id":"x"`, `"id":"x","__proto__":{"polluted":true}`);
    const parsed = parseSlideClipboardPayload(raw);
    expect(parsed).not.toBeNull();
    const slide = parsed!.slides[0] as unknown as Record<string, unknown>;
    expect(slide.injected).toBeUndefined();
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
    expect(parsed!.slides[0].frames).toHaveLength(1);
    const frame = parsed!.slides[0].frames[0] as TextFrame & Record<string, unknown>;
    expect(frame.onclick).toBeUndefined();
    expect(frame.paragraphs[0].runs[0]).toEqual({ text: "hi" });
    expect(frame.paragraphs[0].runs[1]).toEqual({
      text: "ok",
      color: { theme: "accent2", lumMod: 0.5 },
      linkHref: "https://example.com",
    });
    expect(parsed!.media.map((part) => part.mediaId)).toEqual(["ppt/media/b.png"]);
    expect(parsed!.theme).toBeUndefined();
  });
});

describe("resolveSlidePaste", () => {
  it("gives the in-tab copy precedence over the clipboard's rebuilt payload for the same nonce", async () => {
    const model = await fixture();
    const flavours = slideClipboardFlavours(model, [2]);
    const stored = storedSlideCopy(flavours.nonce)!;
    // The clipboard copy is rebuilt (a new object); the stored one wins.
    expect(resolveSlidePaste({ slides: flavours.payload!, text: flavours.text })).toBe(stored);
    clearSlideClipboardStore();
    const fromClipboard = resolveSlidePaste({ slides: flavours.payload!, text: flavours.text });
    expect(fromClipboard).not.toBeNull();
    expect(fromClipboard).not.toBe(stored);
  });

  it("is null for a clipboard holding no slides (text from another application)", () => {
    expect(resolveSlidePaste({ text: "Just some text" })).toBeNull();
    expect(resolveSlidePaste({ slides: "{}", text: "" })).toBeNull();
    expect(resolveSlidePaste({})).toBeNull();
  });
});
