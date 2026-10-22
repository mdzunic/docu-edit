import { mount } from "@vue/test-utils";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import LibraryThumbnail from "./LibraryThumbnail.vue";
import * as documentStore from "../../lib/documentStore";
import type { DocumentModel } from "../../lib/documentModel";

// An empty-but-valid model: renderThumbnail(...)/buildThumbnail(...) return a
// no-op thumbnail for it without throwing, so getDocument being called at all
// is the clean observable of "the deferred render ran".
const MODEL = { sections: [] } as unknown as DocumentModel;

// SPEC-067: the thumbnail render is scheduled through the (uncancellable)
// idle scheduler, so a card that scrolls out and unmounts before its idle
// slot fires must NOT still do the work. In jsdom IntersectionObserver and
// requestIdleCallback are absent, so onMounted renders directly on the
// setTimeout fallback that fake timers drive.
describe("LibraryThumbnail", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.stubGlobal("IntersectionObserver", undefined);
    vi.stubGlobal("requestIdleCallback", undefined);
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("renders the deferred preview when the card stays mounted", () => {
    const getDocument = vi.spyOn(documentStore, "getDocument").mockReturnValue(MODEL);
    mount(LibraryThumbnail, { props: { id: "doc-a", lastModified: 1, format: "docx" } });
    vi.runAllTimers();
    expect(getDocument).toHaveBeenCalledWith("doc-a");
  });

  it("renders a presentation's first slide as a miniature, not the format badge", async () => {
    const deck = {
      kind: "presentation",
      id: "deck-a",
      title: "Deck",
      format: "pptx",
      slideWidthCm: 33.87,
      slideHeightCm: 19.05,
      media: [],
      slides: [
        {
          id: "s1",
          frames: [
            {
              id: "f1",
              kind: "text",
              xCm: 2,
              yCm: 2,
              widthCm: 20,
              heightCm: 4,
              paragraphs: [
                { runs: [{ text: "Nova launch plan", bold: true, sizePt: 40 }] },
              ],
            },
          ],
        },
      ],
    };
    const getPresentation = vi
      .spyOn(documentStore, "getPresentation")
      .mockImplementation((id) => (id === "deck-a" ? (deck as never) : undefined));
    const wrapper = mount(LibraryThumbnail, {
      props: { id: "deck-a", lastModified: 1, format: "pptx" },
    });
    vi.runAllTimers();
    await wrapper.vm.$nextTick();
    expect(getPresentation).toHaveBeenCalledWith("deck-a");
    const miniature = wrapper.find('[data-testid="slide-miniature"]');
    expect(miniature.exists()).toBe(true);
    expect(miniature.text()).toContain("Nova launch plan");
    expect(wrapper.find(".thumbnail-fallback").exists()).toBe(false);
  });

  it("renders a document's images in the preview", async () => {
    const model = {
      sections: [
        {
          blocks: [
            { type: "paragraph", runs: [{ text: "Title", formatting: {} }], headingLevel: 1 },
            {
              type: "image",
              mediaId: "m1",
              widthCm: 8,
              heightCm: 4,
              wrap: "inline",
              naturalWidthPx: 100,
              naturalHeightPx: 50,
            },
          ],
        },
      ],
      defaultStyles: {
        pageSize: { widthCm: 21, heightCm: 29.7 },
        margins: { topCm: 2.54, bottomCm: 2.54, leftCm: 2.54, rightCm: 2.54 },
      },
      media: [{ mediaId: "m1", contentType: "image/png", data: "AAAA" }],
    } as unknown as DocumentModel;
    vi.spyOn(documentStore, "getDocument").mockReturnValue(model);
    const wrapper = mount(LibraryThumbnail, {
      props: { id: "doc-img", lastModified: 1, format: "docx" },
    });
    vi.runAllTimers();
    await wrapper.vm.$nextTick();
    const img = wrapper.get("img.thumbnail-image");
    expect(img.attributes("src")).toBe("data:image/png;base64,AAAA");
    expect(img.attributes("style")).toContain("aspect-ratio: 2");
    expect(wrapper.text()).toContain("Title");
    wrapper.unmount();
  });

  it("renders a text-box shape as a painted canvas with its text overlaid", async () => {
    const model = {
      sections: [
        {
          blocks: [
            {
              type: "shape",
              geometry: "rect",
              widthCm: 8,
              heightCm: 4,
              wrap: "square",
              fill: "#ffe0a0",
              text: {
                vAlign: "middle",
                blocks: [
                  { type: "paragraph", runs: [{ text: "Note to self", formatting: {} }] },
                ],
              },
            },
          ],
        },
      ],
      defaultStyles: {
        pageSize: { widthCm: 21, heightCm: 29.7 },
        margins: { topCm: 2.54, bottomCm: 2.54, leftCm: 2.54, rightCm: 2.54 },
      },
    } as unknown as DocumentModel;
    vi.spyOn(documentStore, "getDocument").mockReturnValue(model);
    const wrapper = mount(LibraryThumbnail, {
      props: { id: "doc-shape", lastModified: 1, format: "docx" },
    });
    vi.runAllTimers();
    await wrapper.vm.$nextTick();
    const shape = wrapper.get(".thumbnail-shape");
    expect(shape.find("canvas.thumbnail-shape-canvas").exists()).toBe(true);
    expect(shape.get(".thumbnail-shape-text").classes()).toContain("valign-middle");
    expect(shape.text()).toContain("Note to self");
    wrapper.unmount();
  });

  it("drops the scheduled render when the card unmounts first", () => {
    const getDocument = vi.spyOn(documentStore, "getDocument").mockReturnValue(MODEL);
    const wrapper = mount(LibraryThumbnail, {
      props: { id: "doc-b", lastModified: 1, format: "docx" },
    });
    wrapper.unmount();
    vi.runAllTimers();
    expect(getDocument).not.toHaveBeenCalled();
  });
});
