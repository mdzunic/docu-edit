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
