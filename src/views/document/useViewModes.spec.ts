import { effectScope, ref, shallowRef, type Ref } from "vue";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { useViewModes, viewModeStorageKey } from "./useViewModes";
import type { DocumentModel } from "../../lib/documentModel";
import { cmToPx } from "../../lib/core/units";

function mount(id = "doc-1"): { view: ReturnType<typeof useViewModes>; documentId: Ref<string>; stop: () => void } {
  const documentId = ref(id);
  const doc = shallowRef<DocumentModel | null>(null);
  const scope = effectScope();
  let view!: ReturnType<typeof useViewModes>;
  scope.run(() => {
    view = useViewModes({ documentId, doc });
  });
  return { view, documentId, stop: () => scope.stop() };
}

describe("useViewModes", () => {
  beforeEach(() => localStorage.clear());
  afterEach(() => localStorage.clear());

  it("opens in Print layout with Focus off by default (AC-48, AC-51)", () => {
    const h = mount();
    expect(h.view.viewMode.value).toBe("print");
    expect(h.view.focusMode.value).toBe(false);
    expect(h.view.isWebLayout.value).toBe(false);
    h.stop();
  });

  it("persists print/web under docu-edit:view-mode:<id> (AC-49)", () => {
    const h = mount("doc-7");
    h.view.setViewMode("web");
    expect(localStorage.getItem("docu-edit:view-mode:doc-7")).toBe("web");
    h.view.setViewMode("print");
    expect(localStorage.getItem("docu-edit:view-mode:doc-7")).toBe("print");
    h.stop();
  });

  it("restores the stored web mode on open (AC-30 analogue)", () => {
    localStorage.setItem(viewModeStorageKey("doc-8"), "web");
    const h = mount("doc-8");
    expect(h.view.viewMode.value).toBe("web");
    expect(h.view.isWebLayout.value).toBe(true);
    h.stop();
  });

  it("never writes Focus to storage and always opens with it off (AC-48, AC-50)", () => {
    const h = mount("doc-3");
    h.view.toggleFocus();
    expect(h.view.focusMode.value).toBe(true);
    expect(localStorage.getItem("docu-edit:view-mode:doc-3")).toBeNull();
    // A fresh mount for the same document opens with Focus off.
    const again = mount("doc-3");
    expect(again.view.focusMode.value).toBe(false);
    h.stop();
    again.stop();
  });

  it("Focus suppresses the web-layout class but keeps the underlying mode", () => {
    const h = mount();
    h.view.setViewMode("web");
    h.view.toggleFocus();
    expect(h.view.isWebLayout.value).toBe(false);
    expect(h.view.rootClass.value["is-focus-mode"]).toBe(true);
    h.view.exitFocus();
    expect(h.view.isWebLayout.value).toBe(true);
    h.stop();
  });

  it("exposes layout/focus command ids (AC-33)", () => {
    const h = mount();
    const ids = h.view.commands.map(([id]) => id);
    expect(ids).toEqual(["view.layout-print", "view.layout-web", "view.focus"]);
    // Dispatching them drives the state.
    const map = new Map(h.view.commands);
    map.get("view.layout-web")?.();
    expect(h.view.viewMode.value).toBe("web");
    map.get("view.focus")?.();
    expect(h.view.focusMode.value).toBe(true);
    h.stop();
  });

  describe("Web-layout content width (AC-40)", () => {
    it("is the first section's content width in px", () => {
      const documentId = ref("doc-1");
      // Page 21cm wide, 2cm + 3cm margins → 16cm content width.
      const model = {
        sections: [{ blocks: [], pageSize: { widthCm: 21, heightCm: 29.7 }, margins: { leftCm: 2, rightCm: 3, topCm: 2, bottomCm: 2 } }],
      } as unknown as DocumentModel;
      const doc = shallowRef<DocumentModel | null>(model);
      const scope = effectScope();
      let view!: ReturnType<typeof useViewModes>;
      scope.run(() => {
        view = useViewModes({ documentId, doc });
      });
      expect(view.webContentWidthCss.value).toBe(`${cmToPx(16)}px`);
      scope.stop();
    });

    it("falls back to auto with no sections", () => {
      const h = mount();
      expect(h.view.webContentWidthCss.value).toBe("auto");
      h.stop();
    });
  });
});
