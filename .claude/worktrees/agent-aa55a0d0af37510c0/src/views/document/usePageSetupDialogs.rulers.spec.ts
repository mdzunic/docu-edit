// The ruler visibility toggles (SPEC-027 AC-14, SPEC-064 AC-7): View → Rulers
// is one control over both rulers, View → Horizontal ruler / Vertical ruler
// flip each alone, and the per-document flags keep the "absent means shown"
// convention. touchDocument is stubbed — it is the library's autosave
// bookkeeping, not what is under test.
import { shallowRef, ref } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createDocument, type DocumentModel } from "../../lib/documentModel";

vi.mock("../../lib/documentStore", () => ({ touchDocument: vi.fn() }));

import { usePageSetupDialogs } from "./usePageSetupDialogs";

function harness(model: DocumentModel = createDocument("d")) {
  const doc = shallowRef<DocumentModel>(model);
  const bodySelection = { anchor: { section: 0, block: 0, offset: 0 }, focus: { section: 0, block: 0, offset: 0 } };
  const deps = {
    doc,
    contentRoot: ref<Element | null>({} as Element),
    editable: ref(true),
    applyEdit: vi.fn(),
    paragraphSelection: () => bodySelection,
    getCaretSectionIndex: () => 0,
  };
  // Only the fields the ruler toggles read are supplied (see the *Deps convention).
  const api = usePageSetupDialogs(deps as unknown as Parameters<typeof usePageSetupDialogs>[0]);
  return { api, doc };
}

describe("ruler visibility toggles", () => {
  it("reads absent flags as both rulers shown, with the group control checked", () => {
    const { api } = harness();
    expect(api.showHorizontalRuler.value).toBe(true);
    expect(api.showVerticalRuler.value).toBe(true);
    expect(api.rulersCheckState.value).toBe(true);
  });

  it("View → Rulers hides both together and shows both again (AC-7)", () => {
    const { api, doc } = harness();
    api.toggleRuler();
    expect(doc.value.showRuler).toBe(false);
    expect(doc.value.showVerticalRuler).toBe(false);
    expect(api.rulersCheckState.value).toBe(false);
    api.toggleRuler();
    expect(doc.value.showRuler).toBe(true);
    expect(doc.value.showVerticalRuler).toBe(true);
    expect(api.rulersCheckState.value).toBe(true);
  });

  it("the per-ruler toggles flip one ruler alone and leave the group control mixed", () => {
    const { api, doc } = harness();
    api.toggleVerticalRuler();
    expect(doc.value.showRuler).toBe(true);
    expect(doc.value.showVerticalRuler).toBe(false);
    expect(api.showHorizontalRuler.value).toBe(true);
    expect(api.showVerticalRuler.value).toBe(false);
    expect(api.rulersCheckState.value).toBe("mixed");

    api.toggleHorizontalRuler();
    expect(doc.value.showRuler).toBe(false);
    expect(api.rulersCheckState.value).toBe(false);

    api.toggleHorizontalRuler();
    api.toggleVerticalRuler();
    expect(api.rulersCheckState.value).toBe(true);
  });

  it("View → Rulers from a mixed state shows both rather than hiding the one that is on", () => {
    const { api, doc } = harness();
    api.toggleHorizontalRuler();
    expect(api.rulersCheckState.value).toBe("mixed");
    api.toggleRuler();
    expect(doc.value.showRuler).toBe(true);
    expect(doc.value.showVerticalRuler).toBe(true);
  });
});
