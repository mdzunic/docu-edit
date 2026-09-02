// The Page setup dialog's Borders tab, driven through the REAL wiring
// (SPEC-074 AC-19..AC-23): openPageSetup's prefill goes into the real
// PageSetupDialog, the dialog's own OK payload comes back out, and
// handlePageSetupApply commits it. PageSetupDialog.spec.ts mounts the dialog
// with hand-written props, so it cannot catch a prefill that only the
// composable produces — this file exists for exactly that seam.
import { shallowRef, ref } from "vue";
import { describe, expect, it, vi } from "vitest";
import { mount } from "@vue/test-utils";
import { createDocument, sectionMargins, type DocumentModel } from "../../lib/documentModel";
import type { CommandResult } from "../../lib/commands";
import PageSetupDialog from "../../components/PageSetupDialog.vue";

vi.mock("../../lib/documentStore", () => ({ touchDocument: vi.fn() }));

import { usePageSetupDialogs } from "./usePageSetupDialogs";

function harness(model: DocumentModel = createDocument("d")) {
  const doc = shallowRef<DocumentModel>(model);
  const caret = { section: 0, block: 0, offset: 0 };
  const applyEdit = vi.fn();
  const deps = {
    doc,
    // handlePageSetupApply focuses the canvas on the way out; only that.
    contentRoot: ref<Element | null>({ focus: () => {} } as unknown as Element),
    editable: ref(true),
    applyEdit,
    paragraphSelection: () => ({ anchor: caret, focus: caret }),
    getCaretSectionIndex: () => 0,
  };
  const api = usePageSetupDialogs(deps as unknown as Parameters<typeof usePageSetupDialogs>[0]);
  return { api, doc, applyEdit };
}

/** The model handed to applyEdit by the last commit. */
function committed(applyEdit: ReturnType<typeof vi.fn>): DocumentModel {
  const calls = applyEdit.mock.calls;
  return (calls[calls.length - 1][0] as CommandResult).model;
}

/** Opens the dialog on the harness's document and mounts it on the prefill
 * openPageSetup actually produced — the wiring under test. */
function openDialog(api: ReturnType<typeof usePageSetupDialogs>) {
  api.openPageSetup();
  expect(api.pageSetupOpen.value).toBe(true);
  return mount(PageSetupDialog, { props: { initial: api.pageSetupInitial.value } });
}

describe("Page setup ↔ Borders tab wiring (SPEC-074 AC-19..AC-23)", () => {
  it("opens with the background switch off for a section that has none", () => {
    const { api } = harness();
    const wrapper = openDialog(api);
    // The regression: an absent background prefilled as "" IS a string, so
    // the switch came up checked and OK wrote "" onto every section.
    expect(api.pageSetupInitial.value.pageBackground).toBeNull();
    expect((wrapper.get('[data-testid="page-background-on"]').element as HTMLInputElement).checked).toBe(false);
  });

  it("OK leaves an untouched document with no page background at all (not \"\")", async () => {
    const { api, doc, applyEdit } = harness();
    const wrapper = openDialog(api);
    // A plain margin change: the user never visits the Borders tab.
    await wrapper.get('[data-testid="page-setup-margin-top"]').setValue("3");
    await wrapper.get('[data-testid="page-setup-apply"]').trigger("click");
    api.handlePageSetupApply(wrapper.emitted("apply")![0][0] as Parameters<typeof api.handlePageSetupApply>[0]);

    const model = committed(applyEdit);
    for (const section of model.sections) {
      expect(section.pageBackground).toBeUndefined();
      expect("pageBackground" in section).toBe(false);
    }
    // ...and the margin edit it was carrying still landed.
    // (Approximate: the field round-trips through the display unit.)
    expect(sectionMargins(model, 0).topCm).toBeCloseTo(3, 2);
  });

  it("still commits a background the user does switch on, and prefills it next time", async () => {
    const { api, doc, applyEdit } = harness();
    const wrapper = openDialog(api);
    await wrapper.get('[data-testid="page-setup-tab-borders"]').trigger("click");
    await wrapper.get('[data-testid="page-background-on"]').setValue(true);
    await wrapper.get('[data-testid="page-background-color"]').setValue("#eeffee");
    await wrapper.get('[data-testid="page-setup-apply"]').trigger("click");
    api.handlePageSetupApply(wrapper.emitted("apply")![0][0] as Parameters<typeof api.handlePageSetupApply>[0]);

    const model = committed(applyEdit);
    expect(model.sections[0].pageBackground).toBe("#eeffee");

    // Reopening on that document prefills it and comes up checked.
    doc.value = model;
    const reopened = openDialog(api);
    expect(api.pageSetupInitial.value.pageBackground).toBe("#eeffee");
    expect((reopened.get('[data-testid="page-background-on"]').element as HTMLInputElement).checked).toBe(true);
  });

  it("switching a background back off removes it from the section", async () => {
    const { api, doc, applyEdit } = harness();
    const withFill = createDocument("d");
    withFill.sections[0].pageBackground = "#eeffee";
    doc.value = withFill;

    const wrapper = openDialog(api);
    await wrapper.get('[data-testid="page-setup-tab-borders"]').trigger("click");
    await wrapper.get('[data-testid="page-background-on"]').setValue(false);
    await wrapper.get('[data-testid="page-setup-apply"]').trigger("click");
    api.handlePageSetupApply(wrapper.emitted("apply")![0][0] as Parameters<typeof api.handlePageSetupApply>[0]);

    expect("pageBackground" in committed(applyEdit).sections[0]).toBe(false);
  });

  it("prefills the caret section's own page border and commits an edit to it", async () => {
    const { api, applyEdit } = harness();
    const wrapper = openDialog(api);
    await wrapper.get('[data-testid="page-setup-tab-borders"]').trigger("click");
    await wrapper.get('[data-testid="page-border-edge-top"]').setValue(true);
    await wrapper.get('[data-testid="page-border-edge-left"]').setValue(true);
    await wrapper.get('[data-testid="page-setup-apply"]').trigger("click");
    api.handlePageSetupApply(wrapper.emitted("apply")![0][0] as Parameters<typeof api.handlePageSetupApply>[0]);

    const borders = committed(applyEdit).sections[0].pageBorders;
    expect(borders?.top).toEqual({ style: "single", widthPt: 1, color: "#000000" });
    expect(borders?.left).toEqual({ style: "single", widthPt: 1, color: "#000000" });
    expect(borders?.bottom).toBeUndefined();
    // No background was switched on, so none was written (the same seam).
    expect("pageBackground" in committed(applyEdit).sections[0]).toBe(false);
  });
});
