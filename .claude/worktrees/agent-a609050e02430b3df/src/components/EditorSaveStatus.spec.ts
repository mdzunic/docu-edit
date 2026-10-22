// SPEC-010/060 — the editor's app-bar save cluster, extracted from
// DocumentView. It shows the save-state text, reveals the "Download a copy"
// escape hatch only in the error state, and routes both it and the Download
// menu's pick out through a single `download` event.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import EditorSaveStatus from "./EditorSaveStatus.vue";

describe("EditorSaveStatus", () => {
  it("shows the save-state text and no download link when saved", () => {
    const wrapper = mount(EditorSaveStatus, {
      props: { saveState: "saved", saveStateText: "All changes saved" },
    });
    expect(wrapper.get('[data-testid="save-state"]').text()).toBe(
      "All changes saved",
    );
    expect(wrapper.find('[data-testid="save-state-download"]').exists()).toBe(
      false,
    );
  });

  it("reveals the download escape hatch in the error state and emits the format", async () => {
    const wrapper = mount(EditorSaveStatus, {
      props: {
        saveState: "error",
        saveStateText: "Not saved",
        downloadFormat: "odt",
      },
    });
    await wrapper.get('[data-testid="save-state-download"]').trigger("click");
    expect(wrapper.emitted("download")).toEqual([["odt"]]);
  });
});
