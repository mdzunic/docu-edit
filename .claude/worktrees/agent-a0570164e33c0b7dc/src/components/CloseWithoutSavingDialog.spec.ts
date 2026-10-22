// SPEC-060 AC-17 — the storage-unavailable confirmation offers exactly
// Download, Close anyway and Cancel, and emits each choice without acting on
// it (the caller owns the download/close/dismiss behaviour).
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import CloseWithoutSavingDialog from "./CloseWithoutSavingDialog.vue";

describe("CloseWithoutSavingDialog (SPEC-060 AC-17)", () => {
  it("offers Download, Close anyway and Cancel", () => {
    const wrapper = mount(CloseWithoutSavingDialog);
    expect(
      wrapper.get('[data-testid="close-without-saving-download"]').text(),
    ).toBe("Download");
    expect(
      wrapper.get('[data-testid="close-without-saving-confirm"]').text(),
    ).toBe("Close anyway");
    expect(
      wrapper.get('[data-testid="close-without-saving-cancel"]').text(),
    ).toBe("Cancel");
  });

  it("emits the chosen action", async () => {
    const wrapper = mount(CloseWithoutSavingDialog);
    await wrapper.get('[data-testid="close-without-saving-download"]').trigger("click");
    await wrapper.get('[data-testid="close-without-saving-confirm"]').trigger("click");
    await wrapper.get('[data-testid="close-without-saving-cancel"]').trigger("click");
    expect(wrapper.emitted("download")).toHaveLength(1);
    expect(wrapper.emitted("confirm")).toHaveLength(1);
    expect(wrapper.emitted("cancel")).toHaveLength(1);
  });
});
