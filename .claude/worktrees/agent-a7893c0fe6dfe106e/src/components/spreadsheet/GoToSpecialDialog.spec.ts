// SPEC-134 — the Go To Special dialog on BaseDialog: shown while the
// controller's `specialOpen`; the six kinds as one radio group, the four
// type filters (checked by default, disabled unless Constants or Formulas),
// the message, Cancel and OK.
import { afterEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import { computed, ref } from "vue";
import GoToSpecialDialog from "./GoToSpecialDialog.vue";

const mounted: VueWrapper[] = [];
afterEach(() => mounted.splice(0).forEach((wrapper) => wrapper.unmount()));

const KINDS = ["blanks", "constants", "formulas", "errors", "last-cell", "visible"];
const TYPES = ["numbers", "text", "logicals", "errors"];

/** A useSheetGoTo stand-in for the special dialog. */
function controller(shown = true) {
  const specialKind = ref("blanks");
  const specialTypes = ref({ numbers: true, text: true, logicals: true, errors: true });
  return {
    specialOpen: ref(shown),
    specialKind,
    specialTypes,
    specialMessage: ref(""),
    typesEnabled: computed(() => specialKind.value === "constants" || specialKind.value === "formulas"),
    setKind: vi.fn((kind: string) => void (specialKind.value = kind)),
    setType: vi.fn((type: "numbers" | "text" | "logicals" | "errors", on: boolean) => void (specialTypes.value = { ...specialTypes.value, [type]: on })),
    confirmSpecial: vi.fn(async () => true),
    cancelSpecial: vi.fn(),
  };
}

async function open(goTo = controller()) {
  const wrapper = mount(GoToSpecialDialog, { props: { goTo }, attachTo: document.body });
  mounted.push(wrapper);
  await flushPromises();
  return { wrapper, goTo };
}

describe("GoToSpecialDialog", () => {
  it("renders nothing while closed", async () => {
    const { wrapper } = await open(controller(false));
    expect(wrapper.find('[data-testid="sheet-goto-special-dialog"]').exists()).toBe(false);
  });

  it("offers the six kinds as one radio group and the four type filters, all checked, disabled for Blanks", async () => {
    const { wrapper } = await open();
    expect(wrapper.get('[data-testid="sheet-goto-special-dialog"]').get("h2").text()).toBe("Go To Special");
    const radios = KINDS.map((kind) => wrapper.get(`[data-testid="sheet-goto-special-${kind}"]`).element as HTMLInputElement);
    expect(radios.map((radio) => radio.type)).toEqual(Array(6).fill("radio"));
    expect(new Set(radios.map((radio) => radio.name)).size).toBe(1);
    expect(radios.map((radio) => radio.checked)).toEqual([true, false, false, false, false, false]);
    const boxes = TYPES.map((type) => wrapper.get(`[data-testid="sheet-goto-type-${type}"]`).element as HTMLInputElement);
    expect(boxes.map((box) => box.checked)).toEqual([true, true, true, true]);
    expect(wrapper.get("fieldset.special-types").attributes("disabled")).toBeDefined();
    expect(wrapper.find('[data-testid="sheet-goto-special-ok"]').exists()).toBe(true);
  });

  it("enables the type filters for Constants and Formulas only", async () => {
    const { wrapper, goTo } = await open();
    for (const [kind, id, enabled] of [
      ["constants", "constants", true],
      ["formulas", "formulas", true],
      ["errors", "errors", false],
      ["lastCell", "last-cell", false],
      ["visible", "visible", false],
    ] as const) {
      await wrapper.get(`[data-testid="sheet-goto-special-${id}"]`).setValue(true);
      expect(goTo.setKind).toHaveBeenLastCalledWith(kind);
      expect(wrapper.get("fieldset.special-types").attributes("disabled") === undefined).toBe(enabled);
    }
  });

  it("passes the filters, OK, Cancel and Escape to the controller", async () => {
    const { wrapper, goTo } = await open();
    await wrapper.get('[data-testid="sheet-goto-special-formulas"]').setValue(true);
    await wrapper.get('[data-testid="sheet-goto-type-text"]').setValue(false);
    expect(goTo.setType).toHaveBeenLastCalledWith("text", false);
    expect((wrapper.get('[data-testid="sheet-goto-type-text"]').element as HTMLInputElement).checked).toBe(false);
    await wrapper.get('[data-testid="sheet-goto-special-ok"]').trigger("click");
    expect(goTo.confirmSpecial).toHaveBeenCalledTimes(1);
    await wrapper.get('[data-testid="sheet-goto-special-cancel"]').trigger("click");
    await wrapper.get('[role="dialog"]').trigger("keydown", { key: "Escape" });
    expect(goTo.cancelSpecial).toHaveBeenCalledTimes(2);
  });

  it("shows the controller's message", async () => {
    const goTo = controller();
    goTo.specialMessage.value = "No cells were found.";
    const { wrapper } = await open(goTo);
    expect(wrapper.get('[data-testid="sheet-goto-special-message"]').text()).toBe("No cells were found.");
  });
});
