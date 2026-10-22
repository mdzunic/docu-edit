// SPEC-134 — the Go To dialog on BaseDialog: shown while the controller's
// `dialogOpen`; its list (names and tables, then the recent origins), the
// reference field (focused, pre-filled text selected), the error, Special…,
// Cancel and OK; a click on an item chooses it, Enter or a double-click on
// one goes, Enter in the field is OK, Escape cancels.
import { afterEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount, type VueWrapper } from "@vue/test-utils";
import { computed, ref } from "vue";
import GoToDialog from "./GoToDialog.vue";

const mounted: VueWrapper[] = [];
afterEach(() => mounted.splice(0).forEach((wrapper) => wrapper.unmount()));

/** A useSheetGoTo stand-in: its refs, and spies for what the dialog calls. */
function controller(open = true) {
  const reference = ref("Sales!A1");
  const error = ref("");
  const chosen = ref<string | null>(null);
  const items = computed(() => [
    { key: "name::TOTALS", kind: "name", text: "Totals" },
    { key: "table:t1", kind: "table", text: "Orders" },
    { key: "recent:0", kind: "recent", text: "Sales!A1" },
  ]);
  return {
    dialogOpen: ref(open),
    reference,
    error,
    chosen,
    items,
    setReference: vi.fn((text: string) => void (reference.value = text)),
    choose: vi.fn((key: string) => void (chosen.value = key)),
    confirm: vi.fn(async () => true),
    activate: vi.fn(async () => true),
    cancel: vi.fn(),
    openSpecial: vi.fn(),
  };
}

async function open(goTo = controller()) {
  const wrapper = mount(GoToDialog, { props: { goTo }, attachTo: document.body });
  mounted.push(wrapper);
  await flushPromises();
  return { wrapper, goTo };
}

describe("GoToDialog", () => {
  it("renders nothing while the dialog is closed", async () => {
    const { wrapper } = await open(controller(false));
    expect(wrapper.find('[data-testid="sheet-goto-dialog"]').exists()).toBe(false);
  });

  it("shows the list, the reference field (focused, its text selected), Special…, Cancel and OK", async () => {
    const { wrapper } = await open();
    expect(wrapper.get('[data-testid="sheet-goto-dialog"]').get("h2").text()).toBe("Go To");
    expect(wrapper.get('[data-testid="sheet-goto-list"]').attributes("role")).toBe("listbox");
    expect(wrapper.findAll('[data-testid="sheet-goto-item"]').map((item) => [item.text(), item.attributes("data-kind")])).toEqual([
      ["Totals", "name"],
      ["Orders", "table"],
      ["Sales!A1", "recent"],
    ]);
    const field = wrapper.get('[data-testid="sheet-goto-reference"]').element as HTMLInputElement;
    expect(document.activeElement).toBe(field);
    expect(field.value).toBe("Sales!A1");
    expect([field.selectionStart, field.selectionEnd]).toEqual([0, "Sales!A1".length]);
    for (const testid of ["sheet-goto-special", "sheet-goto-cancel", "sheet-goto-ok"]) {
      expect(wrapper.find(`[data-testid="${testid}"]`).exists()).toBe(true);
    }
    expect(wrapper.find('[data-testid="sheet-goto-error"]').exists()).toBe(false);
  });

  it("chooses an item on click, goes on Enter or a double-click, and marks the chosen one", async () => {
    const { wrapper, goTo } = await open();
    const items = wrapper.findAll('[data-testid="sheet-goto-item"]');
    await items[0]!.trigger("click");
    expect(goTo.choose).toHaveBeenCalledWith("name::TOTALS");
    expect(items[0]!.attributes("aria-selected")).toBe("true");
    await items[1]!.trigger("keydown", { key: "Enter" });
    await items[2]!.trigger("dblclick");
    expect(goTo.activate.mock.calls).toEqual([["table:t1"], ["recent:0"]]);
  });

  it("types into the controller, confirms on Enter and OK, cancels on Cancel and Escape, opens Special…", async () => {
    const { wrapper, goTo } = await open();
    const field = wrapper.get('[data-testid="sheet-goto-reference"]');
    await field.setValue("B12");
    expect(goTo.setReference).toHaveBeenLastCalledWith("B12");
    await field.trigger("keydown", { key: "Enter" });
    await wrapper.get('[data-testid="sheet-goto-ok"]').trigger("click");
    expect(goTo.confirm).toHaveBeenCalledTimes(2);
    await wrapper.get('[data-testid="sheet-goto-cancel"]').trigger("click");
    await wrapper.get('[role="dialog"]').trigger("keydown", { key: "Escape" });
    expect(goTo.cancel).toHaveBeenCalledTimes(2);
    await wrapper.get('[data-testid="sheet-goto-special"]').trigger("click");
    expect(goTo.openSpecial).toHaveBeenCalledTimes(1);
  });

  it("shows the controller's error as an alert tied to the field", async () => {
    const goTo = controller();
    goTo.error.value = "That isn't a cell reference or a defined name";
    const { wrapper } = await open(goTo);
    const error = wrapper.get('[data-testid="sheet-goto-error"]');
    expect(error.text()).toBe("That isn't a cell reference or a defined name");
    expect(error.attributes("role")).toBe("alert");
    expect(wrapper.get('[data-testid="sheet-goto-reference"]').attributes("aria-describedby")).toBe("sheet-goto-error");
  });
});
