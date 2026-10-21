// SPEC-130: the data-validation alert — a modal role="alertdialog" named by
// its title and described by its text, the style's icon, the buttons each
// style offers (stop: Retry, Cancel; warning: "Continue?", Yes, No, Cancel;
// information: OK, Cancel) with the focused one, each emitting its choice,
// Escape answering "cancel", keys stopping at the panel and Tab cycling its
// buttons.
import { flushPromises, mount } from "@vue/test-utils";
import { afterEach, describe, expect, it } from "vitest";
import ValidationAlert from "./ValidationAlert.vue";

const mounted: Array<() => void> = [];
afterEach(() => {
  for (const unmount of mounted.splice(0)) {
    unmount();
  }
});

async function alertOf(alertStyle: string, title = "Invalid entry", text = "This value doesn't match the data validation rules for this cell.") {
  const host = document.createElement("div");
  document.body.append(host);
  const wrapper = mount(ValidationAlert, { props: { alertStyle, title, text }, attachTo: host });
  mounted.push(() => {
    wrapper.unmount();
    host.remove();
  });
  await flushPromises();
  return wrapper;
}

const testIds = (wrapper: Awaited<ReturnType<typeof alertOf>>): string[] => wrapper.findAll("button").map((button) => button.attributes("data-testid") ?? "");

describe("ValidationAlert", () => {
  it("is a modal alertdialog named by its title and described by its text, line breaks kept", async () => {
    const wrapper = await alertOf("stop", "Priority", "Pick one\nof the list");
    const panel = wrapper.get('[data-testid="validation-alert"]');
    expect(panel.attributes("role")).toBe("alertdialog");
    expect(panel.attributes("aria-modal")).toBe("true");
    expect(document.getElementById(panel.attributes("aria-labelledby")!)?.textContent?.trim()).toBe("Priority");
    expect(document.getElementById(panel.attributes("aria-describedby")!)?.textContent?.trim()).toBe("Pick one\nof the list");
  });

  it("stop: the error icon, Retry (focused) and Cancel", async () => {
    const wrapper = await alertOf("stop");
    expect(wrapper.get('[data-testid="validation-alert-icon"]').attributes("data-icon")).toBe("error");
    expect(testIds(wrapper)).toEqual(["validation-alert-retry", "validation-alert-cancel"]);
    expect(wrapper.get('[data-testid="validation-alert-retry"]').text()).toBe("Retry");
    expect(wrapper.get('[data-testid="validation-alert-cancel"]').text()).toBe("Cancel");
    expect(document.activeElement).toBe(wrapper.get('[data-testid="validation-alert-retry"]').element);
    await wrapper.get('[data-testid="validation-alert-retry"]').trigger("click");
    await wrapper.get('[data-testid="validation-alert-cancel"]').trigger("click");
    expect(wrapper.emitted("answer")).toEqual([["retry"], ["cancel"]]);
  });

  it("warning: the warning icon, the text then \"Continue?\", Yes, No (focused) and Cancel", async () => {
    const wrapper = await alertOf("warning", "Out of range", "Between 1 and 10");
    expect(wrapper.get('[data-testid="validation-alert-icon"]').attributes("data-icon")).toBe("warning");
    const panel = wrapper.get('[data-testid="validation-alert"]');
    expect(panel.text()).toContain("Between 1 and 10");
    expect(panel.text()).toContain("Continue?");
    expect(panel.attributes("aria-describedby")).toBe("validation-alert-text validation-alert-question");
    expect(testIds(wrapper)).toEqual(["validation-alert-yes", "validation-alert-no", "validation-alert-cancel"]);
    expect(wrapper.findAll("button").map((button) => button.text())).toEqual(["Yes", "No", "Cancel"]);
    expect(document.activeElement).toBe(wrapper.get('[data-testid="validation-alert-no"]').element);
    for (const id of ["yes", "no", "cancel"]) {
      await wrapper.get(`[data-testid="validation-alert-${id}"]`).trigger("click");
    }
    expect(wrapper.emitted("answer")).toEqual([["yes"], ["no"], ["cancel"]]);
  });

  it("information: the info icon, OK (focused) and Cancel", async () => {
    const wrapper = await alertOf("information");
    expect(wrapper.get('[data-testid="validation-alert-icon"]').attributes("data-icon")).toBe("info");
    expect(testIds(wrapper)).toEqual(["validation-alert-ok", "validation-alert-cancel"]);
    expect(wrapper.get('[data-testid="validation-alert-ok"]').text()).toBe("OK");
    expect(document.activeElement).toBe(wrapper.get('[data-testid="validation-alert-ok"]').element);
    await wrapper.get('[data-testid="validation-alert-ok"]').trigger("click");
    expect(wrapper.emitted("answer")).toEqual([["ok"]]);
  });

  it("Escape answers cancel; every key stops at the panel; Tab cycles the buttons", async () => {
    const wrapper = await alertOf("warning");
    let reached = 0;
    const listener = (): void => {
      reached += 1;
    };
    window.addEventListener("keydown", listener);
    try {
      const panel = wrapper.get('[data-testid="validation-alert"]');
      await panel.trigger("keydown", { key: "z", ctrlKey: true });
      await panel.trigger("keydown", { key: "Tab" });
      expect(document.activeElement).toBe(wrapper.get('[data-testid="validation-alert-cancel"]').element);
      await panel.trigger("keydown", { key: "Tab" });
      expect(document.activeElement).toBe(wrapper.get('[data-testid="validation-alert-yes"]').element);
      await panel.trigger("keydown", { key: "Tab", shiftKey: true });
      expect(document.activeElement).toBe(wrapper.get('[data-testid="validation-alert-cancel"]').element);
      await panel.trigger("keydown", { key: "Escape" });
      expect(wrapper.emitted("answer")).toEqual([["cancel"]]);
      expect(reached).toBe(0);
    } finally {
      window.removeEventListener("keydown", listener);
    }
  });
});
