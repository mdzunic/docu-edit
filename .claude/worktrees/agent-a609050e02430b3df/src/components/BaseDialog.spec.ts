import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { mount } from "@vue/test-utils";
import BaseDialog from "./BaseDialog.vue";

// Focus restoration is the one behaviour here that is invisible until you
// drive the app from the keyboard: when a dialog closes, the browser has
// already reset focus to <body>, so without an explicit hand-back the next
// Tab restarts at the top of the app — past the wordmark, the title, every
// menu and the whole toolbar. Every dialog in the app is built on this base,
// so it is worth pinning here rather than in each of them.

let trigger: HTMLButtonElement;

beforeEach(() => {
  trigger = document.createElement("button");
  trigger.textContent = "Open";
  document.body.append(trigger);
});

afterEach(() => {
  trigger.remove();
  document.body.innerHTML = "";
});

function open(slots: Record<string, string> = {}) {
  return mount(BaseDialog, {
    attachTo: document.body,
    props: { title: "Insert table", testid: "insert-table-dialog" },
    slots: { default: '<input data-testid="rows">', ...slots },
  });
}

describe("BaseDialog focus", () => {
  it("moves focus into the panel when it opens", async () => {
    trigger.focus();
    const wrapper = open();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(wrapper.element.contains(document.activeElement)).toBe(true);
    wrapper.unmount();
  });

  it("returns focus to whatever had it when the dialog closes", async () => {
    trigger.focus();
    expect(document.activeElement).toBe(trigger);

    const wrapper = open();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(document.activeElement).not.toBe(trigger);

    wrapper.unmount();
    expect(document.activeElement).toBe(trigger);
  });

  it("prefers an [autofocus] element for its initial focus", async () => {
    const wrapper = open({ default: '<input><input autofocus data-testid="second">' });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect((document.activeElement as HTMLElement).dataset.testid).toBe("second");
    wrapper.unmount();
  });

  it("leaves focus alone when the dialog's own action has already placed it", async () => {
    trigger.focus();
    const wrapper = open();
    await new Promise((resolve) => setTimeout(resolve, 0));

    // What a confirming dialog does: put the caret where the user now needs
    // it. Closing must not drag focus back to the trigger.
    const elsewhere = document.createElement("input");
    document.body.append(elsewhere);
    elsewhere.focus();

    wrapper.unmount();
    expect(document.activeElement).toBe(elsewhere);
  });

  it("does not throw when the trigger is gone by the time it closes", async () => {
    trigger.focus();
    const wrapper = open();
    await new Promise((resolve) => setTimeout(resolve, 0));

    // A menu item that opened this dialog and unmounted with its menu.
    trigger.remove();

    expect(() => wrapper.unmount()).not.toThrow();
    expect(document.activeElement).toBe(document.body);
  });
});
