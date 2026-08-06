// SPEC-025: the link dialog — prefill (AC-1, AC-2), dangerous-scheme
// rejection on entry (AC-14..16, part of the AC-18 security coverage), and
// the heading/bookmark target list (AC-19..21).
import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import LinkDialog from "./LinkDialog.vue";

function mountDialog(props: Record<string, unknown> = {}) {
  return mount(LinkDialog, { props, attachTo: document.body });
}

describe("LinkDialog", () => {
  it("prefills the selected text with an empty URL (AC-1)", () => {
    const wrapper = mountDialog({ initialText: "selected words" });
    expect(
      (wrapper.get('[data-testid="link-text-input"]').element as HTMLInputElement).value,
    ).toBe("selected words");
    expect((wrapper.get('[data-testid="link-url-input"]').element as HTMLInputElement).value).toBe(
      "",
    );
  });

  it("starts with both fields empty without a selection (AC-2)", () => {
    const wrapper = mountDialog();
    expect((wrapper.get('[data-testid="link-text-input"]').element as HTMLInputElement).value).toBe(
      "",
    );
    expect((wrapper.get('[data-testid="link-url-input"]').element as HTMLInputElement).value).toBe(
      "",
    );
  });

  it("emits save with the sanitized URL", async () => {
    const wrapper = mountDialog({ initialText: "text" });
    await wrapper.get('[data-testid="link-url-input"]').setValue(" https://example.com ");
    await wrapper.get('[data-testid="link-save"]').trigger("click");
    expect(wrapper.emitted("save")).toEqual([[{ text: "text", href: "https://example.com" }]]);
  });

  it.each(["javascript:alert(1)", "data:text/html,x", "vbscript:msgbox(1)", "JaVaScRiPt:alert(1)"])(
    "rejects %s on entry with an inline error and a disabled save (AC-14..16, AC-18)",
    async (url) => {
      const wrapper = mountDialog();
      await wrapper.get('[data-testid="link-url-input"]').setValue(url);
      expect(wrapper.get('[data-testid="link-error"]').text()).toContain("isn't allowed");
      expect(wrapper.get('[data-testid="link-save"]').attributes("disabled")).toBeDefined();
      await wrapper.get('[data-testid="link-save"]').trigger("click");
      expect(wrapper.emitted("save")).toBeUndefined();
    },
  );

  it("disables save while both URL and target are empty", () => {
    const wrapper = mountDialog();
    expect(wrapper.get('[data-testid="link-save"]').attributes("disabled")).toBeDefined();
  });

  it("lists both headings and bookmarks as targets (AC-19..21)", () => {
    const wrapper = mountDialog({
      headings: [{ anchor: "_h0_Intro", text: "Intro", level: 1 }],
      bookmarks: [{ anchor: "mark1", text: "mark1" }],
    });
    const list = wrapper.get('[data-testid="link-target-list"]');
    expect(list.text()).toContain("Headings");
    expect(list.text()).toContain("Intro");
    expect(list.text()).toContain("Bookmarks");
    expect(list.text()).toContain("mark1");
  });

  it("saves an anchor target, defaulting empty text to its label (AC-19, AC-20)", async () => {
    const wrapper = mountDialog({
      headings: [{ anchor: "_h0_Intro", text: "Intro", level: 1 }],
    });
    await wrapper.get('[data-testid="link-target-_h0_Intro"]').trigger("click");
    expect(wrapper.get('[data-testid="link-save"]').attributes("disabled")).toBeUndefined();
    await wrapper.get('[data-testid="link-save"]').trigger("click");
    expect(wrapper.emitted("save")).toEqual([[{ text: "Intro", anchor: "_h0_Intro" }]]);
  });

  it("clears the anchor selection when a URL is typed, and vice versa", async () => {
    const wrapper = mountDialog({
      bookmarks: [{ anchor: "mark1", text: "mark1" }],
    });
    await wrapper.get('[data-testid="link-target-mark1"]').trigger("click");
    expect(wrapper.get('[data-testid="link-target-mark1"]').attributes("aria-pressed")).toBe(
      "true",
    );
    await wrapper.get('[data-testid="link-url-input"]').setValue("https://example.com");
    expect(wrapper.get('[data-testid="link-target-mark1"]').attributes("aria-pressed")).toBe(
      "false",
    );
    await wrapper.get('[data-testid="link-target-mark1"]').trigger("click");
    expect(
      (wrapper.get('[data-testid="link-url-input"]').element as HTMLInputElement).value,
    ).toBe("");
  });

  it("emits cancel from the cancel button", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="link-cancel"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
  });
});
