// @vitest-environment jsdom
// SPEC-097's slide text UI, mounted: the toolbar's numbering toggle and list
// popover (eight styles previewed through markerText, Escape returns focus to
// the trigger), the link button's session-only rendering, the insert bar's
// whole-frame link button, the slide link dialog's inline URL rejection, and
// the Text panel's controls driving their controller commands.
import { mount } from "@vue/test-utils";
import { computed, ref } from "vue";
import { afterEach, describe, expect, it, vi } from "vitest";
import { markerText } from "../../lib/core/lists";
import PresentationToolbar from "./PresentationToolbar.vue";
import SlideInsertBar from "./SlideInsertBar.vue";
import SlideLinkDialog from "./SlideLinkDialog.vue";
import TextPanel from "./TextPanel.vue";
import { LIST_STYLE_CHOICES } from "./slideTextControls";

afterEach(() => {
  document.body.innerHTML = "";
});

function toolbar(props: Record<string, unknown> = {}) {
  return mount(PresentationToolbar, {
    props: { enabled: true, state: { numbered: true, bulletStyle: "decimal", startAt: 1 }, ...props },
    attachTo: document.body,
  });
}

describe("PresentationToolbar list and link controls (AC-24..29, AC-46)", () => {

  it("slides-numbering reflects the numbered state and emits numbering", async () => {
    const wrapper = toolbar();
    const numbering = wrapper.get('[data-testid="slides-numbering"]');
    expect(numbering.attributes("aria-pressed")).toBe("true");
    await numbering.trigger("click");
    expect(wrapper.emitted("numbering")).toHaveLength(1);
  });

  it("slides-list-menu opens the eight styles previewed through markerText(style, 1)", async () => {
    const wrapper = toolbar();
    await wrapper.get('[data-testid="slides-list-menu"]').trigger("click");
    const popover = wrapper.get('[data-testid="list-style-popover"]');
    expect(LIST_STYLE_CHOICES).toHaveLength(8);
    for (const choice of LIST_STYLE_CHOICES) {
      expect(popover.get(`[data-testid="list-style-${choice.style}"]`).text()).toBe(markerText(choice.style, 1));
    }
    expect(popover.get('[data-testid="list-style-decimal"]').attributes("aria-pressed")).toBe("true");
    await popover.get('[data-testid="list-style-lowerRoman"]').trigger("click");
    expect(wrapper.emitted("list-style")).toEqual([["lowerRoman"]]);
    expect(wrapper.find('[data-testid="list-style-popover"]').exists()).toBe(false);
  });

  it("offers custom and picture bullets and the start-at/indent/hanging inputs", async () => {
    const wrapper = toolbar({ listGeometry: { indentCm: 1.27, hangingCm: 0.635 }, imageAccept: "image/png" });
    await wrapper.get('[data-testid="slides-list-menu"]').trigger("click");
    await wrapper.get('[data-testid="list-bullet-char"]').trigger("click");
    const option = wrapper.get('[data-testid="list-bullet-char-option"]');
    await option.trigger("click");
    expect(wrapper.emitted("bullet-char")?.[0]).toEqual([option.attributes("data-char")]);
    await wrapper.get('[data-testid="slides-list-menu"]').trigger("click");
    expect(wrapper.get('[data-testid="list-bullet-image-input"]').attributes("accept")).toBe("image/png");
    expect((wrapper.get('[data-testid="list-indent"]').element as HTMLInputElement).value).toBe("1.27");
    const startAt = wrapper.get('[data-testid="list-start-at"]');
    await startAt.setValue("4");
    expect(wrapper.emitted("start-at")).toEqual([[4]]);
    await wrapper.get('[data-testid="list-hanging"]').setValue("0.5");
    expect(wrapper.emitted("list-hanging")).toEqual([[0.5]]);
  });

  it("Escape closes the popover and returns focus to its trigger", async () => {
    const wrapper = toolbar();
    await wrapper.get('[data-testid="slides-list-menu"]').trigger("click");
    await wrapper.get('[data-testid="list-style-popover"]').trigger("keydown", { key: "Escape" });
    expect(wrapper.find('[data-testid="list-style-popover"]').exists()).toBe(false);
    expect(document.activeElement).toBe(wrapper.get('[data-testid="slides-list-menu"]').element);
  });

  it("renders slides-link only inside a text session", async () => {
    const wrapper = toolbar();
    expect(wrapper.find('[data-testid="slides-link"]').exists()).toBe(false);
    await wrapper.setProps({ linkEnabled: true });
    await wrapper.get('[data-testid="slides-link"]').trigger("click");
    expect(wrapper.emitted("link")).toHaveLength(1);
  });
});

describe("SlideInsertBar whole-frame link (AC-46)", () => {
  it("renders slides-link only when a linkable frame is selected", async () => {
    const wrapper = mount(SlideInsertBar);
    expect(wrapper.find('[data-testid="slides-link"]').exists()).toBe(false);
    await wrapper.setProps({ canLink: true });
    await wrapper.get('[data-testid="slides-link"]').trigger("click");
    expect(wrapper.emitted("link")).toHaveLength(1);
  });
});

function linkController(initial: unknown = null) {
  return {
    linkDialog: ref({ kind: "run", frameId: "f", range: null, initial }),
    slideChoices: computed(() => [
      { number: 1, label: "Agenda" },
      { number: 2, label: "Plan" },
      { number: 3, label: "Results" },
    ]),
    applyLink: vi.fn(() => ""),
    removeLink: vi.fn(),
    closeLinkDialog: vi.fn(),
  };
}

describe("SlideLinkDialog (AC-45/46)", () => {
  it("lists the deck's slides by title and applies a slide link", async () => {
    const controller = linkController();
    const wrapper = mount(SlideLinkDialog, { props: { controller } });
    expect(wrapper.find('[data-testid="slide-link-dialog"]').exists()).toBe(true);
    const select = wrapper.get('[data-testid="link-slide"]');
    expect(select.findAll("option").map((o) => o.text())).toEqual(["No slide", "Agenda", "Plan", "Results"]);
    await select.setValue("3");
    await wrapper.get('[data-testid="link-tooltip"]').setValue("Jump");
    await wrapper.get('[data-testid="link-apply"]').trigger("click");
    expect(controller.applyLink).toHaveBeenCalledWith({ slide: 3, tooltip: "Jump" });
  });

  it("rejects a dangerous URL inline and keeps Apply disabled", async () => {
    const controller = linkController();
    const wrapper = mount(SlideLinkDialog, { props: { controller } });
    await wrapper.get('[data-testid="link-url"]').setValue("javascript:alert(1)");
    expect(wrapper.get('[data-testid="link-error"]').text()).toContain("isn't allowed");
    expect(wrapper.get('[data-testid="link-apply"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="link-url"]').setValue("https://example.com");
    expect(wrapper.find('[data-testid="link-error"]').exists()).toBe(false);
    await wrapper.get('[data-testid="link-apply"]').trigger("click");
    expect(controller.applyLink).toHaveBeenCalledWith({ href: "https://example.com" });
  });

  it("prefills an existing link; Remove and Cancel route to the controller", async () => {
    const controller = linkController({ href: "https://a.example/", tooltip: "A" });
    const wrapper = mount(SlideLinkDialog, { props: { controller } });
    expect((wrapper.get('[data-testid="link-url"]').element as HTMLInputElement).value).toBe("https://a.example/");
    await wrapper.get('[data-testid="link-remove"]').trigger("click");
    expect(controller.removeLink).toHaveBeenCalled();
    await wrapper.get('[data-testid="link-cancel"]').trigger("click");
    expect(controller.closeLinkDialog).toHaveBeenCalled();
  });
});

function panelController() {
  return {
    panelOpen: ref(true),
    paragraphsEnabled: ref(true),
    bodyEnabled: ref(true),
    lineSpacingMode: ref("multiple"),
    lineSpacingValue: ref(1),
    spaceBeforePt: ref(0),
    spaceAfterPt: ref<number | "">(""),
    letterSpacingPt: ref(0),
    autofit: ref("none"),
    anchor: ref("top"),
    insetsCm: ref({ left: 0.254, top: 0.127, right: 0.254, bottom: 0.127 }),
    columns: ref({ count: 1, gapCm: 0 }),
    linkDialog: ref(null),
    setLineSpacingMode: vi.fn(),
    setLineSpacingValue: vi.fn(),
    setSpaceBefore: vi.fn(),
    setSpaceAfter: vi.fn(),
    setLetterSpacingPt: vi.fn(),
    setAutofit: vi.fn(),
    setAnchor: vi.fn(),
    setInset: vi.fn(),
    setColumnCount: vi.fn(),
    setColumnGap: vi.fn(),
  };
}

describe("TextPanel (AC-34..36)", () => {
  it("is a labelled section only while open", async () => {
    const controller = panelController();
    const wrapper = mount(TextPanel, { props: { controller, slideNumber: 2 } });
    const section = wrapper.get('[data-testid="slides-text-panel"]');
    expect(section.element.tagName).toBe("SECTION");
    expect(section.classes()).toContain("text-panel");
    expect(section.attributes("aria-label")).toBe("Text for slide 2");
    controller.panelOpen.value = false;
    await wrapper.vm.$nextTick();
    expect(wrapper.find('[data-testid="slides-text-panel"]').exists()).toBe(false);
  });

  it("labels the autofit and line-spacing options in the spec's words", () => {
    const wrapper = mount(TextPanel, { props: { controller: panelController() } });
    expect(wrapper.get('[data-testid="text-autofit"]').findAll("option").map((o) => o.text())).toEqual([
      "Do not autofit",
      "Shrink text on overflow",
      "Resize shape to fit text",
    ]);
    expect(
      wrapper
        .get('[data-testid="text-line-spacing-mode"]')
        .findAll("option")
        .filter((o) => o.attributes("hidden") === undefined)
        .map((o) => o.text()),
    ).toEqual(["Multiple", "Exactly", "At least"]);
  });

  it("commits each control to its command; a mixed value shows blank", async () => {
    const controller = panelController();
    const wrapper = mount(TextPanel, { props: { controller } });
    expect((wrapper.get('[data-testid="text-space-after"]').element as HTMLInputElement).value).toBe("");
    await wrapper.get('[data-testid="text-line-spacing-mode"]').setValue("exact");
    expect(controller.setLineSpacingMode).toHaveBeenCalledWith("exact");
    await wrapper.get('[data-testid="text-line-spacing"]').setValue("24");
    expect(controller.setLineSpacingValue).toHaveBeenCalledWith(24);
    await wrapper.get('[data-testid="text-space-after"]').setValue("6");
    expect(controller.setSpaceAfter).toHaveBeenCalledWith(6);
    await wrapper.get('[data-testid="text-letter-spacing"]').setValue("0.5");
    expect(controller.setLetterSpacingPt).toHaveBeenCalledWith(0.5);
    await wrapper.get('[data-testid="text-autofit"]').setValue("shrink");
    expect(controller.setAutofit).toHaveBeenCalledWith("shrink");
    await wrapper.get('[data-testid="text-anchor"]').setValue("middle");
    expect(controller.setAnchor).toHaveBeenCalledWith("middle");
    await wrapper.get('[data-testid="text-inset-left"]').setValue("1");
    expect(controller.setInset).toHaveBeenCalledWith("left", 1);
    await wrapper.get('[data-testid="text-columns"]').setValue("2");
    expect(controller.setColumnCount).toHaveBeenCalledWith(2);
    await wrapper.get('[data-testid="text-column-gap"]').setValue("0.5");
    expect(controller.setColumnGap).toHaveBeenCalledWith(0.5);
  });

  it("Enter commits a number field by blurring it and cancels the key's default", () => {
    // The commit refocuses an open text session's host; an uncancelled Enter
    // would then split the paragraph at the session's caret.
    const wrapper = mount(TextPanel, { props: { controller: panelController() }, attachTo: document.body });
    const input = wrapper.get('[data-testid="text-space-after"]').element as HTMLInputElement;
    input.focus();
    expect(document.activeElement).toBe(input);
    const enter = new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true });
    input.dispatchEvent(enter);
    expect(enter.defaultPrevented).toBe(true);
    expect(document.activeElement).not.toBe(input);
    wrapper.unmount();
  });

  it("disables the paragraph and body groups with nothing targeted", () => {
    const controller = panelController();
    controller.paragraphsEnabled.value = false;
    controller.bodyEnabled.value = false;
    const wrapper = mount(TextPanel, { props: { controller } });
    expect(wrapper.findAll("fieldset").every((f) => f.attributes("disabled") !== undefined)).toBe(true);
  });
});
