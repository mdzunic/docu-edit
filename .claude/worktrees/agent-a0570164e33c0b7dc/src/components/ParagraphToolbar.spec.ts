import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import ParagraphToolbar from "./ParagraphToolbar.vue";

describe("ParagraphToolbar (SPEC-014)", () => {
  it("renders the four alignment buttons with shortcut tooltips (AC-1, AC-3)", () => {
    const wrapper = mount(ParagraphToolbar);
    expect(wrapper.get('[data-testid="align-left"]').attributes("title")).toBe("Align left (Ctrl+L)");
    expect(wrapper.get('[data-testid="align-center"]').attributes("title")).toBe("Align centre (Ctrl+E)");
    expect(wrapper.get('[data-testid="align-right"]').attributes("title")).toBe("Align right (Ctrl+R)");
    expect(wrapper.get('[data-testid="align-justify"]').attributes("title")).toBe("Justify (Ctrl+J)");
  });

  it("renders the indent buttons and the line-spacing menu button (AC-1)", () => {
    const wrapper = mount(ParagraphToolbar);
    expect(wrapper.find('[data-testid="indent-increase"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="indent-decrease"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="line-spacing-button"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="paragraph-direction"]').exists()).toBe(true);
  });

  it("presses the button of the current alignment", () => {
    const wrapper = mount(ParagraphToolbar, { props: { alignment: "center" } });
    expect(wrapper.get('[data-testid="align-center"]').attributes("aria-pressed")).toBe("true");
    expect(wrapper.get('[data-testid="align-left"]').attributes("aria-pressed")).toBe("false");
  });

  it("emits align/indent events on click", async () => {
    const wrapper = mount(ParagraphToolbar);
    await wrapper.get('[data-testid="align-right"]').trigger("click");
    expect(wrapper.emitted("align")).toEqual([["right"]]);
    await wrapper.get('[data-testid="indent-increase"]').trigger("click");
    await wrapper.get('[data-testid="indent-decrease"]').trigger("click");
    expect(wrapper.emitted("indent")).toEqual([[1], [-1]]);
  });

  it("offers Single, 1.15, 1.5, Double and Custom spacing… in the menu (AC-11)", async () => {
    const wrapper = mount(ParagraphToolbar);
    await wrapper.get('[data-testid="line-spacing-button"]').trigger("click");
    const menu = wrapper.get('[data-testid="line-spacing-menu"]');
    expect(menu.get('[data-testid="line-spacing-single"]').text()).toBe("Single");
    expect(menu.get('[data-testid="line-spacing-115"]').text()).toBe("1.15");
    expect(menu.get('[data-testid="line-spacing-15"]').text()).toBe("1.5");
    expect(menu.get('[data-testid="line-spacing-double"]').text()).toBe("Double");
    expect(menu.get('[data-testid="line-spacing-custom"]').text()).toBe("Custom spacing…");
  });

  it("emits the preset's multiple and closes the menu", async () => {
    const wrapper = mount(ParagraphToolbar);
    await wrapper.get('[data-testid="line-spacing-button"]').trigger("click");
    await wrapper.get('[data-testid="line-spacing-115"]').trigger("click");
    expect(wrapper.emitted("spacing-preset")).toEqual([[1.15]]);
    expect(wrapper.find('[data-testid="line-spacing-menu"]').exists()).toBe(false);
  });

  it("labels the space-before toggle from the table: absent reads 'Add' (AC-15)", async () => {
    const wrapper = mount(ParagraphToolbar, { props: { spaceBefore: false } });
    await wrapper.get('[data-testid="line-spacing-button"]').trigger("click");
    expect(wrapper.get('[data-testid="space-before-toggle"]').text()).toBe(
      "Add space before paragraph",
    );
  });

  it("labels the space-before toggle 'Remove' for a non-zero value (AC-16)", async () => {
    const wrapper = mount(ParagraphToolbar, { props: { spaceBefore: true } });
    await wrapper.get('[data-testid="line-spacing-button"]').trigger("click");
    expect(wrapper.get('[data-testid="space-before-toggle"]').text()).toBe(
      "Remove space before paragraph",
    );
  });

  it("emits toggle-space-before and custom-spacing from their menu entries", async () => {
    const wrapper = mount(ParagraphToolbar);
    await wrapper.get('[data-testid="line-spacing-button"]').trigger("click");
    await wrapper.get('[data-testid="space-before-toggle"]').trigger("click");
    expect(wrapper.emitted("toggle-space-before")).toHaveLength(1);
    await wrapper.get('[data-testid="line-spacing-button"]').trigger("click");
    await wrapper.get('[data-testid="line-spacing-custom"]').trigger("click");
    expect(wrapper.emitted("custom-spacing")).toHaveLength(1);
  });

  it("presses the direction button for an RTL paragraph and emits the toggle (AC-24)", async () => {
    const wrapper = mount(ParagraphToolbar, { props: { direction: "rtl" } });
    const button = wrapper.get('[data-testid="paragraph-direction"]');
    expect(button.attributes("aria-pressed")).toBe("true");
    await button.trigger("click");
    expect(wrapper.emitted("toggle-direction")).toHaveLength(1);
  });

  it("disables every control when disabled", async () => {
    const wrapper = mount(ParagraphToolbar, { props: { disabled: true } });
    expect(wrapper.get('[data-testid="align-left"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="indent-increase"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="line-spacing-button"]').trigger("click");
    expect(wrapper.find('[data-testid="line-spacing-menu"]').exists()).toBe(false);
  });
});

describe("ParagraphToolbar lists (SPEC-018)", () => {
  it("renders the two list toggles and presses them from the props (AC-1, AC-2)", () => {
    const wrapper = mount(ParagraphToolbar, { props: { bulleted: true } });
    expect(wrapper.get('[data-testid="list-bulleted"]').attributes("aria-pressed")).toBe("true");
    expect(wrapper.get('[data-testid="list-numbered"]').attributes("aria-pressed")).toBe("false");
  });

  it("emits toggle-list with the button's type", async () => {
    const wrapper = mount(ParagraphToolbar);
    await wrapper.get('[data-testid="list-bulleted"]').trigger("click");
    await wrapper.get('[data-testid="list-numbered"]').trigger("click");
    expect(wrapper.emitted("toggle-list")).toEqual([["bullet"], ["number"]]);
  });

  it("offers the three bullet styles in the bullet dropdown (AC-12)", async () => {
    const wrapper = mount(ParagraphToolbar);
    await wrapper.get('[data-testid="list-bulleted-menu-button"]').trigger("click");
    const menu = wrapper.get('[data-testid="bullet-style-menu"]');
    expect(menu.get('[data-testid="bullet-style-disc"]').text()).toBe("•");
    expect(menu.get('[data-testid="bullet-style-circle"]').text()).toBe("◦");
    expect(menu.get('[data-testid="bullet-style-square"]').text()).toBe("▪");
  });

  it("offers the five numbered styles in the numbered dropdown (AC-12)", async () => {
    const wrapper = mount(ParagraphToolbar);
    await wrapper.get('[data-testid="list-numbered-menu-button"]').trigger("click");
    const menu = wrapper.get('[data-testid="number-style-menu"]');
    expect(menu.get('[data-testid="number-style-decimal"]').text()).toBe("1.");
    expect(menu.get('[data-testid="number-style-lowerAlpha"]').text()).toBe("a.");
    expect(menu.get('[data-testid="number-style-upperAlpha"]').text()).toBe("A.");
    expect(menu.get('[data-testid="number-style-lowerRoman"]').text()).toBe("i.");
    expect(menu.get('[data-testid="number-style-upperRoman"]').text()).toBe("I.");
  });

  it("checks the entry of the current marker style", async () => {
    const wrapper = mount(ParagraphToolbar, { props: { listMarker: "square" } });
    await wrapper.get('[data-testid="list-bulleted-menu-button"]').trigger("click");
    expect(wrapper.get('[data-testid="bullet-style-square"]').attributes("aria-checked")).toBe("true");
    expect(wrapper.get('[data-testid="bullet-style-disc"]').attributes("aria-checked")).toBe("false");
  });

  it("emits list-marker with the picked style and closes the menu (AC-13)", async () => {
    const wrapper = mount(ParagraphToolbar);
    await wrapper.get('[data-testid="list-numbered-menu-button"]').trigger("click");
    await wrapper.get('[data-testid="number-style-lowerRoman"]').trigger("click");
    expect(wrapper.emitted("list-marker")).toEqual([["lowerRoman"]]);
    expect(wrapper.find('[data-testid="number-style-menu"]').exists()).toBe(false);
  });

  it("disables the list controls when disabled", async () => {
    const wrapper = mount(ParagraphToolbar, { props: { disabled: true } });
    expect(wrapper.get('[data-testid="list-bulleted"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="list-bulleted-menu-button"]').trigger("click");
    expect(wrapper.find('[data-testid="bullet-style-menu"]').exists()).toBe(false);
  });
});
