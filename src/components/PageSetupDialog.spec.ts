import { describe, it, expect } from "vitest";
import { mount, type VueWrapper } from "@vue/test-utils";
import PageSetupDialog from "./PageSetupDialog.vue";

/** A form control's current value — test-utils types get() by selector as a
 * plain Element, so the read goes through one cast here. */
function fieldValue(wrapper: VueWrapper, selector: string): string {
  return (wrapper.get(selector).element as HTMLInputElement).value;
}

const A4_INITIAL = {
  pageSize: { widthCm: 21, heightCm: 29.7 },
  margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
  measurementUnit: "cm",
};

describe("PageSetupDialog (SPEC-027 AC-1..9)", () => {
  it("offers the standard paper sizes plus Custom (AC-1)", () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    const options = wrapper
      .get('[data-testid="page-setup-paper-size"]')
      .findAll("option")
      .map((option) => option.text());
    expect(options).toEqual(["A4", "Letter", "Legal", "A3", "A5", "Tabloid", "Custom"]);
  });

  it("populates from the current document: A4 portrait, standard margins, cm (AC-8)", () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    expect(fieldValue(wrapper, '[data-testid="page-setup-paper-size"]')).toBe("A4");
    expect(fieldValue(wrapper, '[data-testid="page-setup-orientation"]')).toBe("portrait");
    expect(fieldValue(wrapper, '[data-testid="page-setup-unit"]')).toBe("cm");
    expect(fieldValue(wrapper, '[data-testid="page-setup-margin-top"]')).toBe("2.54");
    expect(fieldValue(wrapper, '[data-testid="page-setup-margin-right"]')).toBe("2.54");
    // Custom width/height fields are hidden for a standard size.
    expect(wrapper.find('[data-testid="page-setup-width"]').exists()).toBe(false);
  });

  it("populates a non-standard custom size as Custom with its own dimensions (AC-8)", () => {
    const wrapper = mount(PageSetupDialog, {
      props: {
        initial: {
          pageSize: { widthCm: 12, heightCm: 18 },
          margins: { topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 },
          measurementUnit: "cm",
        },
      },
    });
    expect(fieldValue(wrapper, '[data-testid="page-setup-paper-size"]')).toBe("Custom");
    expect(fieldValue(wrapper, '[data-testid="page-setup-width"]')).toBe("12");
    expect(fieldValue(wrapper, '[data-testid="page-setup-height"]')).toBe("18");
  });

  it("OK emits paper size, orientation, margins and unit (AC-1..7)", async () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    await wrapper.get('[data-testid="page-setup-orientation"]').setValue("landscape");
    await wrapper.get('[data-testid="page-setup-apply"]').trigger("click");
    const emitted = wrapper.emitted("apply");
    expect(emitted).toHaveLength(1);
    expect(emitted![0][0]).toEqual({
      pageSize: { widthCm: 29.7, heightCm: 21 },
      margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
      measurementUnit: "cm",
      // SPEC-028 AC-20/21: the dialog reports what the edit targets.
      applyTo: "wholeDocument",
    });
  });

  it("switching to landscape swaps a standard size's width and height (AC-2)", async () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    await wrapper.get('[data-testid="page-setup-orientation"]').setValue("landscape");
    await wrapper.get('[data-testid="page-setup-paper-size"]').setValue("Custom");
    // Custom is seeded from the last standard selection (A4 landscape).
    expect(fieldValue(wrapper, '[data-testid="page-setup-width"]')).toBe("29.7");
    expect(fieldValue(wrapper, '[data-testid="page-setup-height"]')).toBe("21");
  });

  it("switching orientation on a Custom size swaps its width and height (AC-2, AC-4)", async () => {
    const wrapper = mount(PageSetupDialog, {
      props: {
        initial: {
          pageSize: { widthCm: 12, heightCm: 18 },
          margins: { topCm: 1, rightCm: 1, bottomCm: 1, leftCm: 1 },
          measurementUnit: "cm",
        },
      },
    });
    await wrapper.get('[data-testid="page-setup-orientation"]').setValue("landscape");
    expect(fieldValue(wrapper, '[data-testid="page-setup-width"]')).toBe("18");
    expect(fieldValue(wrapper, '[data-testid="page-setup-height"]')).toBe("12");
  });

  it("offers cm, mm, in and pt for the unit selector (AC-4)", () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    const options = wrapper
      .get('[data-testid="page-setup-unit"]')
      .findAll("option")
      .map((option) => option.text());
    expect(options).toEqual(["cm", "mm", "in", "pt"]);
  });

  it("changing the unit re-displays the same physical margin, not a reinterpreted number (AC-6)", async () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    await wrapper.get('[data-testid="page-setup-unit"]').setValue("in");
    // 2.54cm is exactly 1in.
    expect(fieldValue(wrapper, '[data-testid="page-setup-margin-top"]')).toBe("1");
    await wrapper.get('[data-testid="page-setup-unit"]').setValue("mm");
    expect(fieldValue(wrapper, '[data-testid="page-setup-margin-top"]')).toBe("25.4");
  });

  it("OK emits the newly chosen measurement-unit preference (AC-7)", async () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    await wrapper.get('[data-testid="page-setup-unit"]').setValue("pt");
    await wrapper.get('[data-testid="page-setup-apply"]').trigger("click");
    expect(wrapper.emitted("apply")![0][0]).toMatchObject({ measurementUnit: "pt" });
  });

  it("Cancel emits cancel and never apply", async () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    await wrapper.get('[data-testid="page-setup-cancel"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
    expect(wrapper.emitted("apply")).toBeUndefined();
  });

  it("rejects margins that leave under 1cm of text area with the exact message (AC-9)", async () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    await wrapper.get('[data-testid="page-setup-margin-left"]').setValue("10.5");
    await wrapper.get('[data-testid="page-setup-margin-right"]').setValue("10");
    expect(wrapper.get('[data-testid="page-setup-margins-error"]').text()).toBe(
      "Margins are too large for this page size",
    );
    expect(wrapper.get('[data-testid="page-setup-apply"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="page-setup-apply"]').trigger("click");
    expect(wrapper.emitted("apply")).toBeUndefined();
  });

  it("rejects a non-positive Custom width or height", async () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    await wrapper.get('[data-testid="page-setup-paper-size"]').setValue("Custom");
    await wrapper.get('[data-testid="page-setup-width"]').setValue("0");
    expect(wrapper.get('[data-testid="page-setup-size-error"]').text()).toContain("greater than 0");
    expect(wrapper.get('[data-testid="page-setup-apply"]').attributes("disabled")).toBeDefined();
  });

  it("Custom size and unit round-trip through OK (AC-4)", async () => {
    const wrapper = mount(PageSetupDialog, { props: { initial: A4_INITIAL } });
    await wrapper.get('[data-testid="page-setup-paper-size"]').setValue("Custom");
    await wrapper.get('[data-testid="page-setup-unit"]').setValue("in");
    await wrapper.get('[data-testid="page-setup-width"]').setValue("8.5");
    await wrapper.get('[data-testid="page-setup-height"]').setValue("11");
    await wrapper.get('[data-testid="page-setup-apply"]').trigger("click");
    const emitted = wrapper.emitted("apply")![0][0] as { pageSize: { widthCm: number; heightCm: number } };
    expect(emitted.pageSize.widthCm).toBeCloseTo(21.59, 5);
    expect(emitted.pageSize.heightCm).toBeCloseTo(27.94, 5);
  });
});
