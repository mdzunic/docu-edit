// SPEC-117's read-only formula bar: the Name Box is a read-only input named
// "Name Box" showing the active cell's address (or the drag size), and the
// content is a read-only textbox showing the active cell's text with its line
// breaks preserved.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import FormulaBar from "./FormulaBar.vue";

describe("FormulaBar", () => {
  it("shows the active cell's address in a read-only Name Box", () => {
    const wrapper = mount(FormulaBar, { props: { address: "B3", content: "10" } });
    expect(wrapper.find('[data-testid="formula-bar"]').exists()).toBe(true);
    const nameBox = wrapper.get('[data-testid="name-box"]');
    expect(nameBox.element.tagName).toBe("INPUT");
    expect((nameBox.element as HTMLInputElement).readOnly).toBe(true);
    expect(nameBox.attributes("aria-label")).toBe("Name Box");
    expect((nameBox.element as HTMLInputElement).value).toBe("B3");
  });

  it("shows the drag size while a drag is in progress", async () => {
    const wrapper = mount(FormulaBar, { props: { address: "B3", content: "" } });
    await wrapper.setProps({ address: "3R × 2C" });
    expect((wrapper.get('[data-testid="name-box"]').element as HTMLInputElement).value).toBe("3R × 2C");
  });

  it("shows the content as a read-only textbox with line breaks preserved", () => {
    const wrapper = mount(FormulaBar, { props: { address: "A12", content: "First  line\ttabbed\nSecond\nthird" } });
    const content = wrapper.get('[data-testid="formula-bar-content"]');
    expect(content.attributes("role")).toBe("textbox");
    expect(content.attributes("aria-readonly")).toBe("true");
    expect(content.element.textContent).toBe("First  line\ttabbed\nSecond\nthird");
    // pre-wrap keeps the breaks on screen (scoped CSS is not applied in jsdom,
    // so the class carrying it is what the test can pin).
    expect(content.classes()).toContain("formula-bar-content");
  });

  it("is empty for an empty active cell", () => {
    const wrapper = mount(FormulaBar, { props: { address: "Z99", content: "" } });
    expect(wrapper.get('[data-testid="formula-bar-content"]').element.textContent).toBe("");
  });
});
