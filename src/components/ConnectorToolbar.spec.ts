import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import ConnectorToolbar from "./ConnectorToolbar.vue";

// SPEC-077 AC-7/AC-18: the Line colour input shows the owner-resolved hex
// (`lineColorHex`), but every outline emit spreads the MODEL's outline — a
// theme-referenced line colour stays a reference through width/dash edits.
describe("ConnectorToolbar — theme-referenced line colour (SPEC-077 AC-7/AC-18)", () => {
  const themedConnector = {
    id: "c1",
    kind: "straight",
    outline: { color: { theme: "accent4" }, widthPt: 1, dash: "solid" },
  };
  const props = { connector: themedConnector, lineColorHex: "#1f6fb2" };

  it("displays the resolved hex in the colour input", () => {
    const wrapper = mount(ConnectorToolbar, { props });
    const input = wrapper.get('[data-testid="connector-color"]').element as HTMLInputElement;
    expect(input.value).toBe("#1f6fb2");
  });

  it("re-emits the stored reference on a width edit", async () => {
    const wrapper = mount(ConnectorToolbar, { props });
    await wrapper.get('[data-testid="connector-width"]').setValue("3");
    expect(wrapper.emitted("set-outline")).toEqual([
      [{ color: { theme: "accent4" }, widthPt: 3, dash: "solid" }],
    ]);
  });

  it("re-emits the stored reference on a dash edit", async () => {
    const wrapper = mount(ConnectorToolbar, { props });
    await wrapper.get('[data-testid="connector-dash"]').setValue("dash");
    expect(wrapper.emitted("set-outline")).toEqual([
      [{ color: { theme: "accent4" }, widthPt: 1, dash: "dash" }],
    ]);
  });

  it("a colour-input edit stores the picked literal", async () => {
    const wrapper = mount(ConnectorToolbar, { props });
    await wrapper.get('[data-testid="connector-color"]').setValue("#123456");
    expect(wrapper.emitted("set-outline")).toEqual([
      [{ color: "#123456", widthPt: 1, dash: "solid" }],
    ]);
  });
});
