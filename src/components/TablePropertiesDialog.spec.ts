// SPEC-077 AC-22/AC-23 — the table properties dialog's cell-shading theme
// swatch row. The dialog's own SPEC-023 behaviour is exercised end-to-end by
// e2e/SPEC-023.spec.ts; these component tests pin the theme-colour contract:
// a swatch makes OK carry the ThemeColorRef itself, while the colour input
// only ever displays (and, when hand-edited, writes) a literal hex.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import TablePropertiesDialog from "./TablePropertiesDialog.vue";
import { DEFAULT_THEME } from "../lib/core/theme";

function open(initial: Record<string, unknown> = {}, theme?: object) {
  return mount(TablePropertiesDialog, { props: { initial: { scope: "the whole table", ...initial }, theme } });
}

function applied(wrapper: ReturnType<typeof open>) {
  return wrapper.emitted("apply")![0][0] as { shadingFill: unknown };
}

function swatchFor(wrapper: ReturnType<typeof open>, slot: string) {
  // The first matching swatch is the base row's (no lumMod/lumOff variant).
  return wrapper.findAll('[data-testid="table-props-shading-theme-swatch"]').find(
    (s) => s.attributes("data-theme-slot") === slot,
  )!;
}

describe("TablePropertiesDialog — shading theme swatches (SPEC-077 AC-22/AC-23)", () => {
  it("renders no theme row without a theme (AC-21: existing mounts unchanged)", () => {
    expect(open().find('[data-testid="table-props-shading-theme-row"]').exists()).toBe(false);
  });

  it("a literal shading fill still applies as a literal with the row present (AC-21)", async () => {
    const wrapper = open({}, DEFAULT_THEME);
    await wrapper.get('[data-testid="table-props-shading-enabled"]').setValue(true);
    await wrapper.get('[data-testid="table-props-shading-color"]').setValue("#ffcc00");
    await wrapper.get('[data-testid="table-props-apply"]').trigger("click");
    expect(applied(wrapper).shadingFill).toBe("#ffcc00");
  });

  it("a swatch stores the ThemeColorRef on OK, not the hex it displays (AC-23)", async () => {
    const wrapper = open({}, DEFAULT_THEME);
    await swatchFor(wrapper, "accent4").trigger("click");
    // The swatch switches shading on and the input shows the resolved hex...
    expect((wrapper.get('[data-testid="table-props-shading-enabled"]').element as HTMLInputElement).checked).toBe(true);
    expect((wrapper.get('[data-testid="table-props-shading-color"]').element as HTMLInputElement).value).toBe("#ffc000");
    await wrapper.get('[data-testid="table-props-apply"]').trigger("click");
    // ...but OK carries the reference itself.
    expect(applied(wrapper).shadingFill).toEqual({ theme: "accent4" });
  });

  it("hand-editing the colour input after a swatch reverts to a literal", async () => {
    const wrapper = open({}, DEFAULT_THEME);
    await swatchFor(wrapper, "accent1").trigger("click");
    await wrapper.get('[data-testid="table-props-shading-color"]').setValue("#123456");
    await wrapper.get('[data-testid="table-props-apply"]').trigger("click");
    expect(applied(wrapper).shadingFill).toBe("#123456");
  });

  it("a seeded ThemeColorRef displays resolved but survives an untouched OK as the reference", async () => {
    const wrapper = open({ shadingFill: { theme: "accent6" } }, DEFAULT_THEME);
    expect((wrapper.get('[data-testid="table-props-shading-enabled"]').element as HTMLInputElement).checked).toBe(true);
    expect((wrapper.get('[data-testid="table-props-shading-color"]').element as HTMLInputElement).value).toBe("#70ad47");
    await wrapper.get('[data-testid="table-props-apply"]').trigger("click");
    expect(applied(wrapper).shadingFill).toEqual({ theme: "accent6" });
  });

  it("OK emits a plain object, never the draft's reactive proxy (AC-22: the model must survive the word-count worker's and IndexedDB's structured clone)", async () => {
    const wrapper = open({ shadingFill: { theme: "accent6" } }, DEFAULT_THEME);
    await wrapper.get('[data-testid="table-props-apply"]').trigger("click");
    // toEqual sees through a Proxy; structured clone throws on one.
    expect(() => structuredClone(applied(wrapper))).not.toThrow();
  });
});
