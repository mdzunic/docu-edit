// SPEC-056 — the equation editor panel turns palette clicks and keystrokes
// into engine ops (AC-4/5/6/7/12/13) and edits alt text (AC-20). The editing
// rules themselves are proven in src/lib/math/edit.spec.ts; this pins the thin
// component's wiring.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import EquationEditor from "./EquationEditor.vue";
import { parseLinear } from "../lib/math/linearParse";

function editor(body = parseLinear("")) {
  return mount(EquationEditor, { props: { body, linear: "", altText: "" } });
}

describe("EquationEditor wiring", () => {
  it("renders the whole template palette (AC-5)", () => {
    const wrapper = editor();
    expect(wrapper.find('[data-testid="equation-template-frac"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="equation-template-matrix"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="equation-template-func"]').exists()).toBe(true);
  });

  it("emits a template op when a palette button is clicked (AC-5)", async () => {
    const wrapper = editor();
    await wrapper.find('[data-testid="equation-template-frac"]').trigger("click");
    const ops = wrapper.emitted("op");
    expect(ops?.[0]).toEqual([{ kind: "template", id: "frac" }]);
  });

  it("emits a type op for a printable key (AC-4)", async () => {
    const wrapper = editor();
    await wrapper.find('[data-testid="equation-editor-surface"]').trigger("keydown", { key: "a" });
    expect(wrapper.emitted("op")?.[0]).toEqual([{ kind: "type", text: "a" }]);
  });

  it("maps Tab / Shift+Tab and the arrows / Backspace to structural ops (AC-6/7/12/13)", async () => {
    const wrapper = editor();
    const surface = wrapper.find('[data-testid="equation-editor-surface"]');
    await surface.trigger("keydown", { key: "Tab" });
    await surface.trigger("keydown", { key: "Tab", shiftKey: true });
    await surface.trigger("keydown", { key: "ArrowRight" });
    await surface.trigger("keydown", { key: "Backspace" });
    const kinds = (wrapper.emitted("op") ?? []).map((e) => (e[0] as { kind: string }).kind);
    expect(kinds).toEqual(["tab", "shiftTab", "right", "backspace"]);
  });

  it("commits on Enter and Escape", async () => {
    const wrapper = editor();
    const surface = wrapper.find('[data-testid="equation-editor-surface"]');
    await surface.trigger("keydown", { key: "Enter" });
    await surface.trigger("keydown", { key: "Escape" });
    expect(wrapper.emitted("commit")?.length).toBe(2);
  });

  it("emits alt text edits (AC-20)", async () => {
    const wrapper = editor();
    const input = wrapper.find('[data-testid="equation-alt-text"]');
    (input.element as HTMLInputElement).value = "the mass-energy relation";
    await input.trigger("input");
    expect(wrapper.emitted("altText")?.[0]).toEqual(["the mass-energy relation"]);
  });

  it("renders the live equation from its body (AC-4)", () => {
    const wrapper = editor(parseLinear("\\frac{a}{b}"));
    expect(wrapper.find(".math-frac").exists()).toBe(true);
  });
});
