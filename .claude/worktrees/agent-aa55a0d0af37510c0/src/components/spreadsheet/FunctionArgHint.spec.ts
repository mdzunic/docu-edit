// SPEC-127: the argument hint — a tooltip holding the signature with the
// current argument bold in function-arg-current.
import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import { signatureParts } from "../../lib/spreadsheet/formula/functionHelp";
import { getFunction } from "../../lib/spreadsheet/formula/registry";
import FunctionArgHint from "./FunctionArgHint.vue";

describe("FunctionArgHint", () => {
  it("shows the signature with the current argument bold", () => {
    const wrapper = mount(FunctionArgHint, { props: { parts: signatureParts(getFunction("ROUND")!, 1) } });
    const hint = wrapper.get('[data-testid="function-arg-hint"]');
    expect(hint.attributes("role")).toBe("tooltip");
    expect(hint.attributes("id")).toBe("function-arg-hint");
    expect(hint.text()).toBe("ROUND(number, num_digits)");
    const current = wrapper.get('[data-testid="function-arg-current"]');
    expect(current.element.tagName).toBe("B");
    expect(current.text()).toBe("num_digits");
  });

  it("bolds nothing past maxArgs", () => {
    const wrapper = mount(FunctionArgHint, { props: { parts: signatureParts(getFunction("ROUND")!, 2) } });
    expect(wrapper.find('[data-testid="function-arg-current"]').exists()).toBe(false);
  });
});
