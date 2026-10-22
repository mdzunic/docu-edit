import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import MenuSubmenu from "./MenuSubmenu.vue";

describe("MenuSubmenu", () => {
  it("opens and closes its flyout when the trigger is clicked", async () => {
    const wrapper = mount(MenuSubmenu, {
      props: { label: "Drawing", testid: "parent" },
      slots: { default: '<button data-testid="leaf">Position…</button>' },
    });
    const trigger = wrapper.get('[data-testid="parent"]');
    expect(trigger.attributes("aria-expanded")).toBe("false");
    await trigger.trigger("click");
    expect(trigger.attributes("aria-expanded")).toBe("true");
    await trigger.trigger("click");
    expect(trigger.attributes("aria-expanded")).toBe("false");
  });

  it("keeps its flyout open when a nested submenu trigger (data-menu-keep-open) is clicked", async () => {
    // Format → Drawing → Align/Distribute/Align to are MenuSubmenus nested
    // inside the Drawing MenuSubmenu; clicking a child trigger must open that
    // child's flyout, not collapse the parent (QA regression, AC-3/12/13).
    const wrapper = mount(MenuSubmenu, {
      props: { label: "Drawing", testid: "parent" },
      slots: {
        default:
          '<button data-menu-keep-open data-testid="nested-trigger">Align</button>' +
          '<button data-testid="leaf">Position…</button>',
      },
    });
    const trigger = wrapper.get('[data-testid="parent"]');
    await trigger.trigger("click");
    expect(trigger.attributes("aria-expanded")).toBe("true");

    // A click on the nested submenu trigger bubbles to the parent flyout but
    // must leave it open.
    await wrapper.get('[data-testid="nested-trigger"]').trigger("click");
    expect(trigger.attributes("aria-expanded")).toBe("true");

    // A click on a real leaf command still closes the flyout.
    await wrapper.get('[data-testid="leaf"]').trigger("click");
    expect(trigger.attributes("aria-expanded")).toBe("false");
  });
});
