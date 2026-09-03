// SPEC-076 AC-19 — the "Repeat header row" check item in the context menu's
// table section: the same command the Table menu offers, acting on the
// RIGHT-CLICKED row, with its own live check state.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import EditorContextMenu from "./EditorContextMenu.vue";

function mountMenu(props: Record<string, unknown> = {}) {
  return mount(EditorContextMenu, {
    props: { left: 10, top: 10, tableActions: true, ...props },
    attachTo: document.body,
  });
}

describe("EditorContextMenu — Repeat header row (SPEC-076 AC-19)", () => {
  it("appears in the table section as a check item", () => {
    const item = mountMenu().get('[data-testid="context-menu-repeat-header-row"]');
    expect(item.attributes("role")).toBe("menuitemcheckbox");
    expect(item.text()).toContain("Repeat header row");
  });

  it("reflects the right-clicked row's own header state", () => {
    expect(
      mountMenu({ rowIsHeader: false }).get('[data-testid="context-menu-repeat-header-row"]').attributes("aria-checked"),
    ).toBe("false");
    expect(
      mountMenu({ rowIsHeader: true }).get('[data-testid="context-menu-repeat-header-row"]').attributes("aria-checked"),
    ).toBe("true");
  });

  it("emits the table action the view routes to useTableTools", async () => {
    const menu = mountMenu();
    await menu.get('[data-testid="context-menu-repeat-header-row"]').trigger("click");
    expect(menu.emitted("table-action")).toEqual([["repeat-header-row"]]);
  });

  it("is absent outside a table, like the rest of the table section", () => {
    const menu = mountMenu({ tableActions: false });
    expect(menu.find('[data-testid="context-menu-repeat-header-row"]').exists()).toBe(false);
  });
});
