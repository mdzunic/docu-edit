// File → Export (SPEC-006 AC-1, SPEC-007 AC-1/2, SPEC-050, SPEC-034 AC-14):
// the downloaded-copy formats moved off the app bar's Download button into
// the File menu as an Export submenu. These assertions carry over the old
// DownloadMenu component's contract — the four formats, in order, as
// menuitems, each leaving as one `file.export-<format>` command.
import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import DocumentMenuBar from "./DocumentMenuBar.vue";

async function mountWithFileMenuOpen() {
  const wrapper = mount(DocumentMenuBar, {
    props: { editable: true },
    attachTo: document.body,
  });
  await wrapper.get('[data-testid="menu-title-file"]').trigger("click");
  return wrapper;
}

describe("DocumentMenuBar — File → Export", () => {
  it("offers an Export submenu in the File menu that starts closed", async () => {
    const wrapper = await mountWithFileMenuOpen();
    const trigger = wrapper.get('[data-testid="menu-file"] [data-testid="file-menu-export"]');
    expect(trigger.text()).toContain("Export");
    expect(trigger.attributes("role")).toBe("menuitem");
    expect(trigger.attributes("aria-haspopup")).toBe("menu");
    expect(trigger.attributes("aria-expanded")).toBe("false");
    expect(wrapper.find('[data-testid="download-docx"]').exists()).toBe(false);
    wrapper.unmount();
  });

  it("lists Word, OpenDocument, Markdown and PDF in that order as menuitems", async () => {
    const wrapper = await mountWithFileMenuOpen();
    await wrapper.get('[data-testid="file-menu-export"]').trigger("click");
    const flyout = wrapper.get('[data-testid="file-menu-export-flyout"]');
    expect(flyout.attributes("role")).toBe("menu");
    const items = flyout.findAll('[role="menuitem"]');
    expect(items.map((item) => item.text())).toEqual([
      "Word (.docx)",
      "OpenDocument Text (.odt)",
      "Markdown (.md)",
      "PDF (.pdf)",
    ]);
    expect(items.map((item) => item.attributes("data-testid"))).toEqual([
      "download-docx",
      "download-odt",
      "download-md",
      "download-pdf",
    ]);
    wrapper.unmount();
  });

  it.each([
    ["download-docx", "file.export-docx"],
    ["download-odt", "file.export-odt"],
    ["download-md", "file.export-md"],
    ["download-pdf", "file.export-pdf"],
  ])("choosing %s emits the %s command and closes the menus", async (testid, command) => {
    const wrapper = await mountWithFileMenuOpen();
    await wrapper.get('[data-testid="file-menu-export"]').trigger("click");
    await wrapper.get(`[data-testid="${testid}"]`).trigger("click");
    expect(wrapper.emitted("command")).toEqual([[command]]);
    expect(wrapper.get('[data-testid="menu-title-file"]').attributes("aria-expanded")).toBe("false");
    wrapper.unmount();
  });
});
