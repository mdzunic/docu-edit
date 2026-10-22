// SPEC-025 AC-22, AC-26..28: the insert-bookmark dialog's name validation —
// no spaces, unique against existing names, inline error on conflict.
import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import BookmarkDialog from "./BookmarkDialog.vue";
import BookmarkManagerDialog from "./BookmarkManagerDialog.vue";

function mountDialog(props: Record<string, unknown> = {}) {
  return mount(BookmarkDialog, { props, attachTo: document.body });
}

describe("BookmarkDialog", () => {
  it("saves a valid name (AC-22)", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="bookmark-name-input"]').setValue("chapter_2");
    await wrapper.get('[data-testid="bookmark-save"]').trigger("click");
    expect(wrapper.emitted("save")).toEqual([["chapter_2"]]);
  });

  it("rejects a name with spaces with an inline error (AC-26)", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="bookmark-name-input"]').setValue("has space");
    expect(wrapper.get('[data-testid="bookmark-error"]').text()).toContain("spaces");
    expect(wrapper.get('[data-testid="bookmark-save"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="bookmark-save"]').trigger("click");
    expect(wrapper.emitted("save")).toBeUndefined();
  });

  it("rejects a duplicate name with an inline error (AC-27, AC-28)", async () => {
    const wrapper = mountDialog({ existingNames: ["taken"] });
    await wrapper.get('[data-testid="bookmark-name-input"]').setValue("taken");
    expect(wrapper.get('[data-testid="bookmark-error"]').text()).toContain("already exists");
    expect(wrapper.get('[data-testid="bookmark-save"]').attributes("disabled")).toBeDefined();
  });

  it("disables save for an empty name without showing an error yet", () => {
    const wrapper = mountDialog();
    expect(wrapper.find('[data-testid="bookmark-error"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="bookmark-save"]').attributes("disabled")).toBeDefined();
  });
});

describe("BookmarkManagerDialog", () => {
  it("lists bookmarks with jump and delete actions (AC-23..25)", async () => {
    const wrapper = mount(BookmarkManagerDialog, {
      props: { names: ["alpha", "beta"] },
      attachTo: document.body,
    });
    const rows = wrapper.findAll('[data-testid="bookmark-row"]');
    expect(rows).toHaveLength(2);
    expect(rows[0].text()).toContain("alpha");
    await wrapper.get('[data-testid="bookmark-jump-alpha"]').trigger("click");
    expect(wrapper.emitted("jump")).toEqual([["alpha"]]);
    await wrapper.get('[data-testid="bookmark-delete-beta"]').trigger("click");
    expect(wrapper.emitted("remove")).toEqual([["beta"]]);
  });

  it("shows an empty state when there are no bookmarks", () => {
    const wrapper = mount(BookmarkManagerDialog, { props: { names: [] } });
    expect(wrapper.get('[data-testid="bookmark-manager-empty"]').text()).toContain(
      "No bookmarks",
    );
  });
});
