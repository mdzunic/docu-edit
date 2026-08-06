import { beforeEach, describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import SpecialCharacterDialog from "./SpecialCharacterDialog.vue";
import { RECENT_CHARACTERS_STORAGE_KEY } from "../lib/specialCharacters";

describe("SpecialCharacterDialog", () => {
  beforeEach(() => localStorage.clear());

  it("renders the eight categories with the em dash present (AC-1)", () => {
    const wrapper = mount(SpecialCharacterDialog);
    for (const id of [
      "punctuation",
      "currency",
      "maths",
      "arrows",
      "latin",
      "greek",
      "symbols",
      "emoji",
    ]) {
      expect(
        wrapper.find(`[data-testid="special-character-category-${id}"]`).exists(),
      ).toBe(true);
    }
    expect(wrapper.find('[data-testid="special-char-2014"]').exists()).toBe(true);
  });

  it("emits the clicked character and keeps the dialog open (AC-5, AC-6)", async () => {
    const wrapper = mount(SpecialCharacterDialog);
    await wrapper.get('[data-testid="special-char-2014"]').trigger("click");
    expect(wrapper.emitted("insert")).toEqual([["—"]]);
    // Still mounted, no cancel emitted.
    expect(wrapper.emitted("cancel")).toBeUndefined();
    expect(wrapper.find('[data-testid="special-character-dialog"]').exists()).toBe(true);
  });

  it("filters the grid by name substring (AC-2)", async () => {
    const wrapper = mount(SpecialCharacterDialog);
    await wrapper.get('[data-testid="special-character-search"]').setValue("euro sign");
    expect(wrapper.find('[data-testid="special-char-20AC"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="special-char-2014"]').exists()).toBe(false);
  });

  it("filters the grid by hex code point forms (AC-4)", async () => {
    const wrapper = mount(SpecialCharacterDialog);
    for (const q of ["2014", "u+2014", "0x2014", "#2014"]) {
      await wrapper.get('[data-testid="special-character-search"]').setValue(q);
      expect(wrapper.find('[data-testid="special-char-2014"]').exists()).toBe(true);
      expect(wrapper.findAll(".char-cell").length).toBe(1);
    }
  });

  it("shows a Recent row most-recent-first after inserts, de-duplicated (AC-7, AC-8)", async () => {
    localStorage.setItem(
      RECENT_CHARACTERS_STORAGE_KEY,
      JSON.stringify([0x2013, 0x2014]),
    );
    const wrapper = mount(SpecialCharacterDialog);
    const recent = wrapper.get('[data-testid="special-character-recent"]');
    const cells = recent.findAll(".char-cell");
    expect(cells[0].attributes("data-testid")).toBe("recent-char-2013");
    expect(cells[1].attributes("data-testid")).toBe("recent-char-2014");
    // Clicking the em dash moves it to the front.
    await wrapper.get('[data-testid="special-char-2014"]').trigger("click");
    const after = wrapper
      .get('[data-testid="special-character-recent"]')
      .findAll(".char-cell")
      .map((c) => c.attributes("data-testid"));
    expect(after).toEqual(["recent-char-2014", "recent-char-2013"]);
  });

  it("renders glyphless characters as a labelled dotted box (AC-47)", () => {
    const wrapper = mount(SpecialCharacterDialog);
    const nbsp = wrapper.get('[data-testid="special-char-A0"]');
    expect(nbsp.classes()).toContain("glyphless");
    expect(nbsp.attributes("aria-label")).toBe("No-Break Space");
    expect(nbsp.find(".glyphless-box").exists()).toBe(true);
  });

  it("has the four direct-insert entries (AC-11..14)", () => {
    const wrapper = mount(SpecialCharacterDialog);
    for (const hex of ["A0", "2013", "2014", "2011"]) {
      expect(
        wrapper.find(`[data-testid="special-char-direct-${hex}"]`).exists(),
      ).toBe(true);
    }
  });

  it("emits cancel from the close button", async () => {
    const wrapper = mount(SpecialCharacterDialog);
    await wrapper.get('[data-testid="special-character-close"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
  });
});
