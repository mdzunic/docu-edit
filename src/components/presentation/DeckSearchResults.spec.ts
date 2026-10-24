// @vitest-environment jsdom
// SPEC-098 — the deck's grouped results list: one group per slide with its
// 1-based chip, one hit button per hit with the match emphasised, the frame's
// name and scope tag; the empty state with the scope note; progress while a
// big deck walks; activation and Escape through the controller.
import { mount } from "@vue/test-utils";
import { computed, ref } from "vue";
import { describe, expect, it, vi } from "vitest";
import DeckSearchResults from "./DeckSearchResults.vue";

function controller(overrides: Record<string, unknown> = {}) {
  return {
    open: ref(true),
    query: ref("Open"),
    error: ref(""),
    indexing: ref(false),
    hitCount: computed(() => 2),
    currentHitIndex: ref(1),
    scopeTitle: ref("Search covers slide text, shapes, tables and speaker notes. Not searched: text inside 1 group and master and layout text."),
    groups: ref([
      { slideIndex: 0, slideNumber: 1, hits: [{ index: 0, before: "", match: "Open", after: " with the Q3 highlight", frameName: "Speaker notes", scopeTag: "Notes" }] },
      { slideIndex: 1, slideNumber: 2, hits: [{ index: 1, before: "Hire, ", match: "Open", after: " Berlin office", frameName: "Placeholder", scopeTag: "" }] },
    ]),
    bar: ref({ visible: true, query: "Open", matchCount: 2, currentIndex: 1, scopeLabel: "Slides, shapes, tables and notes", scopeTitle: "Search covers …" }),
    barEvents: {},
    goToHit: vi.fn(),
    closeFind: vi.fn(),
    ...overrides,
  };
}

describe("DeckSearchResults", () => {
  it("groups hits by slide with a numbered chip and an emphasised snippet", () => {
    const wrapper = mount(DeckSearchResults, { props: { controller: controller() } });
    expect(wrapper.get('[data-testid="find-bar"]').get('[data-testid="find-scope"]').text()).toBe("Slides, shapes, tables and notes");
    const groups = wrapper.findAll('[data-testid="deck-search-group"]');
    expect(groups.map((group) => group.get('[data-testid="deck-search-slide-chip"]').text())).toEqual(["1", "2"]);
    const hits = wrapper.findAll('[data-testid="deck-search-hit"]');
    expect(hits[0].get("mark").text()).toBe("Open");
    expect(hits[0].text()).toContain("Open with the Q3 highlight");
    expect(hits[0].get('[data-testid="deck-search-hit-scope"]').text()).toBe("Notes");
    expect(hits[1].text()).toContain("Placeholder");
    expect(hits[1].classes()).toContain("is-current");
    expect(hits[1].attributes("aria-current")).toBe("true");
    expect(wrapper.find('[data-testid="deck-search-empty"]').exists()).toBe(false);
  });

  it("activates a hit on click, and Escape anywhere in the dock closes the bar", async () => {
    const c = controller();
    const wrapper = mount(DeckSearchResults, { props: { controller: c } });
    await wrapper.findAll('[data-testid="deck-search-hit"]')[1].trigger("click");
    expect(c.goToHit).toHaveBeenCalledWith(1);
    await wrapper.get('[data-testid="deck-search-hit"]').trigger("keydown", { key: "Escape" });
    expect(c.closeFind).toHaveBeenCalledTimes(1);
  });

  it("shows the empty state with the scope note, and progress while walking", async () => {
    const c = controller({ hitCount: computed(() => 0), groups: ref([]) });
    const wrapper = mount(DeckSearchResults, { props: { controller: c } });
    expect(wrapper.get('[data-testid="deck-search-empty"]').text()).toBe("No matches in this presentation");
    expect(wrapper.get('[data-testid="deck-search-scope"]').text()).toContain("Not searched: text inside 1 group");
    c.indexing.value = true;
    await wrapper.vm.$nextTick();
    expect(wrapper.find('[data-testid="deck-search-progress"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="deck-search-empty"]').exists()).toBe(false);
  });

  it("shows no results list for an empty query or an invalid expression, and nothing when closed", async () => {
    const c = controller({ query: ref("") });
    const wrapper = mount(DeckSearchResults, { props: { controller: c } });
    expect(wrapper.find('[data-testid="deck-search-results"]').exists()).toBe(false);
    c.query.value = "(";
    c.error.value = "Invalid regular expression";
    await wrapper.vm.$nextTick();
    expect(wrapper.find('[data-testid="deck-search-results"]').exists()).toBe(false);
    c.open.value = false;
    await wrapper.vm.$nextTick();
    expect(wrapper.html()).toBe("<!--v-if-->");
  });
});
