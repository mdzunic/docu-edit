// SPEC-111: "What changed" — one row per drop record, grouped by kind with
// counts; nothing at all when nothing was dropped.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import MergeDrops from "./MergeDrops.vue";
import { groupDrops } from "../../lib/pdfMerge";

describe("MergeDrops", () => {
  it("lists every record under its kind, with counts", () => {
    const wrapper = mount(MergeDrops, {
      props: {
        groups: groupDrops([
          { kind: "annotation", source: "a", page: 1, reason: "A link was removed." },
          { kind: "structureTree", source: "b", reason: "Its tags were not carried." },
          { kind: "annotation", source: "a", page: 2, reason: "Another link was removed." },
        ]),
      },
    });
    expect(wrapper.find('[data-testid="merge-drops"]').exists()).toBe(true);
    expect(wrapper.find("summary").text()).toBe("What changed (3)");
    const groups = wrapper.findAll('[data-testid="merge-drop-group"]');
    expect(groups.map((group) => group.attributes("data-kind"))).toEqual(["structureTree", "annotation"]);
    expect(groups.map((group) => group.find('[data-testid="merge-drop-count"]').text())).toEqual(["(1)", "(2)"]);
    expect(wrapper.findAll('[data-testid="merge-drop"]').map((row) => row.text())).toEqual([
      "“b”: Its tags were not carried.",
      "“a”, page 1: A link was removed.",
      "“a”, page 2: Another link was removed.",
    ]);
  });

  it("shows nothing for a merge that dropped nothing", () => {
    const wrapper = mount(MergeDrops, { props: { groups: [] } });
    expect(wrapper.find('[data-testid="merge-drops"]').exists()).toBe(false);
  });
});
