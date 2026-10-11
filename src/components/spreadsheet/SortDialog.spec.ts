// SPEC-128's Sort dialog: its test ids and labels, the header toggle's
// column names and the orders by column type, the level tools acting on the
// selected level, Add level disabled at 64 levels with its reason, the
// duplicate refusal inline, and OK / Cancel.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import { MAX_SORT_KEYS } from "../../lib/spreadsheet/sort";
import SortDialog from "./SortDialog.vue";

const WITH_HEADER = [
  { col: 0, label: "Region", numeric: false },
  { col: 1, label: "Column B", numeric: false },
  { col: 2, label: "Amount", numeric: true },
];
const WITHOUT_HEADER = [
  { col: 0, label: "Column A", numeric: false },
  { col: 1, label: "Column B", numeric: false },
  { col: 2, label: "Column C", numeric: true },
];

function mountDialog(props: Record<string, unknown> = {}) {
  return mount(SortDialog, { props: { withHeader: WITH_HEADER, withoutHeader: WITHOUT_HEADER, hasHeader: true, keyCol: 2, ...props } });
}

const optionTexts = (select: { findAll(selector: string): { text(): string }[] }) => select.findAll("option").map((option) => option.text());

describe("SortDialog (SPEC-128)", () => {
  it("opens with one level on the active cell's column, the header guess checked and Case sensitive off", () => {
    const wrapper = mountDialog();
    expect(wrapper.find('[data-testid="sort-dialog"]').exists()).toBe(true);
    const levels = wrapper.findAll('[data-testid="sort-level"]');
    expect(levels).toHaveLength(1);
    expect(levels[0]!.text()).toContain("Sort by");
    const column = levels[0]!.get<HTMLSelectElement>('[data-testid="sort-level-column"]');
    expect(column.element.value).toBe("2");
    expect(optionTexts(column)).toEqual(["Region", "Column B", "Amount"]);
    expect(optionTexts(levels[0]!.get('[data-testid="sort-level-order"]'))).toEqual(["Smallest to Largest", "Largest to Smallest"]);
    expect(wrapper.get<HTMLInputElement>('[data-testid="sort-has-headers"]').element.checked).toBe(true);
    expect(wrapper.get('[data-testid="sort-has-headers"]').element.closest("label")?.textContent).toContain("My data has headers");
    expect(wrapper.get<HTMLInputElement>('[data-testid="sort-case-sensitive"]').element.checked).toBe(false);
    expect(wrapper.get('[data-testid="sort-case-sensitive"]').element.closest("label")?.textContent).toContain("Case sensitive");
    expect(wrapper.find('[data-testid="sort-dialog-error"]').exists()).toBe(false);
  });

  it("turning headers off names the columns 'Column A' …; a text column orders A to Z", async () => {
    const wrapper = mountDialog({ hasHeader: false, keyCol: 0 });
    const column = wrapper.get('[data-testid="sort-level-column"]');
    expect(optionTexts(column)).toEqual(["Column A", "Column B", "Column C"]);
    expect(optionTexts(wrapper.get('[data-testid="sort-level-order"]'))).toEqual(["A to Z", "Z to A"]);
    await wrapper.get('[data-testid="sort-has-headers"]').setValue(true);
    expect(optionTexts(wrapper.get('[data-testid="sort-level-column"]'))).toEqual(["Region", "Column B", "Amount"]);
  });

  it("adds, copies, moves and deletes levels — the tools act on the selected level — and OK emits the keys in level order", async () => {
    const wrapper = mountDialog();
    expect(wrapper.get('[data-testid="sort-delete-level"]').attributes("disabled")).toBeDefined();
    await wrapper.get('[data-testid="sort-add-level"]').trigger("click");
    let levels = wrapper.findAll('[data-testid="sort-level"]');
    expect(levels).toHaveLength(2);
    expect(levels[1]!.text()).toContain("Then by");
    expect(levels[1]!.attributes("aria-current")).toBe("true");
    expect(levels[1]!.get<HTMLSelectElement>('[data-testid="sort-level-column"]').element.value).toBe("0");
    await levels[1]!.get('[data-testid="sort-level-order"]').setValue("desc");
    await wrapper.get('[data-testid="sort-move-up"]').trigger("click");
    levels = wrapper.findAll('[data-testid="sort-level"]');
    expect(levels[0]!.get<HTMLSelectElement>('[data-testid="sort-level-column"]').element.value).toBe("0");
    expect(wrapper.get('[data-testid="sort-move-up"]').attributes("disabled")).toBeDefined();
    await levels[1]!.trigger("click");
    await wrapper.get('[data-testid="sort-copy-level"]').trigger("click");
    expect(wrapper.findAll('[data-testid="sort-level"]')).toHaveLength(3);
    await wrapper.get('[data-testid="sort-delete-level"]').trigger("click");
    expect(wrapper.findAll('[data-testid="sort-level"]')).toHaveLength(2);
    await wrapper.get('[data-testid="sort-case-sensitive"]').setValue(true);
    await wrapper.get('[data-testid="sort-ok"]').trigger("click");
    expect(wrapper.emitted("confirm")).toEqual([
      [
        {
          keys: [
            { col: 0, descending: true },
            { col: 2, descending: false },
          ],
          hasHeader: true,
          caseSensitive: true,
        },
      ],
    ]);
  });

  it("refuses a duplicated column inline in sort-dialog-error, and shows the view's refusal there too", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="sort-copy-level"]').trigger("click");
    await wrapper.get('[data-testid="sort-ok"]').trigger("click");
    expect(wrapper.emitted("confirm")).toBeUndefined();
    expect(wrapper.get('[data-testid="sort-dialog-error"]').text()).toBe("Column C is used by more than one level — remove the duplicate.");
    const refused = mountDialog({ error: "Sorting isn't possible here — the range contains merged cells that span rows. Unmerge them first." });
    expect(refused.get('[data-testid="sort-dialog-error"]').attributes("role")).toBe("alert");
    expect(refused.get('[data-testid="sort-dialog-error"]').text()).toContain("merged cells");
  });

  it(`disables Add level and Copy level at ${MAX_SORT_KEYS} levels, stating "Up to 64 levels"`, async () => {
    const wrapper = mountDialog();
    for (let i = 1; i < MAX_SORT_KEYS; i += 1) {
      await wrapper.get('[data-testid="sort-add-level"]').trigger("click");
    }
    expect(wrapper.findAll('[data-testid="sort-level"]')).toHaveLength(MAX_SORT_KEYS);
    const add = wrapper.get('[data-testid="sort-add-level"]');
    expect(add.attributes("disabled")).toBeDefined();
    expect(add.attributes("title")).toBe("Up to 64 levels");
    expect(wrapper.get('[data-testid="sort-copy-level"]').attributes("disabled")).toBeDefined();
    expect(wrapper.get('[data-testid="sort-add-level-reason"]').text()).toBe("Up to 64 levels");
  });

  it("Cancel emits cancel", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="sort-cancel"]').trigger("click");
    expect(wrapper.emitted("cancel")).toHaveLength(1);
  });
});
