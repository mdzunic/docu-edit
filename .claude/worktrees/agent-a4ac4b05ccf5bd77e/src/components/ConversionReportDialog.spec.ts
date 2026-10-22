import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import ConversionReportDialog from "./ConversionReportDialog.vue";
import type { DropRecord } from "../lib/opaque";

const items: DropRecord[] = [
  {
    kind: "chart",
    location: "paragraph 3",
    reason: "Not preserved when saving as OpenDocument Text (.odt): charts cannot be converted between formats",
  },
  {
    kind: "customXml",
    location: 'package part "customXml/item1.xml"',
    reason: "Not preserved when saving as OpenDocument Text (.odt): custom XML parts have no equivalent in the target format",
  },
];

describe("ConversionReportDialog", () => {
  it("renders one row per dropped item, naming its type, location and reason (AC-9, AC-11, AC-12)", () => {
    const wrapper = mount(ConversionReportDialog, { props: { items } });
    const rows = wrapper.findAll('[data-testid="conversion-report-item"]');
    expect(rows).toHaveLength(2);
    expect(rows[0].text()).toContain("Chart");
    expect(rows[0].text()).toContain("paragraph 3");
    expect(rows[0].text()).toContain("charts cannot be converted");
    expect(rows[1].text()).toContain("Custom XML part");
    expect(rows[1].text()).toContain('package part "customXml/item1.xml"');
  });

  it("closes with dontShowAgain=false when the checkbox is untouched", async () => {
    const wrapper = mount(ConversionReportDialog, { props: { items } });
    await wrapper.get('[data-testid="conversion-report-close"]').trigger("click");
    expect(wrapper.emitted("close")).toEqual([[false]]);
  });

  it("closes with dontShowAgain=true when the checkbox is ticked (AC-13)", async () => {
    const wrapper = mount(ConversionReportDialog, { props: { items } });
    await wrapper.get('[data-testid="conversion-report-dont-show"]').setValue(true);
    await wrapper.get('[data-testid="conversion-report-close"]').trigger("click");
    expect(wrapper.emitted("close")).toEqual([[true]]);
  });
});
