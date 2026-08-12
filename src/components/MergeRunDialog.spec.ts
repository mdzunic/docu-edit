import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import MergeRunDialog from "./MergeRunDialog.vue";
import { createDocument, type Paragraph, type Table } from "../lib/documentModel";
import { createEachChipRun, createIfChipRun, createMergeChipRun } from "../lib/mergeChips";

function paragraph(...runs: Paragraph["runs"]): Paragraph {
  return { type: "paragraph", runs };
}

beforeEach(() => {
  // jsdom has no createObjectURL/revokeObjectURL (SPEC-046's download path
  // uses the same transient-anchor pattern DocumentView.handleDownload
  // does) — stub them so `run()` can complete without a real browser.
  (URL as unknown as { createObjectURL: () => string }).createObjectURL = vi.fn(() => "blob:mock");
  (URL as unknown as { revokeObjectURL: () => void }).revokeObjectURL = vi.fn();
  vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => {});
});

afterEach(() => {
  vi.restoreAllMocks();
});

function csvFile(text = "client_name,total\nAcme,100\nBeta,\n") {
  return new File([text], "records.csv");
}

function jsonFile(records: unknown[] = [{ client_name: "Acme" }]) {
  return new File([JSON.stringify(records)], "records.json");
}

async function attach(wrapper: ReturnType<typeof mount>, file: File) {
  const input = wrapper.get('[data-testid="merge-run-attach-input"]').element as HTMLInputElement;
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  await wrapper.get('[data-testid="merge-run-attach-input"]').trigger("change");
  // FileReader's load event is a real macrotask in jsdom, not a microtask —
  // flushPromises alone (a single setTimeout(0) tick) is not reliably
  // enough to observe it complete (same fix as TemplateGalleryDialog.spec.ts's
  // import test).
  await flushPromises();
  await flushPromises();
  await new Promise((resolve) => setTimeout(resolve, 20));
}

describe("MergeRunDialog (AC-21..24, AC-38..48)", () => {
  it("attaches a CSV source, shows the record count, and emits attach-source", async () => {
    const model = createDocument("csv-doc");
    model.sections[0].blocks = [paragraph(createMergeChipRun("client_name"))];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, csvFile());

    expect(wrapper.get('[data-testid="merge-run-record-count"]').text()).toContain("2 records");
    expect(wrapper.emitted("attach-source")).toHaveLength(1);
  });

  it("attaches a JSON source too", async () => {
    const model = createDocument("json-doc");
    model.sections[0].blocks = [paragraph(createMergeChipRun("client_name"))];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, jsonFile([{ client_name: "Acme" }, { client_name: "Beta" }]));

    expect(wrapper.get('[data-testid="merge-run-record-count"]').text()).toContain("2 records");
  });

  it("shows a parse error and does not attach on invalid JSON", async () => {
    const model = createDocument("bad-json-doc");
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, new File(["{not json"], "bad.json"));

    expect(wrapper.get('[data-testid="merge-run-source-error"]').text().length).toBeGreaterThan(0);
    expect(wrapper.find('[data-testid="merge-run-record-count"]').exists()).toBe(false);
  });

  it("reports the column-to-field mapping including unmapped fields (AC-24)", async () => {
    const model = createDocument("mapping-doc");
    model.sections[0].blocks = [paragraph(createMergeChipRun("client_name")), paragraph(createMergeChipRun("typo_field"))];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, csvFile());

    const mapping = wrapper.get('[data-testid="merge-run-mapping"]').text();
    expect(mapping).toContain("client_name — mapped");
    expect(mapping).toContain("total — unmapped");
  });

  it("previews the first record with fields filled (AC-23)", async () => {
    const model = createDocument("preview-doc");
    model.sections[0].blocks = [paragraph({ text: "Dear ", formatting: {} }, createMergeChipRun("client_name"), { text: ",", formatting: {} })];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, csvFile());

    expect(wrapper.get('[data-testid="merge-run-preview"]').text()).toContain("Dear Acme,");
  });

  it("blocks the run and lists template errors with locations (AC-35, AC-38)", async () => {
    const model = createDocument("invalid-template-doc");
    model.sections[0].blocks = [paragraph(createIfChipRun("paid"))];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, csvFile());

    expect(wrapper.find('[data-testid="merge-run-template-errors"]').exists()).toBe(true);
    expect(wrapper.get('[data-testid="merge-run-run"]').attributes("disabled")).toBeDefined();
  });

  it("runs a zip merge, downloads it, and reports missing fields without failing (AC-39, AC-40, AC-41, AC-44..46)", async () => {
    const model = createDocument("run-doc");
    model.sections[0].blocks = [paragraph({ text: "Dear ", formatting: {} }, createMergeChipRun("client_name"), { text: ", total ", formatting: {} }, createMergeChipRun("total"))];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    // "total" has no column at all, so it's a genuinely missing field for
    // every record (an empty cell under a present column, by contrast,
    // resolves to "" — present, not missing; SPEC-046's field resolution).
    await attach(wrapper, csvFile("client_name\nAcme\nBeta\n"));
    await wrapper.get('[data-testid="merge-run-run"]').trigger("click");
    await flushPromises();
    await flushPromises();
    await new Promise((resolve) => setTimeout(resolve, 20));

    expect(URL.createObjectURL).toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalled();
    expect(wrapper.get('[data-testid="merge-run-report-summary"]').text()).toContain("2 records merged");
    expect(wrapper.get('[data-testid="merge-run-report-missing"]').text()).toContain('missing "total"');
  });

  it("runs a single-document merge (AC-42, AC-43)", async () => {
    const model = createDocument("run-single-doc");
    model.sections[0].blocks = [paragraph({ text: "Hello ", formatting: {} }, createMergeChipRun("client_name"))];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, csvFile());
    await wrapper.get('[data-testid="merge-run-delivery"]').setValue("single");
    await wrapper.get('[data-testid="merge-run-run"]').trigger("click");
    await flushPromises();
    await flushPromises();
    await new Promise((resolve) => setTimeout(resolve, 20));

    expect(wrapper.get('[data-testid="merge-run-report-summary"]').text()).toContain("2 records merged");
  });

  it("never evaluates merge-field syntax in the data (AC-47, AC-48)", async () => {
    const model = createDocument("injection-doc");
    model.sections[0].blocks = [paragraph(createMergeChipRun("payload"))];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, jsonFile([{ payload: "{{#if x}}{{/if}} <script>alert(1)</script>" }]));
    await wrapper.get('[data-testid="merge-run-run"]').trigger("click");
    await flushPromises();
    await flushPromises();
    await new Promise((resolve) => setTimeout(resolve, 20));

    expect(wrapper.get('[data-testid="merge-run-report-summary"]').text()).toContain("1 record merged");
    expect(wrapper.find('[data-testid="merge-run-error"]').exists()).toBe(false);
  });

  it("reports a non-array {{#each}} target as a data error, never failing the run", async () => {
    const model = createDocument("array-error-doc");
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [{ cells: [{ blocks: [paragraph(createEachChipRun("line_items"))] }, { blocks: [paragraph(createMergeChipRun("sku"))] }] }],
    };
    model.sections[0].blocks = [table];
    const wrapper = mount(MergeRunDialog, { props: { model, initialSource: undefined } });

    await attach(wrapper, csvFile());
    await wrapper.get('[data-testid="merge-run-run"]').trigger("click");
    await flushPromises();
    await flushPromises();
    await new Promise((resolve) => setTimeout(resolve, 20));

    expect(wrapper.get('[data-testid="merge-run-report-array-errors"]').text()).toContain('"line_items" is not a list');
  });
});
