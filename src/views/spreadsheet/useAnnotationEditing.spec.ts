// SPEC-135 AC-11: the sheet on show's annotations are refused where the
// reader left them verbatim — the check loads with an .xlsx workbook only,
// answers per sheet, and is cached per retained part on `opaque`.
import { describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef } from "vue";
import { LINKS_KEPT_AS_READ_TITLE, NOTES_KEPT_AS_READ_TITLE } from "../../lib/spreadsheet/annotations";
import type { Sheet, SpreadsheetFormat, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { useAnnotationEditing } from "./useAnnotationEditing";

function setup(format: SpreadsheetFormat) {
  const model = createBlankWorkbook();
  model.format = format;
  const base = model.sheets[0]!;
  model.sheets = [
    { ...base, id: "s1", name: "Kept", partName: "xl/worksheets/sheet1.xml" },
    { ...base, id: "s2", name: "Modelled", partName: "xl/worksheets/sheet2.xml" },
    { ...base, id: "s3", name: "New" },
  ];
  const wb = shallowRef<SpreadsheetModel | null>(model);
  const index = ref(0);
  const activeSheet = computed(() => wb.value?.sheets[index.value] ?? null);
  const check = vi.fn((_model: Pick<SpreadsheetModel, "opaque">, sheet: Pick<Sheet, "partName">) =>
    sheet.partName === "xl/worksheets/sheet1.xml" ? { notes: false, hyperlinks: false } : { notes: true, hyperlinks: true },
  );
  const load = vi.fn(() => Promise.resolve(check));
  const scope = effectScope();
  const editing = scope.run(() => useAnnotationEditing({ book: { workbook: wb, activeSheet }, load }))!;
  return { wb, index, check, load, editing, stop: () => scope.stop() };
}

const settle = async () => {
  await Promise.resolve();
  await nextTick();
};

describe("useAnnotationEditing (SPEC-135 AC-11)", () => {
  it("refuses notes and links on an .xlsx sheet the check says stay as read; allows them elsewhere; caches per part on opaque", async () => {
    const view = setup("xlsx");
    expect(view.editing.notesReason.value).toBe("");
    await settle();
    expect(view.load).toHaveBeenCalledTimes(1);
    expect(view.editing.notesReason.value).toBe(NOTES_KEPT_AS_READ_TITLE);
    expect(view.editing.linksReason.value).toBe(LINKS_KEPT_AS_READ_TITLE);
    view.index.value = 1;
    expect(view.editing.notesReason.value).toBe("");
    expect(view.editing.linksReason.value).toBe("");
    view.index.value = 2;
    expect(view.editing.notesReason.value).toBe("");
    view.index.value = 0;
    expect(view.editing.notesReason.value).toBe(NOTES_KEPT_AS_READ_TITLE);
    expect(view.check).toHaveBeenCalledTimes(2);
    // The retained-references pass replaces opaque: checked afresh.
    view.wb.value = { ...view.wb.value!, opaque: [] };
    expect(view.editing.notesReason.value).toBe(NOTES_KEPT_AS_READ_TITLE);
    expect(view.check).toHaveBeenCalledTimes(3);
    view.stop();
  });

  it("never loads the check for an .ods or CSV workbook, which keep every annotation they model", async () => {
    for (const format of ["ods", "csv"] as const) {
      const view = setup(format);
      await settle();
      expect(view.load).not.toHaveBeenCalled();
      expect(view.editing.notesReason.value).toBe("");
      expect(view.editing.linksReason.value).toBe("");
      view.stop();
    }
  });
});
