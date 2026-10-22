// View spec: the table controller's command wiring (SPEC-091 AC-7/9/10).
import { computed, shallowRef } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createBlankPresentation } from "../../lib/presentation/slides";
import { insertTableFrame } from "../../lib/presentation/frames";
import type { PresentationModel, SlideFrame } from "../../lib/presentation/model";
import { useTables } from "./useTables";

function setup() {
  const model = shallowRef<PresentationModel | null>(createBlankPresentation());
  const id = insertTableFrame(model.value!, 0, 2, 2)!;
  const currentSlide = computed(() => model.value?.slides[0] ?? null);
  const currentIndex = shallowRef(0);
  const selectedId = shallowRef<string | null>(id);
  const selectedFrame = computed<SlideFrame | null>(
    () => currentSlide.value?.frames.find((f) => f.id === selectedId.value) ?? null,
  );
  const selectFrame = vi.fn((next: string | null) => {
    selectedId.value = next;
  });
  const controller = useTables({
    model,
    currentSlide,
    currentIndex,
    selectedFrame,
    selectFrame,
    record: vi.fn(),
    commit: vi.fn(),
  });
  return { model, controller, id };
}

describe("useTables", () => {
  it("exposes the selected table and open state", () => {
    const { controller } = setup();
    expect(controller.hasTable.value).toBe(true);
    expect(controller.table.value?.rows).toHaveLength(2);
  });

  it("inserts and deletes rows/columns anchored at the current cell (AC-7)", () => {
    const { controller } = setup();
    controller.setCurrentCell(0, 0);
    controller.insertRowBelow();
    expect(controller.table.value?.rows).toHaveLength(3);
    controller.insertColumnRight();
    expect(controller.table.value?.columnWidthsCm).toHaveLength(3);
    controller.deleteRow();
    expect(controller.table.value?.rows).toHaveLength(2);
  });

  it("deleting the last row deletes the whole frame (AC-7c)", () => {
    const { model, controller, id } = setup();
    controller.deleteRow();
    controller.deleteRow(); // now only one row → frame removed
    expect(model.value!.slides[0].frames.find((f) => f.id === id)).toBeUndefined();
  });

  it("merges and splits cells (AC-9)", () => {
    const { controller } = setup();
    controller.setCurrentCell(0, 0);
    controller.mergeRight();
    expect(controller.table.value?.rows[0].cells[0].colSpan).toBe(2);
    expect(controller.table.value?.rows[0].cells[1].covered).toBe(true);
    controller.splitCurrent();
    expect(controller.table.value?.rows[0].cells[0].colSpan).toBeUndefined();
  });

  it("edits cell fill and table look flags (AC-10)", () => {
    const { controller } = setup();
    controller.setCurrentCell(1, 1);
    controller.setFill({ type: "solid", color: "#00ff00" });
    expect(controller.table.value?.rows[1].cells[1].fill).toEqual({ type: "solid", color: "#00ff00" });
    controller.setBandRows(true);
    expect(controller.table.value?.bandRow).toBe(true);
    controller.setHeaderRow(false);
    expect(controller.table.value?.firstRow).toBeUndefined();
  });
});
