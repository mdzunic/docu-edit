// SPEC-096: the Position & Size dialog — each field commits on blur or Enter
// as one apply (the parent turns it into one undo entry), values read/write
// in the measurement unit, rotation is whole degrees, and the aspect lock
// couples width/height edits.
import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import PositionSizeDialog from "./PositionSizeDialog.vue";
import type { ShapeFrame } from "../../lib/presentation/model";

function frame(): ShapeFrame {
  return {
    id: "f1",
    kind: "shape",
    geometry: "rect",
    xCm: 2,
    yCm: 3,
    widthCm: 8,
    heightCm: 4,
    rotationDeg: 30,
  };
}

function mountDialog(f = frame(), unit = "cm") {
  return mount(PositionSizeDialog, { props: { frame: f, unit } });
}

describe("PositionSizeDialog (SPEC-096)", () => {
  it("seeds every field from the frame in the given unit", () => {
    const wrapper = mountDialog();
    expect((wrapper.get('[data-testid="position-x"]').element as HTMLInputElement).value).toBe("2");
    expect((wrapper.get('[data-testid="position-y"]').element as HTMLInputElement).value).toBe("3");
    expect((wrapper.get('[data-testid="position-width"]').element as HTMLInputElement).value).toBe("8");
    expect((wrapper.get('[data-testid="position-height"]').element as HTMLInputElement).value).toBe("4");
    expect((wrapper.get('[data-testid="position-rotation"]').element as HTMLInputElement).value).toBe("30");
  });

  it("seeds in inches when the unit preference says so", () => {
    const wrapper = mountDialog(frame(), "in");
    const x = (wrapper.get('[data-testid="position-x"]').element as HTMLInputElement).value;
    expect(Number(x)).toBeCloseTo(2 / 2.54, 2);
  });

  it("emits one apply per field commit (Enter), parsed back to cm", async () => {
    const wrapper = mountDialog();
    const x = wrapper.get('[data-testid="position-x"]');
    await x.setValue("5");
    await x.trigger("keydown.enter");
    const applies = wrapper.emitted("apply")!;
    expect(applies).toHaveLength(1);
    expect(applies[0][0]).toMatchObject({ xCm: 5, yCm: 3, widthCm: 8, heightCm: 4, rotationDeg: 30 });
  });

  it("commits on blur too", async () => {
    const wrapper = mountDialog();
    const rot = wrapper.get('[data-testid="position-rotation"]');
    await rot.setValue("90");
    await rot.trigger("blur");
    const applies = wrapper.emitted("apply")!;
    expect(applies[0][0]).toMatchObject({ rotationDeg: 90 });
  });

  it("wraps rotation into whole degrees 0–359", async () => {
    const wrapper = mountDialog();
    const rot = wrapper.get('[data-testid="position-rotation"]');
    await rot.setValue("-90");
    await rot.trigger("keydown.enter");
    expect(wrapper.emitted("apply")![0][0]).toMatchObject({ rotationDeg: 270 });
  });

  it("the aspect lock couples a width edit to the height", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="position-aspect-lock"]').setValue(true);
    const width = wrapper.get('[data-testid="position-width"]');
    await width.setValue("4");
    await width.trigger("keydown.enter");
    // 8×4 ratio 2: width 4 → height 2.
    expect(wrapper.emitted("apply")![0][0]).toMatchObject({ widthCm: 4, heightCm: 2 });
  });

  it("dismisses through the base dialog", async () => {
    const wrapper = mountDialog();
    await wrapper.get('[data-testid="position-close"]').trigger("click");
    expect(wrapper.emitted("dismiss")).toHaveLength(1);
  });
});
