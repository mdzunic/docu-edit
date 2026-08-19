// SPEC-057 — a native drawing-group container renders through the paint path:
// its members paint, the container proxy is role="group", and member proxies
// nest as drawing-<i>-member-<j> (AC-39, AC-40). This is also the regression
// guard for the render crash: a drawingGroup block must have a render arm.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import DocBlockEntries from "./DocBlockEntries.vue";
import type { RenderEntry } from "../lib/pageRender";
import type { DrawingGroup, Shape } from "../lib/documentModel";

function shape(overrides: Partial<Shape> = {}): Shape {
  return { type: "shape", geometry: "rect", widthCm: 2, heightCm: 1, wrap: "square", offsetXCm: 0, offsetYCm: 0, fill: "#ff0000", ...overrides };
}

function container(): DrawingGroup {
  return {
    type: "drawingGroup",
    wrap: "square",
    widthCm: 5,
    heightCm: 3,
    offsetXCm: 1,
    offsetYCm: 1,
    members: [
      shape({ offsetXCm: 0, offsetYCm: 0 }),
      // A nested group as the second member, holding two shapes.
      {
        type: "drawingGroup",
        wrap: "square",
        widthCm: 3,
        heightCm: 2,
        offsetXCm: 2,
        offsetYCm: 1,
        members: [shape({ offsetXCm: 0, offsetYCm: 0 }), shape({ offsetXCm: 1, offsetYCm: 0 })],
      },
    ],
  };
}

function groupEntry(block: DrawingGroup, groupSelectedPath: number[] | null = null): RenderEntry {
  return {
    key: "g0",
    memo: [],
    kind: "drawingGroup",
    blockProp: block,
    sectionProp: 0,
    blockIndexProp: 3,
    drawingIndex: 2,
    groupSelectedPath,
    groupModel: { sections: [], media: [] } as unknown as RenderEntry["groupModel"],
    listDefinitions: [],
  } as RenderEntry;
}

describe("DocDrawingGroupBlock rendering (SPEC-057)", () => {
  it("renders the container proxy as role=group with the drawing-<index> testid", () => {
    const w = mount(DocBlockEntries, { props: { entries: [groupEntry(container())] } });
    const proxy = w.get('[data-testid="drawing-2"]');
    expect(proxy.attributes("role")).toBe("group");
  });

  it("nests member proxies as drawing-<i>-member-<j>, recursively", () => {
    const w = mount(DocBlockEntries, { props: { entries: [groupEntry(container())] } });
    expect(w.find('[data-testid="drawing-2-member-0"]').exists()).toBe(true);
    // The nested container is itself a member proxy...
    expect(w.find('[data-testid="drawing-2-member-1"]').exists()).toBe(true);
    // ...and paints its own members one level deeper.
    expect(w.find('[data-testid="drawing-2-member-1-member-0"]').exists()).toBe(true);
    expect(w.find('[data-testid="drawing-2-member-1-member-1"]').exists()).toBe(true);
  });

  it("marks the container selected when its relative path is empty", () => {
    const w = mount(DocBlockEntries, { props: { entries: [groupEntry(container(), [])] } });
    expect(w.find(".doc-group-selected").exists()).toBe(true);
  });

  it("emits group-select with the clicked member's full path and shiftKey", async () => {
    const w = mount(DocBlockEntries, { props: { entries: [groupEntry(container(), [])] } });
    await w.get('[data-testid="drawing-2-member-1-member-0"]').trigger("click");
    const events = w.emitted("group-select");
    expect(events).toBeTruthy();
    const [section, block, payload] = events![0] as [number, number, { path: number[]; shiftKey: boolean }];
    expect(section).toBe(0);
    expect(block).toBe(3);
    expect(payload.path).toEqual([1, 0]);
  });
});
