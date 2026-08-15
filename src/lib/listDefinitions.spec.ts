// SPEC-032 — multilevel list definitions: presets, marker composition and
// the instance-counter scan.
import { describe, expect, it } from "vitest";
import type { Block, DocumentModel, Paragraph } from "./documentModel";
import type { ListDefinition } from "./core/lists";
import { createDocument } from "./documentModel";
import { listMarkers, markersForBlocks } from "./lists";
import {
  createDefinedListCounters,
  definedMarkerText,
  GALLERY_PRESETS,
  markerForDefinedItem,
  presetLevels,
} from "./core/listDefinitions";

function item(defId: string, listId: string, level: number, startAt?: number): Paragraph {
  const paragraph: Paragraph = { type: "paragraph", runs: [{ text: "x", formatting: {} }] };
  paragraph.list = { type: "number", level, defId, listId };
  if (startAt !== undefined) {
    paragraph.list.startAt = startAt;
  }
  return paragraph;
}

function plain(): Paragraph {
  return { type: "paragraph", runs: [{ text: "p", formatting: {} }] };
}

function decimalDef(defId = "d1"): ListDefinition {
  return { defId, levels: presetLevels("decimal") };
}

describe("gallery presets (AC-1..4)", () => {
  it("offers the four presets in order", () => {
    expect(GALLERY_PRESETS.map((preset) => preset.key)).toEqual([
      "decimal",
      "outline",
      "legal",
      "bulletMixed",
    ]);
  });

  it("decimal: full path with trailing dot per level", () => {
    const def = decimalDef();
    expect(definedMarkerText(def, 1, [1])).toBe("1.");
    expect(definedMarkerText(def, 2, [1, 1])).toBe("1.1.");
    expect(definedMarkerText(def, 3, [2, 3, 1])).toBe("2.3.1.");
  });

  it("outline: I./A./1. single components", () => {
    const def: ListDefinition = { defId: "o", levels: presetLevels("outline") };
    expect(definedMarkerText(def, 1, [2])).toBe("II.");
    expect(definedMarkerText(def, 2, [1, 3])).toBe("C.");
    expect(definedMarkerText(def, 3, [1, 1, 4])).toBe("4.");
  });

  it("legal: full path without trailing dot", () => {
    const def: ListDefinition = { defId: "lg", levels: presetLevels("legal") };
    expect(definedMarkerText(def, 3, [1, 2, 3])).toBe("1.2.3");
  });

  it("bullet-mixed: numbers and bullets alternate by level", () => {
    const def: ListDefinition = { defId: "m", levels: presetLevels("bulletMixed") };
    expect(definedMarkerText(def, 1, [1])).toBe("1.");
    expect(definedMarkerText(def, 2, [1, 1])).toBe("•");
    expect(definedMarkerText(def, 3, [1, 1, 2])).toBe("b.");
  });
});

describe("definedMarkerText (AC-5..8)", () => {
  it("wraps the number in prefix/suffix text", () => {
    const def = decimalDef();
    def.levels[0] = { style: "upperRoman", prefix: "(", suffix: ")" };
    expect(definedMarkerText(def, 1, [5])).toBe("(V)");
  });

  it("includeLevels narrows the shown path ('1.1' vs '1')", () => {
    const def = decimalDef();
    def.levels[1] = { style: "decimal" }; // includeLevels absent -> own level only
    expect(definedMarkerText(def, 2, [3, 2])).toBe("2");
    def.levels[1] = { style: "decimal", includeLevels: 2 };
    expect(definedMarkerText(def, 2, [3, 2])).toBe("3.2");
  });

  it("formats each component in its own level's style", () => {
    const def = decimalDef();
    def.levels[0] = { style: "upperRoman", suffix: "." };
    def.levels[1] = { style: "lowerAlpha", includeLevels: 2, suffix: "." };
    expect(definedMarkerText(def, 2, [4, 2])).toBe("IV.b.");
  });
});

describe("markerForDefinedItem — instance counters", () => {
  it("counts up within one instance and resets deeper levels", () => {
    const def = decimalDef();
    const counters = createDefinedListCounters();
    const at = (level: number) =>
      markerForDefinedItem(def, { type: "number", level, defId: "d1", listId: "a" }, counters);
    expect(at(1)).toBe("1.");
    expect(at(2)).toBe("1.1.");
    expect(at(2)).toBe("1.2.");
    expect(at(1)).toBe("2.");
    expect(at(2)).toBe("2.1.");
  });

  it("keeps independent instances independent", () => {
    const def = decimalDef();
    const counters = createDefinedListCounters();
    const a = markerForDefinedItem(def, { type: "number", level: 1, defId: "d1", listId: "a" }, counters);
    const b = markerForDefinedItem(def, { type: "number", level: 1, defId: "d1", listId: "b" }, counters);
    const a2 = markerForDefinedItem(def, { type: "number", level: 1, defId: "d1", listId: "a" }, counters);
    expect([a, b, a2]).toEqual(["1.", "1.", "2."]);
  });

  it("an explicit startAt restarts the level at that value (AC-16)", () => {
    const def = decimalDef();
    const counters = createDefinedListCounters();
    const at = (startAt?: number) =>
      markerForDefinedItem(
        def,
        { type: "number", level: 1, defId: "d1", listId: "a", startAt },
        counters,
      );
    expect(at()).toBe("1.");
    expect(at(42)).toBe("42.");
    expect(at()).toBe("43.");
  });
});

describe("markersForBlocks with definitions", () => {
  it("continues an instance across an interruption (AC-15's basis)", () => {
    const def = decimalDef();
    const blocks: Block[] = [item("d1", "a", 1), item("d1", "a", 1), plain(), item("d1", "a", 1)];
    expect(markersForBlocks(blocks, [def])).toEqual(["1.", "2.", "", "3."]);
  });

  it("restarts at a fresh instance even when adjacent (AC-14's basis)", () => {
    const def = decimalDef();
    const blocks: Block[] = [item("d1", "a", 1), item("d1", "b", 1), item("d1", "b", 1)];
    expect(markersForBlocks(blocks, [def])).toEqual(["1.", "1.", "2."]);
  });

  it("a defined item interrupts an ad-hoc run without joining it", () => {
    const adHoc: Paragraph = { type: "paragraph", runs: [{ text: "x", formatting: {} }] };
    adHoc.list = { type: "number", level: 1 };
    const adHoc2: Paragraph = { type: "paragraph", runs: [{ text: "y", formatting: {} }] };
    adHoc2.list = { type: "number", level: 1 };
    const blocks: Block[] = [adHoc, item("d1", "a", 1), adHoc2];
    expect(markersForBlocks(blocks, [decimalDef()])).toEqual(["1.", "1.", "1."]);
  });

  it("threads one instance's counters across sections via listMarkers", () => {
    const model: DocumentModel = createDocument("multi-section");
    model.listDefinitions = [decimalDef()];
    model.sections = [
      { blocks: [item("d1", "a", 1)] },
      { blocks: [item("d1", "a", 1)] },
    ];
    expect(listMarkers(model)).toEqual([["1."], ["2."]]);
  });

  it("changing a definition's level changes every item of that list only (AC-17, AC-18)", () => {
    const defA = decimalDef("dA");
    const defB = decimalDef("dB");
    const blocks: Block[] = [item("dA", "a", 1), plain(), item("dB", "b", 1)];
    defA.levels[0] = { style: "upperRoman", suffix: "." };
    expect(markersForBlocks(blocks, [defA, defB])).toEqual(["I.", "", "1."]);
  });
});
