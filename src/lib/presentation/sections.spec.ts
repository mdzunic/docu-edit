// SPEC-099 — the deck-section commands: every command, the "Default Section"
// rule, and normaliseSections' idempotence after each structural command.
import { describe, expect, it } from "vitest";
import type { PresentationModel } from "./model";
import { addSlide, createBlankPresentation, deleteSlide, duplicateSlide, moveSlide, stripRetention } from "./slides";
import {
  addSection,
  DEFAULT_SECTION_NAME,
  MAX_SECTION_NAME_LENGTH,
  moveSection,
  normaliseSections,
  removeSection,
  renameSection,
  sectionKeyForInsertion,
  sectionsOf,
  setSlideSection,
} from "./sections";

/** A blank deck with `n` slides, each with a recognisable title id. */
function deck(n: number): PresentationModel {
  const model = createBlankPresentation();
  for (let i = 1; i < n; i++) {
    addSlide(model, i - 1);
  }
  delete model.structureDirty;
  return model;
}

function keys(model: PresentationModel): (string | undefined)[] {
  return model.slides.map((slide) => slide.sectionKey);
}

function runs(model: PresentationModel): [string, number, number][] {
  return sectionsOf(model).map((run) => [run.section.name, run.startIndex, run.endIndex]);
}

/** Two sections: "Intro" over slides 0..1, "Body" over 2..4. */
function sectioned(): PresentationModel {
  const model = deck(5);
  addSection(model, 0, "Intro");
  addSection(model, 2, "Body");
  delete model.structureDirty;
  return model;
}

function expectNormalised(model: PresentationModel): void {
  const before = structuredClone(model);
  normaliseSections(model);
  expect(model).toEqual(before);
  if (model.sections) {
    expect(model.sections.length).toBeGreaterThan(0);
    if (model.slides.length > 0) {
      expect(model.sections.map((s) => s.key)).toContain(model.slides[0].sectionKey);
    }
  }
}

describe("sectionsOf", () => {
  it("is empty for a deck with no sections", () => {
    expect(sectionsOf(deck(3))).toEqual([]);
  });

  it("reports each section's run, an empty one with start === end", () => {
    const model = sectioned();
    model.sections!.push({ key: "de-sec9", name: "Empty" });
    expect(runs(model)).toEqual([
      ["Intro", 0, 2],
      ["Body", 2, 5],
      ["Empty", 5, 5],
    ]);
  });
});

describe("addSection", () => {
  it("at index 0 of an unsectioned deck covers every slide", () => {
    const model = deck(3);
    const key = addSection(model, 0, "  All  ");
    expect(model.sections).toEqual([{ key, name: "All" }]);
    expect(keys(model)).toEqual([key, key, key]);
    expect(model.structureDirty).toBe(true);
    expectNormalised(model);
  });

  it("at an index > 0 with nothing covering the head first mints a leading Default Section", () => {
    const model = deck(4);
    const key = addSection(model, 2, "Later");
    expect(model.sections!.map((s) => s.name)).toEqual([DEFAULT_SECTION_NAME, "Later"]);
    const lead = model.sections![0].key;
    expect(keys(model)).toEqual([lead, lead, key, key]);
    expect(runs(model)).toEqual([
      [DEFAULT_SECTION_NAME, 0, 2],
      ["Later", 2, 4],
    ]);
    expectNormalised(model);
  });

  it("splits the run holding the index and sits right after its section", () => {
    const model = sectioned();
    addSection(model, 3, "Tail");
    expect(runs(model)).toEqual([
      ["Intro", 0, 2],
      ["Body", 2, 3],
      ["Tail", 3, 5],
    ]);
    expect(model.structureDirty).toBe(true);
    expectNormalised(model);
  });

  it("past the last slide appends an empty section", () => {
    const model = sectioned();
    addSection(model, 5, "Appendix");
    expect(runs(model)).toEqual([
      ["Intro", 0, 2],
      ["Body", 2, 5],
      ["Appendix", 5, 5],
    ]);
  });

  it("mints deterministic de-sec keys and falls back to the default name for a blank one", () => {
    const model = deck(2);
    const first = addSection(model, 0, "   ");
    const second = addSection(model, 1, "Two");
    expect(first).toBe("de-sec1");
    expect(second).toBe("de-sec2");
    expect(model.sections![0].name).toBe(DEFAULT_SECTION_NAME);
  });
});

describe("renameSection", () => {
  it("commits a trimmed name clamped to 120 characters and rejects an empty one", () => {
    const model = sectioned();
    const key = model.sections![1].key;
    expect(renameSection(model, key, "  Results  ")).toBe(true);
    expect(model.sections![1].name).toBe("Results");
    expect(model.structureDirty).toBe(true);
    expect(renameSection(model, key, "x".repeat(200))).toBe(true);
    expect(model.sections![1].name).toHaveLength(MAX_SECTION_NAME_LENGTH);
    expect(renameSection(model, key, "   ")).toBe(false);
    expect(model.sections![1].name).toHaveLength(MAX_SECTION_NAME_LENGTH);
    expect(renameSection(model, "nope", "Name")).toBe(false);
  });
});

describe("moveSection", () => {
  it("moves the section and its whole run, returning the run's new first slide", () => {
    const model = sectioned();
    const ids = model.slides.map((slide) => slide.id);
    const body = model.sections![1].key;
    expect(moveSection(model, body, -1)).toBe(0);
    expect(model.sections!.map((s) => s.name)).toEqual(["Body", "Intro"]);
    expect(model.slides.map((slide) => slide.id)).toEqual([ids[2], ids[3], ids[4], ids[0], ids[1]]);
    expect(runs(model)).toEqual([
      ["Body", 0, 3],
      ["Intro", 3, 5],
    ]);
    expect(model.structureDirty).toBe(true);
    expectNormalised(model);
  });

  it("clamps at either end and reports -1 for an unknown key", () => {
    const model = sectioned();
    expect(moveSection(model, model.sections![0].key, -3)).toBe(0);
    expect(model.structureDirty).toBeUndefined();
    expect(moveSection(model, "nope", 1)).toBe(-1);
  });
});

describe("removeSection", () => {
  it("sectionOnly hands the run to the preceding section", () => {
    const model = sectioned();
    const intro = model.sections![0].key;
    expect(removeSection(model, model.sections![1].key, "sectionOnly")).toBe(2);
    expect(model.sections!.map((s) => s.name)).toEqual(["Intro"]);
    expect(keys(model)).toEqual([intro, intro, intro, intro, intro]);
    expect(model.structureDirty).toBe(true);
    expectNormalised(model);
  });

  it("sectionOnly on the first section hands its run to the next one", () => {
    const model = sectioned();
    const body = model.sections![1].key;
    removeSection(model, model.sections![0].key, "sectionOnly");
    expect(keys(model)).toEqual([body, body, body, body, body]);
  });

  it("removing the only section leaves a deck with none — never []", () => {
    const model = deck(2);
    const key = addSection(model, 0, "Only");
    removeSection(model, key, "sectionOnly");
    expect(model.sections).toBeUndefined();
    expect(keys(model)).toEqual([undefined, undefined]);
  });

  it("withSlides deletes the run's slides too", () => {
    const model = sectioned();
    const kept = model.slides.slice(0, 2).map((slide) => slide.id);
    expect(removeSection(model, model.sections![1].key, "withSlides")).toBe(1);
    expect(model.slides.map((slide) => slide.id)).toEqual(kept);
    expect(model.sections!.map((s) => s.name)).toEqual(["Intro"]);
    expectNormalised(model);
  });
});

describe("setSlideSection", () => {
  it("moves a boundary slide into the neighbouring section", () => {
    const model = sectioned();
    const body = model.sections![1].key;
    setSlideSection(model, 1, body);
    expect(runs(model)).toEqual([
      ["Intro", 0, 1],
      ["Body", 1, 5],
    ]);
    expect(model.structureDirty).toBe(true);
  });
});

describe("normaliseSections", () => {
  it("drops [] and every slide key when the deck has no sections", () => {
    const model = deck(2);
    model.sections = [];
    model.slides[0].sectionKey = "stale";
    normaliseSections(model);
    expect(model.sections).toBeUndefined();
    expect(keys(model)).toEqual([undefined, undefined]);
  });

  it("an orphan key adopts its predecessor's; an orphan slides[0] the first entry", () => {
    const model = sectioned();
    const [intro, body] = model.sections!.map((s) => s.key);
    model.slides[0].sectionKey = "gone";
    model.slides[3].sectionKey = "gone";
    delete model.slides[4].sectionKey;
    normaliseSections(model);
    expect(keys(model)).toEqual([intro, intro, body, body, body]);
  });

  it("re-orders sections to first appearance, an emptied section keeping its slot", () => {
    const model = deck(4);
    model.sections = [
      { key: "a", name: "A" },
      { key: "e", name: "Empty" },
      { key: "b", name: "B" },
    ];
    model.slides.forEach((slide, i) => (slide.sectionKey = i < 2 ? "b" : "a"));
    normaliseSections(model);
    expect(model.sections.map((s) => s.key)).toEqual(["b", "e", "a"]);
    expectNormalised(model);
  });

  it("keeps runs contiguous: a key whose run already closed adopts its predecessor's", () => {
    const model = deck(3);
    model.sections = [
      { key: "a", name: "A" },
      { key: "b", name: "B" },
    ];
    model.slides[0].sectionKey = "a";
    model.slides[1].sectionKey = "b";
    model.slides[2].sectionKey = "a";
    normaliseSections(model);
    expect(keys(model)).toEqual(["a", "b", "b"]);
  });

  it("drops duplicate entries", () => {
    const model = deck(1);
    model.sections = [
      { key: "a", name: "A" },
      { key: "a", name: "Again" },
    ];
    model.slides[0].sectionKey = "a";
    normaliseSections(model);
    expect(model.sections).toEqual([{ key: "a", name: "A" }]);
  });
});

describe("the structural commands end normalised (slides.ts tail)", () => {
  it("addSlide joins the new slide to its predecessor's section", () => {
    const model = sectioned();
    const body = model.sections![1].key;
    const at = addSlide(model, 3);
    expect(model.slides[at].sectionKey).toBe(body);
    expectNormalised(model);
    const first = addSlide(model, -1);
    expect(model.slides[first].sectionKey).toBe(model.sections![0].key);
    expectNormalised(model);
  });

  it("duplicateSlide keeps the copy in the source's section", () => {
    const model = sectioned();
    const at = duplicateSlide(model, 1);
    expect(model.slides[at].sectionKey).toBe(model.sections![0].key);
    expect(runs(model)).toEqual([
      ["Intro", 0, 3],
      ["Body", 3, 6],
    ]);
    expectNormalised(model);
  });

  it("deleteSlide empties a section without dropping it", () => {
    const model = sectioned();
    deleteSlide(model, 0);
    deleteSlide(model, 0);
    expect(runs(model)).toEqual([
      ["Intro", 0, 0],
      ["Body", 0, 3],
    ]);
    expectNormalised(model);
  });

  it("moveSlide keeps a slide beside its own section and otherwise joins the one it lands in", () => {
    const model = sectioned();
    const [intro, body] = model.sections!.map((s) => s.key);
    // Slide 3 (Body) to the very front: it lands before Intro's run → Intro.
    moveSlide(model, 3, 0);
    expect(keys(model)).toEqual([intro, intro, intro, body, body]);
    expectNormalised(model);
    // The last Intro slide one step down: it now sits between Body slides →
    // it joins Body.
    moveSlide(model, 2, 3);
    expect(keys(model)).toEqual([intro, intro, body, body, body]);
    expectNormalised(model);
    // Body's first slide one step up: beside Intro only → joins Intro.
    moveSlide(model, 2, 1);
    expect(keys(model)).toEqual([intro, intro, intro, body, body]);
    expectNormalised(model);
  });

  it("a deck without sections is left exactly as before", () => {
    const model = deck(3);
    moveSlide(model, 0, 2);
    duplicateSlide(model, 0);
    expect(model.sections).toBeUndefined();
    expect(keys(model).every((key) => key === undefined)).toBe(true);
  });
});

describe("sectionKeyForInsertion", () => {
  it("takes the predecessor's section, the first section at index 0", () => {
    const model = sectioned();
    const [intro, body] = model.sections!.map((s) => s.key);
    expect(sectionKeyForInsertion(model, 0)).toBe(intro);
    expect(sectionKeyForInsertion(model, 3)).toBe(body);
    expect(sectionKeyForInsertion(deck(2), 1)).toBeUndefined();
  });
});

describe("stripRetention keeps sections (AC-15)", () => {
  it("carries sections and every sectionKey into the cross-format copy", () => {
    const model = sectioned();
    const stripped = stripRetention(model);
    expect(stripped.sections).toEqual(model.sections);
    expect(keys(stripped)).toEqual(keys(model));
  });

  it("the model stays plain JSON", () => {
    const model = sectioned();
    model.sections![0].collapsed = true;
    expect(JSON.parse(JSON.stringify(model))).toEqual(model);
  });
});
