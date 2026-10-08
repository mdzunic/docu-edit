// SPEC-099 — the slide sorter's controller: the open preference, the grid
// rows, every mutation as record(label) → mutate → commit() with the spec's
// labels, block reorder (a no-op records nothing), the clipboard events
// (copy writes both flavours, paste imports or says there are no slides), the
// keyboard, and Reuse slides through the library.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { defineComponent, ref, shallowRef } from "vue";
import { mount } from "@vue/test-utils";
import type { PresentationModel, Slide, TextFrame } from "../../lib/presentation/model";
import { addSection } from "../../lib/presentation/sections";
import { clearSlideClipboardStore, SLIDES_MIME } from "../../lib/presentation/slideClipboard";
import { usePreferences } from "../../composables/usePreferences";
import { describeImport, NO_SLIDES_MESSAGE } from "./sorterFields";
import { useSlideSorter, type SlideSorter } from "./useSlideSorter";

const library = vi.hoisted(() => new Map<string, PresentationModel>());

vi.mock("../../lib/documentStore", () => {
  const decks = library;
  return {
    listDocuments: () =>
      [...decks.values()].map((model) => ({
        id: model.id,
        title: model.title,
        kind: "presentation",
        format: model.format,
        lastModified: 0,
        starred: false,
      })),
    getPresentation: (id: string) => decks.get(id),
  };
});

function slide(title: string): Slide {
  const frame: TextFrame = {
    id: crypto.randomUUID(),
    kind: "text",
    role: "title",
    xCm: 1,
    yCm: 1,
    widthCm: 20,
    heightCm: 3,
    paragraphs: [{ runs: [{ text: title }] }],
  };
  return { id: crypto.randomUUID(), frames: [frame] };
}

function deck(id: string, names: string[]): PresentationModel {
  return {
    kind: "presentation",
    id,
    title: id,
    format: "pptx",
    slides: names.map(slide),
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
  };
}

function titles(model: PresentationModel): string[] {
  return model.slides.map((s) => ((s.frames[0] as TextFrame).paragraphs[0]?.runs[0]?.text ?? ""));
}

let unmount: (() => void) | undefined;

function harness(model: PresentationModel = deck("d1", ["A", "B", "C", "D", "E"])) {
  const deckRef = shallowRef<PresentationModel | null>(model);
  const currentIndex = ref(0);
  const labels: string[] = [];
  const record = vi.fn((label?: string) => labels.push(label ?? ""));
  const commit = vi.fn();
  const reportError = vi.fn();
  let sorter!: SlideSorter;
  const wrapper = mount(
    defineComponent({
      setup() {
        sorter = useSlideSorter({ model: deckRef, currentIndex, record, commit, reportError });
        return () => null;
      },
    }),
  );
  unmount = () => wrapper.unmount();
  return { model, currentIndex, labels, record, commit, sorter };
}

function clipboardEvent(type: string, data: DataTransfer): ClipboardEvent {
  const event = new Event(type, { bubbles: true, cancelable: true }) as ClipboardEvent;
  Object.defineProperty(event, "clipboardData", { value: data });
  return event;
}

/** A minimal DataTransfer (jsdom has none). */
function transfer(initial: Record<string, string> = {}): DataTransfer {
  const store = new Map(Object.entries(initial));
  return {
    setData: (type: string, value: string) => void store.set(type, value),
    getData: (type: string) => store.get(type) ?? "",
  } as unknown as DataTransfer;
}

const { setPreference } = usePreferences();

/** Shows the sorter (the preference is module state, so set it explicitly). */
function openSorter(): void {
  setPreference("presentationSorterOpen", true);
}

beforeEach(() => {
  localStorage.clear();
  setPreference("presentationSorterOpen", false);
  clearSlideClipboardStore();
});
afterEach(() => {
  unmount?.();
  unmount = undefined;
});

describe("open and rows", () => {
  it("toggles through the presentationSorterOpen preference", () => {
    const { sorter } = harness();
    expect(sorter.open.value).toBe(false);
    sorter.toggle();
    expect(sorter.open.value).toBe(true);
    expect(JSON.parse(localStorage.getItem("docu-edit:preferences") ?? "{}").presentationSorterOpen).toBe(true);
  });

  it("is one headerless group without sections, one group per section with them", () => {
    const { sorter, model } = harness();
    expect(sorter.rows.value).toHaveLength(1);
    expect(sorter.rows.value[0].key).toBeNull();
    addSection(model, 0, "One");
    addSection(model, 2, "Two");
    const { sorter: sectioned } = harness(model);
    expect(sectioned.rows.value.map((row) => [row.name, row.cards.map((card) => card.index)])).toEqual([
      ["One", [0, 1]],
      ["Two", [2, 3, 4]],
    ]);
  });
});

describe("mutations are one labelled undo step each (AC-14)", () => {
  it("duplicate, delete and move the selection", () => {
    const { sorter, model, labels, commit } = harness();
    sorter.selectAt(1);
    sorter.selectAt(2, new MouseEvent("click", { shiftKey: true }));
    sorter.duplicate();
    expect(titles(model)).toEqual(["A", "B", "B", "C", "C", "D", "E"]);
    sorter.remove();
    expect(titles(model)).toEqual(["A", "B", "C", "D", "E"]);
    sorter.selectAt(0);
    sorter.selectAt(1, new MouseEvent("click", { shiftKey: true }));
    sorter.reorder(4);
    expect(titles(model)).toEqual(["C", "D", "A", "B", "E"]);
    expect(sorter.selection.indexes.value).toEqual([2, 3]);
    expect(sorter.announcement.value).toBe("2 slides moved to positions 3–4");
    expect(labels).toEqual(["Duplicate slides", "Delete slides", "Move slides"]);
    expect(commit).toHaveBeenCalledTimes(3);
  });

  it("a drop where the block already sits records nothing", () => {
    const { sorter, labels } = harness();
    sorter.selectAt(1);
    sorter.reorder(1);
    sorter.reorder(2);
    expect(labels).toEqual([]);
  });

  it("section commands use their labels; an empty rename is rejected without a step", () => {
    const { sorter, model, labels } = harness();
    sorter.selectAt(2);
    sorter.addSection();
    expect(model.sections?.map((s) => s.name)).toEqual(["Default Section", "New section"]);
    const key = model.sections![1].key;
    expect(sorter.renaming.value).toBe(key);
    expect(sorter.renameSection(key, "   ")).toBe(false);
    expect(sorter.renameSection(key, " Results ")).toBe(true);
    sorter.moveSection(key, -1);
    expect(titles(model)).toEqual(["C", "D", "E", "A", "B"]);
    sorter.removeSection(key, "sectionOnly");
    expect(model.sections?.map((s) => s.name)).toEqual(["Default Section"]);
    expect(labels).toEqual(["New section", "Rename section", "Move section", "Remove section"]);
  });

  it("collapsing a section is view state: no undo step", () => {
    const { sorter, model, labels, commit } = harness();
    addSection(model, 0, "All");
    sorter.toggleCollapsed(model.sections![0].key);
    expect(model.sections![0].collapsed).toBe(true);
    expect(labels).toEqual([]);
    expect(commit).toHaveBeenCalledTimes(1);
    sorter.toggleCollapsed(model.sections![0].key);
    expect(model.sections![0].collapsed).toBeUndefined();
  });
});

describe("the slide clipboard (AC-9)", () => {
  it("copy writes text/plain and SLIDES_MIME; paste lands the slides after the selection", () => {
    const { sorter, model, labels } = harness();
    openSorter();
    sorter.selectAt(1);
    sorter.selectAt(2, new MouseEvent("click", { shiftKey: true }));
    const data = transfer();
    const copy = clipboardEvent("copy", data);
    window.dispatchEvent(copy);
    expect(copy.defaultPrevented).toBe(true);
    expect(data.getData("text/plain")).toBe("B\n\nC");
    expect(JSON.parse(data.getData(SLIDES_MIME)).mimeType).toBe(SLIDES_MIME);

    sorter.selectAt(4);
    window.dispatchEvent(clipboardEvent("paste", data));
    expect(titles(model)).toEqual(["A", "B", "C", "D", "E", "B", "C"]);
    expect(sorter.selection.indexes.value).toEqual([5, 6]);
    expect(labels).toEqual(["Paste slides"]);
    expect(sorter.announcement.value).toBe("2 slides pasted");
  });

  it("cut copies then deletes the selection", () => {
    const { sorter, model, labels } = harness();
    openSorter();
    sorter.selectAt(0);
    const data = transfer();
    window.dispatchEvent(clipboardEvent("cut", data));
    expect(titles(model)).toEqual(["B", "C", "D", "E"]);
    expect(labels).toEqual(["Delete slides"]);
    expect(data.getData("text/plain")).toBe("A");
  });

  it("a paste with no slides payload is a no-op that says so", () => {
    const { sorter, model, labels } = harness();
    openSorter();
    window.dispatchEvent(clipboardEvent("paste", transfer({ "text/plain": "hello" })));
    expect(titles(model)).toHaveLength(5);
    expect(labels).toEqual([]);
    expect(sorter.announcement.value).toBe(NO_SLIDES_MESSAGE);
  });

  it("does nothing while the sorter is closed or the target is a text input", () => {
    const { labels } = harness();
    const data = transfer();
    const event = clipboardEvent("copy", data);
    window.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    openSorter();
    const input = document.createElement("input");
    document.body.append(input);
    const typed = clipboardEvent("copy", data);
    input.dispatchEvent(typed);
    expect(typed.defaultPrevented).toBe(false);
    input.remove();
    expect(labels).toEqual([]);
  });
});

describe("keyboard (AC-4/5)", () => {
  it("Cmd/Ctrl+D duplicates, Delete removes, Ctrl+A selects all — kept from the stage's handlers", () => {
    const { sorter, model } = harness();
    openSorter();
    const later = vi.fn();
    window.addEventListener("keydown", sorter.onGlobalKeydown);
    window.addEventListener("keydown", later);
    try {
      sorter.selectAt(0);
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "d", ctrlKey: true, cancelable: true }));
      expect(titles(model)).toEqual(["A", "A", "B", "C", "D", "E"]);
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", cancelable: true }));
      expect(titles(model)).toHaveLength(5);
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "a", ctrlKey: true, cancelable: true }));
      expect(sorter.selection.indexes.value).toHaveLength(5);
      expect(later).not.toHaveBeenCalled();
      // Undo passes through to the stage's handler.
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true, cancelable: true }));
      expect(later).toHaveBeenCalledTimes(1);
    } finally {
      window.removeEventListener("keydown", sorter.onGlobalKeydown);
      window.removeEventListener("keydown", later);
    }
  });
});

describe("Reuse slides (AC-13)", () => {
  it("lists other presentations, inserts the picked slides after the selection and reports", () => {
    const source = deck("other", ["X", "Y", "Z"]);
    library.set("other", source);
    const legacy = deck("legacy", ["L"]);
    delete (legacy as Partial<PresentationModel>).slideWidthCm;
    library.set("legacy", legacy);
    const { sorter, model, labels } = harness();
    library.set("d1", model);
    sorter.selectAt(1);
    sorter.reuse.show();
    expect(sorter.reuse.decks.value.map((entry) => entry.id)).toEqual(["other"]);
    expect(sorter.reuse.formatting.value).toBe("destination");
    sorter.reuse.toggleSlide(0);
    sorter.reuse.toggleSlide(2);
    sorter.reuse.insert();
    expect(titles(model)).toEqual(["A", "B", "X", "Z", "C", "D", "E"]);
    expect(labels).toEqual(["Reuse slides"]);
    expect(sorter.reuse.open.value).toBe(false);
    expect(sorter.report.value?.title).toBe("Inserted 2 slides");
    expect(sorter.report.value?.lines[0]).toBe("Kept 2 slides: “X”, “Z”.");
    sorter.dismissReport();
    expect(sorter.report.value).toBeNull();
    library.clear();
  });
});

describe("describeImport", () => {
  it("is honest about what could not come along", () => {
    const report = describeImport(
      {
        slides: 1,
        droppedFrames: 2,
        droppedFrameNames: ["Diagram", "Group"],
        lostTransitions: 1,
        lostAnimations: 0,
        mediaReused: 1,
        rescaled: true,
      },
      ["Intro"],
      "Pasted",
    );
    expect(report).toEqual({
      title: "Pasted 1 slide",
      lines: [
        "Kept 1 slide: “Intro”.",
        "Reused 1 image already in this deck.",
        "Resized to fit this deck's slide size.",
        "Left out 2 objects this deck can't hold: Diagram, Group.",
        "1 transition couldn't come along.",
      ],
    });
  });
});
