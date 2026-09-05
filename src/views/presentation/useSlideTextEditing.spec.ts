// @vitest-environment jsdom
import { computed, ref } from "vue";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { PresentationModel } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";
import { TYPING_BURST_MS, useSlideTextEditing } from "./useSlideTextEditing";

function setup(paragraphs: unknown = [{ runs: [{ text: "hi" }] }]) {
  const model = ref({
    kind: "presentation",
    id: "d",
    title: "t",
    format: "pptx",
    slideWidthCm: 10,
    slideHeightCm: 5,
    slides: [
      {
        id: "s1",
        frames: [
          { id: "f1", kind: "text", role: "body", xCm: 0, yCm: 0, widthCm: 10, heightCm: 5, paragraphs },
        ],
      },
    ],
  } as unknown as PresentationModel);
  const currentSlide = computed(() => model.value.slides[0]);
  const selection = useFrameSelection({ currentSlide });
  const records: string[] = [];
  const editing = useSlideTextEditing({
    model,
    currentSlide,
    selection,
    record: (label) => records.push(label ?? ""),
    commit: () => {},
  });
  const host = document.createElement("div");
  document.body.appendChild(host);
  editing.startEditing("f1");
  editing.registerHost(host);
  return { editing, host, records, model };
}

function cell(text: string) {
  return { paragraphs: [{ runs: [{ text }] }] };
}
function setupTable() {
  const model = ref({
    kind: "presentation",
    id: "d",
    title: "t",
    format: "pptx",
    slideWidthCm: 10,
    slideHeightCm: 5,
    slides: [
      {
        id: "s1",
        frames: [
          {
            id: "t1",
            kind: "table",
            xCm: 0,
            yCm: 0,
            widthCm: 10,
            heightCm: 5,
            columnWidthsCm: [5, 5],
            rows: [
              { cells: [cell("A"), cell("B")] },
              { cells: [cell("C"), cell("D")] },
            ],
          },
        ],
      },
    ],
  } as unknown as PresentationModel);
  const currentSlide = computed(() => model.value.slides[0]);
  const selection = useFrameSelection({ currentSlide });
  const editing = useSlideTextEditing({
    model,
    currentSlide,
    selection,
    record: () => {},
    commit: () => {},
  });
  const host = document.createElement("div");
  document.body.appendChild(host);
  return { editing, host, model };
}

describe("undo-burst coalescing (AC-28/29)", () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => {
    vi.useRealTimers();
    document.body.innerHTML = "";
  });

  it("coalesces a run of typing into a single undo entry", () => {
    const { editing, host, records } = setup();
    host.innerHTML = "<div>hia</div>";
    editing.onInput();
    host.innerHTML = "<div>hiab</div>";
    editing.onInput();
    expect(records).toEqual(["Typing"]);
  });

  it("a pause ≥ 1s starts a new burst", () => {
    const { editing, host, records } = setup();
    host.innerHTML = "<div>hia</div>";
    editing.onInput();
    vi.advanceTimersByTime(TYPING_BURST_MS + 10);
    host.innerHTML = "<div>hiab</div>";
    editing.onInput();
    expect(records).toEqual(["Typing", "Typing"]);
  });

  it("a formatting command closes the burst and is its own entry", () => {
    const { editing, host, records } = setup();
    host.innerHTML = "<div>hia</div>";
    editing.onInput();
    editing.toggleBold();
    host.innerHTML = "<div>hiab</div>";
    editing.onInput();
    expect(records).toEqual(["Typing", "Bold", "Typing"]);
  });

  it("replaces the slide object on edit so outline-derived views update live (AC-43/44)", () => {
    const { editing, host, model } = setup();
    const before = model.value.slides[0];
    host.innerHTML = "<div>hello there</div>";
    editing.onInput();
    // A fresh slide identity (its frame too) is what makes the filmstrip and the
    // read-only stage re-render off the shallowRef.
    expect(model.value.slides[0]).not.toBe(before);
    expect(model.value.slides[0].frames[0]).not.toBe(before.frames[0]);
    const frame = model.value.slides[0].frames[0];
    expect(frame.kind === "text" && frame.paragraphs[0].runs[0].text).toBe("hello there");
    expect(model.value.slides[0].dirty).toBe(true);
  });

  it("edits a table cell through the same machinery (SPEC-091 AC-6a)", () => {
    const { editing, host, model } = setupTable();
    editing.startEditingCell("t1", 0, 1);
    editing.registerHost(host);
    // startEditing rendered the cell's own paragraph ("B") into the host.
    expect(host.textContent).toBe("B");
    host.innerHTML = "<div>Bravo</div>";
    editing.onInput();
    // Select the cell's text so the run-format command has a range to act on.
    const range = document.createRange();
    range.selectNodeContents(host);
    const sel = document.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    editing.toggleBold();
    const table = model.value.slides[0].frames[0];
    const edited = table.kind === "table" ? table.rows[0].cells[1] : null;
    expect(edited?.paragraphs[0].runs.map((r) => r.text).join("")).toBe("Bravo");
    expect(edited?.paragraphs[0].runs.every((r) => r.bold === true)).toBe(true);
    // The unedited neighbour cell is untouched.
    const other = table.kind === "table" ? table.rows[0].cells[0] : null;
    expect(other?.paragraphs[0].runs[0].text).toBe("A");
  });

  it("Tab moves to the next non-covered cell (SPEC-091 AC-6b)", () => {
    const { editing, host } = setupTable();
    editing.startEditingCell("t1", 0, 0);
    editing.registerHost(host);
    editing.onKeydown(new KeyboardEvent("keydown", { key: "Tab" }));
    expect(editing.editingCell.value).toEqual({ row: 0, col: 1 });
  });

  it("Tab in the last cell appends a row and lands in it (SPEC-091 AC-6b)", () => {
    const { editing, host, model } = setupTable();
    editing.startEditingCell("t1", 1, 1);
    editing.registerHost(host);
    editing.onKeydown(new KeyboardEvent("keydown", { key: "Tab" }));
    const table = model.value.slides[0].frames[0];
    expect(table.kind === "table" && table.rows.length).toBe(3);
    expect(editing.editingCell.value).toEqual({ row: 2, col: 0 });
  });

  it("a single click anchors the active cell without starting a session (SPEC-091 AC-9c)", () => {
    const { editing } = setupTable();
    editing.focusCell("t1", 1, 0);
    expect(editing.activeCell.value).toEqual({ row: 1, col: 0 });
    expect(editing.editingFrameId.value).toBeNull();
  });

  it("keeps bullets on the surviving lines when Enter adds a paragraph (AC-8/11)", () => {
    const { editing, host, model } = setup([
      { runs: [{ text: "A" }], bullet: "disc" },
      { runs: [{ text: "B" }], bullet: "disc" },
      { runs: [{ text: "C" }], bullet: "disc" },
    ]);
    // startEditing rendered the bulleted blocks with data-p markers into `host`.
    // Simulate the browser inserting a fresh paragraph (Enter at end of A + a
    // typed word) — a plain <div> with no marker.
    const inserted = document.createElement("div");
    inserted.textContent = "word";
    host.children[0].after(inserted);
    editing.onInput();
    const frame = model.value.slides[0].frames[0];
    const paras = frame.kind === "text" ? frame.paragraphs : [];
    expect(paras.map((p) => [p.runs.map((r) => r.text).join(""), p.bullet])).toEqual([
      ["A", "disc"],
      ["word", "disc"],
      ["B", "disc"],
      ["C", "disc"], // was silently de-bulleted by the index-based merge
    ]);
  });
});
