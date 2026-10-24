// @vitest-environment jsdom
import { computed, ref } from "vue";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { PresentationModel } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";
import { TYPING_BURST_MS, useSlideTextEditing } from "./useSlideTextEditing";
import { setSelectionRange } from "../../components/presentation/slideTextDom";

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

function numbered(...texts: string[]) {
  return texts.map((text) => ({ runs: [{ text }], bullet: "number" }));
}
function paragraphsOf(model: ReturnType<typeof setup>["model"]) {
  const frame = model.value.slides[0].frames[0];
  return frame.kind === "text" ? frame.paragraphs : [];
}
function selectParagraphs(host: HTMLElement, from: number, to: number) {
  setSelectionRange(host, { para: from, offset: 0 }, { para: to, offset: 1 });
}
function tab(shiftKey = false) {
  return new KeyboardEvent("keydown", { key: "Tab", shiftKey, cancelable: true });
}

describe("Tab / Shift+Tab demote and promote (SPEC-097 AC-23)", () => {
  afterEach(() => {
    document.body.innerHTML = "";
  });

  it("Tab demotes every paragraph the selection touches as ONE 'Indent' entry and prevents the default", () => {
    const { editing, host, records, model } = setup(numbered("one", "two", "three"));
    selectParagraphs(host, 1, 2);
    const event = tab();
    editing.onKeydown(event);
    expect(event.defaultPrevented).toBe(true);
    expect(records).toEqual(["Indent"]);
    const paras = paragraphsOf(model);
    expect(paras.map((p) => p.level ?? 0)).toEqual([0, 1, 1]);
    // A demoted "3." moves to its new level's default marker ("a.").
    expect(paras[2].bulletStyle).toBe("lowerAlpha");
    expect(model.value.slides[0].dirty).toBe(true);
  });

  it("never inserts a literal tab character and never leaves the session", () => {
    const { editing, host, model } = setup(numbered("one", "two"));
    selectParagraphs(host, 1, 1);
    editing.onKeydown(tab());
    expect(paragraphsOf(model).map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["one", "two"]);
    expect(host.textContent).not.toContain("\t");
    expect(editing.editingFrameId.value).toBe("f1");
  });

  it("Shift+Tab promotes, restoring the level's default marker", () => {
    const { editing, host, records, model } = setup([
      { runs: [{ text: "one" }], bullet: "number" },
      { runs: [{ text: "two" }], bullet: "number", level: 1, bulletStyle: "lowerAlpha" },
    ]);
    selectParagraphs(host, 1, 1);
    const event = tab(true);
    editing.onKeydown(event);
    expect(event.defaultPrevented).toBe(true);
    expect(records).toEqual(["Indent"]);
    const paras = paragraphsOf(model);
    expect(paras[1].level).toBeUndefined();
    expect(paras[1].bulletStyle).toBeUndefined();
  });

  it("clamps to levels 0–8", () => {
    const { editing, host, model } = setup([
      { runs: [{ text: "top" }] },
      { runs: [{ text: "deep" }], level: 8 },
    ]);
    selectParagraphs(host, 0, 0);
    editing.onKeydown(tab(true));
    expect(paragraphsOf(model)[0].level).toBeUndefined();
    selectParagraphs(host, 1, 1);
    editing.onKeydown(tab());
    expect(paragraphsOf(model)[1].level).toBe(8);
  });

  it("a table-cell session keeps Tab for cell navigation", () => {
    const { editing, host, model } = setupTable();
    editing.startEditingCell("t1", 0, 0);
    editing.registerHost(host);
    editing.onKeydown(tab());
    expect(editing.editingCell.value).toEqual({ row: 0, col: 1 });
    const table = model.value.slides[0].frames[0];
    expect(table.kind === "table" && table.rows[0].cells[0].paragraphs[0].level).toBeUndefined();
  });

  it("the toolbar indent buttons shift levels through the same demote/promote", () => {
    const { editing, host, records, model } = setup(numbered("one", "two"));
    selectParagraphs(host, 1, 1);
    editing.indent(1);
    expect(paragraphsOf(model)[1]).toMatchObject({ level: 1, bulletStyle: "lowerAlpha" });
    editing.indent(-1);
    expect(paragraphsOf(model)[1].level).toBeUndefined();
    expect(records).toEqual(["Indent", "Outdent"]);
  });
});

describe("SPEC-097 hooks for the text-depth composable", () => {
  afterEach(() => {
    document.body.innerHTML = "";
  });

  it("Cmd/Ctrl+K runs the registered link shortcut (and nothing without one)", () => {
    const { editing } = setup();
    const unhandled = new KeyboardEvent("keydown", { key: "k", ctrlKey: true, cancelable: true });
    editing.onKeydown(unhandled);
    expect(unhandled.defaultPrevented).toBe(false);
    const handler = vi.fn();
    editing.setLinkShortcut(handler);
    const event = new KeyboardEvent("keydown", { key: "k", metaKey: true, cancelable: true });
    editing.onKeydown(event);
    expect(event.defaultPrevented).toBe(true);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it("runTextCommand records ONE entry, marks the slide edited and honours a range override", () => {
    const { editing, records, model } = setup([{ runs: [{ text: "hello world" }] }]);
    const ran = editing.runTextCommand(
      "Letter spacing",
      (owner, range) => {
        expect(range).toEqual({ start: { para: 0, offset: 6 }, end: { para: 0, offset: 11 } });
        owner.paragraphs[0].runs[0].letterSpacingPt = 2;
      },
      { start: { para: 0, offset: 6 }, end: { para: 0, offset: 11 } },
    );
    expect(ran).toBe(true);
    expect(records).toEqual(["Letter spacing"]);
    expect(model.value.slides[0].dirty).toBe(true);
  });

  it("runParagraphCommand hands the command the touched paragraph span", () => {
    const { editing, host } = setup([{ runs: [{ text: "a" }] }, { runs: [{ text: "b" }] }, { runs: [{ text: "c" }] }]);
    setSelectionRange(host, { para: 1, offset: 0 }, { para: 2, offset: 1 });
    const spans: unknown[] = [];
    editing.runParagraphCommand("Paragraph spacing", (_owner, range) => spans.push(range));
    expect(spans).toEqual([{ startPara: 1, endPara: 2 }]);
  });

  it("tells the after-edit hook the frame after typing, a command and the end of a session that changed text", () => {
    const { editing, host } = setup();
    const seen: string[] = [];
    editing.setAfterEdit((frameId) => seen.push(frameId));
    host.innerHTML = "<div>hia</div>";
    editing.onInput();
    editing.toggleBold();
    expect(seen).toEqual(["f1", "f1"]);
    editing.stopEditing();
    expect(seen).toEqual(["f1", "f1", "f1"]);
    // A command on the selected frame, outside any session.
    editing.toggleItalic();
    expect(seen).toEqual(["f1", "f1", "f1", "f1"]);
    editing.setAfterEdit(null);
    editing.toggleItalic();
    expect(seen).toHaveLength(4);
  });

  it("a session that changed nothing ends without telling the hook", () => {
    const { editing, model } = setup();
    const afterEdit = vi.fn();
    editing.setAfterEdit(afterEdit);
    editing.stopEditing();
    expect(afterEdit).not.toHaveBeenCalled();
    expect(model.value.slides[0].frames[0]).toMatchObject({ heightCm: 5 });
  });

  it("commandTarget is null with nothing targeted", () => {
    const { editing } = setupTable(); // nothing selected, no session
    expect(editing.commandTarget()).toBeNull();
    expect(editing.runTextCommand("Add link", () => {})).toBe(false);
  });
});

describe("SPEC-098: a find hit opens a session with its range selected", () => {
  it("startEditing(…, range) selects the range once the host registers", () => {
    const { editing, host, model } = setup([{ runs: [{ text: "one" }] }, { runs: [{ text: "Open Berlin office" }] }]);
    editing.stopEditing();
    editing.registerHost(null);
    editing.startEditing("f1", false, undefined, { start: { para: 1, offset: 0 }, end: { para: 1, offset: 4 } });
    editing.registerHost(host);
    expect(document.getSelection()?.toString()).toBe("Open");
    expect(model.value.slides[0].dirty).toBeUndefined();
  });

  it("a session that changed nothing ends without marking the slide edited", () => {
    const { editing, model } = setup();
    editing.stopEditing();
    expect(model.value.slides[0].dirty).toBeUndefined();
  });

  it("a shape hit opens a session on the shape's text, centred as on the stage, and reads typing back into it", () => {
    const { editing, host, model } = setup();
    editing.stopEditing();
    editing.registerHost(null);
    const slide = model.value.slides[0];
    slide.frames.push({ id: "sh", kind: "shape", geometry: "ellipse", xCm: 0, yCm: 0, widthCm: 4, heightCm: 2, paragraphs: [{ runs: [{ text: "Press me" }] }] });
    editing.startEditing("sh", false, undefined, { start: { para: 0, offset: 0 }, end: { para: 0, offset: 5 } });
    expect(editing.editingFrameId.value).toBe("sh");
    editing.registerHost(host);
    expect(document.getSelection()?.toString()).toBe("Press");
    expect(host.querySelector<HTMLElement>(".ce-p")?.style.textAlign).toBe("center");
    expect(editing.toolbarState.value?.align).toBe("center");
    expect(slide.dirty).toBeUndefined();
    host.innerHTML = "<div>Push me</div>";
    editing.onInput();
    editing.stopEditing();
    const shape = model.value.slides[0].frames.find((f) => f.id === "sh");
    expect(shape?.kind === "shape" && shape.paragraphs).toEqual([{ runs: [{ text: "Push me" }] }]);
    expect(model.value.slides[0].dirty).toBe(true);
  });

  it("a shape with no text opens no session", () => {
    const { editing, model } = setup();
    editing.stopEditing();
    model.value.slides[0].frames.push({ id: "sh", kind: "shape", geometry: "rect", xCm: 0, yCm: 0, widthCm: 4, heightCm: 2 });
    editing.startEditing("sh");
    expect(editing.editingFrameId.value).toBeNull();
  });

  it("a session that typed still reads its DOM back on stop", () => {
    const { editing, host, model } = setup();
    host.innerHTML = "<div>typed</div>";
    editing.onInput();
    editing.stopEditing();
    expect(model.value.slides[0].dirty).toBe(true);
    const frame = model.value.slides[0].frames[0];
    expect(frame.kind === "text" && frame.paragraphs[0].runs[0].text).toBe("typed");
  });
});
