// @vitest-environment jsdom
import { computed, ref } from "vue";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { PresentationModel } from "../../lib/presentation/model";
import { useFrameSelection } from "./useFrameSelection";
import { TYPING_BURST_MS, useSlideTextEditing } from "./useSlideTextEditing";

function setup() {
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
          { id: "f1", kind: "text", role: "body", xCm: 0, yCm: 0, widthCm: 10, heightCm: 5, paragraphs: [{ runs: [{ text: "hi" }] }] },
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
});
