// @vitest-environment jsdom
// SPEC-097: the slide text-depth controller — list, spacing, body and link
// commands each record ONE undo entry with the right label, the autofit scale
// is written only when a mode is picked (or inside a later command's edit),
// never by merely reading/rendering, the Text panel's open state persists,
// and the stage's slide-link-follow event navigates.
import { computed, effectScope, nextTick, ref, type EffectScope } from "vue";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { PresentationModel, ShapeFrame, Slide, SlideFrame, TextFrame } from "../../lib/presentation/model";
import { PREFERENCES_STORAGE_KEY } from "../../lib/preferences";
import { resizedHeightCm } from "../../lib/presentation/textBody";
import { setSelectionRange } from "../../components/presentation/slideTextDom";
import { useFrameSelection } from "./useFrameSelection";
import { useSlideTextEditing } from "./useSlideTextEditing";
import { stageFit } from "./slideTextDepthFields";
import {
  BULLET_CHAR_ERROR,
  BULLET_IMAGE_TYPE_ERROR,
  LINK_URL_ERROR,
  SLIDE_LINK_FOLLOW_EVENT,
  useSlideTextDepth,
  type FrameFit,
} from "./useSlideTextDepth";

function textFrame(id: string, role: TextFrame["role"], texts: string[]): TextFrame {
  return {
    id,
    kind: "text",
    role,
    xCm: 1,
    yCm: 1,
    widthCm: 10,
    heightCm: 4,
    paragraphs: texts.map((text) => ({ runs: [{ text }] })),
  };
}

function deck(): PresentationModel {
  const shape: ShapeFrame = {
    id: "shape",
    kind: "shape",
    geometry: "rect",
    xCm: 12,
    yCm: 1,
    widthCm: 5,
    heightCm: 3,
    paragraphs: [{ runs: [{ text: "Shape text" }] }],
  } as ShapeFrame;
  const image = { id: "img", kind: "image", xCm: 1, yCm: 8, widthCm: 4, heightCm: 3, mediaId: "ppt/media/image4.png" };
  const slides: Slide[] = [
    { id: "s1", frames: [textFrame("title", "title", ["Agenda"]), textFrame("body", "body", ["one", "two", "three"]), shape, image as SlideFrame] },
    { id: "s2", frames: [textFrame("t2", "title", ["Second"])] },
    { id: "s3", frames: [textFrame("t3", "title", ["Results"])] },
  ];
  return {
    kind: "presentation",
    id: "d",
    title: "Deck",
    format: "pptx",
    slideWidthCm: 25.4,
    slideHeightCm: 14.288,
    slides,
    media: [{ mediaId: "ppt/media/image4.png", contentType: "image/png", data: "" }],
  } as PresentationModel;
}

/** jsdom's File has no arrayBuffer(); the composable reads only `.type` and
 * `.arrayBuffer()` (useFrameInsertion.spec.ts's stub). */
function imageFile(type: string, bytes: number[]): File {
  return { type, arrayBuffer: async () => new Uint8Array(bytes).buffer } as unknown as File;
}

let scope: EffectScope | null = null;

function harness(fit: FrameFit | null = null) {
  const model = ref<PresentationModel | null>(deck());
  const currentIndex = ref(0);
  const currentSlide = computed(() => model.value?.slides[currentIndex.value] ?? null);
  const records: string[] = [];
  const record = (label?: string) => records.push(label ?? "");
  const commit = vi.fn();
  const measured = { fit };
  const measureFit = vi.fn(() => measured.fit);
  scope = effectScope();
  const built = scope.run(() => {
    const selection = useFrameSelection({ currentSlide });
    const editing = useSlideTextEditing({ model, currentSlide, selection, record, commit });
    const depth = useSlideTextDepth({ model, currentSlide, currentIndex, record, commit, selection, editing, measureFit });
    return { selection, editing, depth };
  });
  if (!built) throw new Error("scope did not run");
  const frame = <K extends SlideFrame>(id: string) => model.value?.slides[currentIndex.value].frames.find((f) => f.id === id) as K;
  return { model, currentIndex, records, commit, measured, measureFit, frame, ...built };
}

/** Starts a session on the body frame with characters [from, to) of `para` selected. */
function editBody(h: ReturnType<typeof harness>, para: number, from: number, to: number): HTMLElement {
  const host = document.createElement("div");
  document.body.appendChild(host);
  h.editing.startEditing("body");
  h.editing.registerHost(host);
  setSelectionRange(host, { para, offset: from }, { para, offset: to });
  return host;
}

beforeEach(() => {
  localStorage.clear();
});
afterEach(() => {
  scope?.stop();
  scope = null;
  document.body.innerHTML = "";
});

describe("the Text panel's open state (presentationTextOpen)", () => {
  it("starts closed and persists each toggle in the preferences record", () => {
    const { depth } = harness();
    expect(depth.panelOpen.value).toBe(false);
    depth.togglePanel();
    expect(depth.panelOpen.value).toBe(true);
    expect(JSON.parse(localStorage.getItem(PREFERENCES_STORAGE_KEY) ?? "{}").presentationTextOpen).toBe(true);
    depth.togglePanel();
    expect(depth.panelOpen.value).toBe(false);
    expect(JSON.parse(localStorage.getItem(PREFERENCES_STORAGE_KEY) ?? "{}").presentationTextOpen).toBe(false);
  });
});

describe("list commands (AC-24..28)", () => {
  it("the numbering toggle numbers the whole selected frame as one 'Numbering' entry, then clears it", () => {
    const h = harness();
    h.selection.select("body");
    h.depth.toggleNumbering();
    expect(h.records).toEqual(["Numbering"]);
    expect(h.frame<TextFrame>("body").paragraphs.map((p) => p.bullet)).toEqual(["number", "number", "number"]);
    expect(h.model.value?.slides[0].dirty).toBe(true);
    expect(h.depth.formatState.value?.numbered).toBe(true);
    h.depth.toggleNumbering();
    expect(h.records).toEqual(["Numbering", "Numbering"]);
    expect(h.frame<TextFrame>("body").paragraphs.every((p) => p.bullet === undefined)).toBe(true);
  });

  it("a marker style is one 'List style' entry and decides the list type", () => {
    const h = harness();
    h.selection.select("body");
    h.depth.applyListStyle("upperRoman");
    expect(h.records).toEqual(["List style"]);
    expect(h.frame<TextFrame>("body").paragraphs[0]).toMatchObject({ bullet: "number", bulletStyle: "upperRoman" });
    h.depth.applyListStyle("square");
    expect(h.frame<TextFrame>("body").paragraphs[0]).toMatchObject({ bullet: "disc", bulletStyle: "square" });
  });

  it("a custom bullet glyph must be one grapheme", () => {
    const h = harness();
    h.selection.select("body");
    h.depth.setBulletChar("ab");
    expect(h.depth.listError.value).toBe(BULLET_CHAR_ERROR);
    expect(h.records).toEqual([]);
    h.depth.setBulletChar("★");
    expect(h.depth.listError.value).toBe("");
    expect(h.records).toEqual(["List style"]);
    expect(h.frame<TextFrame>("body").paragraphs[2]).toMatchObject({ bullet: "disc", bulletChar: "★" });
  });

  it("a picture bullet mints its image by SPEC-082's naming rule inside the one entry", async () => {
    const h = harness();
    h.selection.select("body");
    await h.depth.setBulletImage(imageFile("image/png", [137, 80, 78, 71]));
    expect(h.records).toEqual(["List style"]);
    const media = h.model.value?.media ?? [];
    expect(media.map((m) => m.mediaId)).toEqual(["ppt/media/image4.png", "ppt/media/image5.png"]);
    expect(media[1].contentType).toBe("image/png");
    expect(h.frame<TextFrame>("body").paragraphs.map((p) => p.bulletImageId)).toEqual([
      "ppt/media/image5.png",
      "ppt/media/image5.png",
      "ppt/media/image5.png",
    ]);
  });

  it("rejects a non-image picture bullet without an undo entry", async () => {
    const h = harness();
    h.selection.select("body");
    await h.depth.setBulletImage(imageFile("text/plain", [120]));
    expect(h.depth.listError.value).toBe(BULLET_IMAGE_TYPE_ERROR);
    expect(h.records).toEqual([]);
    expect(h.model.value?.media).toHaveLength(1);
  });

  it("start-at is a 'Numbering' entry; indent and hanging are 'Indent' entries", () => {
    const h = harness();
    h.selection.select("body");
    h.depth.toggleNumbering();
    h.depth.setStartAt(3);
    h.depth.setListIndent(2);
    h.depth.setListHanging(0.8);
    expect(h.records).toEqual(["Numbering", "Numbering", "Indent", "Indent"]);
    expect(h.frame<TextFrame>("body").paragraphs[0]).toMatchObject({ startAt: 3, indentCm: 2, hangingCm: 0.8 });
    expect(h.depth.startAt.value).toBe(3);
    expect(h.depth.listGeometry.value).toEqual({ indentCm: 2, hangingCm: 0.8 });
  });

  it("list geometry reads the derived defaults and blanks when mixed", () => {
    const h = harness();
    h.selection.select("body");
    h.depth.toggleNumbering();
    expect(h.depth.listGeometry.value).toEqual({ indentCm: 1.27, hangingCm: 0.635 });
    h.frame<TextFrame>("body").paragraphs[0].indentCm = 3;
    expect(h.depth.listGeometry.value).toEqual({ indentCm: "", hangingCm: 0.635 });
  });
});

describe("spacing commands (AC-35/36)", () => {
  it("line spacing modes switch as 'Line spacing' entries and the panel reads them back", () => {
    const h = harness();
    h.selection.select("body");
    expect(h.depth.lineSpacingMode.value).toBe("multiple");
    expect(h.depth.lineSpacingValue.value).toBe(1);
    h.depth.setLineSpacingMode("exact");
    expect(h.frame<TextFrame>("body").paragraphs[0]).toMatchObject({ lineSpacingPt: 22, lineSpacingRule: "exact" });
    h.depth.setLineSpacingValue(24);
    expect(h.depth.lineSpacingMode.value).toBe("exact");
    expect(h.depth.lineSpacingValue.value).toBe(24);
    h.depth.setLineSpacingMode("atLeast");
    expect(h.frame<TextFrame>("body").paragraphs[0].lineSpacingRule).toBeUndefined();
    expect(h.frame<TextFrame>("body").paragraphs[0].lineSpacingPt).toBe(24);
    h.depth.setLineSpacingMode("multiple");
    h.depth.setLineSpacingValue(1.5);
    expect(h.frame<TextFrame>("body").paragraphs[0]).toMatchObject({ lineSpacing: 1.5 });
    expect(h.frame<TextFrame>("body").paragraphs[0].lineSpacingPt).toBeUndefined();
    expect(h.records).toEqual(["Line spacing", "Line spacing", "Line spacing", "Line spacing", "Line spacing"]);
  });

  it("space before/after are 'Paragraph spacing' entries over the touched paragraphs only", () => {
    const h = harness();
    editBody(h, 1, 0, 2);
    h.depth.setSpaceAfter(6);
    h.depth.setSpaceBefore(3);
    expect(h.records).toEqual(["Paragraph spacing", "Paragraph spacing"]);
    const paras = h.frame<TextFrame>("body").paragraphs;
    expect(paras.map((p) => p.spaceAfterPt)).toEqual([undefined, 6, undefined]);
    expect(paras[1].spaceBeforePt).toBe(3);
  });

  it("letter spacing is one 'Letter spacing' entry; a caret spaces its whole paragraph", () => {
    const h = harness();
    editBody(h, 0, 1, 1);
    h.depth.setLetterSpacingPt(0.5);
    expect(h.records).toEqual(["Letter spacing"]);
    expect(h.frame<TextFrame>("body").paragraphs[0].runs).toEqual([{ text: "one", letterSpacingPt: 0.5 }]);
    expect(h.frame<TextFrame>("body").paragraphs[1].runs).toEqual([{ text: "two" }]);
  });

  it("paragraph controls are disabled with nothing targeted", () => {
    const h = harness();
    expect(h.depth.paragraphsEnabled.value).toBe(false);
    h.depth.setSpaceAfter(6);
    expect(h.records).toEqual([]);
  });
});

describe("body commands (AC-36)", () => {
  it("anchor, insets and columns are their own entries on a text frame", () => {
    const h = harness();
    h.selection.select("body");
    h.depth.setAnchor("middle");
    h.depth.setInset("left", 1);
    h.depth.setColumnCount(2);
    h.depth.setColumnGap(0.5);
    expect(h.records).toEqual(["Text anchor", "Text insets", "Text columns", "Text columns"]);
    expect(h.frame<TextFrame>("body").body).toEqual({
      anchor: "middle",
      insetsCm: { left: 1, top: 0.127, right: 0.254, bottom: 0.127 },
      columns: { count: 2, gapCm: 0.5 },
    });
    expect(h.depth.anchor.value).toBe("middle");
    expect(h.depth.columns.value).toEqual({ count: 2, gapCm: 0.5 });
    expect(h.model.value?.slides[0].dirty).toBe(true);
  });

  it("edits a selected shape's body (whose anchor defaults to middle)", () => {
    const h = harness();
    h.selection.select("shape");
    expect(h.depth.bodyEnabled.value).toBe(true);
    expect(h.depth.anchor.value).toBe("middle");
    h.depth.setAnchor("bottom");
    expect(h.frame<ShapeFrame>("shape").body).toEqual({ anchor: "bottom" });
  });

  it("an image frame has no text body", () => {
    const h = harness();
    h.selection.select("img");
    expect(h.depth.bodyEnabled.value).toBe(false);
    h.depth.setAnchor("bottom");
    expect(h.records).toEqual([]);
  });
});

describe("autofit write-back (AC-40/41)", () => {
  const shrinkFit: FrameFit = { scale: { fontScalePct: 62.5, lineScalePct: 90 }, contentHeightCm: 6 };

  it("reading and rendering never measure or write", () => {
    const h = harness(shrinkFit);
    h.selection.select("body");
    void h.depth.autofit.value;
    void h.depth.insetsCm.value;
    void h.depth.lineSpacingValue.value;
    expect(h.measureFit).not.toHaveBeenCalled();
    expect(h.frame<TextFrame>("body").body).toBeUndefined();
    expect(h.model.value?.slides[0].dirty).toBeUndefined();
  });

  it("picking 'shrink' stores the measured scale inside the one 'Text autofit' entry", async () => {
    const h = harness(shrinkFit);
    h.selection.select("body");
    h.depth.setAutofit("shrink");
    expect(h.records).toEqual(["Text autofit"]);
    expect(h.frame<TextFrame>("body").body).toEqual({ autofit: "shrink", fontScalePct: 62.5, lineScalePct: 90 });
    expect(h.model.value?.slides[0].dirty).toBe(true);
    await nextTick();
    await nextTick();
    expect(h.records).toEqual(["Text autofit"]);
  });

  it("picking 'resizeShape' sets the height to the measured content height", () => {
    const h = harness(shrinkFit);
    h.selection.select("body");
    const frame = h.frame<TextFrame>("body");
    h.depth.setAutofit("resizeShape");
    expect(h.frame<TextFrame>("body").heightCm).toBe(resizedHeightCm(frame, 6));
    expect(h.records).toEqual(["Text autofit"]);
  });

  it("'none' clears the stored scale; an unrendered frame keeps what is stored", () => {
    const h = harness(null);
    h.selection.select("body");
    h.depth.setAutofit("shrink");
    expect(h.frame<TextFrame>("body").body).toEqual({ autofit: "shrink" });
    h.measured.fit = shrinkFit;
    h.depth.setAutofit("shrink");
    h.depth.setAutofit("none");
    expect(h.frame<TextFrame>("body").body).toBeUndefined();
  });

  it("the default stage measurer reads nothing where there is no layout (keeps what is stored)", () => {
    expect(stageFit(textFrame("x", "body", ["one"]))).toBeNull();
  });

  it("a later command on a shrinking frame re-measures after the re-render, inside the same entry", async () => {
    const h = harness(shrinkFit);
    h.selection.select("body");
    h.depth.setAutofit("shrink");
    h.measured.fit = { scale: { fontScalePct: 50, lineScalePct: 80 }, contentHeightCm: 8 };
    h.depth.setSpaceAfter(12);
    await nextTick();
    await nextTick();
    expect(h.frame<TextFrame>("body").body).toEqual({ autofit: "shrink", fontScalePct: 50, lineScalePct: 80 });
    expect(h.records).toEqual(["Text autofit", "Paragraph spacing"]);
  });
});

describe("links (AC-43..46)", () => {
  it("Cmd/Ctrl+K in a session opens the dialog over the selected characters", () => {
    const h = harness();
    editBody(h, 1, 0, 3);
    h.editing.onKeydown(new KeyboardEvent("keydown", { key: "k", ctrlKey: true, cancelable: true }));
    expect(h.depth.linkDialog.value).toEqual({
      kind: "run",
      frameId: "body",
      range: { start: { para: 1, offset: 0 }, end: { para: 1, offset: 3 } },
      initial: null,
    });
  });

  it("a caret widens to its word", () => {
    const h = harness();
    editBody(h, 2, 2, 2);
    h.depth.openLinkDialog();
    expect(h.depth.linkDialog.value?.range).toEqual({ start: { para: 2, offset: 0 }, end: { para: 2, offset: 5 } });
  });

  it("applies an in-deck slide link as one 'Add link' entry, then removes it", () => {
    const h = harness();
    editBody(h, 1, 0, 3);
    h.depth.openLinkDialog();
    expect(h.depth.applyLink({ slide: 3 })).toBe("");
    expect(h.records).toEqual(["Add link"]);
    expect(h.depth.linkDialog.value).toBeNull();
    expect(h.frame<TextFrame>("body").paragraphs[1].runs).toEqual([{ text: "two", linkSlide: 3 }]);
    h.depth.openLinkDialog();
    expect(h.depth.linkDialog.value?.initial).toEqual({ slide: 3, tooltip: undefined });
    h.depth.removeLink();
    expect(h.records).toEqual(["Add link", "Remove link"]);
    expect(h.frame<TextFrame>("body").paragraphs[1].runs).toEqual([{ text: "two" }]);
  });

  it("rejects a dangerous URL without touching the model", () => {
    const h = harness();
    editBody(h, 0, 0, 3);
    h.depth.openLinkDialog();
    expect(h.depth.applyLink({ href: "javascript:alert(1)" })).toBe(LINK_URL_ERROR);
    expect(h.records).toEqual([]);
    expect(h.depth.linkDialog.value).not.toBeNull();
  });

  it("links a whole selected image frame from the insert bar", () => {
    const h = harness();
    h.selection.select("img");
    expect(h.depth.linkableFrame.value?.id).toBe("img");
    expect(h.depth.canLinkText.value).toBe(false);
    h.depth.openFrameLinkDialog();
    expect(h.depth.linkDialog.value).toMatchObject({ kind: "frame", frameId: "img" });
    h.depth.applyLink({ href: "https://example.com/", tooltip: "Site" });
    expect(h.frame<SlideFrame>("img")).toMatchObject({ linkHref: "https://example.com/", linkTooltip: "Site" });
    expect(h.records).toEqual(["Add link"]);
  });

  it("offers a frame link only outside a text session", () => {
    const h = harness();
    editBody(h, 0, 0, 1);
    expect(h.depth.linkableFrame.value).toBeNull();
    expect(h.depth.canLinkText.value).toBe(true);
  });

  it("lists the deck's slides by title", () => {
    const h = harness();
    expect(h.depth.slideChoices.value).toEqual([
      { number: 1, label: "Agenda" },
      { number: 2, label: "Second" },
      { number: 3, label: "Results" },
    ]);
  });

  it("slide-link-follow navigates the editor, stopping the session first", () => {
    const h = harness();
    editBody(h, 0, 0, 1);
    document.body.dispatchEvent(new CustomEvent(SLIDE_LINK_FOLLOW_EVENT, { bubbles: true, detail: { slide: 3 } }));
    expect(h.currentIndex.value).toBe(2);
    expect(h.editing.editingFrameId.value).toBeNull();
    document.dispatchEvent(new CustomEvent(SLIDE_LINK_FOLLOW_EVENT, { detail: { slide: 9 } }));
    expect(h.currentIndex.value).toBe(2);
  });

  it("stops listening once its scope is disposed", () => {
    const h = harness();
    scope?.stop();
    document.dispatchEvent(new CustomEvent(SLIDE_LINK_FOLLOW_EVENT, { detail: { slide: 2 } }));
    expect(h.currentIndex.value).toBe(0);
  });
});
