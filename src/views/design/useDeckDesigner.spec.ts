// SPEC-115 — the deck designer's session against a stubbed gateway (no test
// in this repo requires a running model): a valid reply checked before the
// plan shows it; the repair-once-then-stop path; cancellation storing
// nothing; the flag-off path making no request at all; the source document
// untouched, asserted against a frozen model; provenance present with the
// prompt absent by default; the build storing the plan the user last saw as a
// new, numbered library entry; and Redesign's Apply as exactly one undo
// entry — or none when nothing changed — that one Undo reverts.
import { nextTick, ref, shallowRef } from "vue";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { ModelRequest, ModelResult } from "../../lib/ai/modelGateway";
import { ModelConsentDeniedError, ModelTimeoutError } from "../../lib/ai/errors";
import { BUILT_IN_THEMES } from "../../lib/core/theme";
import { createDocument, type DocumentModel, type Paragraph } from "../../lib/documentModel";
import { layoutFor, synthesiseStandardLayouts } from "../../lib/presentation/layouts";
import type { PresentationModel, Slide, TextFrame } from "../../lib/presentation/model";
import { createBlankPresentation, titleBodyFrames } from "../../lib/presentation/slides";
import type { OutlineSource } from "../../lib/design/outlineDesigner";
import { usePresentationUndo } from "../presentation/usePresentationUndo";
import { outlineFromDeck, outlineFromDocument } from "./outlineSource";
import { useDeckDesigner, type DesignerLibrary, type DeckDesignerDeps, type RedesignHost } from "./useDeckDesigner";

const control = vi.hoisted(() => ({ failBuild: false }));

// Build-time failure is defensive (a validated plan always builds), so one
// test forces SPEC-106's validator to refuse the edited plan.
vi.mock("../../lib/design/build", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../lib/design/build")>();
  const { DocumentGenerationError } = await import("../../lib/generate/errors");
  return {
    ...actual,
    buildPlannedDeck: (...args: Parameters<typeof actual.buildPlannedDeck>) => {
      if (control.failBuild) {
        throw new DocumentGenerationError("slides[1].title: expected string or run array");
      }
      return actual.buildPlannedDeck(...args);
    },
  };
});

const VALID_PLAN = {
  title: "Rotterdam migration",
  theme: "Ion",
  size: "16:9",
  slides: [
    { layout: "title", title: "Rotterdam migration", bullets: ["Board update"] },
    { layout: "titleAndBody", title: "Status", bullets: ["On track", "Phase 2 done", "Phase 3 next"] },
    { layout: "titleAndBody", title: "Risks", bullets: ["Vendor lead times"] },
    { layout: "sectionHeader", title: "Costs", bullets: [] },
    { layout: "titleOnly", title: "Questions?", bullets: [], note: "Leave ten minutes." },
  ],
};

type Reply = string | Error | ((request: ModelRequest) => Promise<string>);

/** A gateway whose `stream` answers each request with the next reply,
 * delivering text in two chunks as SSE would. */
function stubGateway(replies: Reply[]) {
  const requests: ModelRequest[] = [];
  const stream = vi.fn(async (request: ModelRequest, onDelta: (chunk: string) => void): Promise<ModelResult> => {
    requests.push(request);
    const reply = replies.shift();
    if (reply === undefined) {
      throw new Error("no more stubbed replies");
    }
    if (reply instanceof Error) {
      throw reply;
    }
    const text = typeof reply === "string" ? reply : await reply(request);
    onDelta(text.slice(0, 5));
    onDelta(text.slice(5));
    return { text, model: "stub-model", elapsedMs: 1 };
  });
  const loadGateway = vi.fn(async () => ({ stream }));
  return { requests, stream, loadGateway };
}

function paragraph(text: string, headingLevel?: number): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }], ...(headingLevel ? { headingLevel } : {}) };
}

function sourceDocument(): DocumentModel {
  const doc = createDocument("doc-1");
  doc.title = "Migration notes";
  doc.sections = [
    {
      blocks: [
        paragraph("Where we are", 1),
        paragraph("A short update for the board."),
        paragraph("Status", 2),
        paragraph("Phase one finished in August."),
        paragraph("Next", 2),
        paragraph("Phase two starts in November."),
      ],
    },
  ];
  return doc;
}

function deepFreeze<T>(value: T): T {
  if (value !== null && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    Object.values(value).forEach(deepFreeze);
  }
  return value;
}

function fakeLibrary(outlines: Record<string, () => OutlineSource> = {}, takenTitles = ["Rotterdam migration"]) {
  const stored: PresentationModel[] = [];
  const trashed: string[] = [];
  const library: DesignerLibrary = {
    sources: () => Object.keys(outlines).map((id) => ({ id, title: outlines[id]().title, kind: "document" as const })),
    outlineOf: (id) => outlines[id]?.(),
    titles: () => takenTitles,
    store: (model) => stored.push(model),
    trash: (id) => trashed.push(id),
  };
  return { library, stored, trashed };
}

function newDesigner(overrides: Partial<DeckDesignerDeps> = {}, replies: Reply[] = [JSON.stringify(VALID_PLAN)]) {
  const gateway = stubGateway(replies);
  const lib = fakeLibrary();
  const onClose = vi.fn();
  const navigate = vi.fn();
  const toast = vi.fn();
  const designer = useDeckDesigner({
    mode: "new",
    onClose,
    navigate,
    toast,
    isEnabled: () => true,
    providerMode: () => "local",
    now: () => new Date("2026-10-24T12:00:00.000Z"),
    library: lib.library,
    loadGateway: gateway.loadGateway,
    ...overrides,
  });
  return { designer, gateway, ...lib, onClose, navigate, toast };
}

function titles(model: PresentationModel): string[] {
  return model.slides.map((slide) => {
    const frame = slide.frames.find((f) => f.kind === "text" && f.role === "title") as TextFrame | undefined;
    return frame ? frame.paragraphs.map((p) => p.runs.map((run) => run.text).join("")).join("") : "";
  });
}

beforeEach(() => {
  control.failBuild = false;
});

describe("the New flow against a stubbed gateway", () => {
  it("sends under 'Design a deck' with json: true, streams the reply, and checks it before the plan step shows it", async () => {
    const { designer, gateway } = newDesigner();
    designer.brief.value = "A board update on the Rotterdam migration";
    const sending = designer.send();
    expect(designer.step.value).toBe("plan");
    expect(designer.phase.value).toBe("streaming");
    await sending;
    expect(gateway.requests).toHaveLength(1);
    expect(gateway.requests[0]).toMatchObject({ purpose: "Design a deck", json: true });
    expect(gateway.requests[0].user).toContain("Brief: A board update on the Rotterdam migration");
    expect(designer.streamText.value).toBe(JSON.stringify(VALID_PLAN));
    expect(designer.phase.value).toBe("ready");
    expect(designer.plan.slides.map((slide) => slide.title)).toEqual(VALID_PLAN.slides.map((slide) => slide.title));
    expect(designer.generatedModel.value).toBe("stub-model");
  });

  it("repairs a malformed reply exactly once, then stops with the error, the raw text and nothing stored", async () => {
    const { designer, gateway, stored } = newDesigner({}, ["Sure! Here is your deck:", "{\"slides\": "]);
    designer.brief.value = "Board update";
    await designer.send();
    expect(gateway.stream).toHaveBeenCalledTimes(2);
    const repair = gateway.requests[1];
    expect(repair.purpose).toBe("Design a deck");
    expect(repair.user).toContain(gateway.requests[0].user);
    expect(repair.user).toContain("Your previous reply was:\n<<<\nSure! Here is your deck:\n>>>");
    expect(repair.user).toMatch(/It was rejected: reply: not valid JSON \(.+\)\nReply again with the corrected JSON object only\.$/);
    expect(designer.phase.value).toBe("failed");
    expect(designer.error.value).toMatch(/^reply: not valid JSON/);
    expect(designer.raw.value).toBe("{\"slides\": ");
    expect(designer.plan.slides).toEqual([]);
    await nextTick();
    expect(gateway.stream).toHaveBeenCalledTimes(2);
    expect(stored).toEqual([]);
    designer.build();
    expect(stored).toEqual([]);
  });

  it("a reply validatePlan rejects takes the repair too, and a good repair is accepted", async () => {
    const { designer, gateway } = newDesigner({}, [JSON.stringify({ title: "x", slides: [] }), JSON.stringify(VALID_PLAN)]);
    designer.brief.value = "Board update";
    await designer.send();
    expect(gateway.stream).toHaveBeenCalledTimes(2);
    expect(gateway.requests[1].user).toContain("It was rejected: slides: expected a non-empty array of slides");
    expect(designer.phase.value).toBe("ready");
    expect(designer.plan.slides).toHaveLength(5);
  });

  it("'Try again' after the final failure starts a new request — never a silent one", async () => {
    const { designer, gateway } = newDesigner({}, ["nope", "nope", JSON.stringify(VALID_PLAN)]);
    designer.brief.value = "Board update";
    await designer.send();
    expect(designer.phase.value).toBe("failed");
    designer.tryAgain();
    await vi.waitFor(() => expect(designer.phase.value).toBe("ready"));
    expect(gateway.stream).toHaveBeenCalledTimes(3);
    expect(gateway.requests[2].user).not.toContain("Your previous reply was");
  });

  it("cancellation aborts through the AbortSignal and stores nothing", async () => {
    let signal: AbortSignal | undefined;
    const hanging = (request: ModelRequest) =>
      new Promise<string>((_resolve, reject) => {
        signal = request.signal;
        request.signal?.addEventListener("abort", () => reject(new DOMException("cancelled", "AbortError")));
      });
    const { designer, stored, onClose } = newDesigner({}, [hanging]);
    designer.brief.value = "Board update";
    const sending = designer.send();
    await vi.waitFor(() => expect(signal).toBeDefined());
    designer.cancel();
    await sending;
    expect(signal?.aborted).toBe(true);
    expect(onClose).toHaveBeenCalledTimes(1);
    expect(stored).toEqual([]);
    expect(designer.error.value).toBeNull();
  });

  it("with the experiment off, no request is possible: the gateway is never even loaded", async () => {
    const { designer, gateway } = newDesigner({ isEnabled: () => false });
    designer.brief.value = "Board update";
    await designer.send();
    expect(gateway.loadGateway).not.toHaveBeenCalled();
    expect(gateway.stream).not.toHaveBeenCalled();
    expect(designer.offline.value).toContain("switched off");
    expect(designer.phase.value).toBe("failed");
  });

  it("a declined consent sends nothing and leaves the model-free designer on offer", async () => {
    const { designer } = newDesigner({}, [new ModelConsentDeniedError("http://localhost:11434")]);
    designer.brief.value = "Board update";
    await designer.send();
    expect(designer.offline.value).toContain("nothing was sent");
    expect(designer.error.value).toBeNull();
  });

  it("a timeout reports SPEC-113's ModelTimeoutError with the configured limit named", async () => {
    const { designer } = newDesigner({}, [new ModelTimeoutError(45_000)]);
    designer.brief.value = "Board update";
    await designer.send();
    expect(designer.phase.value).toBe("failed");
    expect(designer.error.value).toContain("No complete answer within 45 s");
  });

  it("builds the plan the user last saw, stores it as a new numbered entry, opens it and offers Undo", async () => {
    const { designer, stored, navigate, toast, trashed } = newDesigner();
    designer.brief.value = "Board update";
    await designer.send();
    designer.setSlideField(2, "title", "Risks we accept");
    designer.deleteSlide(3);
    designer.build();
    expect(stored).toHaveLength(1);
    const deck = stored[0];
    expect(deck.slides).toHaveLength(4);
    expect(titles(deck)).toEqual(["Rotterdam migration", "Status", "Risks we accept", "Questions?"]);
    expect(deck.title).toBe("Rotterdam migration (2)");
    expect(deck.theme?.name).toBe("Ion");
    expect(deck.slides[3].notes).toEqual(["Leave ten minutes."]);
    expect(navigate).toHaveBeenCalledWith(`/slides/${deck.id}`);
    expect(toast).toHaveBeenCalledWith("Created “Rotterdam migration (2)”", expect.objectContaining({ label: "Undo" }));
    toast.mock.calls[0][1].run();
    expect(trashed).toEqual([deck.id]);
  });

  it("records provenance — GeneratedBy, GeneratedAt, GeneratedModel — and not the prompt unless asked", async () => {
    const first = newDesigner();
    first.designer.brief.value = "Board update on Rotterdam";
    await first.designer.send();
    first.designer.build();
    expect(first.stored[0].properties?.customProperties).toEqual([
      { name: "GeneratedBy", type: "text", value: "docu-edit" },
      { name: "GeneratedAt", type: "text", value: "2026-10-24T12:00:00.000Z" },
      { name: "GeneratedModel", type: "text", value: "stub-model" },
    ]);

    const second = newDesigner();
    second.designer.brief.value = "Board update on Rotterdam";
    second.designer.keepPrompt.value = true;
    await second.designer.send();
    second.designer.build();
    expect(second.stored[0].properties?.customProperties).toContainEqual({
      name: "GeneratedPrompt",
      type: "text",
      value: "Board update on Rotterdam",
    });
  });

  it("switching a row with several bullets to title lists the bullets that will not be built, before Build", async () => {
    const { designer, stored } = newDesigner();
    designer.brief.value = "Board update";
    await designer.send();
    expect(designer.pendingDrops.value).toEqual([]);
    designer.setSlideField(1, "layout", "title");
    expect(designer.pendingDrops.value.map((record) => record.reason)).toEqual([
      "a title slide keeps one line under its title; “Phase 2 done” dropped",
      "a title slide keeps one line under its title; “Phase 3 next” dropped",
    ]);
    designer.build();
    const subtitle = stored[0].slides[1].frames.find((f) => f.kind === "text" && f.role === "subtitle") as TextFrame;
    expect(subtitle.paragraphs.map((p) => p.runs[0].text)).toEqual(["On track"]);
  });

  it("plan editing: Alt+arrow moves, add and delete; the select can only take the closed vocabulary", async () => {
    const { designer } = newDesigner();
    designer.brief.value = "Board update";
    await designer.send();
    expect(designer.moveSlide(1, 1)).toBe(2);
    expect(designer.plan.slides.map((slide) => slide.title).slice(1, 3)).toEqual(["Risks", "Status"]);
    expect(designer.moveSlide(0, -1)).toBe(0);
    designer.addSlide();
    expect(designer.plan.slides).toHaveLength(6);
    designer.setSlideField(5, "layout", "other");
    expect(designer.plan.slides[5].layout).toBe("titleAndBody");
    designer.setPlanField("theme", "Neon");
    expect(designer.plan.theme).toBe("Ion");
  });

  it("a deckFromDescription failure at Build shows the error with the raw text, stores nothing and keeps Back", async () => {
    const { designer, gateway, stored } = newDesigner();
    designer.brief.value = "Board update";
    await designer.send();
    control.failBuild = true;
    designer.build();
    expect(designer.step.value).toBe("build");
    expect(designer.buildPhase.value).toBe("failed");
    expect(designer.error.value).toBe("slides[1].title: expected string or run array");
    expect(designer.raw.value).toBe(JSON.stringify(VALID_PLAN));
    expect(stored).toEqual([]);
    expect(gateway.stream).toHaveBeenCalledTimes(1);
    designer.back();
    expect(designer.step.value).toBe("plan");
    expect(designer.phase.value).toBe("ready");
  });
});

describe("the source is the user's choice, and is never touched", () => {
  it("is carried only when picked, and the frozen source survives a request, an outline and a build", async () => {
    const frozen = deepFreeze(sourceDocument());
    const pristine = structuredClone(frozen);
    const deck = threeSlideDeck();
    deck.theme = { ...BUILT_IN_THEMES[3] };
    const frozenDeck = deepFreeze(deck);
    const pristineDeck = structuredClone(frozenDeck);
    const gateway = stubGateway([JSON.stringify(VALID_PLAN)]);
    const lib = fakeLibrary({ "doc-1": () => outlineFromDocument(frozen), "deck-1": () => outlineFromDeck(frozenDeck) });
    const designer = useDeckDesigner({
      mode: "new",
      onClose: vi.fn(),
      navigate: vi.fn(),
      toast: vi.fn(),
      isEnabled: () => true,
      providerMode: () => "local",
      library: lib.library,
      loadGateway: gateway.loadGateway,
    });
    designer.brief.value = "Board update";
    expect(designer.request.value.user).not.toContain("Phase one");
    designer.sourceChoice.value = "doc-1";
    expect(designer.request.value.user).toContain("Source document “Migration notes”:");
    expect(designer.request.value.user).toContain("## Status\nPhase one finished in August.");
    expect(designer.bytes.value).toBe(
      new TextEncoder().encode(designer.request.value.system).byteLength +
        new TextEncoder().encode(designer.request.value.user).byteLength,
    );
    await designer.send();

    designer.outline();
    expect(designer.generatedModel.value).toBe("outline-designer");
    expect(designer.plan.slides.map((slide) => slide.layout)).toEqual(["title", "sectionHeader", "titleAndBody", "titleAndBody"]);
    designer.build();
    expect(lib.stored).toHaveLength(1);
    expect(lib.stored[0].properties?.customProperties).toContainEqual({ name: "GeneratedModel", type: "text", value: "outline-designer" });
    expect(frozen).toEqual(pristine);

    // A deck as the source: outlined with its own theme, and left as it was.
    const fromDeck = useDeckDesigner({
      mode: "new",
      onClose: vi.fn(),
      navigate: vi.fn(),
      toast: vi.fn(),
      isEnabled: () => true,
      providerMode: () => "off",
      library: lib.library,
      loadGateway: gateway.loadGateway,
    });
    fromDeck.sourceChoice.value = "deck-1";
    fromDeck.outline();
    expect(fromDeck.plan.theme).toBe(BUILT_IN_THEMES[3].name);
    expect(fromDeck.plan.slides.map((slide) => slide.title)).toEqual(["Q3 review", "Alpha", "Beta", "Slide 3"]);
    fromDeck.build();
    expect(lib.stored).toHaveLength(2);
    expect(frozenDeck).toEqual(pristineDeck);
  });

  it("states a source longer than MAX_SOURCE_CHARS is cut", () => {
    const { designer } = newDesigner();
    designer.sourceChoice.value = "paste";
    designer.pastedText.value = "word ".repeat(9_000);
    expect(designer.truncation.value).toBe("The source is 45,000 characters long; only the first 40,000 will be sent.");
  });
});

// --- Redesign -------------------------------------------------------------------

function titledSlide(deck: PresentationModel, id: string, title: string, body: string): Slide {
  const [titleFrame, bodyFrame] = titleBodyFrames(deck.slideWidthCm, deck.slideHeightCm);
  titleFrame.paragraphs = [{ runs: [{ text: title }] }];
  bodyFrame.paragraphs = [{ runs: [{ text: body }] }];
  return { id, frames: [titleFrame, bodyFrame], notes: [`${title} notes`] };
}

function threeSlideDeck(): PresentationModel {
  const deck = createBlankPresentation();
  deck.title = "Q3 review";
  deck.layouts = synthesiseStandardLayouts(deck.slideWidthCm, deck.slideHeightCm);
  deck.slides = [
    titledSlide(deck, "a", "Alpha", "Alpha body text"),
    titledSlide(deck, "b", "Beta", "Beta body text"),
    titledSlide(deck, "c", "", "Gamma body text"),
  ];
  deck.slides[0].layoutKey = "standard:titleAndBody";
  return deck;
}

function contentOf(deck: PresentationModel) {
  return deck.slides.map((slide) => ({
    id: slide.id,
    notes: slide.notes,
    text: slide.frames.map((frame) => (frame.kind === "text" ? frame.paragraphs.map((p) => p.runs.map((r) => r.text).join("")) : [])),
  }));
}

function redesigner(replies: Reply[], deck = threeSlideDeck()) {
  const model = shallowRef<PresentationModel | null>(deck);
  const undo = usePresentationUndo({ model, onApply: () => {} });
  const record = vi.fn((label?: string) => undo.record(label));
  const commit = vi.fn();
  const host: RedesignHost = { model, record, commit, currentIndex: ref(0) };
  const gateway = stubGateway(replies);
  const onClose = vi.fn();
  const navigate = vi.fn();
  const toast = vi.fn();
  const lib = fakeLibrary();
  const designer = useDeckDesigner({
    mode: "redesign",
    host,
    onClose,
    navigate,
    toast,
    isEnabled: () => true,
    providerMode: () => "local",
    library: lib.library,
    loadGateway: gateway.loadGateway,
  });
  return { designer, deck, undo, record, commit, gateway, onClose, navigate, toast, lib, host };
}

describe("Redesign", () => {
  it("asks under 'Redesign a deck' with the slides' index, title and layout — no body text, no notes", async () => {
    const { designer, gateway } = redesigner([JSON.stringify({ theme: null, slides: [] })]);
    designer.brief.value = "Tighter";
    await designer.send();
    const request = gateway.requests[0];
    expect(request.purpose).toBe("Redesign a deck");
    expect(request.user).toContain("Deck title: Q3 review\nCurrent theme: Office\nSlides (3):");
    expect(JSON.parse(request.user.slice(request.user.indexOf("[")))).toEqual([
      { index: 0, title: "Alpha", layout: "titleAndBody" },
      { index: 1, title: "Beta", layout: "other" },
      { index: 2, title: "", layout: "other" },
    ]);
    expect(request.user).not.toContain("body text");
    expect(request.user).not.toContain("notes");
  });

  it("cuts a slide title at 200 characters in the request", () => {
    const deck = threeSlideDeck();
    (deck.slides[1].frames[0] as TextFrame).paragraphs = [{ runs: [{ text: "t".repeat(260) }] }];
    const { designer } = redesigner([], deck);
    expect(designer.request.value.user).toContain(`"title": "${"t".repeat(200)}"`);
    expect(designer.request.value.user).not.toContain("t".repeat(201));
  });

  it("a reply naming source 2 then 0 lists 3 rows — 3, 1, 2 — and Apply is exactly one undo entry", async () => {
    const { designer, deck, undo, record, commit, onClose, navigate, toast, lib } = redesigner([
      JSON.stringify({ theme: "Ion", slides: [{ source: 2, layout: "sectionHeader" }, { source: 0 }] }),
    ]);
    const before = contentOf(deck);
    await designer.send();
    expect(designer.rows.value.map((row) => row.source + 1)).toEqual([3, 1, 2]);
    expect(designer.rows.value.map((row) => row.layout)).toEqual(["sectionHeader", "titleAndBody", null]);
    expect(designer.dropped.value.map((drop) => drop.kind)).toEqual(["slideOrder"]);
    expect(designer.redesignTheme.value).toBe("Ion");

    designer.applyRedesign();
    expect(record).toHaveBeenCalledTimes(1);
    expect(commit).toHaveBeenCalledTimes(1);
    expect(undo.canUndo.value).toBe(true);
    expect(deck.slides.map((slide) => slide.id)).toEqual(["c", "a", "b"]);
    expect(deck.slides).toHaveLength(3);
    expect(layoutFor(deck, deck.slides[0])?.kind).toBe("sectionHeader");
    expect(deck.theme?.name).toBe("Ion");
    // Slide count, every slide's text and notes: unchanged.
    const after = new Map(contentOf(deck).map((entry) => [entry.id, entry]));
    for (const entry of before) {
      expect(after.get(entry.id)).toEqual(entry);
    }
    // No library entry, no navigation, no toast, no provenance.
    expect(lib.stored).toEqual([]);
    expect(navigate).not.toHaveBeenCalled();
    expect(toast).not.toHaveBeenCalled();
    expect(deck.properties?.customProperties ?? []).toEqual([]);
    expect(onClose).toHaveBeenCalledTimes(1);

    // One Undo restores the order, the layouts and the theme.
    undo.undo();
    expect(deck.slides.map((slide) => slide.id)).toEqual(["a", "b", "c"]);
    expect(deck.slides.map((slide) => slide.layoutKey)).toEqual(["standard:titleAndBody", undefined, undefined]);
    expect(deck.theme).toBeUndefined();
    expect(undo.canUndo.value).toBe(false);
  });

  it("records nothing when nothing changed, and closes", async () => {
    const { designer, record, commit, onClose, undo } = redesigner([]);
    designer.outline();
    expect(designer.phase.value).toBe("ready");
    designer.setRedesignTheme("Office");
    designer.applyRedesign();
    expect(record).not.toHaveBeenCalled();
    expect(commit).not.toHaveBeenCalled();
    expect(undo.canUndo.value).toBe(false);
    expect(onClose).toHaveBeenCalledTimes(1);
  });

  it("without a model: its plan step seeded from the open deck, with no request made", () => {
    const { designer, gateway } = redesigner([]);
    designer.outline();
    expect(gateway.loadGateway).not.toHaveBeenCalled();
    expect(designer.step.value).toBe("plan");
    expect(designer.rows.value.map((row) => [row.source, row.title, row.layout])).toEqual([
      [0, "Alpha", "titleAndBody"],
      [1, "Beta", null],
      [2, "", null],
    ]);
    expect(designer.redesignTheme.value).toBe("");
    expect(designer.moveRow(2, -1)).toBe(1);
    expect(designer.rows.value.map((row) => row.source)).toEqual([0, 2, 1]);
  });

  it("an unparseable reply is repaired once; a second failure leaves the by-hand plan on offer", async () => {
    const { designer, gateway } = redesigner(["not json", "[]"]);
    await designer.send();
    expect(gateway.stream).toHaveBeenCalledTimes(2);
    expect(gateway.requests[1].user).toContain("It was rejected: reply: not valid JSON");
    expect(designer.phase.value).toBe("failed");
    expect(designer.error.value).toBe("redesign: expected a JSON object");
    expect(designer.raw.value).toBe("[]");
    designer.outline();
    expect(designer.rows.value).toHaveLength(3);
  });

  it("a reply with the wrong count is normalised, never repaired", async () => {
    const { designer, gateway } = redesigner([JSON.stringify({ slides: [0, 1, 2, 3, 4].map((source) => ({ source })) })]);
    await designer.send();
    expect(gateway.stream).toHaveBeenCalledTimes(1);
    expect(designer.rows.value.map((row) => row.source)).toEqual([0, 1, 2]);
  });
});
