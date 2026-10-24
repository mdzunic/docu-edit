// SPEC-098 — the deck hit model: ordering, every searched scope, the scope's
// exclusions (pinned by throwing getters on everything the walk must never
// read), replace with mark inheritance and back-to-front application, and
// per-slide, per-scope dirty marking in both formats.
import { describe, expect, it } from "vitest";
import { compileQuery, type CompiledQuery } from "../core/findQuery";
import { buildPptxFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import type { PresentationFormat, PresentationModel, Slide, SlideFrame, SlideParagraph, SlideRun } from "./model";
import {
  DECK_SCOPE_LABEL,
  deckScopeTitle,
  findInDeck,
  hitIndexAtOrAfter,
  hitsBySlide,
  isHitCurrent,
  replaceAllInDeck,
  replaceHit,
  type DeckHit,
} from "./deckSearch";

const q = (query: string, options: Partial<{ matchCase: boolean; wholeWords: boolean; regex: boolean }> = {}): CompiledQuery =>
  compileQuery(query, { matchCase: false, wholeWords: false, regex: false, ...options }) as CompiledQuery;

const para = (...runs: (string | SlideRun)[]): SlideParagraph => ({ runs: runs.map((r) => (typeof r === "string" ? { text: r } : r)) });
const box = { xCm: 0, yCm: 0, widthCm: 10, heightCm: 2 };
const text = (id: string, ...paragraphs: SlideParagraph[]): SlideFrame => ({ kind: "text", id, ...box, paragraphs });
const shape = (id: string, ...paragraphs: SlideParagraph[]): SlideFrame => ({ kind: "shape", id, ...box, geometry: "rect", paragraphs });
const table = (id: string, rows: (string | null)[][]): SlideFrame => ({
  kind: "table",
  id,
  ...box,
  columnWidthsCm: rows[0].map(() => 2),
  rows: rows.map((cells) => ({ cells: cells.map((cell) => (cell === null ? { paragraphs: [], covered: true as const } : { paragraphs: [para(cell)] })) })),
});
const slide = (id: string, frames: SlideFrame[], notes?: string[]): Slide => ({ id, frames, ...(notes ? { notes } : {}) });
const deck = (slides: Slide[], format: PresentationFormat = "pptx"): PresentationModel => ({
  kind: "presentation",
  id: "deck",
  title: "Deck",
  format,
  slides,
  slideWidthCm: 33.867,
  slideHeightCm: 19.05,
});

/** A getter that fails the test if the walk ever reads it. */
function forbid<T extends object>(target: T, key: string): T {
  Object.defineProperty(target, key, {
    enumerable: false,
    get() {
      throw new Error(`deck search read ${key}`);
    },
  });
  return target;
}

const address = (hit: DeckHit) => [hit.slideIndex, hit.scope, hit.frameId ?? "-", hit.row ?? "-", hit.column ?? "-", hit.para, hit.start, hit.end];

describe("findInDeck — order and scope (AC-7..AC-12)", () => {
  it("orders hits by slide, then z-order, then paragraph, then offset, with notes after the frames", () => {
    const model = deck([
      slide("s1", [text("back", para("cat"), para("a cat, a cat")), shape("front", para("cat"))], ["cat notes"]),
      slide("s2", [table("t", [["cat", null], ["x", "the cat"]])]),
    ]);
    expect(findInDeck(model, q("cat")).map(address)).toEqual([
      [0, "frame", "back", "-", "-", 0, 0, 3],
      [0, "frame", "back", "-", "-", 1, 2, 5],
      [0, "frame", "back", "-", "-", 1, 9, 12],
      [0, "frame", "front", "-", "-", 0, 0, 3],
      [0, "notes", "-", "-", "-", 0, 0, 3],
      [1, "cell", "t", 0, 0, 0, 0, 3],
      [1, "cell", "t", 1, 1, 0, 4, 7],
    ]);
  });

  it("matches across run boundaries within a paragraph, on paragraphText offsets", () => {
    const model = deck([slide("s", [text("f", para("Op", { text: "en day", bold: true }))])]);
    const [hit] = findInDeck(model, q("open"));
    expect([hit.start, hit.end]).toEqual([0, 4]);
  });

  it("never reads group children, foreign markup, slide sources, preserved XML, opaque or media", () => {
    const group = forbid({ kind: "group", id: "g", ...box } as unknown as SlideFrame, "children");
    const foreign = forbid({ kind: "foreign", id: "x", ...box, label: "Object" } as unknown as SlideFrame, "xml");
    const s = forbid(forbid(slide("s", [group, foreign, text("t", para("needle"))]), "sourceXml"), "preserved");
    const model = forbid(forbid(deck([s]), "opaque"), "media");
    expect(findInDeck(model, q("needle")).map((hit) => hit.frameId)).toEqual(["t"]);
  });

  it("skips covered table cells", () => {
    const frame = table("t", [["a", null]]);
    if (frame.kind === "table") frame.rows[0].cells[1].paragraphs = [para("a")]; // a covered slot's text is never read
    const model = deck([slide("s", [frame])]);
    expect(findInDeck(model, q("a")).map((hit) => hit.column)).toEqual([0]);
  });

  it("finds the fixture's two \"Open\" hits — slide 1's notes and slide 2's body — and nothing grouped", async () => {
    const model = await readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "mini-deck");
    const hits = findInDeck(model, q("Open"));
    expect(hits.map((hit) => [hit.slideIndex, hit.scope, hit.snippet])).toEqual([
      [0, "notes", "Open with the Q3 highlight"],
      [1, "frame", "Open Berlin office"],
    ]);
    expect(findInDeck(model, q("Grouped"))).toEqual([]);
  });

  it("is pure: it changes no identity and marks nothing", () => {
    const model = deck([slide("s", [text("t", para("cat"))], ["cat"])]);
    const slides = model.slides;
    const first = model.slides[0];
    const snapshot = JSON.stringify(model);
    findInDeck(model, q("cat"));
    expect(model.slides).toBe(slides);
    expect(model.slides[0]).toBe(first);
    expect(JSON.stringify(model)).toBe(snapshot);
  });
});

describe("the query options (SPEC-017's compileQuery)", () => {
  const model = deck([slide("s", [text("t", para("Cat cat category"))])]);
  it("is case-insensitive by default and honours Match case", () => {
    expect(findInDeck(model, q("cat"))).toHaveLength(3);
    expect(findInDeck(model, q("cat", { matchCase: true }))).toHaveLength(2);
  });
  it("honours Whole words and regular expressions", () => {
    expect(findInDeck(model, q("cat", { wholeWords: true }))).toHaveLength(2);
    expect(findInDeck(model, q("c.t\\b", { regex: true })).map((hit) => hit.start)).toEqual([0, 4]);
  });
  it("steps past zero-width matches instead of looping", () => {
    expect(findInDeck(model, q("x*", { regex: true }))).toEqual([]);
  });
});

describe("snippets, groups and positions (AC-21..AC-23)", () => {
  it("trims the snippet to ±40 characters and records where the match starts in it", () => {
    const long = `${"a".repeat(60)} needle ${"b".repeat(60)}`;
    const [hit] = findInDeck(deck([slide("s", [text("t", para(long))])]), q("needle"));
    expect(hit.snippet).toBe(`…${"a".repeat(39)} needle ${"b".repeat(39)}…`);
    expect(hit.snippet.slice(hit.snippetOffset, hit.snippetOffset + 6)).toBe("needle");
    const short = findInDeck(deck([slide("s", [text("t", para("a needle"))])]), q("needle"))[0];
    expect([short.snippet, short.snippetOffset]).toEqual(["a needle", 2]);
  });

  it("groups hits by slide, keeping each hit's index into the list", () => {
    const hits = findInDeck(deck([slide("a", [text("t", para("x x"))]), slide("b", []), slide("c", [], ["x"])]), q("x"));
    expect(hitsBySlide(hits).map((group) => [group.slideIndex, group.hits.map((entry) => entry.index)])).toEqual([
      [0, [0, 1]],
      [2, [2]],
    ]);
  });

  it("finds the first hit at or after a slide/frame, wrapping to the first", () => {
    const hits = findInDeck(deck([slide("a", [text("t1", para("x")), text("t2", para("x"))]), slide("b", []), slide("c", [text("t3", para("x"))])]), q("x"));
    expect(hitIndexAtOrAfter(hits, 0, "t2")).toBe(1);
    expect(hitIndexAtOrAfter(hits, 1)).toBe(2);
    expect(hitIndexAtOrAfter(hits, 3)).toBe(0);
    expect(hitIndexAtOrAfter([], 0)).toBe(-1);
  });
});

describe("replaceHit / replaceAllInDeck (AC-33..AC-39)", () => {
  it("replaces one hit, the replacement taking the marks of the run containing its start", () => {
    const model = deck([slide("s", [text("t", para("say ", { text: "hel", bold: true }, { text: "lo", italic: true }, " now"))])]);
    const [hit] = findInDeck(model, q("hello"));
    expect(replaceHit(model, hit, "hi")).toBe(true);
    const frame = model.slides[0].frames[0];
    expect(frame.kind === "text" && frame.paragraphs[0].runs).toEqual([{ text: "say " }, { text: "hi", bold: true }, { text: " now" }]);
  });

  it("re-coalesces a replacement into identical neighbours", () => {
    const model = deck([slide("s", [text("t", para("a", { text: "X", bold: true }, "b"))])]);
    const [hit] = findInDeck(model, q("X"));
    // The replacement inherits X's bold, so nothing merges with the plain runs…
    replaceHit(model, hit, "Y");
    const frame = model.slides[0].frames[0];
    expect(frame.kind === "text" && frame.paragraphs[0].runs).toEqual([{ text: "a" }, { text: "Y", bold: true }, { text: "b" }]);
    // …while one plain run split by a replacement stays one run.
    const plain = deck([slide("s", [text("t", para("one two three"))])]);
    replaceHit(plain, findInDeck(plain, q("two"))[0], "2");
    const p = plain.slides[0].frames[0];
    expect(p.kind === "text" && p.paragraphs[0].runs).toEqual([{ text: "one 2 three" }]);
  });

  it("returns false and changes nothing for a stale hit", () => {
    const model = deck([slide("s", [text("t", para("cat"))])]);
    const [hit] = findInDeck(model, q("cat"));
    const frame = model.slides[0].frames[0];
    if (frame.kind === "text") frame.paragraphs[0] = para("dog");
    const before = JSON.stringify(model);
    expect(isHitCurrent(model, hit)).toBe(false);
    expect(replaceHit(model, hit, "x")).toBe(false);
    expect(JSON.stringify(model)).toBe(before);
    expect(model.slides[0].dirty).toBeUndefined();
  });

  it("applies a paragraph's hits back to front so earlier offsets stay valid, and counts them", () => {
    const model = deck([slide("s", [text("t", para("a-a-a"))], ["a a"])]);
    const hits = findInDeck(model, q("a"));
    expect(replaceAllInDeck(model, hits, "bbb")).toBe(5);
    const frame = model.slides[0].frames[0];
    expect(frame.kind === "text" && frame.paragraphs[0].runs).toEqual([{ text: "bbb-bbb-bbb" }]);
    expect(model.slides[0].notes).toEqual(["bbb bbb"]);
  });

  it("marks a frame or cell replacement with markSlideEdited, and leaves slides with no hit untouched (pptx)", () => {
    const model = deck([slide("s1", [text("t", para("cat"))]), slide("s2", [table("tb", [["cat"]])]), slide("s3", [text("u", para("dog"))])]);
    replaceAllInDeck(model, findInDeck(model, q("cat")), "cow");
    expect(model.slides.map((s) => [s.dirty, s.notesDirty])).toEqual([
      [true, undefined],
      [true, undefined],
      [undefined, undefined],
    ]);
  });

  it("marks a notes-only replacement notesDirty on pptx — the slide part stays verbatim", () => {
    const model = deck([slide("s", [text("t", para("dog"))], ["cat"])]);
    model.slides[0].partName = "ppt/slides/slide1.xml";
    expect(replaceHit(model, findInDeck(model, q("cat"))[0], "cow")).toBe(true);
    expect(model.slides[0]).toMatchObject({ notes: ["cow"], notesDirty: true });
    expect(model.slides[0].dirty).toBeUndefined();
  });

  it("rebuilds the page for a notes replacement on odp (writeOdp never consults notesDirty)", () => {
    const model = deck([slide("s", [text("t", para("dog"))], ["cat"]), slide("s2", [])], "odp");
    model.slides[0].sourceXml = "<draw:page/>";
    model.slides[1].sourceXml = "<draw:page/>";
    replaceAllInDeck(model, findInDeck(model, q("cat")), "cow");
    expect(model.slides[0].dirty).toBe(true);
    expect(model.slides[0].sourceXml).toBeUndefined();
    expect(model.slides[0].notesDirty).toBeUndefined();
    expect(model.slides[1]).toEqual({ id: "s2", frames: [], sourceXml: "<draw:page/>" });
  });

  it("strips trailing empty notes lines a replacement leaves (SPEC-083's shape)", () => {
    const model = deck([slide("s", [], ["keep", "Open", ""])]);
    replaceHit(model, findInDeck(model, q("Open"))[0], "");
    expect(model.slides[0].notes).toEqual(["keep"]);
  });

  it("drops notes a replacement empties entirely ([] is never stored)", () => {
    const model = deck([slide("s", [], ["cat"])]);
    replaceHit(model, findInDeck(model, q("cat"))[0], "");
    expect(model.slides[0].notes).toBeUndefined();
    expect(model.slides[0].notesDirty).toBe(true);
  });
});

describe("the scope statement (AC-20)", () => {
  it("names the deck's search scope and counts the exclusions from the model", () => {
    expect(DECK_SCOPE_LABEL).toBe("Slides, shapes, tables and notes");
    const group = { kind: "group", id: "g", ...box, children: [] } as SlideFrame;
    const foreign = { kind: "foreign", id: "f", ...box, label: "Object", xml: "<x/>" } as SlideFrame;
    expect(deckScopeTitle(deck([slide("a", [group, foreign]), slide("b", [{ ...group, id: "g2" }])]))).toBe(
      "Search covers slide text, shapes, tables and speaker notes. Not searched: text inside 2 groups, 1 embedded object and master and layout text.",
    );
    expect(deckScopeTitle(deck([slide("a", [])]))).toBe(
      "Search covers slide text, shapes, tables and speaker notes. Not searched: master and layout text.",
    );
  });
});
