// SPEC-106 — the JSON deck-description format: each `layout` row of the
// Reference table produces exactly what the matching PresentationBuilder
// calls produce (1:1), unknown top-level and per-slide keys are validation
// errors named by path, an object with no `slides` describes an empty deck,
// and the published schema/deck-description.json has the runtime validator's
// shape (beside SPEC-045's schema/document-description.json).
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { firstStructuralDifference } from "../equality";
import type { PresentationModel, TextFrame } from "../presentation/model";
import { MAX_INDENT_LEVEL } from "../presentation/text";
import { DECK_LAYOUT_KEYS, DECK_LAYOUTS, deckFromDescription, type DeckDescription } from "./deckDescription";
import { DocumentGenerationError } from "./errors";
import { PresentationBuilder } from "./presentation";

const PNG_4X2 = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAAAAABawyK/AAAAC0lEQVR4nGNggAEAAAoAAX+AdF4AAAAASUVORK5CYII=";

function expectSame(a: PresentationModel, b: PresentationModel): void {
  expect(firstStructuralDifference(a, b)).toBeNull();
}

describe("deckFromDescription — one row per layout (AC-12)", () => {
  it('"title" is addTitleSlide (+ setNotes)', () => {
    const fromJson = deckFromDescription({
      slides: [{ layout: "title", title: "Q3 review", subtitle: [{ text: "Prepared ", italic: true }, "30 September"], notes: ["Smile."] }],
    });
    const fromBuilder = new PresentationBuilder()
      .addTitleSlide("Q3 review", [{ text: "Prepared ", italic: true }, "30 September"])
      .setNotes(["Smile."])
      .toDeck();
    expectSame(fromJson, fromBuilder);
    expect((fromJson.slides[0].frames[1] as TextFrame).paragraphs[0].runs[0]).toEqual({ text: "Prepared ", italic: true });
  });

  it('"bullets" is addBulletSlide, its style the method\'s options', () => {
    const bullets = ["Revenue up 12%", ["Europe up 20%", { text: "APAC flat", level: 3 }], { runs: [{ text: "Costs", bold: true }] }];
    const fromJson = deckFromDescription({ slides: [{ layout: "bullets", title: "Highlights", bullets, style: "number", notes: ["Lead with Europe."] }] });
    const fromBuilder = new PresentationBuilder()
      .addBulletSlide("Highlights", bullets as never, { style: "number" })
      .setNotes(["Lead with Europe."])
      .toDeck();
    expectSame(fromJson, fromBuilder);
    expect((fromJson.slides[0].frames[1] as TextFrame).paragraphs).toHaveLength(4);
  });

  it('"image" is addImageSlide', () => {
    const image = { data: PNG_4X2, contentType: "image/png", width: { cm: 12 } } as const;
    expectSame(
      deckFromDescription({ slides: [{ layout: "image", title: "Pipeline", image }] }),
      new PresentationBuilder().addImageSlide("Pipeline", image).toDeck(),
    );
  });

  it('"blank" is addBlankSlide plus addTextBox / addImage / setBackground per entry', () => {
    const rect = { x: { cm: 2 }, y: { cm: 3 }, width: { cm: 10 }, height: { cm: 2 } };
    const placed = { data: PNG_4X2, contentType: "image/png", x: { cm: 14 }, y: { cm: 3 }, width: { cm: 8 } } as const;
    const fromJson = deckFromDescription({
      slides: [
        {
          layout: "blank",
          textBoxes: [{ text: "Plain", ...rect }, { runs: [{ text: "Rich", bold: true }], ...rect }],
          images: [placed],
          background: "#112233",
          notes: ["n"],
        },
      ],
    });
    const fromBuilder = new PresentationBuilder()
      .addBlankSlide()
      .addTextBox("Plain", rect)
      .addTextBox([{ text: "Rich", bold: true }], rect)
      .addImage(placed)
      .setBackground("#112233")
      .setNotes(["n"])
      .toDeck();
    expectSame(fromJson, fromBuilder);
    expect(fromJson.slides[0].frames.map((frame) => frame.kind)).toEqual(["text", "text", "image"]);
    expect(fromJson.slides[0].background).toEqual({ type: "solid", color: "#112233" });
  });

  it("carries title, format and size, as the builder's options do", () => {
    const description: DeckDescription = { title: "Q3", format: "odp", size: { width: { cm: 30 }, height: { cm: 20 } }, slides: [] };
    expectSame(deckFromDescription(description), new PresentationBuilder({ title: "Q3", format: "odp", size: { width: { cm: 30 }, height: { cm: 20 } } }).toDeck());
    expect(deckFromDescription({ size: "4:3" }).slideWidthCm).toBe(25.4);
  });
});

describe("validation (AC-13)", () => {
  it("an object with no slides describes an empty deck", () => {
    const deck = deckFromDescription({});
    expect(deck.slides).toEqual([]);
    expect(deck.kind).toBe("presentation");
    expect(deckFromDescription({ slides: [] }).slides).toEqual([]);
  });

  it("rejects an unknown top-level key", () => {
    expect(() => deckFromDescription({ body: [] })).toThrow(DocumentGenerationError);
    expect(() => deckFromDescription({ body: [] })).toThrow(/body: unknown property/);
  });

  it("rejects an unknown per-slide key, including another layout's field", () => {
    expect(() => deckFromDescription({ slides: [{ layout: "title", title: "a", bullets: [] }] })).toThrow("slides[0].bullets: unknown property");
    expect(() => deckFromDescription({ slides: [{ layout: "blank", title: "a" }] })).toThrow("slides[0].title: unknown property");
    expect(() => deckFromDescription({ slides: [{ layout: "image", title: "a", image: { data: PNG_4X2, contentType: "image/png" }, background: "#000000" }] })).toThrow(
      "slides[0].background: unknown property",
    );
    expect(() => deckFromDescription({ slides: [{ layout: "blank", textBoxes: [{ text: "a", x: { cm: 0 }, y: { cm: 0 }, width: { cm: 1 }, height: { cm: 1 }, fill: "#000000" }] }] })).toThrow(
      "slides[0].textBoxes[0].fill: unknown property",
    );
  });

  it("names the offending path from the description's root", () => {
    const slides = [
      { layout: "title", title: "a" },
      { layout: "title", title: "b" },
      { layout: "bullets", title: "c", bullets: [{ level: 1, text: 5 }] },
    ];
    expect(() => deckFromDescription({ slides })).toThrow("slides[2].bullets[0].text: expected a string");
    slides[2] = { layout: "bullets", title: "c", bullets: [7] } as never;
    expect(() => deckFromDescription({ slides })).toThrow(/^slides\[2\]\.bullets\[0\]: expected string or run array/);
    expect(() => deckFromDescription({ slides: [{ layout: "bullets", title: "c" }] })).toThrow("slides[0].bullets: expected an array");
    expect(() => deckFromDescription({ slides: [{ title: "c" }] })).toThrow("slides[0].layout: expected a string");
    expect(() => deckFromDescription({ slides: [{ layout: "chart" }] })).toThrow(/^slides\[0\]\.layout: expected one of "title", "bullets", "image", "blank"/);
    expect(() => deckFromDescription({ slides: [{ layout: "image", title: "x" }] })).toThrow("slides[0].image: expected an object");
    expect(() => deckFromDescription({ slides: [{ layout: "blank", images: [{ data: PNG_4X2, contentType: "image/png", x: { cm: 0 }, y: { cm: 0 } }] }] })).toThrow(
      /^slides\[0\]\.images\[0\]\.width: expected a length/,
    );
    expect(() => deckFromDescription({ slides: [{ layout: "blank", textBoxes: [{ text: "a", runs: [] }] }] })).toThrow('slides[0].textBoxes[0]: specify either "text" or "runs", not both');
    expect(() => deckFromDescription({ slides: [{ layout: "blank", notes: [1] }] })).toThrow("slides[0].notes[0]: expected a string");
    expect(() => deckFromDescription({ format: "key" })).toThrow(/^format: expected one of "pptx", "odp"/);
    expect(() => deckFromDescription({ slides: {} })).toThrow("slides: expected an array");
    expect(() => deckFromDescription("deck")).toThrow(DocumentGenerationError);
  });
});

describe("PresentationBuilder.fromDescription (AC-13)", () => {
  it("is the static factory: the description's deck, ready for more calls", () => {
    const description = { title: "Q3", slides: [{ layout: "title", title: "Q3 review" }] };
    const builder = PresentationBuilder.fromDescription(description);
    expect(builder).toBeInstanceOf(PresentationBuilder);
    expectSame(builder.toDeck(), deckFromDescription(description));
    builder.addBulletSlide("More", ["x"]);
    expect(builder.toDeck().slides).toHaveLength(2);
    expect(builder.fileName()).toBe("Q3.pptx");
  });

  it("reports what the description dropped", () => {
    const builder = PresentationBuilder.fromDescription({
      slides: [{ layout: "title", title: [{ text: "x", highlight: "#ffff00" }] }],
    });
    expect(builder.dropped().map((record) => [record.kind, record.location])).toEqual([["runFormatting", "slide 1"]]);
  });
});

// The published schema describes the same closed shape the runtime
// validator enforces — checked structurally, the way SPEC-045's schema is
// described beside it (the project carries no JSON Schema engine).
describe("schema/deck-description.json", () => {
  const schemaDir = resolve(dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "schema");
  const schema = JSON.parse(readFileSync(resolve(schemaDir, "deck-description.json"), "utf8"));
  const documentSchema = JSON.parse(readFileSync(resolve(schemaDir, "document-description.json"), "utf8"));
  const defs = schema.definitions;

  it("is a closed draft-07 object with the description's four top-level keys", () => {
    expect(schema.$schema).toBe(documentSchema.$schema);
    expect(schema.type).toBe("object");
    expect(schema.additionalProperties).toBe(false);
    expect(Object.keys(schema.properties).toSorted()).toEqual(["format", "size", "slides", "title"]);
    expect(schema.required).toBeUndefined();
    expect(schema.properties.format.enum).toEqual(["pptx", "odp"]);
    expect(defs.size.oneOf[0].enum).toEqual(["16:9", "4:3"]);
  });

  it("has one closed slide shape per layout, keyed exactly as the validator", () => {
    const refs = defs.slide.oneOf.map((entry: { $ref: string }) => entry.$ref.replace("#/definitions/", ""));
    const byLayout = new Map<string, { properties: Record<string, unknown>; additionalProperties: boolean; required: string[] }>(
      refs.map((name: string) => [defs[name].properties.layout.const, defs[name]]),
    );
    expect([...byLayout.keys()]).toEqual([...DECK_LAYOUTS]);
    for (const layout of DECK_LAYOUTS) {
      const shape = byLayout.get(layout)!;
      expect(shape.additionalProperties, layout).toBe(false);
      expect(Object.keys(shape.properties).toSorted(), layout).toEqual([...DECK_LAYOUT_KEYS[layout]].toSorted());
      expect(shape.required, layout).toContain("layout");
    }
    expect(byLayout.get("bullets")!.required).toEqual(["layout", "title", "bullets"]);
    expect(byLayout.get("image")!.required).toEqual(["layout", "title", "image"]);
  });

  it("shares SPEC-045's length and run shapes, and the bullet level range", () => {
    expect(defs.length).toEqual(documentSchema.definitions.length);
    expect(defs.runObject.properties).toEqual(documentSchema.definitions.runObject.properties);
    expect(defs.runObject.additionalProperties).toBe(false);
    expect(defs.runObject.properties.sizePt).toBeUndefined();
    const row = defs.bulletRow.oneOf[1];
    expect(row.properties.level).toEqual({ type: "integer", minimum: 0, maximum: MAX_INDENT_LEVEL });
  });
});
