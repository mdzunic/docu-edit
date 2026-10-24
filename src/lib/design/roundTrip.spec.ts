// SPEC-115 — a designed deck is an ordinary deck: written by the existing
// writers in both formats (no serialisation of its own) and read back by the
// existing readers with its slides, titles, layout kinds and theme intact —
// and with its provenance. SPEC-041's custom properties travel in both
// formats (docProps/custom.xml in .pptx, meta.xml in .odp), so GeneratedBy,
// GeneratedAt, GeneratedModel and an opted-in GeneratedPrompt leave with the
// file, as the dialog says they will.
import { describe, expect, it } from "vitest";
import { unzip } from "../core/zip";
import { readOdp, readPptx } from "../generate/readDeck";
import { layoutFor } from "../presentation/layouts";
import type { PresentationModel, TextFrame } from "../presentation/model";
import { writeOdp } from "../presentation/odp/writer";
import { writePptx } from "../presentation/pptx/writer";
import { createBlankPresentation } from "../presentation/slides";
import { buildPlannedDeck } from "./build";
import { validatePlan } from "./plan";

const PLAN = validatePlan({
  title: "Rotterdam migration",
  theme: "Facet",
  size: "16:9",
  slides: [
    { layout: "title", title: "Rotterdam migration", bullets: ["Board update"] },
    { layout: "sectionHeader", title: "Where we are" },
    { layout: "titleAndBody", title: "Status", bullets: ["On track", { text: "Phase 2 done", level: 1 }], note: "Lead with the date." },
    { layout: "titleOnly", title: "Questions?" },
  ],
}).plan;

function titles(deck: PresentationModel): string[] {
  return deck.slides.map((slide) => {
    const frame = slide.frames.find((f) => f.kind === "text" && f.role === "title") as TextFrame | undefined;
    return frame ? frame.paragraphs.map((p) => p.runs.map((run) => run.text).join("")).join("") : "";
  });
}

describe.each(["pptx", "odp"] as const)("a designed deck round-trips as .%s", (format) => {
  it("keeps its slides, kinds, theme and provenance", async () => {
    const deck = buildPlannedDeck(PLAN, {
      generatedAt: "2026-10-24T12:00:00.000Z",
      generatedModel: "stub-model",
      prompt: "A board update",
    });
    deck.format = format;
    const bytes = format === "pptx" ? (await writePptx(deck)).bytes : (await writeOdp(deck)).bytes;
    const back = format === "pptx" ? await readPptx(bytes, deck.title) : await readOdp(bytes, deck.title);

    expect(titles(back)).toEqual(["Rotterdam migration", "Where we are", "Status", "Questions?"]);
    // ODF page layouts carry no type: SPEC-089 classifies them from their
    // slots, so a section header (a title and one body) reads back as
    // titleAndBody there. .pptx keeps every kind.
    expect(back.slides.map((slide) => layoutFor(back, slide)?.kind)).toEqual([
      "title",
      format === "pptx" ? "sectionHeader" : "titleAndBody",
      "titleAndBody",
      "titleOnly",
    ]);
    expect(back.theme?.name).toBe("Facet");
    expect(back.slides[2].notes).toEqual(["Lead with the date."]);
    const custom = back.properties?.customProperties ?? [];
    expect(custom).toEqual(
      expect.arrayContaining([
        { name: "GeneratedBy", type: "text", value: "docu-edit" },
        { name: "GeneratedAt", type: "text", value: "2026-10-24T12:00:00.000Z" },
        { name: "GeneratedModel", type: "text", value: "stub-model" },
        { name: "GeneratedPrompt", type: "text", value: "A board update" },
      ]),
    );
  });
});

describe("the .pptx writer's custom-properties part", () => {
  it("is written only when the deck has custom properties", async () => {
    const plain = createBlankPresentation();
    const names = [...(await unzip((await writePptx(plain)).bytes)).keys()];
    expect(names).not.toContain("docProps/custom.xml");

    const designed = buildPlannedDeck(PLAN, { generatedAt: "2026-10-24T12:00:00.000Z", generatedModel: "outline-designer" });
    const entries = await unzip((await writePptx(designed)).bytes);
    const text = (name: string) => new TextDecoder().decode(entries.get(name));
    expect(text("docProps/custom.xml")).toContain('name="GeneratedModel"><vt:lpwstr>outline-designer</vt:lpwstr>');
    expect(text("[Content_Types].xml")).toContain('PartName="/docProps/custom.xml"');
    expect(text("_rels/.rels")).toContain('Target="docProps/custom.xml"');
  });
});
