// SPEC-106 — HTML export through the generation API (SPEC-088's follow-up):
// `PresentationBuilder.toHtml` and the top-level `exportDeckHtml` wrap
// presentation/htmlExport.ts's `exportPresentationHtml` unchanged — the same
// self-contained document — with `includeNotes` defaulting to false and
// passed through explicitly, and a deck with no slides refused at
// validation time with a DocumentGenerationError naming the empty deck,
// never SPEC-088's raw throw.
import { describe, expect, it } from "vitest";
import { exportPresentationHtml } from "../presentation/htmlExport";
import { exportDeckHtml } from "./deckHtml";
import { DocumentGenerationError } from "./errors";
import { PresentationBuilder } from "./presentation";
import { parseSlideMarkdown } from "./slideMarkdown";

const PNG_4X2 = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAAAAABawyK/AAAAC0lEQVR4nGNggAEAAAoAAX+AdF4AAAAASUVORK5CYII=";

function sample(): PresentationBuilder {
  return new PresentationBuilder({ title: "Q3 <review>" })
    .addTitleSlide("Q3 review", "Prepared 30 September")
    .setNotes(["Lead with Europe."])
    .addBulletSlide("Highlights", ["Revenue up 12%", ["Europe up 20%"]])
    .addImageSlide("Pipeline", { data: PNG_4X2, contentType: "image/png" });
}

describe("PresentationBuilder.toHtml (AC-20)", () => {
  it("is exportPresentationHtml's output, unchanged, with includeNotes false by default", () => {
    const builder = sample();
    const html = builder.toHtml();
    expect(html).toBe(exportPresentationHtml(builder.toDeck(), { includeNotes: false }));
    expect(html).toContain('data-notes="0"');
    expect(html).not.toContain("Lead with Europe.");
  });

  it("passes includeNotes through", () => {
    const builder = sample();
    const html = builder.toHtml({ includeNotes: true });
    expect(html).toBe(exportPresentationHtml(builder.toDeck(), { includeNotes: true }));
    expect(html).toContain('data-notes="1"');
    expect(html).toContain("Lead with Europe.");
    expect(builder.toHtml({ includeNotes: false })).toBe(builder.toHtml());
  });

  it("is one self-contained document: every slide, images inline, no external request", () => {
    const html = sample().toHtml();
    expect(html.startsWith("<!doctype html>")).toBe(true);
    expect(html.match(/<section /g)).toHaveLength(3);
    expect(html).toContain(`data:image/png;base64,${PNG_4X2}`);
    expect(html).not.toMatch(/(?:src|href)="(?:https?:)?\/\//);
    expect(html).toContain("<title>Q3 &lt;review&gt;</title>");
  });

  it("refuses a deck with no slides at validation time, naming the empty deck", () => {
    const empty = new PresentationBuilder({ title: "Nothing yet" });
    expect(() => empty.toHtml()).toThrow(DocumentGenerationError);
    expect(() => empty.toHtml()).toThrow(/^deck: an empty deck \(no slides\) cannot be exported to HTML/);
    expect(() => empty.toHtml()).not.toThrow("Cannot export a presentation with no slides");
  });

  it("validates its options", () => {
    expect(() => sample().toHtml({ includeNotes: "yes" } as never)).toThrow("options.includeNotes: expected a boolean");
    expect(() => sample().toHtml({ notes: true } as never)).toThrow("options.notes: unknown property");
  });
});

describe("exportDeckHtml (AC-20)", () => {
  it("exports any deck — here one parsed from slide Markdown — the same way", () => {
    const deck = parseSlideMarkdown("# From Markdown\n---\n## Rows\n- one\n\nNotes: Speak.");
    expect(exportDeckHtml(deck)).toBe(exportPresentationHtml(deck, { includeNotes: false }));
    expect(exportDeckHtml(deck, { includeNotes: true })).toBe(exportPresentationHtml(deck, { includeNotes: true }));
  });

  it("refuses an empty deck and a non-deck", () => {
    expect(() => exportDeckHtml(parseSlideMarkdown(""))).toThrow(DocumentGenerationError);
    expect(() => exportDeckHtml(null as never)).toThrow("deck: expected an object");
    expect(() => exportDeckHtml({ slides: "none" } as never)).toThrow("deck.slides: expected an array");
  });
});
