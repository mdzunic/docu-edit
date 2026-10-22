import { describe, it, expect } from "vitest";
import { createDocument, type Paragraph } from "./documentModel";
import { documentLibraryStats, WORDS_PER_PAGE } from "./libraryStats";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

describe("documentLibraryStats (SPEC-067)", () => {
  it("counts body words", () => {
    const model = createDocument("doc");
    model.sections[0].blocks = [paragraph("one two three four five")];
    expect(documentLibraryStats(model).words).toBe(5);
  });

  it("estimates at least one page for an empty document", () => {
    const model = createDocument("empty");
    expect(documentLibraryStats(model)).toEqual({ words: 0, pages: 1 });
  });

  it("rounds the page estimate up from the word count", () => {
    const model = createDocument("long");
    const words = WORDS_PER_PAGE * 2 + 1;
    model.sections[0].blocks = [paragraph(Array(words).fill("word").join(" "))];
    const stats = documentLibraryStats(model);
    expect(stats.words).toBe(words);
    expect(stats.pages).toBe(3);
  });
});
