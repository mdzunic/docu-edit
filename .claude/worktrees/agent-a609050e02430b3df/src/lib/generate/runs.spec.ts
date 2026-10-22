import { describe, expect, it } from "vitest";
import { buildRun, buildRuns, buildRunsFromContent, contentFromTextOrRuns } from "./runs";
import { DocumentGenerationError } from "./errors";

describe("buildRun", () => {
  it("builds an unformatted run from a plain string", () => {
    expect(buildRun("hello", "run")).toEqual({ text: "hello", formatting: {} });
  });

  it("builds a formatted run from an object", () => {
    expect(
      buildRun(
        { text: "hi", bold: true, italic: true, underline: true, strikethrough: true, color: "#FF0000" },
        "run",
      ),
    ).toEqual({
      text: "hi",
      formatting: { bold: true, italic: true, underline: true, strikethrough: true, color: "#ff0000" },
    });
  });

  it("requires text on an object run", () => {
    expect(() => buildRun({ bold: true }, "run")).toThrow(DocumentGenerationError);
  });

  it("rejects an unknown property (AC-9)", () => {
    expect(() => buildRun({ text: "hi", bold2: true }, "run")).toThrow(/unknown property/);
  });

  it("rejects a non-hex color", () => {
    expect(() => buildRun({ text: "hi", color: "red" }, "run")).toThrow(DocumentGenerationError);
  });

  it("rejects a non-positive fontSizePt", () => {
    expect(() => buildRun({ text: "hi", fontSizePt: 0 }, "run")).toThrow(DocumentGenerationError);
  });
});

describe("buildRuns", () => {
  it("builds an array of runs, mixing strings and objects", () => {
    expect(buildRuns(["a", { text: "b", bold: true }], "runs")).toEqual([
      { text: "a", formatting: {} },
      { text: "b", formatting: { bold: true } },
    ]);
  });

  it("rejects a non-array", () => {
    expect(() => buildRuns("not an array", "runs")).toThrow(DocumentGenerationError);
  });
});

describe("buildRunsFromContent", () => {
  it("accepts a plain string", () => {
    expect(buildRunsFromContent("hi", "content")).toEqual([{ text: "hi", formatting: {} }]);
  });

  it("accepts an array of runs", () => {
    expect(buildRunsFromContent(["a", "b"], "content")).toEqual([
      { text: "a", formatting: {} },
      { text: "b", formatting: {} },
    ]);
  });

  it("accepts an object with text", () => {
    expect(buildRunsFromContent({ text: "hi" }, "content")).toEqual([{ text: "hi", formatting: {} }]);
  });

  it("accepts an object with runs", () => {
    expect(buildRunsFromContent({ runs: ["a"] }, "content")).toEqual([{ text: "a", formatting: {} }]);
  });

  it("rejects both text and runs together", () => {
    expect(() => buildRunsFromContent({ text: "a", runs: ["b"] }, "content")).toThrow(/either "text" or "runs"/);
  });

  it("rejects an unknown property on the object form", () => {
    expect(() => buildRunsFromContent({ text: "a", foo: 1 }, "content")).toThrow(/unknown property/);
  });

  it("returns no runs for an object with neither text nor runs (an empty paragraph)", () => {
    expect(buildRunsFromContent({}, "content")).toEqual([]);
  });
});

describe("contentFromTextOrRuns", () => {
  it("resolves the raw text/runs content out of a larger object, ignoring its other keys", () => {
    expect(contentFromTextOrRuns({ type: "paragraph", text: "hi", alignment: "left" }, "block")).toBe("hi");
    expect(contentFromTextOrRuns({ type: "paragraph", runs: ["a", "b"] }, "block")).toEqual(["a", "b"]);
    expect(contentFromTextOrRuns({ type: "paragraph" }, "block")).toEqual([]);
  });

  it("rejects both text and runs together", () => {
    expect(() => contentFromTextOrRuns({ text: "a", runs: ["b"] }, "block")).toThrow(/either "text" or "runs"/);
  });
});
