// SPEC-028 — formatting-mark segmentation.
import { describe, expect, it } from "vitest";
import { markSegments } from "./marks";

describe("markSegments", () => {
  it("splits spaces and tabs out of plain text", () => {
    expect(markSegments("a b\tc")).toEqual([
      { text: "a", kind: "text" },
      { text: " ", kind: "space" },
      { text: "b", kind: "text" },
      { text: "\t", kind: "tab" },
      { text: "c", kind: "text" },
    ]);
  });

  it("emits one segment per consecutive space", () => {
    expect(markSegments("a  b").filter((segment) => segment.kind === "space")).toHaveLength(2);
  });

  it("keeps soft line breaks in plain segments", () => {
    expect(markSegments("a\nb")).toEqual([{ text: "a\nb", kind: "text" }]);
  });

  it("concatenates back to the exact input (byte-identical DOM text)", () => {
    const text = "  leading\t and nbsp \t trailing  ";
    expect(markSegments(text).map((segment) => segment.text).join("")).toBe(text);
  });

  it("returns no segments for empty text", () => {
    expect(markSegments("")).toEqual([]);
  });
});
