import { describe, expect, it } from "vitest";
import { formatAppliesTo, parseAppliesTo } from "./cfRangesInput";

describe("parseAppliesTo / formatAppliesTo", () => {
  it("formats ranges back as space-separated A1 text", () => {
    expect(formatAppliesTo([{ top: 1, left: 1, bottom: 12, right: 1 }])).toBe("B2:B13");
    expect(
      formatAppliesTo([
        { top: 1, left: 1, bottom: 12, right: 1 },
        { top: 0, left: 3, bottom: 0, right: 3 },
      ]),
    ).toBe("B2:B13 D1");
  });

  it("parses space-separated ranges", () => {
    expect(parseAppliesTo("B2:B13 D1")).toEqual({
      ranges: [
        { top: 1, left: 1, bottom: 12, right: 1 },
        { top: 0, left: 3, bottom: 0, right: 3 },
      ],
    });
  });

  it("parses comma-separated ranges, tolerating extra whitespace", () => {
    expect(parseAppliesTo("B2:B13, D1")).toEqual({
      ranges: [
        { top: 1, left: 1, bottom: 12, right: 1 },
        { top: 0, left: 3, bottom: 0, right: 3 },
      ],
    });
  });

  it("refuses an empty range list", () => {
    expect(parseAppliesTo("   ")).toEqual({ error: "Enter at least one range" });
  });

  it("refuses an invalid entry, naming it", () => {
    expect(parseAppliesTo("B2:B13 not-a-range")).toEqual({ error: 'Not a valid range: "not-a-range"' });
  });
});
