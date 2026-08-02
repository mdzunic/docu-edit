import { describe, it, expect } from "vitest";
import { formatTimestamp } from "./format";

describe("formatTimestamp", () => {
  it("formats a date as 'YYYY-MM-DD HH:mm' in local time", () => {
    expect(formatTimestamp(new Date(2026, 7, 2, 14, 35, 59))).toBe(
      "2026-08-02 14:35",
    );
  });

  it("zero-pads month, day, hours and minutes", () => {
    expect(formatTimestamp(new Date(2026, 0, 5, 9, 7, 0))).toBe(
      "2026-01-05 09:07",
    );
  });
});
