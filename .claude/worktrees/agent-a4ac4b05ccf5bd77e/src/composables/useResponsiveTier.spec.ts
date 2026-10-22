import { describe, expect, it } from "vitest";
import {
  COMPACT_MAX_WIDTH,
  FULL_MIN_WIDTH,
  tierForWidth,
} from "./useResponsiveTier";

describe("tierForWidth", () => {
  it("is compact below 640px (AC-1)", () => {
    expect(tierForWidth(0)).toBe("compact");
    expect(tierForWidth(390)).toBe("compact");
    expect(tierForWidth(639)).toBe("compact");
  });

  it("is medium from 640 to 1023px inclusive (AC-1)", () => {
    expect(tierForWidth(640)).toBe("medium");
    expect(tierForWidth(768)).toBe("medium");
    expect(tierForWidth(1023)).toBe("medium");
  });

  it("is full at 1024px and above (AC-1)", () => {
    expect(tierForWidth(1024)).toBe("full");
    expect(tierForWidth(1440)).toBe("full");
    expect(tierForWidth(3840)).toBe("full");
  });

  it("keeps the boundaries at the documented widths", () => {
    expect(COMPACT_MAX_WIDTH).toBe(640);
    expect(FULL_MIN_WIDTH).toBe(1024);
    // The boundaries are exclusive/inclusive exactly as the tier table states:
    // compact is <640, full is ≥1024, medium is everything between.
    expect(tierForWidth(COMPACT_MAX_WIDTH - 1)).toBe("compact");
    expect(tierForWidth(COMPACT_MAX_WIDTH)).toBe("medium");
    expect(tierForWidth(FULL_MIN_WIDTH - 1)).toBe("medium");
    expect(tierForWidth(FULL_MIN_WIDTH)).toBe("full");
  });
});
