import { describe, expect, it } from "vitest";
import { decimalSeparatorFor } from "./language";

describe("SPEC-054 — decimalSeparatorFor (AC-19)", () => {
  it("returns '.' for English and undefined (the product default)", () => {
    expect(decimalSeparatorFor("en-US")).toBe(".");
    expect(decimalSeparatorFor("en-GB")).toBe(".");
    expect(decimalSeparatorFor(undefined)).toBe(".");
  });

  it("returns ',' for locales that write decimals with a comma", () => {
    expect(decimalSeparatorFor("de-DE")).toBe(",");
    expect(decimalSeparatorFor("fr-FR")).toBe(",");
    expect(decimalSeparatorFor("cs-CZ")).toBe(",");
  });

  it("falls back to '.' for an unresolvable tag", () => {
    expect(decimalSeparatorFor("not-a-real-tag-xyz")).toBe(".");
  });
});
