import { describe, it, expect } from "vitest";
import { documentExists } from "./documents";

describe("documentExists", () => {
  it("treats every routable id as existing (SPEC-001: no persistence yet)", () => {
    expect(documentExists("abc")).toBe(true);
    expect(documentExists("123e4567-e89b")).toBe(true);
    expect(documentExists("with spaces and $ymbols")).toBe(true);
  });

  it("rejects the empty string, which the /doc/:id route can never produce", () => {
    expect(documentExists("")).toBe(false);
  });
});
