import { describe, it, expect } from "vitest";
import { getDocument, getOrCreateDocument } from "./documentStore";
import { DEFAULT_TITLE } from "./documentModel";

describe("documentStore", () => {
  it("has no document for an id that was never created", () => {
    expect(getDocument("never-seen")).toBeUndefined();
  });

  it("creates and stores a fresh document for a new id", () => {
    const doc = getOrCreateDocument("store-new-id");
    expect(doc.id).toBe("store-new-id");
    expect(doc.title).toBe(DEFAULT_TITLE);
    expect(getDocument("store-new-id")).toBe(doc);
  });

  it("returns the same document instance on repeated lookups", () => {
    const first = getOrCreateDocument("store-repeat-id");
    first.title = "Renamed";
    const second = getOrCreateDocument("store-repeat-id");
    expect(second).toBe(first);
    expect(second.title).toBe("Renamed");
  });
});
