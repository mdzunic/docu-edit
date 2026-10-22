import { describe, it, expect } from "vitest";
import { deleteRecord, loadAllRecords, putRecord } from "./idb";
import { createDocument } from "./documentModel";

// jsdom has no `indexedDB` global, which is exactly the "storage unavailable"
// environment (AC-21): every operation must reject cleanly (never throw
// synchronously or hang) so documentStore can fall back to in-memory mode.
// The happy path against a real IndexedDB is covered by e2e/SPEC-003.spec.ts.
describe("idb without IndexedDB available", () => {
  it("loadAllRecords rejects", async () => {
    await expect(loadAllRecords()).rejects.toBeTruthy();
  });

  it("putRecord rejects", async () => {
    const record = {
      id: "x",
      format: "new" as const,
      lastModified: 0,
      model: createDocument("x"),
    };
    await expect(putRecord(record)).rejects.toBeTruthy();
  });

  it("deleteRecord rejects", async () => {
    await expect(deleteRecord("x")).rejects.toBeTruthy();
  });
});
