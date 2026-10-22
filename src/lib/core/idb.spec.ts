// SPEC-137 AC-66: core/idb.ts's range reads, range deletes and writeBatch,
// against an in-memory IndexedDB (fakeIndexedDb.ts) that enforces a real
// transaction's activity rules — so a batch that issued an operation outside
// a success callback would throw here as it would in a browser.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { openIdbDatabase } from "./idb";
import { FakeIndexedDb } from "./fakeIndexedDb";

let fake: FakeIndexedDb;

beforeEach(() => {
  fake = new FakeIndexedDb();
  fake.install();
});

afterEach(() => {
  fake.uninstall();
});

function open() {
  return openIdbDatabase("test", 1, [{ name: "documents", keyPath: "id" }]);
}

async function seed(db: ReturnType<typeof open>, ids: string[]) {
  await db.writeBatch("documents", ids.map((id) => ({ put: { id } })));
  fake.log.length = 0;
}

describe("getAllInRange", () => {
  it("reads the closed range by default, either bound open on request, and either side unbounded", async () => {
    const db = open();
    await seed(db, ["a", "b", "c", "d", "e"]);
    const ids = async (lower?: string, upper?: string, bounds?: { lowerOpen?: boolean; upperOpen?: boolean }) =>
      (await db.getAllInRange<{ id: string }>("documents", lower, upper, bounds)).map((record) => record.id);
    expect(await ids("b", "d")).toEqual(["b", "c", "d"]);
    expect(await ids("b", "d", { lowerOpen: true })).toEqual(["c", "d"]);
    expect(await ids("b", "d", { upperOpen: true })).toEqual(["b", "c"]);
    expect(await ids("b", "d", { lowerOpen: true, upperOpen: true })).toEqual(["c"]);
    expect(await ids(undefined, "c", { upperOpen: true })).toEqual(["a", "b"]);
    expect(await ids("c", undefined, { lowerOpen: true })).toEqual(["d", "e"]);
    expect(await ids()).toEqual(["a", "b", "c", "d", "e"]);
  });
});

describe("deleteRange", () => {
  it("deletes every key in [lower, upper] and nothing else — a prefix range stops at its colon", async () => {
    const db = open();
    await seed(db, ["sheetrows:d:s:0000", "sheetrows:d:s:0001", "sheetrows:d:sx:0000", "sheetrows:d:s", "sheetsource:d"]);
    await db.deleteRange("documents", "sheetrows:d:s:", "sheetrows:d:s:\uffff");
    expect([...fake.stores.get("documents")!.keys()].toSorted()).toEqual(["sheetrows:d:s", "sheetrows:d:sx:0000", "sheetsource:d"]);
  });
});

describe("writeBatch", () => {
  it("issues each operation only after the previous one's success, all in one transaction", async () => {
    const db = open();
    await seed(db, ["old", "range:1", "range:2", "keep"]);
    await db.writeBatch("documents", [
      { put: { id: "a", n: 1 } },
      { delete: "old" },
      { deleteRange: ["range:", "range:\uffff"] },
      { put: { id: "b", n: 2 } },
    ]);
    const events = fake.log.filter((event) => event.kind === "issue" || event.kind === "success");
    expect(events.map((event) => `${event.kind}:${event.key}`)).toEqual([
      "issue:a",
      "success:a",
      "issue:old",
      "success:old",
      "issue:range:..range:\uffff",
      "success:range:..range:\uffff",
      "issue:b",
      "success:b",
    ]);
    expect(new Set(events.map((event) => event.transaction)).size).toBe(1);
    expect(fake.log.at(-1)?.kind).toBe("complete");
    expect([...fake.stores.get("documents")!.keys()].toSorted()).toEqual(["a", "b", "keep"]);
  });

  it("rejects and rolls the whole batch back when one operation fails", async () => {
    const db = open();
    await seed(db, ["x"]);
    fake.failPutPrefix = "bad";
    await expect(
      db.writeBatch("documents", [{ put: { id: "a" } }, { delete: "x" }, { put: { id: "bad" } }, { put: { id: "c" } }]),
    ).rejects.toThrow();
    expect([...fake.stores.get("documents")!.keys()]).toEqual(["x"]);
    expect(fake.log.some((event) => event.kind === "issue" && event.key === "c")).toBe(false);
  });

  it("resolves at once for an empty batch, opening no transaction", async () => {
    const db = open();
    await seed(db, []);
    await db.writeBatch("documents", []);
    expect(fake.log).toEqual([]);
  });
});
