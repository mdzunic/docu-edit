// SPEC-103 — the version-chain arithmetic SPEC-040 wrote for the word
// processor, moved into core so the deck's history chains, budgets and prunes
// by the same rules: chain gathering, the four baseline triggers in order,
// prune-candidate selection, the byte measure and the digest.
import { describe, expect, it } from "vitest";
import {
  CHAIN_DELTA_CAP,
  HISTORY_BUDGET_BYTES,
  HISTORY_PAUSED_MESSAGE,
  RETAINED_RECENT_COUNT,
  WHOLESALE_REWRITE_RATIO,
  decideBaseline,
  emptyVersionsIndex,
  gatherChainInfo,
  isRetained,
  measureBytes,
  oldestPrunable,
  oldestRetainedDescendant,
  sha256Hex,
  versionBlobKey,
  versionPayloadKey,
  versionsIndexKey,
  type VersionIndexEntry,
} from "./versionChain";

function entry(id: string, kind: "baseline" | "delta", parentId: string | null, size: number, name?: string): VersionIndexEntry {
  return { id, timestamp: 0, author: "You", words: 0, kind, parentId, size, ...(name ? { name } : {}) };
}

/** A linear chain: one baseline then `deltas` deltas, each on the previous. */
function chain(deltas: number, baselineSize = 1000, deltaSize = 10): VersionIndexEntry[] {
  const entries = [entry("v0", "baseline", null, baselineSize)];
  for (let i = 1; i <= deltas; i++) {
    entries.push(entry(`v${i}`, "delta", `v${i - 1}`, deltaSize));
  }
  return entries;
}

describe("constants", () => {
  it("keep SPEC-040's numbers, the budget as a 50 MB literal", () => {
    expect(HISTORY_BUDGET_BYTES).toBe(50 * 1024 * 1024);
    expect(CHAIN_DELTA_CAP).toBe(50);
    expect(WHOLESALE_REWRITE_RATIO).toBe(0.5);
    expect(RETAINED_RECENT_COUNT).toBe(10);
    expect(HISTORY_PAUSED_MESSAGE).toBe(
      "Version history paused — 50 MB used. Delete a version, or remove a name, to resume",
    );
  });

  it("lay out every product's records under the same three key prefixes", () => {
    expect(versionsIndexKey("d")).toBe("versions:d");
    expect(versionPayloadKey("d", "v")).toBe("version:d:v");
    expect(versionBlobKey("d", "abc")).toBe("blob:d:abc");
    expect(emptyVersionsIndex("d")).toEqual({ id: "versions:d", store: "index", docId: "d", versions: [], paused: false });
  });
});

describe("gatherChainInfo", () => {
  it("walks back from the tail to its baseline, summing the deltas", () => {
    expect(gatherChainInfo(chain(3, 1000, 10), "v3")).toEqual({
      baselineId: "v0",
      baselineBytes: 1000,
      deltaCount: 3,
      accumulatedDeltaBytes: 30,
    });
  });

  it("stops at the nearest baseline, not the oldest", () => {
    const entries = [...chain(2), entry("b2", "baseline", null, 500), entry("d", "delta", "b2", 7)];
    expect(gatherChainInfo(entries, "d")).toEqual({ baselineId: "b2", baselineBytes: 500, deltaCount: 1, accumulatedDeltaBytes: 7 });
  });

  it("is null for an unknown tail or a broken link", () => {
    expect(gatherChainInfo(chain(1), "nope")).toBeNull();
    expect(gatherChainInfo([entry("d", "delta", "gone", 5)], "d")).toBeNull();
    expect(gatherChainInfo([entry("d", "delta", null, 5)], "d")).toBeNull();
  });
});

describe("decideBaseline — the four triggers, in order", () => {
  const info = { baselineId: "v0", baselineBytes: 1000, deltaCount: 3, accumulatedDeltaBytes: 30 };

  it("1: no chain to extend", () => {
    expect(decideBaseline(null, 1)).toBe(true);
  });

  it("2: the chain is at the 50-delta cap", () => {
    expect(decideBaseline({ ...info, deltaCount: CHAIN_DELTA_CAP }, 1)).toBe(true);
    expect(decideBaseline({ ...info, deltaCount: CHAIN_DELTA_CAP - 1 }, 1)).toBe(false);
  });

  it("3: the chain would outweigh a copy of its baseline", () => {
    expect(decideBaseline({ ...info, accumulatedDeltaBytes: 900 }, 100)).toBe(true);
    expect(decideBaseline({ ...info, accumulatedDeltaBytes: 899 }, 100)).toBe(false);
  });

  it("4: the delta alone is a wholesale rewrite", () => {
    expect(decideBaseline(info, 500)).toBe(true);
    expect(decideBaseline(info, 499)).toBe(false);
  });

  it("stays a delta when none holds", () => {
    expect(decideBaseline(info, 10)).toBe(false);
  });
});

describe("prune-candidate selection", () => {
  it("retains named versions and the newest ten", () => {
    const entries = chain(11); // 12 versions
    expect(isRetained(entries, entries[0])).toBe(false);
    expect(isRetained(entries, entries[1])).toBe(false);
    expect(isRetained(entries, entries[2])).toBe(true);
    expect(isRetained(entries, entries[11])).toBe(true);
    const named = entries.map((e, i) => (i === 0 ? { ...e, name: "Keep" } : e));
    expect(isRetained(named, named[0])).toBe(true);
  });

  it("picks the oldest unretained version, or none", () => {
    const entries = chain(11);
    expect(oldestPrunable(entries)?.id).toBe("v0");
    const named = entries.map((e, i) => (i === 0 ? { ...e, name: "Keep" } : e));
    expect(oldestPrunable(named)?.id).toBe("v1");
    expect(oldestPrunable(chain(9))).toBeNull();
  });

  it("finds the oldest retained version whose chain passes through a pruned one", () => {
    const entries = chain(11);
    // v1 is prunable; v2 is its child and retained (inside the newest ten).
    expect(oldestRetainedDescendant(entries, "v0")?.id).toBe("v2");
    expect(oldestRetainedDescendant(entries, "v1")?.id).toBe("v2");
    expect(oldestRetainedDescendant(entries, "v11")).toBeNull();
  });
});

describe("measureBytes and sha256Hex", () => {
  it("measures the UTF-8 bytes of a value's JSON", () => {
    expect(measureBytes("abc")).toBe(5);
    expect(measureBytes({ a: 1 })).toBe(7);
    expect(measureBytes("é")).toBe(4);
  });

  it("digests a string's UTF-8 bytes as lower-case hex SHA-256", async () => {
    expect(await sha256Hex("abc")).toBe("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
    expect(await sha256Hex("")).toBe("e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855");
  });
});
