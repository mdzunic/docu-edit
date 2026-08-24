import { describe, it, expect } from "vitest";
import type { LibraryEntry } from "./documentStore";
import { kindOfFormat } from "./documentKinds";
import {
  filterCounts,
  filterLibrary,
  highlightSegments,
  searchLibrary,
  sortLibrary,
  subsequenceMatch,
} from "./libraryView";

function entry(over: Partial<LibraryEntry> & { id: string }): LibraryEntry {
  const format = over.format ?? "new";
  return {
    id: over.id,
    title: over.title ?? over.id,
    format,
    kind: over.kind ?? kindOfFormat(format),
    lastModified: over.lastModified ?? 0,
    starred: over.starred ?? false,
    trashedAt: over.trashedAt,
  };
}

describe("subsequenceMatch (SPEC-067 AC-8)", () => {
  it("matches characters in order, not necessarily adjacent", () => {
    expect(subsequenceMatch("rpt", "Report")).toEqual([0, 2, 5]);
  });

  it("is case-insensitive", () => {
    expect(subsequenceMatch("RE", "Report")).toEqual([0, 1]);
  });

  it("returns null when the query is not a subsequence", () => {
    expect(subsequenceMatch("xyz", "Report")).toBeNull();
    expect(subsequenceMatch("tr", "Report")).toBeNull();
  });

  it("matches everything with an empty query", () => {
    expect(subsequenceMatch("", "Report")).toEqual([]);
  });
});

describe("highlightSegments (SPEC-067 AC-9)", () => {
  it("splits matched characters into their own segments", () => {
    const indices = subsequenceMatch("rp", "Report")!;
    expect(highlightSegments("Report", indices)).toEqual([
      { text: "R", match: true },
      { text: "e", match: false },
      { text: "p", match: true },
      { text: "ort", match: false },
    ]);
  });

  it("merges consecutive matches", () => {
    const indices = subsequenceMatch("re", "Report")!;
    expect(highlightSegments("Report", indices)).toEqual([
      { text: "Re", match: true },
      { text: "port", match: false },
    ]);
  });

  it("returns a single unmatched segment when nothing matched", () => {
    expect(highlightSegments("Report", [])).toEqual([
      { text: "Report", match: false },
    ]);
  });
});

describe("filterLibrary (SPEC-067)", () => {
  const entries = [
    entry({ id: "a", title: "Annual Report", format: "docx", starred: true }),
    entry({ id: "b", title: "Budget", format: "odt" }),
    entry({ id: "c", title: "Rough notes", format: "md" }),
  ];

  it("filters by subsequence search (AC-8, AC-10)", () => {
    expect(searchLibrary(entries, "rpt").map((e) => e.id)).toEqual(["a"]);
    // Clearing restores everything.
    expect(searchLibrary(entries, "").map((e) => e.id)).toEqual(["a", "b", "c"]);
  });

  it("filters by format chips (AC-16)", () => {
    const out = filterLibrary(entries, {
      search: "",
      formats: new Set(["odt", "md"]),
      starredOnly: false,
    });
    expect(out.map((e) => e.id)).toEqual(["b", "c"]);
  });

  it("filters by starred (AC-17)", () => {
    const out = filterLibrary(entries, {
      search: "",
      formats: new Set(),
      starredOnly: true,
    });
    expect(out.map((e) => e.id)).toEqual(["a"]);
  });
});

describe("filterCounts (SPEC-067 AC-16, AC-17)", () => {
  it("counts formats and starred over the search-filtered set", () => {
    const entries = [
      entry({ id: "a", title: "Report one", format: "docx", starred: true }),
      entry({ id: "b", title: "Report two", format: "docx" }),
      entry({ id: "c", title: "Memo", format: "odt", starred: true }),
    ];
    const all = filterCounts(entries, "");
    expect(all.formats.docx).toBe(2);
    expect(all.formats.odt).toBe(1);
    expect(all.starred).toBe(2);

    const searched = filterCounts(entries, "report");
    expect(searched.formats.docx).toBe(2);
    expect(searched.formats.odt).toBe(0);
    expect(searched.starred).toBe(1);
  });
});

describe("sortLibrary (SPEC-067 AC-11..14, AC-21)", () => {
  const entries = [
    entry({ id: "old", title: "Alpha", lastModified: 100 }),
    entry({ id: "new", title: "Zeta", lastModified: 300 }),
    entry({ id: "mid", title: "Mu", lastModified: 200, starred: true }),
  ];
  const sizeOf = (id: string) => ({ old: 30, new: 10, mid: 20 })[id] ?? 0;

  it("sorts by last modified newest first by default, starred first (AC-11, AC-21)", () => {
    const out = sortLibrary(entries, { key: "modified", direction: "desc" }, sizeOf);
    expect(out.map((e) => e.id)).toEqual(["mid", "new", "old"]);
  });

  it("sorts by title A–Z (AC-12)", () => {
    const out = sortLibrary(entries, { key: "title", direction: "asc" }, sizeOf);
    // starred "mid" first, then Alpha, Zeta
    expect(out.map((e) => e.id)).toEqual(["mid", "old", "new"]);
  });

  it("sorts by size and honours direction (AC-13, AC-14)", () => {
    const desc = sortLibrary(entries, { key: "size", direction: "desc" }, sizeOf);
    expect(desc.map((e) => e.id)).toEqual(["mid", "old", "new"]);
    const asc = sortLibrary(entries, { key: "size", direction: "asc" }, sizeOf);
    expect(asc.map((e) => e.id)).toEqual(["mid", "new", "old"]);
  });
});
