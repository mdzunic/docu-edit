import { describe, it, expect } from "vitest";
import type { LibraryEntry } from "./documentStore";
import { kindOfFormat } from "./documentKinds";
import {
  FILTERABLE_FORMATS,
  FORMAT_CHIPS,
  chipActive,
  chipCount,
  chipUnavailable,
  filterCounts,
  filterLibrary,
  highlightSegments,
  searchLibrary,
  sortLibrary,
  subsequenceMatch,
  toggleChip,
} from "./libraryView";
import { formatBadgeLabel } from "./format";
import type { DocumentFormat } from "./documentStore";

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

const sizeOf = (id: string) => ({ old: 30, new: 10, mid: 20 })[id] ?? 0;

describe("sortLibrary (SPEC-067 AC-11..14, AC-21)", () => {
  const entries = [
    entry({ id: "old", title: "Alpha", lastModified: 100 }),
    entry({ id: "new", title: "Zeta", lastModified: 300 }),
    entry({ id: "mid", title: "Mu", lastModified: 200, starred: true }),
  ];

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

describe("chipUnavailable", () => {
  const chip = FORMAT_CHIPS[0];
  const images = FORMAT_CHIPS.find((c) => c.id === "images")!;
  // Annotated, not cast: every format DocumentFormat declares carries a count
  // here, so a format added to the union fails this file rather than silently
  // counting as zero (which is how the media split slipped past).
  const empty: Record<DocumentFormat, number> = {
    docx: 0,
    odt: 0,
    md: 0,
    pptx: 0,
    odp: 0,
    xlsx: 0,
    ods: 0,
    csv: 0,
    tsv: 0,
    pdf: 0,
    png: 0,
    jpeg: 0,
    gif: 0,
    svg: 0,
    webp: 0,
    new: 0,
  };

  it("disables a filter that stands for nothing in the library", () => {
    // A one-document library offered nine filters, eight reading "0" — nine
    // advertised routes, eight of which could only produce an empty list.
    expect(chipUnavailable(chip, new Set(), empty)).toBe(true);
  });

  it("leaves a filter with matches alone", () => {
    const counts = { ...empty, [chip.formats[0]]: 3 };
    expect(chipUnavailable(chip, new Set(), counts)).toBe(false);
  });

  it("keeps an ACTIVE filter enabled even at zero, so it can be switched off", () => {
    // Otherwise a filter that empties the list would trap the user in it.
    expect(chipUnavailable(chip, new Set(chip.formats), empty)).toBe(false);
  });

  it("sums a grouped chip over every format it stands for (SPEC-107)", () => {
    // "Images" is one click over five formats, so a library holding nothing
    // but .webp must still offer it — counting only the first format would
    // grey out the one filter that works.
    expect(chipUnavailable(images, new Set(), empty)).toBe(true);
    for (const format of images.formats) {
      expect(chipUnavailable(images, new Set(), { ...empty, [format]: 2 })).toBe(false);
    }
  });
});

describe("media in the library view (SPEC-107)", () => {
  const MEDIA = ["pdf", "png", "jpeg", "gif", "svg", "webp"] as const;

  it("makes every media format filterable, behind two kind-grouped chips", () => {
    for (const format of MEDIA) {
      expect(FILTERABLE_FORMATS).toContain(format);
    }
    const pdfs = FORMAT_CHIPS.find((chip) => chip.id === "pdf");
    const images = FORMAT_CHIPS.find((chip) => chip.id === "images");
    expect(pdfs).toMatchObject({ label: "PDFs", formats: ["pdf"] });
    expect(images).toMatchObject({ label: "Images", formats: ["png", "jpeg", "gif", "svg", "webp"] });
    // No media format gets a chip of its own.
    expect(FORMAT_CHIPS.filter((chip) => chip.formats.some((format) => (MEDIA as readonly string[]).includes(format)))).toHaveLength(2);
  });

  it('one click on "Images" filters to every image format, and a second clears it', () => {
    const images = FORMAT_CHIPS.find((chip) => chip.id === "images")!;
    const entries = [
      entry({ id: "d", format: "docx" }),
      entry({ id: "p", format: "pdf" }),
      entry({ id: "a", format: "png" }),
      entry({ id: "b", format: "webp" }),
      entry({ id: "c", format: "svg" }),
    ];
    const on = toggleChip(images, new Set());
    expect(chipActive(images, on)).toBe(true);
    expect(filterLibrary(entries, { search: "", formats: on, starredOnly: false }).map((e) => e.id)).toEqual(["a", "b", "c"]);
    expect(chipCount(images, filterCounts(entries, "").formats)).toBe(3);
    expect(toggleChip(images, on).size).toBe(0);
  });

  it("badges each media format", () => {
    expect(MEDIA.map((format) => formatBadgeLabel(format))).toEqual(["PDF", "PNG", "JPEG", "GIF", "SVG", "WEBP"]);
  });
});

/** A workbook's library entry (SPEC-116). */
function workbookEntry(id: string, format: DocumentFormat, packageVariant?: LibraryEntry["packageVariant"]): LibraryEntry {
  return { id, title: id, format, kind: "spreadsheet", lastModified: 1, starred: false, packageVariant };
}

describe("workbook chips (SPEC-116)", () => {

  it("adds an xlsx and an ods chip right after odp, and makes both filterable", () => {
    const ids = FORMAT_CHIPS.map((chip) => chip.id);
    const odpAt = ids.indexOf("odp");
    expect(ids.slice(odpAt, odpAt + 3)).toEqual(["odp", "xlsx", "ods"]);
    expect(FORMAT_CHIPS.find((chip) => chip.id === "xlsx")).toEqual({ id: "xlsx", label: "XLSX", formats: ["xlsx"] });
    expect(FORMAT_CHIPS.find((chip) => chip.id === "ods")).toEqual({ id: "ods", label: "ODS", formats: ["ods"] });
    expect(FILTERABLE_FORMATS).toContain("xlsx");
    expect(FILTERABLE_FORMATS).toContain("ods");
  });

  it("adds a csv and a tsv chip right after ods, both filterable (SPEC-126)", () => {
    const ids = FORMAT_CHIPS.map((chip) => chip.id);
    const odsAt = ids.indexOf("ods");
    expect(ids.slice(odsAt, odsAt + 3)).toEqual(["ods", "csv", "tsv"]);
    expect(FORMAT_CHIPS.find((chip) => chip.id === "csv")).toEqual({ id: "csv", label: "CSV", formats: ["csv"] });
    expect(FORMAT_CHIPS.find((chip) => chip.id === "tsv")).toEqual({ id: "tsv", label: "TSV", formats: ["tsv"] });
    expect(FILTERABLE_FORMATS).toContain("csv");
    expect(FILTERABLE_FORMATS).toContain("tsv");
  });

  it("filters and counts a workbook's variants with their format", () => {
    const entries = [
      workbookEntry("book", "xlsx"),
      workbookEntry("macros", "xlsx", "macroEnabled"),
      workbookEntry("template", "xlsx", "template"),
      workbookEntry("calc", "ods"),
      workbookEntry("calc-template", "ods", "template"),
    ];
    const xlsx = filterLibrary(entries, { search: "", formats: new Set(["xlsx"]), starredOnly: false });
    expect(xlsx.map((found) => found.id)).toEqual(["book", "macros", "template"]);
    const ods = filterLibrary(entries, { search: "", formats: new Set(["ods"]), starredOnly: false });
    expect(ods.map((found) => found.id)).toEqual(["calc", "calc-template"]);
    const counts = filterCounts(entries, "").formats;
    expect(counts.xlsx).toBe(3);
    expect(counts.ods).toBe(2);
  });
});
