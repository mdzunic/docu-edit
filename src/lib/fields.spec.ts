import { describe, it, expect } from "vitest";
import { createDocument, type DocumentModel, type Paragraph, type FieldSpec } from "./documentModel";
import { insertBookmark } from "./links";
import { caretSelection } from "./commands";
import {
  DATE_TIME_CATALOGUE,
  REF_SOURCE_NOT_FOUND,
  canonicalDateTimeFormatFromDocx,
  evaluateField,
  formatFromOdtDateStyleName,
  isReadOnlyField,
  isSupportedDateTimeFormat,
  odtDateStyleName,
  renderDateTime,
  type FieldContext,
} from "./fields";
import { now, setClock } from "./clock";

// A fixed local instant: Wednesday 4 March 2026, 14:30:00 — the spec's own
// worked example (en-GB). Built from local components so getDate()/getHours()
// are timezone-independent.
const INSTANT = new Date(2026, 2, 4, 14, 30, 0);

function docWith(paragraphs: Paragraph[], patch: Partial<DocumentModel> = {}): DocumentModel {
  const model = createDocument("fields-test");
  model.sections = [{ blocks: paragraphs }];
  model.defaultStyles.language = "en-GB";
  return { ...model, ...patch };
}

function para(text: string, extra: Partial<Paragraph> = {}): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }], ...extra };
}

function ctx(model: DocumentModel, extra: Partial<FieldContext> = {}): FieldContext {
  return { model, now: INSTANT, fileName: "report.docx", ...extra };
}

describe("clock (SPEC-053 Determinism, AC-49)", () => {
  it("setClock pins now() to a fixed instant and null restores the wall clock", () => {
    setClock(INSTANT);
    expect(now().getTime()).toBe(INSTANT.getTime());
    setClock(null);
    expect(now().getTime()).not.toBe(INSTANT.getTime());
  });
});

describe("renderDateTime — the closed catalogue (SPEC-053 AC-2)", () => {
  const cases: Array<[string, string]> = [
    ["d MMMM yyyy", "4 March 2026"],
    ["d MMM yyyy", "4 Mar 2026"],
    ["dd/MM/yyyy", "04/03/2026"],
    ["MM/dd/yyyy", "03/04/2026"],
    ["yyyy-MM-dd", "2026-03-04"],
    ["EEEE, d MMMM yyyy", "Wednesday, 4 March 2026"],
    ["HH:mm", "14:30"],
    ["h:mm a", "2:30 PM"],
    ["HH:mm:ss", "14:30:00"],
  ];
  for (const [format, expected] of cases) {
    it(`renders "${format}" as "${expected}" (en-GB)`, () => {
      expect(renderDateTime(format, INSTANT, "en-GB")).toBe(expected);
    });
  }

  it("renders every catalogue entry (no gaps)", () => {
    for (const entry of DATE_TIME_CATALOGUE) {
      expect(renderDateTime(entry.format, INSTANT, "en-GB")).not.toBeNull();
    }
  });

  it("returns null for a picture outside the catalogue", () => {
    expect(renderDateTime("MMMM/yy", INSTANT, "en-GB")).toBeNull();
    expect(isSupportedDateTimeFormat("MMMM/yy")).toBe(false);
  });
});

describe("evaluateField (SPEC-053)", () => {
  it("date field renders the clock; time field too", () => {
    const model = docWith([para("")]);
    expect(evaluateField({ kind: "date", format: "d MMMM yyyy", cached: "old" }, ctx(model))).toBe("4 March 2026");
    expect(evaluateField({ kind: "time", format: "HH:mm", cached: "old" }, ctx(model))).toBe("14:30");
  });

  it("a date field with an unsupported picture shows its cached value read-only (AC-42)", () => {
    const field: FieldSpec = { kind: "date", format: "MMMM/yy", cached: "March/26" };
    expect(evaluateField(field, ctx(docWith([para("")])))).toBe("March/26");
    expect(isReadOnlyField(field)).toBe(true);
  });

  it("metadata fields read SPEC-041 properties", () => {
    const model = docWith([para("")], {
      properties: { author: "Ada", title: "Spec", subject: "Fields", keywords: "docx,odt" },
    });
    expect(evaluateField({ kind: "author", cached: "" }, ctx(model))).toBe("Ada");
    expect(evaluateField({ kind: "docTitle", cached: "" }, ctx(model))).toBe("Spec");
    expect(evaluateField({ kind: "docSubject", cached: "" }, ctx(model))).toBe("Fields");
    expect(evaluateField({ kind: "docKeywords", cached: "" }, ctx(model))).toBe("docx,odt");
  });

  it("docProperty reads a custom property by name", () => {
    const model = docWith([para("")], {
      properties: { customProperties: [{ name: "Client", type: "text", value: "Acme" }] },
    });
    expect(evaluateField({ kind: "docProperty", name: "Client", cached: "" }, ctx(model))).toBe("Acme");
    expect(evaluateField({ kind: "docProperty", name: "Missing", cached: "" }, ctx(model))).toBe("");
  });

  it("fileName honours the path option", () => {
    const model = docWith([para("")]);
    expect(evaluateField({ kind: "fileName", cached: "" }, ctx(model, { fileName: "/docs/report.docx" }))).toBe(
      "report.docx",
    );
    expect(
      evaluateField({ kind: "fileName", path: true, cached: "" }, ctx(model, { fileName: "/docs/report.docx" })),
    ).toBe("/docs/report.docx");
  });

  it("wordCount reads the SPEC-036 body counter", () => {
    const model = docWith([para("one two three four")]);
    expect(evaluateField({ kind: "wordCount", cached: "" }, ctx(model))).toBe("4");
  });

  it("page/numpages render through the page context", () => {
    const model = docWith([para("")]);
    expect(evaluateField({ kind: "page", format: "decimal" }, ctx(model, { pageNumber: 3, totalPages: 9 }))).toBe("3");
    expect(evaluateField({ kind: "numpages", format: "decimal" }, ctx(model, { pageNumber: 3, totalPages: 9 }))).toBe(
      "9",
    );
    expect(
      evaluateField({ kind: "page", format: "lowerRoman" }, ctx(model, { pageNumber: 4, totalPages: 9 })),
    ).toBe("iv");
  });

  it("unsupported field displays its cached result and is read-only (AC-41, AC-42, AC-43)", () => {
    const field: FieldSpec = { kind: "unsupported", code: 'SEQ Figure \\* ARABIC', cached: "7" };
    expect(evaluateField(field, ctx(docWith([para("")])))).toBe("7");
    expect(isReadOnlyField(field)).toBe(true);
  });
});

describe("cross-references (SPEC-053 AC-23..27)", () => {
  function docWithBookmark(): DocumentModel {
    let model = docWith([para("Intro"), para("Chapter about widgets", { headingLevel: 1 })]);
    // A bookmark named at the start of block 1.
    model = insertBookmark(model, caretSelection({ section: 0, block: 1, offset: 0 }), "widgets").model;
    return model;
  }

  it("shows the target's text", () => {
    const model = docWithBookmark();
    expect(evaluateField({ kind: "ref", target: "widgets", show: "text", cached: "" }, ctx(model))).toBe(
      "Chapter about widgets",
    );
  });

  it("shows the target's page number when a page lookup is provided (AC-24)", () => {
    const model = docWithBookmark();
    const field: FieldSpec = { kind: "pageRef", target: "widgets", show: "page", cached: "" };
    expect(evaluateField(field, ctx(model, { pageOfBookmark: () => 5 }))).toBe("5");
  });

  it("shows above/below position relative to the reference (AC-25)", () => {
    const model = docWithBookmark();
    const field: FieldSpec = { kind: "ref", target: "widgets", show: "position", cached: "" };
    // Reference sitting after the target (block 1) → target is above.
    expect(evaluateField(field, ctx(model, { here: { section: 0, block: 1, offset: 5 } }))).toBe("above");
    // Reference sitting before the target (block 0) → target is below.
    expect(evaluateField(field, ctx(model, { here: { section: 0, block: 0, offset: 0 } }))).toBe("below");
  });

  it("a deleted target renders the Word error string (AC-27)", () => {
    const model = docWith([para("Intro")]);
    expect(evaluateField({ kind: "ref", target: "gone", show: "text", cached: "x" }, ctx(model))).toBe(
      REF_SOURCE_NOT_FOUND,
    );
    expect(evaluateField({ kind: "pageRef", target: "gone", show: "page", cached: "x" }, ctx(model))).toBe(
      REF_SOURCE_NOT_FOUND,
    );
  });
});

describe("format round-trip helpers", () => {
  it("maps docx pictures to canonical patterns and back", () => {
    expect(canonicalDateTimeFormatFromDocx("dddd, d MMMM yyyy")).toBe("EEEE, d MMMM yyyy");
    expect(canonicalDateTimeFormatFromDocx("d MMMM yyyy")).toBe("d MMMM yyyy");
    // An unknown picture is kept verbatim.
    expect(canonicalDateTimeFormatFromDocx("MMMM/yy")).toBe("MMMM/yy");
  });

  it("odt date-style names round-trip to their format", () => {
    for (const entry of DATE_TIME_CATALOGUE) {
      expect(formatFromOdtDateStyleName(odtDateStyleName(entry))).toBe(entry.format);
    }
    expect(formatFromOdtDateStyleName("N42")).toBeNull();
  });
});
