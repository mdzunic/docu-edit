import { describe, it, expect } from "vitest";
import { createDocument } from "../documentModel";
import { documentStats, emptyCountStats } from "../wordCount";
import { computeDocumentStats, projectModelForCounting } from "./wordCountClient";

// jsdom has no global Worker, so these exercise the synchronous fallback
// path (isWorkerSupported() === false) — the same code the worker itself
// runs (src/lib/wordCount.ts's documentStats), just awaited instead of
// posted across a thread. The worker script itself is covered end-to-end by
// e2e/SPEC-047.perf.spec.ts, which runs in a real browser.
describe("computeDocumentStats", () => {
  it("returns empty stats for a null model", async () => {
    await expect(computeDocumentStats(null, { includeNotes: false })).resolves.toEqual(emptyCountStats());
  });

  it("matches documentStats for a plain document", async () => {
    const doc = createDocument();
    doc.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "hello world", formatting: {} }] }];
    const expected = documentStats(doc, { includeNotes: false });
    await expect(computeDocumentStats(doc, { includeNotes: false })).resolves.toEqual(expected);
  });

  it("includes notes when asked, matching documentStats", async () => {
    const doc = createDocument();
    doc.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "body text", formatting: {} }] }];
    doc.notes = [
      {
        noteId: "n1",
        kind: "footnote",
        blocks: [{ type: "paragraph", runs: [{ text: "note text", formatting: {} }] }],
      },
    ];
    const expected = documentStats(doc, { includeNotes: true });
    await expect(computeDocumentStats(doc, { includeNotes: true })).resolves.toEqual(expected);
  });
});

// SPEC-047 AC-15: computeDocumentStats sends a projected model over
// postMessage for the includeNotes: false path (the live count, recomputed
// on a throttle while typing) rather than the real model, to keep the
// structured-clone cheap. jsdom has no Worker, so computeDocumentStats
// itself never exercises that projection in a unit test — these pin
// projectModelForCounting directly against documentStats so drift (a field
// wordCount.ts starts reading that the projection drops) shows up here.
describe("projectModelForCounting", () => {
  it("matches documentStats across paragraphs, tracked deletions, tables and atoms with no countable text", () => {
    const doc = createDocument();
    doc.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "Hello world.", formatting: {} }] },
      {
        type: "paragraph",
        runs: [
          { text: "kept ", formatting: {} },
          { text: "deleted", formatting: {}, changes: { delete: { changeId: "c1", author: "A", date: "2024-01-01" } } },
        ],
      },
      {
        type: "table",
        columnWidthsCm: [5, 5],
        widthMode: "auto",
        rows: [
          {
            cells: [
              { blocks: [{ type: "paragraph", runs: [{ text: "cell one", formatting: {} }] }] },
              { blocks: [{ type: "paragraph", runs: [{ text: "cell two", formatting: {} }] }] },
            ],
          },
        ],
      },
      {
        type: "image",
        mediaId: "img1",
        widthCm: 2,
        heightCm: 2,
        wrap: "inline",
        naturalWidthPx: 100,
        naturalHeightPx: 100,
      },
      { type: "break", kind: "page" },
      { type: "toc", options: { levels: [1, 2], showPageNumbers: true, hyperlinksOnly: false }, entries: [] },
    ];

    const expected = documentStats(doc, { includeNotes: false });
    const projected = projectModelForCounting(doc);
    expect(documentStats(projected, { includeNotes: false })).toEqual(expected);
    // Genuinely smaller, not just re-shaped: no media/opaque/sourceXml/notes,
    // and every run/table cell dropped its formatting-only fields.
    expect(projected.media).toBeUndefined();
    expect(projected.notes).toEqual([]);
  });

  it("drops model.notes — never read by the includeNotes: false path it's used for", () => {
    const doc = createDocument();
    doc.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "body", formatting: {} }] }];
    doc.notes = [
      { noteId: "n1", kind: "footnote", blocks: [{ type: "paragraph", runs: [{ text: "note", formatting: {} }] }] },
    ];
    const projected = projectModelForCounting(doc);
    expect(documentStats(projected, { includeNotes: false })).toEqual(documentStats(doc, { includeNotes: false }));
  });
});
