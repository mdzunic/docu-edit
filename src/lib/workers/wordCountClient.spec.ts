import { describe, it, expect } from "vitest";
import { createDocument } from "../documentModel";
import { documentStats, emptyCountStats } from "../wordCount";
import { computeDocumentStats } from "./wordCountClient";

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
