// SPEC-137 "Saving stays responsive": building a 25,000-row sheet's XML is
// hundreds of thousands of small string operations — easily tens of ms of
// unbroken synchronous work, past this spec's main-thread-block-during-save
// budget. The xlsx and ODS sheet serialisers (xlsx/sheetXml.ts's
// serializeWorksheet, ods/tableXml.ts's serializeTable) check this module's
// clock every few rows and yield to the main thread whenever
// SERIALISE_SLICE_MS has elapsed, so a save of a large sheet reads as many
// short tasks instead of one long one — the same idea as
// core/zipWriter.ts's crc32Chunked, applied to the row loop that builds a
// sheet's XML rather than the zip entry's compressed bytes.
import { yieldToMainThread } from "../core/idleScheduler";

/** How long a serialiser's row loop may run before it must yield. */
export const SERIALISE_SLICE_MS = 8;

/** A cheap, resumable yield point for a long synchronous loop: `due()` is a
 * plain clock read (so a loop can check it every row, or every few rows,
 * without the check itself costing anything worth yielding over); `yield()`
 * hands control back to the browser (core/idleScheduler.ts's
 * yieldToMainThread — a real macrotask boundary, not a microtask) and resets
 * the slice. */
export interface SliceClock {
  due(): boolean;
  yield(): Promise<void>;
}

/** The slicing clock a serialiser's row loop drives: `due()` once
 * `sliceMs` of wall time has elapsed since the clock was created or last
 * yielded; `yield()` resets it. */
export function createSliceClock(sliceMs: number = SERIALISE_SLICE_MS): SliceClock {
  let last = performance.now();
  return {
    due(): boolean {
      return performance.now() - last >= sliceMs;
    },
    async yield(): Promise<void> {
      await yieldToMainThread();
      last = performance.now();
    },
  };
}
