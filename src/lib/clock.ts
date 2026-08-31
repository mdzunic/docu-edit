// The injectable clock moved to src/lib/core/clock.ts in SPEC-093 (the shape of
// SPEC-079's core/drawingMl.ts move) so both products share one seam. This
// module re-exports it so the word-processor consumers (useFields.ts,
// usePaginationView.ts, fields.spec.ts) and their imports stay untouched.
export { now, setClock } from "./core/clock";
