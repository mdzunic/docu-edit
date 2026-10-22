// The one error type every Keynote-import failure surfaces as (SPEC-087
// AC-12). It lives in its own module so the whole decode stack (snappy.ts,
// protobuf.ts, iwa.ts) can raise it without importing reader.ts — which
// imports them. reader.ts re-exports it for the open funnel.

/** A `.key` file (or a file that took the Keynote path) could not be
 * imported. `message` is always one of reader.ts's user-facing strings by
 * the time it leaves `parseKeynote`; the decode stack throws it with
 * internal messages that `parseKeynote` normalises. */
export class KeynoteOpenError extends Error {}

