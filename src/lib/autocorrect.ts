// Moved to src/lib/core/autocorrect.ts — the product-neutral core layer shared
// with future editors (e.g. Slides). This shim keeps existing import paths
// working; new code should import from ./core/autocorrect directly.
export * from "./core/autocorrect";
