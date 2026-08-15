// Moved to src/lib/core/editLock.ts — the product-neutral core layer shared
// with future editors (e.g. Slides). This shim keeps existing import paths
// working; new code should import from ./core/editLock directly.
export * from "./core/editLock";
