// Moved to src/lib/core/language.ts — the product-neutral core layer shared
// with future editors (e.g. Slides). This shim keeps existing import paths
// working; new code should import from ./core/language directly.
export * from "./core/language";
