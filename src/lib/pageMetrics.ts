// Moved to src/lib/core/pageMetrics.ts — DOM line-box measurement is
// product-neutral (any paginated/fitted layout measures rendered blocks the
// same way). This shim keeps existing import paths working.
export * from "./core/pageMetrics";
