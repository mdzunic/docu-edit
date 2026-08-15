// Moved to src/lib/core/chartCanvas.ts — painting a chart's axes, series
// and labels needs only the core ChartVisual, not the document model's
// anchored Chart block. This shim keeps existing import paths working.
export * from "./core/chartCanvas";
