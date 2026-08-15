// Moved to src/lib/core/shapeCanvas.ts — painting a shape's geometry, fill
// and outline needs only the core ShapeVisual, not the document model's
// anchored Shape block. This shim keeps existing import paths working.
export * from "./core/shapeCanvas";
