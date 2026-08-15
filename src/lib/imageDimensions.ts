// Moved to src/lib/core/imageDimensions.ts — decoding intrinsic pixel sizes
// from image bytes is product-neutral (any editing surface inserting media
// needs it). This shim keeps existing import paths working.
export * from "./core/imageDimensions";
