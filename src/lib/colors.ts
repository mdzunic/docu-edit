// Moved to src/lib/core/colors.ts — colour parsing, palettes and the
// OOXML/ODF highlight palette are product-neutral (a Slides product's
// pickers use the same data). This shim keeps existing import paths
// working; new code should import from ./core/colors directly.
export * from "./core/colors";
