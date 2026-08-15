// Moved to src/lib/core/zip.ts — the ZIP container is OPC/ODF plumbing
// shared by every package format (.docx/.odt today, .pptx/.odp for a future
// Slides product), not WordprocessingML-specific. This shim keeps existing
// import paths working; new code should import from ../core/zip directly.
export * from "../core/zip";
