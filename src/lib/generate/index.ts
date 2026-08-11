// Public entry point for the document generation API (SPEC-045). See
// docs/generate-api.md for the full builder method reference and the JSON
// document-description format, and examples/ for runnable end-to-end
// programs.
export { DocumentBuilder } from "./builder";
export type { DocumentBuilderOptions, GenerateOptions, OutputFormat } from "./builder";

export { docFromDescription } from "./description";
export type { DocumentDescription } from "./description";

export { readDocx, readOdt } from "./read";

export { DocumentGenerationError } from "./errors";

export type { LengthInput, LengthCm, LengthPt, LengthIn, LengthEmu } from "./length";

export type { RunSpec, RunSpecObject } from "./runs";

export type { HeadingOptions, ImageSpec, ListOptions, ParagraphOptions, TableSpec, TocSpec } from "./blocks";
