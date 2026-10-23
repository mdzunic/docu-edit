// Public entry point for the generation API: documents (SPEC-045) and decks
// (SPEC-106). See docs/generate-api.md for the full builder method
// reference, the JSON description formats and the slide-Markdown grammar,
// and examples/ for runnable end-to-end programs. The two builders never
// import each other — this index re-exports both, and each stays separately
// importable (`./builder`, `./presentation`), so a text-only script never
// loads the deck side and vice versa.
export { DocumentBuilder } from "./builder";
export type { DocumentBuilderOptions, GenerateOptions, OutputFormat } from "./builder";

export { docFromDescription } from "./description";
export type { DocumentDescription } from "./description";

export { readDocx, readOdt } from "./read";

export { DocumentGenerationError } from "./errors";

export type { LengthInput, LengthCm, LengthPt, LengthIn, LengthEmu } from "./length";

export type { RunSpec, RunSpecObject } from "./runs";

export type { HeadingOptions, ImageSpec, ListOptions, ParagraphOptions, TableSpec, TocSpec } from "./blocks";

export { PresentationBuilder } from "./presentation";
export type { DeckGenerateOptions, PresentationBuilderOptions } from "./presentation";
export type {
  BulletOptions,
  BulletRow,
  PlacedImageSpec,
  RectSpec,
  SlideImageSpec,
  SlideTextContent,
} from "./deckSlides";

export { deckFromDescription } from "./deckDescription";
export type { DeckDescription, DeckSlideDescription, TextBoxDescription } from "./deckDescription";

export { parseSlideMarkdown, parseSlideMarkdownWithReport, writeSlideMarkdown } from "./slideMarkdown";
export type { SlideMarkdownOptions, SlideMarkdownParse } from "./slideMarkdown";

export type { DeckDropKind, DeckDropRecord } from "./deckDrops";

export { exportDeckHtml } from "./deckHtml";
export type { DeckHtmlOptions } from "./deckHtml";

export { readPptx, readOdp } from "./readDeck";
