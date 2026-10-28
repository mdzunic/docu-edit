// Public entry point for the generation API: documents (SPEC-045), decks
// (SPEC-106) and workbooks (SPEC-140). See docs/generate-api.md for the full
// builder method reference, the JSON description formats and the
// slide-Markdown grammar, docs/generate-workbook-api.md for workbooks, and
// examples/ for runnable end-to-end programs. The three builders never
// import each other — this index re-exports all three, and each stays
// separately importable (`./builder`, `./presentation`, `./workbook`), so a
// text-only script never loads the deck or workbook side and vice versa.
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

export { WorkbookBuilder } from "./workbook";
export type {
  AddSheetOptions,
  BorderSpec,
  CellInput,
  CellInputObject,
  CellStyleSpec,
  ColumnWidthSpec,
  FreezeSpec,
  RowHeightSpec,
  SetRowsOptions,
  WorkbookBuilderOptions,
  WorkbookGenerateOptions,
  WorkbookOutputFormat,
} from "./workbook";

export { workbookFromDescription } from "./workbookDescription";
export type { WorkbookDescription, WorkbookSheetDescription } from "./workbookDescription";

export { readWorkbook } from "./readWorkbook";
