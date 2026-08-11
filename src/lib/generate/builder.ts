// The fluent generation API (SPEC-045 AC-1, AC-2): a thin front end over
// the canonical `Doc` model (documentModel.ts's `DocumentModel`) that
// composes it block by block and delegates to the existing writers
// (docx/writer.ts, odf/writer.ts) for serialisation — no new serialisation
// code, per the spec's own Reference section. Every `add*`/`set*` method
// validates its arguments before mutating anything (AC-6), so a
// `DocumentGenerationError` always surfaces before any byte is written.
//
// Generated headings are never given a style reference (`styleId`): the
// model only requires a paragraph to reference a style the saved file
// defines (see lib/commands.ts's applyParagraphStyle), which means adding
// one here would also mean shipping a style catalogue — extra machinery a
// generation API has no need for. `headingLevel` plus baked run/paragraph
// formatting (blocks.ts's buildHeadingParagraph) reproduces the built-in
// Heading style's appearance and round-trips independently of any style
// catalogue (`w:outlineLvl` / `text:outline-level`).
import {
  createDocument,
  type DocumentModel,
  type HeaderFooterKind,
  type ParagraphAlignment,
} from "../documentModel";
import { DOCX_MIME_TYPE, writeDocx } from "../docx/writer";
import { ODT_MIMETYPE, writeOdt } from "../odf/writer";
import {
  HEADING_OPTION_KEYS,
  LIST_OPTION_KEYS,
  PARAGRAPH_ALIGNMENTS,
  PARAGRAPH_OPTION_KEYS,
  buildHeadingParagraph,
  buildHeaderFooterContent,
  buildImageBlock,
  buildListParagraphs,
  buildPageBreak,
  buildParagraph,
  buildTable,
  buildTocBlock,
  finalizeTocEntries,
  type BuildContext,
  type HeadingOptions,
  type ListOptions,
  type ParagraphOptions,
} from "./blocks";
import { docFromDescription } from "./description";
import { DocumentGenerationError } from "./errors";
import type { RunSpec } from "./runs";
import {
  expectIntInRange,
  expectEnum,
  expectNonNegativeNumber,
  expectPlainObject,
  expectString,
  rejectUnknownKeys,
} from "./validate";

export type OutputFormat = "docx" | "odt";

export interface DocumentBuilderOptions {
  title?: string;
}

export interface GenerateOptions {
  /** The dcterms:created/modified timestamp; defaults to now (passed
   * straight through to writeDocx/writeOdt). */
  modified?: Date;
}

function parseHeadingOptions(raw: unknown, path: string): HeadingOptions {
  if (raw === undefined) {
    return {};
  }
  const obj = expectPlainObject(raw, path);
  rejectUnknownKeys(obj, HEADING_OPTION_KEYS, path);
  const options: HeadingOptions = {};
  if (obj.alignment !== undefined) {
    options.alignment = expectEnum(obj.alignment, PARAGRAPH_ALIGNMENTS, `${path}.alignment`);
  }
  return options;
}

function parseParagraphOptions(raw: unknown, path: string): ParagraphOptions {
  if (raw === undefined) {
    return {};
  }
  const obj = expectPlainObject(raw, path);
  rejectUnknownKeys(obj, PARAGRAPH_OPTION_KEYS, path);
  const options: ParagraphOptions = {};
  if (obj.alignment !== undefined) {
    options.alignment = expectEnum(obj.alignment, PARAGRAPH_ALIGNMENTS, `${path}.alignment`);
  }
  if (obj.spaceBeforePt !== undefined) {
    options.spaceBeforePt = expectNonNegativeNumber(obj.spaceBeforePt, `${path}.spaceBeforePt`);
  }
  if (obj.spaceAfterPt !== undefined) {
    options.spaceAfterPt = expectNonNegativeNumber(obj.spaceAfterPt, `${path}.spaceAfterPt`);
  }
  return options;
}

function parseListOptions(raw: unknown, path: string): ListOptions {
  if (raw === undefined) {
    return {};
  }
  const obj = expectPlainObject(raw, path);
  rejectUnknownKeys(obj, LIST_OPTION_KEYS, path);
  const options: ListOptions = {};
  if (obj.style !== undefined) {
    options.style = expectEnum(obj.style, ["bullet", "number"] as const, `${path}.style`);
  }
  if (obj.startAt !== undefined) {
    options.startAt = expectIntInRange(obj.startAt, `${path}.startAt`, 1, 9999);
  }
  return options;
}

/** A fluent front end over the canonical document model (AC-1): every
 * `add*` method appends one block (or, for a list, one per item) to the
 * document's single implicit section and returns `this` for chaining.
 * Headers/footers (`setHeader`/`setFooter`) attach to that same section.
 * Nothing is serialised until `toBytes`/`toBlob`/`toBuffer` is called. */
export class DocumentBuilder {
  #doc: DocumentModel;

  constructor(options: DocumentBuilderOptions = {}) {
    const obj = expectPlainObject(options, "options");
    rejectUnknownKeys(obj, ["title"], "options");
    this.#doc = createDocument();
    this.#doc.sections[0].blocks = [];
    if (obj.title !== undefined) {
      this.#doc.title = expectString(obj.title, "options.title");
    }
  }

  /** Interprets a JSON document description (AC-4) into a builder ready for
   * output — the JSON format and this class produce identical models for
   * identical content. */
  static fromDescription(description: unknown): DocumentBuilder {
    const builder = new DocumentBuilder();
    builder.#doc = docFromDescription(description);
    return builder;
  }

  #ctx(): BuildContext {
    return { doc: this.#doc };
  }

  get #blocks() {
    return this.#doc.sections[0].blocks;
  }

  #path(): string {
    return `body[${this.#blocks.length}]`;
  }

  addHeading(content: string | RunSpec[], level: number, options?: { alignment?: ParagraphAlignment }): this {
    const path = this.#path();
    this.#blocks.push(buildHeadingParagraph(this.#ctx(), content, level, parseHeadingOptions(options, path), path));
    return this;
  }

  addParagraph(
    content: string | RunSpec[],
    options?: { alignment?: ParagraphAlignment; spaceBeforePt?: number; spaceAfterPt?: number },
  ): this {
    const path = this.#path();
    this.#blocks.push(buildParagraph(this.#ctx(), content, parseParagraphOptions(options, path), path));
    return this;
  }

  addList(items: Array<string | RunSpec[]>, options?: { style?: "bullet" | "number"; startAt?: number }): this {
    const path = this.#path();
    this.#blocks.push(...buildListParagraphs(this.#ctx(), items, parseListOptions(options, path), path));
    return this;
  }

  addTable(spec: unknown): this {
    this.#blocks.push(buildTable(this.#ctx(), spec, this.#path()));
    return this;
  }

  addImage(spec: unknown): this {
    this.#blocks.push(buildImageBlock(this.#ctx(), spec, this.#path()));
    return this;
  }

  addPageBreak(): this {
    this.#blocks.push(buildPageBreak());
    return this;
  }

  addTableOfContents(options?: unknown): this {
    this.#blocks.push(buildTocBlock(options, this.#path()));
    return this;
  }

  setHeader(content: unknown, kind: HeaderFooterKind = "default"): this {
    const section = this.#doc.sections[0];
    section.headers = { ...section.headers, [kind]: buildHeaderFooterContent(this.#ctx(), content, "header") };
    return this;
  }

  setFooter(content: unknown, kind: HeaderFooterKind = "default"): this {
    const section = this.#doc.sections[0];
    section.footers = { ...section.footers, [kind]: buildHeaderFooterContent(this.#ctx(), content, "footer") };
    return this;
  }

  /** The assembled canonical model (mostly useful for tests and for
   * feeding the in-browser editor directly, bypassing serialisation
   * entirely — see docs/generate-api.md). Every TOC block's `entries` is
   * resolved from the document's current headings at this point. */
  toDoc(): DocumentModel {
    const doc = structuredClone(this.#doc);
    finalizeTocEntries(doc);
    return doc;
  }

  /** The raw package bytes (AC-2) — what `toBlob`/`toBuffer` wrap. */
  async toBytes(format: OutputFormat, options: GenerateOptions = {}): Promise<Uint8Array> {
    const doc = this.toDoc();
    if (format === "docx") {
      return (await writeDocx(doc, options)).bytes;
    }
    if (format === "odt") {
      return (await writeOdt(doc, options)).bytes;
    }
    throw new DocumentGenerationError(`format: expected "docx" or "odt", got "${String(format)}"`);
  }

  /** A `.docx`/`.odt` Blob (AC-2, AC-16): works in both the browser and
   * Node (Node has had a global `Blob` since v18). */
  async toBlob(format: OutputFormat, options: GenerateOptions = {}): Promise<Blob> {
    const bytes = await this.toBytes(format, options);
    return new Blob([bytes], { type: format === "docx" ? DOCX_MIME_TYPE : ODT_MIMETYPE });
  }

  /** A `.docx`/`.odt` Buffer (AC-2, AC-17) for Node callers. Throws if
   * `Buffer` is not global (a browser with no Node polyfill) — use `toBlob`
   * or `toBytes` there instead. */
  async toBuffer(format: OutputFormat, options: GenerateOptions = {}): Promise<Buffer> {
    if (typeof Buffer === "undefined") {
      throw new DocumentGenerationError(
        "Buffer is not available in this environment — use toBlob or toBytes instead",
      );
    }
    const bytes = await this.toBytes(format, options);
    return Buffer.from(bytes);
  }
}
