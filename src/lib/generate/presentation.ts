// The fluent deck builder (SPEC-106): the presentation twin of builder.ts's
// `DocumentBuilder`, composing a `PresentationModel` slide by slide and
// delegating every byte to SPEC-080's writers (`writePptx`/`writeOdp`) — no
// serialisation code of its own; a construct the writers cannot emit is fixed
// in the writer. Every `add*`/`set*` method validates its arguments through
// deckSlides.ts before mutating anything and returns `this`, so a
// `DocumentGenerationError` naming the offending path always surfaces before
// any byte is written.
//
// Imports lib/presentation and lib/core only (plus generate/ siblings): this
// builder never imports documentModel.ts or the .docx/.odt layers, and the
// text builder never imports it — index.ts re-exports both, and each stays
// separately importable.
import { ODP_MIMETYPE, odpFileName, writeOdp } from "../presentation/odp/writer";
import { PPTX_MIME_TYPE, pptxFileName, writePptx } from "../presentation/pptx/writer";
import type { PresentationFormat, PresentationModel } from "../presentation/model";
import type { DeckDropRecord } from "./deckDrops";
import { exportDeckHtml, type DeckHtmlOptions } from "./deckHtml";
import { deckContextFromDescription } from "./deckDescription";
import {
  addBlankSlide,
  addBulletSlide,
  addBullets,
  addImage,
  addImageSlide,
  addTextBox,
  addTitleSlide,
  createDeckContext,
  nextSlidePath,
  PRESENTATION_FORMATS,
  setBackground,
  setNotes,
  type BulletOptions,
  type BulletRow,
  type DeckContext,
  type DeckSizeName,
  type PlacedImageSpec,
  type RectSpec,
  type SlideImageSpec,
  type SlideTextContent,
} from "./deckSlides";
import { DocumentGenerationError } from "./errors";
import type { LengthInput } from "./length";
import { expectEnum, expectPlainObject, rejectUnknownKeys } from "./validate";

export interface PresentationBuilderOptions {
  title?: string;
  /** The deck's own format — what `toBytes()`/`fileName()` use when called
   * without one. Defaults to `"pptx"`. */
  format?: PresentationFormat;
  /** `"16:9"` (default, 33.867 × 19.05 cm), `"4:3"` (25.4 × 19.05 cm) or an
   * explicit size. */
  size?: DeckSizeName | { width: LengthInput; height: LengthInput };
}

export interface DeckGenerateOptions {
  /** The stamped created/modified timestamp; defaults to now. Fix it for
   * byte-identical output. */
  modified?: Date;
}

/** A fluent front end over the slide model: every `add*Slide` method appends
 * one slide, every other `add*`/`set*` method acts on the slide most
 * recently added (and throws when there is none), and all of them return
 * `this`. Nothing is serialised until an output method is called. */
export class PresentationBuilder {
  #ctx: DeckContext;

  constructor(options: PresentationBuilderOptions = {}) {
    const obj = expectPlainObject(options, "options");
    rejectUnknownKeys(obj, ["title", "format", "size"], "options");
    this.#ctx = createDeckContext(obj, "options");
  }

  /** Interprets a JSON deck description (deckDescription.ts) into a builder
   * ready for more calls or for output — the description and this class
   * produce identical models for identical content. */
  static fromDescription(description: unknown): PresentationBuilder {
    const builder = new PresentationBuilder();
    builder.#ctx = deckContextFromDescription(description);
    return builder;
  }

  /** A title slide: the title placeholder and, below it, the subtitle. */
  addTitleSlide(title: SlideTextContent, subtitle?: SlideTextContent): this {
    addTitleSlide(this.#ctx, title, subtitle, nextSlidePath(this.#ctx));
    return this;
  }

  /** A title + body slide whose body holds `bullets`: strings, `{ text |
   * runs, level }` rows, and nested arrays one level deeper. */
  addBulletSlide(title: SlideTextContent, bullets: BulletRow[], options?: BulletOptions): this {
    addBulletSlide(this.#ctx, title, bullets, options, nextSlidePath(this.#ctx));
    return this;
  }

  /** A title slide over an image fitted into the body rect at its own
   * aspect, centred. */
  addImageSlide(title: SlideTextContent, image: SlideImageSpec): this {
    addImageSlide(this.#ctx, title, image, nextSlidePath(this.#ctx));
    return this;
  }

  /** An empty slide for the per-slide methods to fill. */
  addBlankSlide(): this {
    addBlankSlide(this.#ctx);
    return this;
  }

  /** Appends bullet rows to the current slide's body (adding a body
   * placeholder at the editor's geometry when it has none). */
  addBullets(items: BulletRow[], options?: BulletOptions): this {
    addBullets(this.#ctx, items, options);
    return this;
  }

  /** A free text box holding one paragraph at `rect`. */
  addTextBox(content: SlideTextContent, rect: RectSpec): this {
    addTextBox(this.#ctx, content, rect);
    return this;
  }

  /** An image at an explicit rect, clamped into the slide. */
  addImage(spec: PlacedImageSpec): this {
    addImage(this.#ctx, spec);
    return this;
  }

  /** The current slide's speaker notes, one string per paragraph. */
  setNotes(paragraphs: string[]): this {
    setNotes(this.#ctx, paragraphs);
    return this;
  }

  /** The current slide's solid `#rrggbb` background. */
  setBackground(color: string): this {
    setBackground(this.#ctx, color);
    return this;
  }

  /** Everything the builder accepted but the slide model cannot hold (a
   * run's `strikethrough`/`highlight`, an image's alt text), one record per
   * slide and kind, in call order. */
  dropped(): DeckDropRecord[] {
    return this.#ctx.dropped.map((record) => ({ ...record }));
  }

  /** The assembled deck — a deep copy, plain JSON. */
  toDeck(): PresentationModel {
    return structuredClone(this.#ctx.deck);
  }

  #format(format: unknown): PresentationFormat {
    return format === undefined ? this.#ctx.deck.format : expectEnum(format, PRESENTATION_FORMATS, "format");
  }

  /** The raw `.pptx`/`.odp` package bytes, in `format` or the deck's own. */
  async toBytes(format?: PresentationFormat, options: DeckGenerateOptions = {}): Promise<Uint8Array> {
    const target = this.#format(format);
    const writeOptions = generateOptions(options);
    // The writers key media and structure off the model alone; the format
    // field is the only thing a cross-format write changes (the app's own
    // "Download a copy as…" does the same).
    const deck = { ...this.toDeck(), format: target };
    return (target === "odp" ? await writeOdp(deck, writeOptions) : await writePptx(deck, writeOptions)).bytes;
  }

  /** A `.pptx`/`.odp` Blob (browser, or Node 18+). */
  async toBlob(format?: PresentationFormat, options: DeckGenerateOptions = {}): Promise<Blob> {
    const target = this.#format(format);
    const bytes = await this.toBytes(target, options);
    return new Blob([bytes as BlobPart], { type: target === "odp" ? ODP_MIMETYPE : PPTX_MIME_TYPE });
  }

  /** A `.pptx`/`.odp` Buffer for Node callers. Throws where `Buffer` is not
   * global (a browser) — use `toBlob` or `toBytes` there. */
  async toBuffer(format?: PresentationFormat, options: DeckGenerateOptions = {}): Promise<Buffer> {
    if (typeof Buffer === "undefined") {
      throw new DocumentGenerationError("Buffer is not available in this environment — use toBlob or toBytes instead");
    }
    return Buffer.from(await this.toBytes(format, options));
  }

  /** The download file name for `format` (or the deck's own): the title plus
   * the extension, "Untitled presentation" when the title is blank. */
  fileName(format?: PresentationFormat): string {
    return this.#format(format) === "odp" ? odpFileName(this.#ctx.deck.title) : pptxFileName(this.#ctx.deck.title);
  }

  /** The deck as one self-contained HTML document (SPEC-088's export).
   * `includeNotes` defaults to `false`; a deck with no slides is refused. */
  toHtml(options: DeckHtmlOptions = {}): string {
    return exportDeckHtml(this.#ctx.deck, options);
  }
}

function generateOptions(options: unknown): { modified?: Date } {
  const obj = expectPlainObject(options, "options");
  rejectUnknownKeys(obj, ["modified"], "options");
  if (obj.modified === undefined) {
    return {};
  }
  if (!(obj.modified instanceof Date) || Number.isNaN(obj.modified.getTime())) {
    throw new DocumentGenerationError("options.modified: expected a valid Date");
  }
  return { modified: obj.modified };
}
